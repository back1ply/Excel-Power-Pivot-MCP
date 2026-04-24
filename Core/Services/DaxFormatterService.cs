using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Polly;
using Polly.Retry;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Resilient DAX formatter with retry logic, caching, and graceful degradation.
///
/// NOTE: The Dax.Formatter library creates its own HttpClient internally.
/// While we cannot override its HttpClient instance, we wrap its usage with
/// retry policies and caching to prevent excessive requests.
///
/// For future improvement: Consider using HttpClient directly with the DAX Formatter
/// API instead of the library to gain full control over HttpClient lifecycle.
/// </summary>
public class DaxFormatterService : IDaxFormatterService
{
    private readonly McpConfiguration _config;
    private readonly ILogger<DaxFormatterService> _logger;
    private readonly AsyncRetryPolicy _retryPolicy;
    private readonly Dictionary<string, CacheEntry> _cache;
    private readonly object _cacheLock = new();

    // Cache statistics
    private long _cacheHits;
    private long _cacheMisses;
    private long _totalRequests;
    private long _successfulFormats;
    private long _failedFormats;

    // Cache TTL for failed formatting attempts (shorter than successful results)
    private const int FailedFormatCacheTtlMinutes = 5;

    // Shared DaxFormatterClient instance to reduce object creation overhead
    // The library creates HttpClient internally, but at least we can reuse the wrapper
    private static readonly Lazy<Dax.Formatter.DaxFormatterClient> _sharedClient =
        new(() => new Dax.Formatter.DaxFormatterClient(), LazyThreadSafetyMode.ExecutionAndPublication);

    public DaxFormatterService(McpConfiguration config, ILogger<DaxFormatterService> logger)
    {
        _config = config;
        _logger = logger;
        _cache = new Dictionary<string, CacheEntry>();

        // Build retry policy using Polly (already in dependencies)
        _retryPolicy = Policy
            .Handle<HttpRequestException>() // Network errors
            .Or<TaskCanceledException>()    // Timeout-induced cancellations
            .WaitAndRetryAsync(
                retryCount: _config.DaxFormatterRetryCount,
                sleepDurationProvider: attempt =>
                {
                    var delay = _config.DaxFormatterRetryDelayMs * Math.Pow(2, attempt - 1);
                    return TimeSpan.FromMilliseconds(
                        Math.Min(delay, _config.DaxFormatterMaxRetryDelayMs));
                },
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    _logger.LogWarning(exception,
                        "DAX Formatter retry {RetryCount}/{MaxRetries} after {DelayMs}ms",
                        retryCount, _config.DaxFormatterRetryCount, timeSpan.TotalMilliseconds);
                });
    }

    public async Task<(string formatted, string? warning)> FormatDaxAsync(
        string expression,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(expression))
            return (expression, null);

        Interlocked.Increment(ref _totalRequests);

        // Check cache first (if enabled)
        if (_config.DaxFormatterCacheEnabled)
        {
            var cacheKey = ComputeCacheKey(expression);
            if (TryGetFromCache(cacheKey, out var cachedResult))
            {
                Interlocked.Increment(ref _cacheHits);
                _logger.LogDebug("DAX Formatter cache hit");
                return cachedResult;
            }
            Interlocked.Increment(ref _cacheMisses);
        }

        // Attempt formatting with retry logic
        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(TimeSpan.FromSeconds(_config.DaxFormatterTimeoutSeconds));

            var result = await _retryPolicy.ExecuteAsync(async () =>
            {
                // Use shared client instance to reduce overhead
                // NOTE: The library creates HttpClient internally - we cannot control that
                // But we can at least reuse the client wrapper and rely on our caching layer
                var request = new Dax.Formatter.Models.DaxFormatterSingleRequest { Dax = expression };

                var response = await _sharedClient.Value.FormatAsync(request);

                // If there are DAX syntax errors, throw - the LLM needs to fix its DAX
                if (response?.Errors?.Count > 0)
                {
                    var errors = string.Join("; ",
                        response.Errors.Select(e =>
                            $"Line {e?.Line}, Col {e?.Column}: {e?.Message}"));
                    throw new InvalidOperationException($"DAX syntax error: {errors}");
                }

                if (response is { Formatted: not null })
                {
                    return response.Formatted;
                }

                // Formatter returned success but no formatted text - unusual but gracefully handle
                return expression;
            });

            Interlocked.Increment(ref _successfulFormats);

            // Cache successful result
            if (_config.DaxFormatterCacheEnabled)
            {
                var cacheKey = ComputeCacheKey(expression);
                AddToCache(cacheKey, (result, null));
            }

            return (result, null);
        }
        catch (InvalidOperationException)
        {
            // DAX syntax errors should propagate to LLM
            throw;
        }
        catch (OperationCanceledException) when (ct.IsCancellationRequested)
        {
            // User-requested cancellation should propagate
            throw;
        }
        catch (Exception ex)
        {
            // All other errors: graceful degradation
            Interlocked.Increment(ref _failedFormats);

            var warningMessage = ex switch
            {
                OperationCanceledException => "DAX formatter timed out - syntax not pre-validated",
                HttpRequestException => "DAX formatter unavailable (network error) - syntax not pre-validated",
                _ => $"DAX formatter error ({ex.GetType().Name}) - syntax not pre-validated"
            };

            _logger.LogWarning(ex, "DAX Formatter: {WarningMessage}", warningMessage);

            // Cache failures too (avoid repeated failed attempts for bad network)
            if (_config.DaxFormatterCacheEnabled)
            {
                var cacheKey = ComputeCacheKey(expression);
                AddToCache(cacheKey, (expression, warningMessage),
                    ttlMinutes: FailedFormatCacheTtlMinutes);
            }

            return (expression, warningMessage);
        }
    }

    public void ClearCache()
    {
        lock (_cacheLock)
        {
            _cache.Clear();
            _logger.LogInformation("DAX Formatter cache cleared");
        }
    }

    public Dictionary<string, object> GetCacheStats()
    {
        lock (_cacheLock)
        {
            return new Dictionary<string, object>
            {
                ["cacheEnabled"] = _config.DaxFormatterCacheEnabled,
                ["cacheSize"] = _cache.Count,
                ["cacheMaxEntries"] = _config.DaxFormatterCacheMaxEntries,
                ["cacheTtlMinutes"] = _config.DaxFormatterCacheTtlMinutes,
                ["cacheHits"] = _cacheHits,
                ["cacheMisses"] = _cacheMisses,
                ["hitRate"] = _cacheMisses > 0
                    ? (double)_cacheHits / (_cacheHits + _cacheMisses)
                    : 0.0,
                ["totalRequests"] = _totalRequests,
                ["successfulFormats"] = _successfulFormats,
                ["failedFormats"] = _failedFormats,
                ["successRate"] = _totalRequests > 0
                    ? (double)_successfulFormats / _totalRequests
                    : 0.0
            };
        }
    }

    private static string ComputeCacheKey(string expression)
    {
        // Use SHA256 hash for consistent, compact cache keys
        var bytes = Encoding.UTF8.GetBytes(expression.Trim());
        var hash = SHA256.HashData(bytes);
        return Convert.ToBase64String(hash);
    }

    private bool TryGetFromCache(string key, out (string formatted, string? warning) result)
    {
        lock (_cacheLock)
        {
            if (_cache.TryGetValue(key, out var entry))
            {
                // Check if entry is still valid
                if (DateTime.UtcNow < entry.ExpiresAt)
                {
                    result = entry.Result;
                    return true;
                }

                // Expired - remove it
                _cache.Remove(key);
            }

            result = default;
            return false;
        }
    }

    private void AddToCache(string key, (string formatted, string? warning) result, int? ttlMinutes = null)
    {
        lock (_cacheLock)
        {
            // Enforce cache size limit (LRU eviction)
            if (_cache.Count >= _config.DaxFormatterCacheMaxEntries)
            {
                // Remove oldest entry
                var oldestKey = _cache
                    .OrderBy(kvp => kvp.Value.CreatedAt)
                    .First()
                    .Key;
                _cache.Remove(oldestKey);
            }

            var ttl = ttlMinutes ?? _config.DaxFormatterCacheTtlMinutes;
            var entry = new CacheEntry
            {
                Result = result,
                CreatedAt = DateTime.UtcNow,
                ExpiresAt = DateTime.UtcNow.AddMinutes(ttl)
            };

            _cache[key] = entry;
        }
    }

    private sealed class CacheEntry
    {
        public required (string formatted, string? warning) Result { get; init; }
        public required DateTime CreatedAt { get; init; }
        public required DateTime ExpiresAt { get; init; }
    }
}

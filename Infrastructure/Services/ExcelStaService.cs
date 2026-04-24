using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Polly;
using Polly.Retry;
using System.Runtime.InteropServices;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Dispatcher interface for executing actions on the STA thread with Excel COM.
/// </summary>
public interface IExcelDispatcher
{
    /// <summary>
    /// Execute an action on the STA thread.
    /// </summary>
    Task RunAsync(Action action, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Execute a function on the STA thread and return the result.
    /// </summary>
    Task<T> RunAsync<T>(Func<T> func, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Execute an action with a specific timeout.
    /// </summary>
    Task RunAsync(Action action, TimeSpan timeout, CancellationToken cancellationToken = default);
    
    /// <summary>
    /// Execute a function with a specific timeout.
    /// </summary>
    Task<T> RunAsync<T>(Func<T> func, TimeSpan timeout, CancellationToken cancellationToken = default);
}

/// <summary>
/// STA thread service for Excel COM interop.
/// 
/// CRITICAL: This service runs a dedicated STA thread with a Windows message pump,
/// which is required for Excel COM automation. Without message pumping, modal dialogs
/// (save prompts, connection errors, etc.) will deadlock the server.
/// </summary>
public class ExcelStaService : IHostedService, IExcelDispatcher, IDisposable
{
    private readonly Thread _staThread;
    private readonly ILogger<ExcelStaService> _logger;
    private readonly McpConfiguration _config;
    
    // Defines the context for marshaling calls to the STA thread
    private SynchronizationContext? _staContext;
    private readonly TaskCompletionSource _threadStartedTcs = new();
    
    public ExcelStaService(ILogger<ExcelStaService> logger, McpConfiguration config)
    {
        _logger = logger;
        _config = config;
        
        // Define resilience policy for Excel COM operations
        // Retries on "Call was rejected by callee" (RPC_E_CALL_REJECTED / 0x80010001)
        _retryPolicy = Policy
            .Handle<COMException>(ex => (uint)ex.ErrorCode == 0x80010001)
            .WaitAndRetry(
                retryCount: 3,
                sleepDurationProvider: attempt => TimeSpan.FromMilliseconds(200 * attempt),
                onRetry: (ex, time) => _logger.LogWarning("Excel busy (RPC_E_CALL_REJECTED), retrying in {Time}ms...", time.TotalMilliseconds)
            );

        _staThread = new Thread(ThreadLoop)
        {
            IsBackground = true,
            Name = "ExcelStaThread"
        };
        _staThread.SetApartmentState(ApartmentState.STA);
    }

    private readonly Policy _retryPolicy;

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _staThread.Start();
        return _threadStartedTcs.Task;
    }

    // Timeout for STA thread graceful shutdown
    private static readonly TimeSpan StaThreadExitTimeout = TimeSpan.FromSeconds(5);

    public async Task StopAsync(CancellationToken cancellationToken)
    {
        _logger.LogInformation("Excel STA Service stopping");

        // Marshal the exit call to the STA thread
        if (_staContext != null)
        {
            _staContext.Post(_ => Application.Exit(), null);
        }

        // Wait for STA thread to actually exit (with timeout)
        // This prevents orphaned threads when modal dialogs block Application.Exit()
        var joined = await Task.Run(() => _staThread.Join((int)StaThreadExitTimeout.TotalMilliseconds), cancellationToken);
        
        if (!joined)
        {
            _logger.LogWarning("STA thread did not exit within {TimeoutMs}ms - possible modal dialog blocking", StaThreadExitTimeout.TotalMilliseconds);
        }
    }

    /// <summary>
    /// Main STA thread loop using standard Windows Forms message pump.
    /// This is superior to a custom loop with DoEvents because:
    /// 1. It handles all COM interop requirements correctly
    /// 2. It avoids deadlocks with modal dialogs
    /// 3. It correctly handles reentrancy
    /// </summary>
    private void ThreadLoop()
    {
        try
        {
            _logger.LogInformation("Excel STA Thread started");
            
            // Initialize OLE for this thread (required for COM)
            Application.OleRequired();
            
            // Install the standard Windows Forms synchronization context
            // This allows us to marshal calls to this thread from anywhere
            var context = new WindowsFormsSynchronizationContext();
            SynchronizationContext.SetSynchronizationContext(context);
            _staContext = context;
            
            // Signal that we are ready to accept commands
            _threadStartedTcs.TrySetResult();
            
            // Run the standard message pump
            // This will block until Application.Exit() is called
            Application.Run(new ApplicationContext());
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Fatal error in STA thread");
            _threadStartedTcs.TrySetException(ex);
        }
        finally
        {
            _logger.LogInformation("Excel STA Thread stopped");
        }
    }

    #region IExcelDispatcher Implementation

    /// <summary>
    /// Helper to setup timeout and cancellation token handling for a TaskCompletionSource.
    /// Returns the registration and timeout CTS for disposal.
    /// </summary>
    private static (CancellationTokenRegistration registration, CancellationTokenSource? timeoutCts) SetupTimeoutHandling<T>(
        TaskCompletionSource<T> tcs,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        // Register cancellation
        CancellationTokenRegistration registration = default;
        if (cancellationToken.CanBeCanceled)
        {
            registration = cancellationToken.Register(() =>
                tcs.TrySetCanceled(cancellationToken));
        }

        // Handle timeout
        CancellationTokenSource? timeoutCts = null;
        if (timeout != Timeout.InfiniteTimeSpan)
        {
            timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(timeout);

            // Re-register if we have a timeout so we catch that cancellation too
            registration.Dispose();
            registration = timeoutCts.Token.Register(() =>
            {
                if (!tcs.Task.IsCompleted)
                {
                    tcs.TrySetException(new TimeoutException($"Operation timed out after {timeout.TotalSeconds:F1} seconds"));
                }
            });
        }

        return (registration, timeoutCts);
    }

    public Task RunAsync(Action action, CancellationToken cancellationToken = default)
    {
        return RunAsync(action, Timeout.InfiniteTimeSpan, cancellationToken);
    }
    
    public Task RunAsync(Action action, TimeSpan timeout, CancellationToken cancellationToken = default)
    {
        if (_staContext == null)
        {
            throw new InvalidOperationException("Excel STA service is not running or hasn't started yet.");
        }

        var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var (registration, timeoutCts) = SetupTimeoutHandling(tcs, timeout, cancellationToken);

        // Post the work to the STA thread
        _staContext.Post(_ =>
        {
            // If already cancelled/timed out, don't run
            if (tcs.Task.IsCompleted) return;

            try
            {
                // Execute with retry policy for "Application Busy" errors
                _retryPolicy.Execute(() => action());
                tcs.TrySetResult(true);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
            finally
            {
                registration.Dispose();
                try
                {
                    timeoutCts?.Dispose();
                }
                catch (ObjectDisposedException)
                {
                    // Race condition with timeout callback - expected
                }
            }
        }, null);

        return tcs.Task;
    }

    public Task<T> RunAsync<T>(Func<T> func, CancellationToken cancellationToken = default)
    {
        return RunAsync(func, Timeout.InfiniteTimeSpan, cancellationToken);
    }
    
    public Task<T> RunAsync<T>(Func<T> func, TimeSpan timeout, CancellationToken cancellationToken = default)
    {
        if (_staContext == null)
        {
            throw new InvalidOperationException("Excel STA service is not running or hasn't started yet.");
        }

        var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        var (registration, timeoutCts) = SetupTimeoutHandling(tcs, timeout, cancellationToken);

        _staContext.Post(_ =>
        {
            if (tcs.Task.IsCompleted) return;

            try
            {
                // Execute with retry policy
                var result = _retryPolicy.Execute(() => func());
                tcs.TrySetResult(result);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
            finally
            {
                registration.Dispose();
                try
                {
                    timeoutCts?.Dispose();
                }
                catch (ObjectDisposedException)
                {
                    // Race condition with timeout callback - expected
                }
            }
        }, null);

        return tcs.Task;
    }

    #endregion

    public void Dispose()
    {
        // Stop is handled by StopAsync via IHostedService
        GC.SuppressFinalize(this);
    }
}

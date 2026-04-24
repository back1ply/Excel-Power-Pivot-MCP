# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest release | ✅ |
| Older releases | ❌ |

## Reporting a Vulnerability

Open a [GitHub issue](https://github.com/back1ply/Excel-Power-Pivot-MCP/issues) and label it `security`.

Include:
- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Any suggested fix (optional)

## Scope

This project runs locally on Windows and communicates with Excel via COM interop. It does not expose network services, store credentials, or transmit data externally (except DAX formatting requests sent to daxformatter.com).

Areas of concern:
- COM object handling and memory safety
- MCP stdio transport input parsing
- DAX query input passed to Excel's ADOMD.NET engine

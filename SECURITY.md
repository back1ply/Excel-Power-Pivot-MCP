# Security Policy

## Supported Versions

| Version | Supported |
|---------|-----------|
| Latest release | ✅ |
| Older releases | ❌ |

## Reporting a Vulnerability

**Do not open a public GitHub issue for security vulnerabilities.**

Report via email: **back1ply@gmail.com**

Include:
- Description of the vulnerability
- Steps to reproduce
- Potential impact
- Any suggested fix (optional)

Expect an initial response within 48 hours. Once confirmed, a fix will be prioritized and a patched release issued. You will be credited in the release notes unless you prefer otherwise.

## Scope

This project runs locally on Windows and communicates with Excel via COM interop. It does not expose network services, store credentials, or transmit data externally (except DAX formatting requests sent to daxformatter.com).

Areas of concern:
- COM object handling and memory safety
- MCP stdio transport input parsing
- DAX query input passed to Excel's ADOMD.NET engine

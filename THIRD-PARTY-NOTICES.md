# Third-party components

LedgerLens uses the following components. Upstream licenses and notices are included in `licenses` where supplied; each component retains its own terms.

| Component | Version | Upstream |
| --- | --- | --- |
| Excel-DNA | 1.9.0 | https://github.com/Excel-DNA/ExcelDna |
| Newtonsoft.Json | 13.0.3 | https://github.com/JamesNK/Newtonsoft.Json |
| Microsoft WebView2 SDK | 1.0.3800.47 | https://learn.microsoft.com/en-us/microsoft-edge/webview2/ |
| .NET runtime | 10.0.12 | https://github.com/dotnet/runtime |
| Microsoft.Office.Interop.Excel | 15.0.4795.1001 | https://www.nuget.org/packages/Microsoft.Office.Interop.Excel/ |

Microsoft Excel and WebView2 Runtime are installed prerequisites; Excel is not redistributed. The Office interop package is a repackaged Microsoft assembly and is not a Microsoft-supported NuGet package. This prototype has been validated with the installed licensed Microsoft 365 host.

Development-only dependencies, including Playwright, Microsoft's Office manifest validator, and xUnit, are enumerated with exact versions in `package-lock.json` and .NET `packages.lock.json` files. They are not installed or executed by the Windows product launcher. The web preview loads Office.js from Microsoft's official CDN only when `?office=1` is selected.

Public financial evidence is retrieved from SEC EDGAR. Company names identify the issuers; no endorsement or AlphaSense affiliation is implied.

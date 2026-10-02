# Setup and operations

Keep the entire extracted release together. `native` contains both XLL bitness variants and their managed/WebView2 dependencies; the XLL alone is not a complete installation. `service` includes a self-contained .NET 10 Windows x64 runtime. The launcher detects Excel bitness. Only x64 Excel has been runtime-tested.

Launch through `Start LedgerLens.cmd` or `Start-LedgerLens.ps1`. A writable local folder is needed for `.runtime` and `workbooks`. The launcher saves a new workbook each time. Excel's AddIns manager loads the native plugin. To remove it, save your work, open Excel Options → Add-ins → Manage Excel Add-ins → Go, and uncheck LedgerLens; stop the local service with `Stop-LedgerLens.ps1` before moving the release folder.

The default service port is 17843. Use `-Port 17853` if occupied. A service is reused only if its executable path, PID, port, and authenticated health check match this folder. Do not run two copies on the same port. After restarting the service, close and relaunch the LedgerLens Excel session so it receives the new session token.

When updating from 1.0.0, save your workbook, close that LedgerLens Excel session, run `Stop-LedgerLens.ps1` from the old release folder, and launch the new release. Keep the old folder until you have confirmed the upgrade. The 1.0.1 review used a separate service port and test Excel instance; it did not change the user's already-open workbook.

## OpenAI key

The service reads `OPENAI_API_KEY` from the current Windows user's environment settings, then the process environment. On the development machine, the refreshed key was tested successfully. To configure a different machine without writing the key into shell history:

```powershell
$secret = Read-Host 'OpenAI API key' -AsSecureString
$pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
try {
    [Environment]::SetEnvironmentVariable('OPENAI_API_KEY', [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer), 'User')
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    $secret.Dispose()
}
```

This stores the key in the Windows user environment, not an encrypted credential vault. The application does not print it. Never send `.runtime`, environment exports, or a populated secret file with the portfolio package. Each reviewer should use their own key. A billing-enabled API project with access to the configured model is required for live AI; calculated research requires no key.

## Troubleshooting

| Symptom | Check |
| --- | --- |
| XLL blocked | Review the organization's XLL policy. Do not weaken global Office security settings. Use an approved local development folder/environment. |
| Blank research pane | Verify WebView2 Runtime is installed, then inspect `.runtime/excel.log`. |
| `#N/A` briefly | Async formulas are pending. Persistent errors: use `LL.STATUS` or Health. |
| Unknown formula | Start through the launcher and ensure the matching XLL is enabled. |
| Service unavailable | Inspect `.runtime/service.log` and `service-error.log`; check the port. |
| Unauthorized after restart | Close and relaunch LedgerLens Excel to obtain the new local token. |
| AI unavailable | Inspect the sanitized UI error, key configuration, model access, and API billing; select calculated analysis to continue. |
| SEC sync fails | Existing validated evidence remains available. Check network/proxy access to data.sec.gov and retry later. |

Offline mode is an explicit demonstration state. It preserves historical evidence and marks its freshness. Live SEC sync and AI need external connectivity; they do not bypass corporate proxy or authentication policy. Existing workbook values persist independently of service lifetime.

The native remote-debugging port is used only by `scripts/native-session.ps1`. The product launcher does not enable it. Test scripts close only the Excel instance they created. Avoid editing that instance while tests are running.

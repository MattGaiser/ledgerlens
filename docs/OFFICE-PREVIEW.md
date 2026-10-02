# Office.js preview and compatibility boundary

The Office.js adapter supports inserting a sourced numeric value into one empty cell and saving a cited research answer to a new sheet. It checks read-only documents, merged cells, protection, existing values/formulas, and citation membership. Nine adapter tests pass. The XML manifest passes Microsoft's validator.

**No real Office.js host or Mac Excel session was validated.** This is an adapter/manifest preview. Native `LL.*` functions, the C# ribbon, and guarded model refresh are Windows features. The release does not include a Mac-native backend binary.

For a Windows developer preview, first establish a trusted local ASP.NET development certificate using the [.NET development certificate instructions](https://learn.microsoft.com/en-us/dotnet/core/tools/dotnet-dev-certs). The installed development machine already has one. Then:

```powershell
.\Stop-LedgerLens.ps1
.\Start-LedgerLens.ps1 -NoExcel -OfficePreview
```

The service uses HTTPS on port 17844 by default. The launcher writes `.runtime/office-session-manifest.xml` with the current session bootstrap fragment. Treat that manifest as a local credential and regenerate it after service restart. Do not publish it or put it on a broadly accessible network share. The checked-in `office/manifest.xml` contains no credential.

Use Microsoft's [Windows development sideload workflow](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) with a catalog accessible only to the developer. This is a testing workflow, not the project's production deployment mechanism. The current manifest requests ExcelApi 1.13, so older hosts may not support the preview.

For Mac, the intended path is running the portable service from source with .NET 10, configuring trusted loopback HTTPS and a process-level API key, generating a local session manifest, then using [Microsoft's Mac sideload procedure](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac). Those steps require additional host validation before claiming support. Production would replace development certificates and fragment bootstrapping with an approved HTTPS deployment and user authentication.

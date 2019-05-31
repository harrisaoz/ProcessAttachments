.net core Command Line Application: alert-by-smtp
=

Project Definition
-

```
ProcessAttachments.sln
```

Build/Deploy
-

* Clone from GitHub https://github.com/harrisaoz/ProcessAttachments
* Build and publish from Powershell (may also work with minor changes using Bash shell or Command shell) using the following commands:

Build all projects

```
dotnet build -c Release
```

Publish the Timesheet Processor

```
dotnet publish -c Release -r win-x64 --self-contained true -o .\publish\timesheets ProcessTimesheets
```

Publish the Invoice Downloader

```
dotnet publish -c Release -r win-x64 --self-contained true -o .\publish\invoices DownloadInvoices
```

* Copy the resulting folder to the target host

Configuration
-

* Edit *project*.dll.config - on the target host - to configure the required
parameters
* Ensure that the project*.dll.config fileis in the same folder as the
 *project*.exe file.

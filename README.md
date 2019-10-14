# .net core Command Line Applications: ProcessTimesheets and DownloadInvoices

## Project Definition

```
ProcessAttachments.sln
```

## Build, Deploy, Configure

Build and publish from Powershell Core.

Before running any of the Powershell deployment functions, source ./deploy.ps1:

```
. ./deploy.ps1
```

### Build all projects

```
dotnet build -c Release
```

### Publish the Tool (ProcessTimesheets or DownloadInvoices)

```
Publish-Tool -ToolName (ProcessTimesheets | DownloadInvoices)
```

Note that this creates the config file under ./publish/{Toolname}/Properties/{Toolname}.dll.config

### Deploy to Target System

The initial deployment must copy the entire published folder (e.g.
publish\ProcessTimesheets), and this takes some time (depending on network and VPN
throughput).  It is typical for this initial deployment to take at least 20 seconds
(but not more than a minute).

Note that Deploy-Application will fail if the target folder contains any files.

However, subsequent (minor change) re-deployments (if not changing .net dependencies) 
can be more concise in delivering only the relevant changed parts.  This is achieved
by using the -Include parameter to Redeploy-Application, providing a filter expression
such as ImapAttachmentProcessing* or ProcessTimesheets*.

#### New Deployment

Local deployment:

```
$psdrive = New-LocalSystemDrive -TargetHost localhost
```

Remote deployment:

```
$psdrive = New-RemoteSystemDrive
```

All deployments (local/remote):

```
New-Deployment -PSDrive $psdrive `
  -ToolName (ProcessTimesheets|DownloadInvoices) -DeployTo <folder-name>
```

#### Existing Deployment

To update components of an existing deployment:

```
Set-DeploymentComponents -PSDrive $psdrive `
  -ToolName (ProcessTimesheets|DownloadInvoices) -DeployTo <folder-name> `
  [-Components <component1,...>]
```

Note that updating components does NOT modify the deployed configuration file.

### Configuration

* Edit *project*.dll.config - on the target host - to configure the required
parameters
* Ensure that the project*.dll.config file is in the same folder as the
 *project*.exe file.
 

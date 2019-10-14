enum ToolName {
    ProcessTimesheets
    DownloadInvoices
}

function New-RemoteSystemDrive {
    param(
        [Parameter(Mandatory)][string]$TargetHost,
        [Parameter(Mandatory)][PSCredential]$Credentials
    )

    $driveRoot = "\\${TargetHost}\c$"
    $driveName = "${TargetHost}c"
    return New-PSDrive -PSProvider FileSystem `
        -Name $driveName -Credential $Credentials -Root $driveRoot -Scope Global;
}

function New-LocalSystemDrive {
    param(
        [Parameter(Mandatory)][string]$TargetHost
    )

    $driveRoot = "\\${TargetHost}\c$"
    $driveName = "${TargetHost}c"
    return New-PSDrive -PSProvider FileSystem -Name $driveName -Root $driveRoot -Scope Global;
}

function Set-DeploymentComponents
{
    param(
        [Parameter(Mandatory)][ToolName]$ToolName,
        [Parameter(Mandatory)][string]$DeployTo,
        [Parameter(Mandatory)][System.Management.Automation.PSDriveInfo]$PSDrive,
        [array]$Components = $null
    )

    # Assume that we will deploy to the System (C) drive on the target host.
    $drive = $PSDrive;

    $publishDir = ".\publish\${ToolName}";
    $targetDir = "${drive}:$DeployTo";

    $pubDir = Get-Item -Path $targetDir -ErrorAction SilentlyContinue;
    if ($pubDir.Exists) {
        if ($null -eq $components) {
            Copy-Item -Path "$publishDir\*" -Exclude *.config -Destination $targetDir -Verbose;
        } else {
            $components | ForEach-Object {
                $component = "$publishDir\$_";
                Copy-Item -Path "${component}*" -Exclude *.config -Destination $targetDir -Verbose;
            }
        }
    } else {
        Write-Host -ForegroundColor Red 'Deployment does not exist.';
    }
}

function New-Deployment
{
    param(
        [Parameter(Mandatory)][ToolName]$ToolName,
        [Parameter(Mandatory)][string]$DeployTo,
        [Parameter(Mandatory)][System.Management.Automation.PSDriveInfo]$PSDrive
    )

    # Assume that we will deploy to the System (C) drive on the target host.
    $drive = $PSDrive;

    $publishDir = ".\publish\${ToolName}";
    $confFile = ".\publish\${ToolName}\Properties\${ToolName}.dll.config";
    $targetDir = "${drive}:$DeployTo";

    $pubDir = Get-Item -Path $targetDir -ErrorAction SilentlyContinue;
    if ($pubDir.Exists) {
        Write-Host -ForegroundColor Red 'Deployment already exists.';
    } else {
        New-Item -ItemType Directory -Path "$targetDir" -Force -Verbose;
        Copy-Item -Path "$publishDir\*" -Destination $targetDir -Exclude Properties -Verbose;
        Copy-Item -LiteralPath $confFile $targetDir -Verbose;
    }
}

function Publish-Tool {
    param (
        [Parameter(Mandatory)][ToolName]$ToolName
    )

    dotnet publish -c Release -r win-x64 --self-contained true -o ".\publish\${ToolName}" $ToolName
}

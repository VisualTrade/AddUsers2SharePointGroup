<#
.SYNOPSIS
    Builds and publishes the "Add Users" Outlook add-in (VSTO/ClickOnce) to Azure Blob Storage.

.DESCRIPTION
    - Auto-increments the four-part version stored in build\version.txt (or uses -Version).
    - Locates MSBuild through vswhere.
    - Runs msbuild /t:Restore, then /t:Publish with ClickOnce properties pointing at the
      public blob container. UpdateInterval=0 makes installed clients check for a new
      version every time the add-in loads.
    - Uploads AddUsers\bin\Release\app.publish\** to the container, preserving the
      'Application Files/...' folder structure.
    - Fixes blob content types for .vsto / .manifest / .deploy / setup.exe so the VSTO
      runtime accepts the downloads.
    - Prints the final install URL.

    Prerequisites: Visual Studio with the Office/SharePoint development workload,
    Azure CLI logged in (az login), and the resources created by setup-azure.ps1.

.EXAMPLE
    .\build\publish.ps1 -StorageAccountName adduserssa

.EXAMPLE
    .\build\publish.ps1 -StorageAccountName adduserssa -Version 2.0.0.0
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-z0-9]{3,24}$')]
    [string]$StorageAccountName,

    [string]$ResourceGroup = 'rg-addusers',

    [string]$Container = 'addusers',

    # Explicit four-part version. When omitted, the version in build\version.txt is
    # incremented. Either way the published version is written back to version.txt.
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'

function Assert-ExitCode([string]$Step) {
    if ($LASTEXITCODE -ne 0) { throw "$Step failed with exit code $LASTEXITCODE." }
}

$repoRoot    = Split-Path -Parent $PSScriptRoot
$projectFile = Join-Path $repoRoot 'AddUsers\AddUsers.csproj'
$publishDir  = Join-Path $repoRoot 'AddUsers\bin\Release\app.publish'
$versionFile = Join-Path $PSScriptRoot 'version.txt'
$baseUrl     = "https://$StorageAccountName.blob.core.windows.net/$Container"

if (-not (Test-Path $projectFile)) { throw "Project file not found: $projectFile" }

# --- Version -----------------------------------------------------------------
if ($Version) {
    $newVersion = $Version
}
elseif (Test-Path $versionFile) {
    $current = (Get-Content $versionFile -TotalCount 1).Trim()
    if ($current -notmatch '^(\d+)\.(\d+)\.(\d+)\.(\d+)$') {
        throw "Expected a four-part version (e.g. 1.0.0.0) in $versionFile, found '$current'."
    }
    $newVersion = '{0}.{1}.{2}.{3}' -f $Matches[1], $Matches[2], $Matches[3], ([int]$Matches[4] + 1)
}
else {
    $newVersion = '1.0.0.0'
}

# --- Tooling -----------------------------------------------------------------
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswhere)) {
    throw 'vswhere.exe not found. Install Visual Studio 2019/2022 with the "Office/SharePoint development" workload.'
}

# Prefer a VS instance that has the Office workload (it carries the VSTO publish targets),
# fall back to any MSBuild with a warning.
$msbuild = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Workload.Office `
    -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuild) {
    $msbuild = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild `
        -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
    if ($msbuild) {
        Write-Warning 'No Visual Studio instance with the Office/SharePoint workload found; using the newest MSBuild. The VSTO publish targets may be missing.'
    }
}
if (-not $msbuild) {
    throw 'MSBuild not found via vswhere. Install Visual Studio with the "Office/SharePoint development" workload.'
}

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw 'Azure CLI (az) not found on PATH. Install it from https://aka.ms/installazurecli.'
}
az account show --output none --only-show-errors
if ($LASTEXITCODE -ne 0) { throw "Azure CLI is not logged in. Run 'az login' first." }

Write-Host "Publishing Add Users $newVersion"
Write-Host "  MSBuild:     $msbuild"
Write-Host "  Install URL: $baseUrl/"

# --- Build & publish ---------------------------------------------------------
# Remove stale publish output so only the new version is uploaded.
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }

& $msbuild $projectFile /t:Restore /p:Configuration=Release /nologo /verbosity:minimal
Assert-ExitCode 'msbuild /t:Restore'

# PublishUrl and IsWebBootstrapper are required by the VSTO publish targets when
# building from the command line (VS normally sets them in the publish wizard):
# without PublishUrl the GenerateOfficeDocumentInstallationPath task fails, and
# without IsWebBootstrapper=true the setup.exe bootstrapper is never generated.
& $msbuild $projectFile /t:Publish /nologo /verbosity:minimal `
    /p:Configuration=Release `
    "/p:PublishUrl=$baseUrl/" `
    "/p:InstallUrl=$baseUrl/" `
    /p:IsWebBootstrapper=true `
    /p:UpdateEnabled=true `
    /p:UpdateInterval=0 `
    /p:MapFileExtensions=true `
    /p:ApplicationVersion=$newVersion `
    '/p:ProductName=Add Users' `
    /p:PublisherName=VisualTrade
Assert-ExitCode 'msbuild /t:Publish'

if (-not (Test-Path (Join-Path $publishDir 'AddUsers.vsto'))) {
    throw "Publish output not found at $publishDir (expected AddUsers.vsto)."
}

# Record the version only after a successful build.
Set-Content -Path $versionFile -Value $newVersion -Encoding Ascii

# --- Upload ------------------------------------------------------------------
$accountKey = az storage account keys list `
    --resource-group $ResourceGroup `
    --account-name $StorageAccountName `
    --query '[0].value' --output tsv --only-show-errors
Assert-ExitCode 'az storage account keys list'

# Because installed clients check for updates on every add-in load, the upload order
# matters: the new version's payload must be fully in place (with correct content
# types) BEFORE the root pointers are re-pointed at it. A client that checks mid-publish
# then still sees the old, intact version instead of 404ing on a half-uploaded new one.

# Blob storage guesses content types from extensions and gets these wrong; the
# ClickOnce/VSTO runtime is strict about them.
$contentTypes = @{
    '.vsto'     = 'application/x-ms-vsto'
    '.manifest' = 'application/x-ms-manifest'
    '.deploy'   = 'application/octet-stream'
    '.exe'      = 'application/octet-stream'   # setup.exe bootstrapper
}

# 1) New version payload.
$appFilesDir = Join-Path $publishDir 'Application Files'
Write-Host "Uploading new version payload to container '$Container'..."
az storage blob upload-batch `
    --account-name $StorageAccountName `
    --account-key $accountKey `
    --destination $Container `
    --destination-path 'Application Files' `
    --source $appFilesDir `
    --overwrite `
    --no-progress `
    --output none --only-show-errors
Assert-ExitCode 'az storage blob upload-batch (Application Files)'

# 2) Payload content types, still before any client is pointed at the new version.
Write-Host 'Setting payload content types...'
Get-ChildItem -Path $appFilesDir -Recurse -File | ForEach-Object {
    $contentType = $contentTypes[$_.Extension.ToLowerInvariant()]
    if ($contentType) {
        $blobName = $_.FullName.Substring($publishDir.Length + 1) -replace '\\', '/'
        az storage blob update `
            --account-name $StorageAccountName `
            --account-key $accountKey `
            --container-name $Container `
            --name $blobName `
            --content-type $contentType `
            --output none --only-show-errors
        Assert-ExitCode "az storage blob update ($blobName)"
    }
}

# 3) Root pointers last, content types set at upload time. AddUsers.vsto goes very
#    last: it is the blob that re-points installed clients at the new version.
Write-Host 'Uploading root install pointers...'
az storage blob upload `
    --account-name $StorageAccountName `
    --account-key $accountKey `
    --container-name $Container `
    --file (Join-Path $publishDir 'setup.exe') `
    --name 'setup.exe' `
    --content-type 'application/octet-stream' `
    --overwrite `
    --output none --only-show-errors
Assert-ExitCode 'az storage blob upload (setup.exe)'

az storage blob upload `
    --account-name $StorageAccountName `
    --account-key $accountKey `
    --container-name $Container `
    --file (Join-Path $publishDir 'AddUsers.vsto') `
    --name 'AddUsers.vsto' `
    --content-type 'application/x-ms-vsto' `
    --overwrite `
    --output none --only-show-errors
Assert-ExitCode 'az storage blob upload (AddUsers.vsto)'

# --- Done --------------------------------------------------------------------
Write-Host ''
Write-Host "Published Add Users $newVersion" -ForegroundColor Green
Write-Host ''
Write-Host "  First-time install (installs prerequisites): $baseUrl/setup.exe"
Write-Host "  Install URL:                                 $baseUrl/AddUsers.vsto"
Write-Host ''
Write-Host 'Installed clients pick up this version automatically the next time Outlook starts.'

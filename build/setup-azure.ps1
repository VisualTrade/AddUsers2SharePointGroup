<#
.SYNOPSIS
    One-time Azure setup for hosting the "Add Users" VSTO/ClickOnce deployment.

.DESCRIPTION
    Creates a resource group, a StorageV2 account, and a blob container with public
    (anonymous) read access for blobs. Public blob read is required because the VSTO
    runtime on end-user machines downloads the deployment manifests and payload
    without credentials. Only blob reads are public; the container cannot be listed.

    Run once, then publish with build\publish.ps1 using the same parameter values.

.EXAMPLE
    az login
    .\build\setup-azure.ps1 -StorageAccountName adduserssa
#>
[CmdletBinding()]
param(
    # Storage account names are global across Azure: 3-24 chars, lowercase letters and digits.
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-z0-9]{3,24}$')]
    [string]$StorageAccountName,

    [string]$ResourceGroup = 'rg-addusers',

    [string]$Container = 'addusers',

    [string]$Location = 'eastus'
)

$ErrorActionPreference = 'Stop'

function Assert-ExitCode([string]$Step) {
    if ($LASTEXITCODE -ne 0) { throw "$Step failed with exit code $LASTEXITCODE." }
}

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw 'Azure CLI (az) not found on PATH. Install it from https://aka.ms/installazurecli.'
}
az account show --output none --only-show-errors
if ($LASTEXITCODE -ne 0) { throw "Azure CLI is not logged in. Run 'az login' first." }

Write-Host "Creating resource group '$ResourceGroup' in $Location..."
az group create --name $ResourceGroup --location $Location --output none --only-show-errors
Assert-ExitCode 'az group create'

Write-Host "Creating storage account '$StorageAccountName'..."
az storage account create `
    --name $StorageAccountName `
    --resource-group $ResourceGroup `
    --location $Location `
    --sku Standard_LRS `
    --kind StorageV2 `
    --min-tls-version TLS1_2 `
    --allow-blob-public-access true `
    --output none --only-show-errors
Assert-ExitCode 'az storage account create'

$accountKey = az storage account keys list `
    --resource-group $ResourceGroup `
    --account-name $StorageAccountName `
    --query '[0].value' --output tsv --only-show-errors
Assert-ExitCode 'az storage account keys list'

# The allow-blob-public-access flag can take a moment to propagate on a brand-new
# account, so retry the container creation a few times.
Write-Host "Creating container '$Container' with public read access for blobs..."
$maxAttempts = 3
for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    az storage container create `
        --name $Container `
        --account-name $StorageAccountName `
        --account-key $accountKey `
        --public-access blob `
        --output none --only-show-errors
    if ($LASTEXITCODE -eq 0) { break }
    if ($attempt -eq $maxAttempts) {
        throw "az storage container create failed after $maxAttempts attempts. If the error is 'PublicAccessNotPermitted', check whether an Azure Policy in your subscription forbids public blob access."
    }
    Write-Warning 'Container creation failed (public-access flag may still be propagating); retrying in 15 seconds...'
    Start-Sleep -Seconds 15
}

# 'container create' exits 0 without changing anything when the container already
# exists — including one that was created private. Converge the ACL explicitly so a
# pre-existing private container cannot silently break anonymous installs with 404s.
az storage container set-permission `
    --name $Container `
    --account-name $StorageAccountName `
    --account-key $accountKey `
    --public-access blob `
    --output none --only-show-errors
Assert-ExitCode 'az storage container set-permission'

Write-Host ''
Write-Host 'Azure hosting is ready.' -ForegroundColor Green
Write-Host ''
Write-Host "  Deployment base URL: https://$StorageAccountName.blob.core.windows.net/$Container/"
Write-Host ''
Write-Host 'Next step:'
Write-Host "  .\build\publish.ps1 -StorageAccountName $StorageAccountName -ResourceGroup $ResourceGroup -Container $Container"

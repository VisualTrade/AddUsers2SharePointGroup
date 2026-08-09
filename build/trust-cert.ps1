<#
.SYNOPSIS
    Trusts the "Add Users" signing certificate on this machine (run as Administrator).

.DESCRIPTION
    ClickOnce/VSTO refuses to install from an Internet-zone URL (like Azure blob
    storage) unless the manifest-signing certificate chains to a trusted root AND,
    for a silent install, is a trusted publisher. For a self-signed certificate that
    means importing its PUBLIC key (.cer — contains no private key) into:

      - LocalMachine\Root              (makes the certificate chain valid)
      - LocalMachine\TrustedPublisher  (makes the install run without prompts)

    Without this, users get: "Customized functionality in this application will not
    work because the certificate used to sign the deployment manifest ... or its
    location is not trusted."

    Deploy this to end-user machines via Intune/GPO/SCCM, or run it once per machine
    from an elevated PowerShell. (GPO can also distribute the .cer directly via
    Computer Configuration > Policies > Windows Settings > Security Settings >
    Public Key Policies.)

    Not needed if you re-sign the add-in with a CA-issued code-signing certificate —
    then only the TrustedPublisher import (or clicking the install prompt) applies.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\trust-cert.ps1
#>
[CmdletBinding()]
param(
    # Public certificate exported from AddUsers_TemporaryKey.pfx. Defaults to the .cer
    # next to this script, or to the current directory when the code runs outside a
    # script file ($PSScriptRoot is empty then, e.g. lines pasted into a console).
    [string]$CerPath
)

$ErrorActionPreference = 'Stop'

if (-not $CerPath) {
    $baseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $CerPath = Join-Path $baseDir 'AddUsersTemporaryKey.cer'
}

if (-not (Test-Path $CerPath)) {
    throw "Certificate file not found: $CerPath. Place AddUsersTemporaryKey.cer next to this script or pass -CerPath."
}

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'This script must run as Administrator (machine certificate stores are protected).'
}

$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CerPath)
Write-Host "Trusting certificate: $($cert.Subject) (thumbprint $($cert.Thumbprint))"

Import-Certificate -FilePath $CerPath -CertStoreLocation Cert:\LocalMachine\Root | Out-Null
Write-Host '  Imported into LocalMachine\Root'

Import-Certificate -FilePath $CerPath -CertStoreLocation Cert:\LocalMachine\TrustedPublisher | Out-Null
Write-Host '  Imported into LocalMachine\TrustedPublisher'

Write-Host ''
Write-Host 'Done. The Add Users ClickOnce install will now be trusted on this machine.' -ForegroundColor Green

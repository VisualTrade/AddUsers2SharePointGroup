<#
.SYNOPSIS
    Trusts the "Add Users" signing certificate on this machine (run as Administrator).

.DESCRIPTION
    ClickOnce/VSTO refuses to install from an Internet-zone URL (like Azure blob
    storage) unless the manifest-signing certificate chains to a trusted root AND,
    for a silent install, is a trusted publisher. This script imports the PUBLIC
    certificate (no private key) into:

      - LocalMachine\Root              (makes the certificate chain valid)
      - LocalMachine\TrustedPublisher  (makes the install run without prompts)

    Without this, users get: "Customized functionality in this application will not
    work because the certificate used to sign the deployment manifest ... or its
    location is not trusted."

    The certificate is EMBEDDED below, so this script is the only file that needs
    copying to a machine. Deploy via Intune/GPO/SCCM, or run once per machine from
    an elevated PowerShell. Not needed once the add-in is re-signed with a CA-issued
    code-signing certificate.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\trust-cert.ps1

.EXAMPLE
    # Trust a different exported certificate instead of the embedded one:
    powershell -ExecutionPolicy Bypass -File .\trust-cert.ps1 -CerPath .\SomeOther.cer
#>
[CmdletBinding()]
param(
    # Optional: a .cer file to trust INSTEAD of the certificate embedded below.
    [string]$CerPath
)

$ErrorActionPreference = 'Stop'

# Public certificate for "CN=AddUsers Temporary Key" (DER, base64-encoded).
# Thumbprint: 5F2DAAE0B745C9CBF2FAD0E596F0A14C6B93058A
$EmbeddedCertBase64 = @'
MIIDEjCCAfqgAwIBAgIQTuQ8IKzvaJ9IojqQRXxwbTANBgkqhkiG9w0BAQsFADAhMR8wHQYDVQQDDBZBZGRVc2VycyBUZW1wb3JhcnkgS2V5MB4XDTI2MDgwODE3MTQxMloXDTMxMDgwODE3MjQxMlowITEfMB0GA1UEAwwWQWRkVXNlcnMgVGVtcG9yYXJ5IEtleTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBANh0HeONu1eJ4BX+e6L8Lp8pAHOnQmiQFdcZ/28jvYlsRmDjd0xg/DNKjYa7t0AqqryL9Yx/YtYdhu4vd/kcZbq8Sqrdr28xJcuCvYZU68x15fVrOmiVA5NAsqoc2ZdAB64YAr0hI10pKKPsx1LkpowY7kPNlnBWe8evUb5Fe40I7QZfAIoigV6izVPJ7DDSaU0bHMTs5muUhnxNBV516CFlOFugoPF47ySwYT83njS6Vu7uheDHnrqR5PFHhRU7Omy6JHjlVR5CsVeqIAO8xYmYuT4Se67VACnnVoTJjHz8snRNYL/UWPb6iWzVu/H3unP/aL8hG2mfzy9FKC1hl30CAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBTctRSfPz7A4fccPNBPzRWVQP/yDDANBgkqhkiG9w0BAQsFAAOCAQEAR1gn6PV27BouVmJzoPwp9u3rRK884gomx5XV9CWG7trXgd16G8ooW/UjYCDmRiuWQn0x4dorDCxNGJB+0wmGENB8hOq7LvfKZ8nGETlC88q/HdglemdnUNa3hkuZyYj9OMcEzx43qBxVTAMXQbSJ7ixb8LOhhUe4uL5YA0TVBTKHuyJoXPn4WL9ogCDMnHUa6l5g10BtCTaPSGIGxCRXrbj8tW2c6bL2Q9Ivamk/D7rAJOUJhqtxZh3/Ct3XX8JWECf89WBKF7Yf5xb8giw4E3ysRxK8e4rrQqhH7m/SzCioaicF/etzi7uBV0+isaeMnhU9c0++VJFl2ooPtRTq8w==
'@

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'This script must run as Administrator (machine certificate stores are protected).'
}

if ($CerPath) {
    if (-not (Test-Path $CerPath)) {
        throw "Certificate file not found: $CerPath"
    }
    $certBytes = [IO.File]::ReadAllBytes($CerPath)
    $source = $CerPath
}
else {
    $certBytes = [Convert]::FromBase64String(($EmbeddedCertBase64 -replace '\s', ''))
    $source = 'embedded certificate'
}

try {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, $certBytes)
}
catch {
    throw "The certificate data ($source) is not a valid certificate — the file is likely a corrupt or HTML-wrapped download. ($($_.Exception.Message))"
}

Write-Host "Trusting certificate: $($cert.Subject) (thumbprint $($cert.Thumbprint))"

foreach ($storeName in 'Root', 'TrustedPublisher') {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'LocalMachine')
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    try {
        $store.Add($cert)   # no-op if already present
    }
    finally {
        $store.Close()
    }
    Write-Host "  Imported into LocalMachine\$storeName"
}

Write-Host ''
Write-Host 'Done. The Add Users ClickOnce install will now be trusted on this machine.' -ForegroundColor Green

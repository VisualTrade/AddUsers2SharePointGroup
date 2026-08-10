#Requires -Version 5
<#
.SYNOPSIS
    Registers a per-user scheduled task that keeps the "Add Users" add-in updated.

.DESCRIPTION
    VSTO's built-in update check runs INSIDE Outlook at startup, and applying an
    update there can fail with "Access to the path ... is denied" because the
    running Outlook process (or a lingering outlook.exe from the previous session)
    holds locks in the ClickOnce store. Applying the update outside the host is the
    reliable pattern.

    This script registers a scheduled task for the CURRENT user (no admin rights
    needed) that runs at logon and once daily:

        VSTOInstaller.exe /install <deployment-url> /silent

    At logon Outlook is normally not running yet, so the update applies cleanly;
    Outlook then starts with the new version already in place and its own startup
    check finds nothing to do. The silent install works because the signing
    certificate is in Trusted Publishers (see trust-cert.ps1). If the task fires
    while Outlook happens to be running and hits the same locks, it fails silently
    and retries at the next trigger — no user-facing errors.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\register-update-task.ps1

.EXAMPLE
    # Remove the task again:
    powershell -ExecutionPolicy Bypass -File .\register-update-task.ps1 -Remove
#>
[CmdletBinding()]
param(
    [string]$InstallUrl = 'https://addusersvisualtrade.blob.core.windows.net/addusers/AddUsers.vsto',

    [string]$TaskName = 'AddUsers add-in update',

    # Daily retry time for machines that stay logged on for days.
    [string]$DailyTime = '09:00',

    [switch]$Remove
)

$ErrorActionPreference = 'Stop'

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Scheduled task '$TaskName' removed (if it existed)."
    return
}

$vstoInstaller = @(
    "$env:CommonProgramFiles\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe",
    "${env:CommonProgramFiles(x86)}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"
) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if (-not $vstoInstaller) {
    throw 'VSTOInstaller.exe not found - install the VSTO Runtime first (setup.exe does this).'
}

$action = New-ScheduledTaskAction -Execute $vstoInstaller -Argument "/install `"$InstallUrl`" /silent"
$triggers = @(
    New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
    New-ScheduledTaskTrigger -Daily -At $DailyTime
)
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $triggers `
    -Settings $settings `
    -Description 'Applies pending ClickOnce updates for the Add Users Outlook add-in outside of Outlook, where the update cannot hit file locks.' `
    -Force | Out-Null

Write-Host "Scheduled task '$TaskName' registered for $env:USERNAME (at logon + daily $DailyTime)." -ForegroundColor Green
Write-Host "It runs: `"$vstoInstaller`" /install `"$InstallUrl`" /silent"
Write-Host 'Updates now apply before Outlook starts; the in-Outlook check will simply find itself current.'

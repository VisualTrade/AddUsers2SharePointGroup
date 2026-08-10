#Requires -Version 7.0
<#
.SYNOPSIS
    Diagnoses why the "Add Users" add-in cannot add users to a SharePoint group.

.DESCRIPTION
    Signs in interactively with the SAME app registration the add-in uses, then
    answers, with server-side evidence:

      1. What delegated scopes does the access token ACTUALLY contain?
         (If consent was granted before a permission was added to the app
         registration, `.default` tokens silently lack the new scope.)
      2. Who does SharePoint think you are on the target site? Site admin?
      3. What are your effective base permissions under this (scope-capped) token —
         in particular ManagePermissions, which AllSites.Manage strips?
      4. The target group's settings and, decisively, SharePoint's own verdict:
         CanCurrentUserEditMembership / CanCurrentUserManageGroup.
      5. Whether EnsureUser works for the recipient that failed (-TestEmail).
      6. Optionally (-AttemptAdd) the real add call, capturing the exact error.

    Prints a verdict with the most likely cause and fix, and writes a JSON report
    (no tokens, only claims) you can share for analysis.

.EXAMPLE
    ./debug-permissions.ps1 -SiteUrl https://contoso.sharepoint.com/sites/projects `
        -ClientId 11111111-2222-3333-4444-555555555555 -GroupName 'Project Members' `
        -TestEmail person.that.failed@contoso.com

.EXAMPLE
    # Additionally attempt the real group add (makes an actual change on success):
    ./debug-permissions.ps1 -SiteUrl ... -ClientId ... -GroupName ... `
        -TestEmail person@contoso.com -AttemptAdd
#>
[CmdletBinding()]
param(
    # The site URL configured in the add-in (the site that owns the group).
    [Parameter(Mandatory)][string]$SiteUrl,

    # Application (client) ID of the Entra app registration the add-in uses.
    [Parameter(Mandatory)][string]$ClientId,

    # Directory (tenant) ID; defaults to the multi-tenant 'organizations' authority,
    # matching the add-in's behavior when Tenant ID is left empty.
    [string]$TenantId = 'organizations',

    # Title of the SharePoint group the add-in is configured to add users to.
    [Parameter(Mandatory)][string]$GroupName,

    # Email of a recipient that fails in the add-in; used for the EnsureUser probe.
    # Defaults to the signed-in user's own address (a harmless self-lookup).
    [string]$TestEmail,

    # Actually attempt adding -TestEmail to the group. THIS MAKES A REAL CHANGE
    # when it succeeds; without it the script is read-only apart from EnsureUser.
    [switch]$AttemptAdd,

    # Where to write the JSON report (default: AddUsers-PermissionReport-<time>.json
    # in the current directory).
    [string]$ReportPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$SiteUrl = $SiteUrl.Trim().TrimEnd('/')
$siteHost = ([Uri]$SiteUrl).Host
$resource = "https://$siteHost"
if (-not $ReportPath) {
    $ReportPath = Join-Path (Get-Location) ("AddUsers-PermissionReport-{0:yyyyMMdd-HHmmss}.json" -f (Get-Date))
}

$report = [ordered]@{
    generated = (Get-Date).ToString('o')
    siteUrl   = $SiteUrl
    clientId  = $ClientId
    tenantId  = $TenantId
    groupName = $GroupName
}

function Write-Step([string]$Text) { Write-Host "`n=== $Text ===" -ForegroundColor Cyan }
function Write-Good([string]$Text) { Write-Host "  [OK]   $Text" -ForegroundColor Green }
function Write-Bad([string]$Text)  { Write-Host "  [FAIL] $Text" -ForegroundColor Red }
function Write-Info([string]$Text) { Write-Host "  $Text" }

function ConvertTo-Base64Url([byte[]]$Bytes) {
    [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function Get-InteractiveToken {
    param(
        [Parameter(Mandatory)][string]$Scope,
        [string]$Prompt = 'select_account'
    )

    $verifierBytes = [byte[]]::new(32)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($verifierBytes)
    $verifier = ConvertTo-Base64Url $verifierBytes
    $challenge = ConvertTo-Base64Url ([System.Security.Cryptography.SHA256]::HashData([Text.Encoding]::ASCII.GetBytes($verifier)))

    # Bind a localhost listener; the app registration's http://localhost redirect
    # URI matches any port.
    $listener = $null
    $redirect = $null
    foreach ($attempt in 1..10) {
        $port = Get-Random -Minimum 42000 -Maximum 64000
        $candidate = [System.Net.HttpListener]::new()
        $candidate.Prefixes.Add("http://localhost:$port/")
        try {
            $candidate.Start()
            $listener = $candidate
            $redirect = "http://localhost:$port/"
            break
        }
        catch { $candidate.Close() }
    }
    if (-not $listener) { throw 'Could not bind a localhost port for the sign-in redirect.' }

    try {
        $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize" +
            "?client_id=$ClientId" +
            '&response_type=code&response_mode=query' +
            "&redirect_uri=$([Uri]::EscapeDataString($redirect))" +
            "&scope=$([Uri]::EscapeDataString($Scope))" +
            "&code_challenge=$challenge&code_challenge_method=S256" +
            "&prompt=$Prompt"

        Write-Info "Opening browser for sign-in (scope: $Scope, prompt: $Prompt)..."
        Start-Process $authUrl | Out-Null

        $contextTask = $listener.GetContextAsync()
        if (-not $contextTask.Wait(300000)) {
            throw 'Timed out waiting for the browser sign-in (5 minutes). Close the browser tab and rerun.'
        }
        $context = $contextTask.Result
        $code = $context.Request.QueryString['code']
        $authError = $context.Request.QueryString['error']
        $authErrorDesc = $context.Request.QueryString['error_description']

        $html = if ($code) { '<html><body><h3>Sign-in complete.</h3>You can close this tab and return to PowerShell.</body></html>' }
                else       { "<html><body><h3>Sign-in failed.</h3><pre>$authError`n$authErrorDesc</pre></body></html>" }
        $buffer = [Text.Encoding]::UTF8.GetBytes($html)
        $context.Response.ContentType = 'text/html'
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.Close()

        if (-not $code) {
            throw "Authorization failed: $authError - $authErrorDesc"
        }
    }
    finally {
        $listener.Stop()
        $listener.Close()
    }

    try {
        $tokenResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body @{
                client_id     = $ClientId
                grant_type    = 'authorization_code'
                code          = $code
                redirect_uri  = $redirect
                code_verifier = $verifier
                scope         = $Scope
            }
    }
    catch {
        # ErrorDetails is null on transport failures and empty-body proxy responses.
        $detail = if ($_.ErrorDetails) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw "Token request failed: $detail"
    }
    return $tokenResponse.access_token
}

function ConvertFrom-Jwt([string]$Token) {
    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
    [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
}

function Get-Claim($Claims, [string]$Name) {
    $prop = $Claims.PSObject.Properties[$Name]
    if ($prop) { $prop.Value } else { $null }
}

$script:AccessToken = $null

function Invoke-Spo {
    param(
        [Parameter(Mandatory)][string]$Path,   # relative, e.g. /_api/web/currentuser
        [string]$Method = 'GET',
        $Body = $null,
        [switch]$VerboseOData                  # SP.User POSTs need odata=verbose metadata
    )
    $odata = $VerboseOData ? 'application/json;odata=verbose' : 'application/json;odata=nometadata'
    $params = @{
        Uri                = "$SiteUrl$Path"
        Method             = $Method
        Headers            = @{ Authorization = "Bearer $script:AccessToken"; Accept = $odata }
        SkipHttpErrorCheck = $true
    }
    if ($null -ne $Body) {
        $params.ContentType = $odata
        $params.Body = ($Body | ConvertTo-Json -Depth 5)
    }
    $response = Invoke-WebRequest @params
    $parsed = $null
    if ($response.Content) {
        try { $parsed = $response.Content | ConvertFrom-Json } catch { $parsed = $response.Content }
    }
    [pscustomobject]@{
        Status = [int]$response.StatusCode
        Ok     = ([int]$response.StatusCode -ge 200 -and [int]$response.StatusCode -lt 300)
        Data   = $parsed
    }
}

function Get-SpoErrorText($Result) {
    if ($null -eq $Result.Data) { return "HTTP $($Result.Status)" }
    $text = $null
    if ($Result.Data -is [string]) { $text = $Result.Data }
    else {
        $odataError = $Result.Data.PSObject.Properties['odata.error'] ?? $Result.Data.PSObject.Properties['error']
        if ($odataError) {
            $val = $odataError.Value
            $msg = $val.PSObject.Properties['message']
            if ($msg) {
                $text = ($msg.Value -is [string]) ? $msg.Value : $msg.Value.value
                $codeProp = $val.PSObject.Properties['code']
                if ($codeProp) { $text = "$($codeProp.Value): $text" }
            }
        }
    }
    if (-not $text) { $text = ($Result.Data | ConvertTo-Json -Depth 4 -Compress) }
    "HTTP $($Result.Status) - $text"
}

# SPBasePermissions bits relevant to this add-in's operations.
$permissionBits = [ordered]@{
    ViewListItems        = 0
    ManageLists          = 11
    Open                 = 16
    ViewPages            = 17
    CreateGroups         = 24
    ManagePermissions    = 25
    BrowseDirectories    = 26
    BrowseUserInfo       = 27
    ManageWeb            = 30
    UseRemoteAPIs        = 37
    EnumeratePermissions = 62
}

# ---------------------------------------------------------------------------
# 1. Acquire a token exactly like the add-in does (resource/.default)
# ---------------------------------------------------------------------------
Write-Step "Sign-in 1: $resource/.default (what the add-in requests)"
$script:AccessToken = Get-InteractiveToken -Scope "$resource/.default"
$claims = ConvertFrom-Jwt $script:AccessToken

$tokenInfo = [ordered]@{
    audience    = Get-Claim $claims 'aud'
    scopes      = Get-Claim $claims 'scp'
    user        = (Get-Claim $claims 'upn') ?? (Get-Claim $claims 'unique_name')
    tenant      = Get-Claim $claims 'tid'
    appId       = Get-Claim $claims 'appid'
    expiresUtc  = ([DateTimeOffset]::FromUnixTimeSeconds([long](Get-Claim $claims 'exp'))).UtcDateTime.ToString('o')
}
# Snapshot copy: the re-consent path below updates $scopes, and the report must
# keep the evidence of what the add-in-mirroring .default token actually carried.
$report.defaultToken = [ordered]@{} + $tokenInfo
Write-Info "Signed in as:  $($tokenInfo.user)  (tenant $($tokenInfo.tenant))"
Write-Info "Token scopes:  $($tokenInfo.scopes)"

$consentFixedByRun = $false
$scopes = [string]$tokenInfo.scopes
if ($scopes -match 'AllSites\.(Manage|FullControl)') {
    Write-Good "Token carries $($Matches[0]) - consent is in place."
}
else {
    Write-Bad "Token does NOT contain any AllSites scope. `.default` only includes ALREADY-CONSENTED permissions - consent likely predates the permission change on the app registration."
    Write-Step 'Sign-in 2: explicit AllSites.Manage with prompt=consent (attempting to establish consent)'
    try {
        $script:AccessToken = Get-InteractiveToken -Scope "$resource/AllSites.Manage" -Prompt 'consent'
        $claims = ConvertFrom-Jwt $script:AccessToken
        $scopes = [string](Get-Claim $claims 'scp')
        $report.reconsentToken = [ordered]@{ scopes = $scopes }
        if ($scopes -match 'AllSites\.(Manage|FullControl)') {
            $consentFixedByRun = $true
            Write-Good "Consent established; token now carries: $scopes"
        }
        else {
            Write-Bad "Even after re-consent the token lacks AllSites scopes: $scopes"
        }
    }
    catch {
        $report.reconsentError = "$_"
        Write-Bad "Re-consent failed: $_"
        Write-Info 'If the error mentions AADSTS65001/consent, your tenant requires an admin to consent even for AllSites.Manage. Continuing with the limited token to gather the rest of the evidence.'
        $script:AccessToken = Get-InteractiveToken -Scope "$resource/.default"
    }
}

# ---------------------------------------------------------------------------
# 2. Who am I on this site?
# ---------------------------------------------------------------------------
Write-Step "Site identity: $SiteUrl"
$web = Invoke-Spo '/_api/web?$select=Title,Url'
if (-not $web.Ok) {
    $report.webError = Get-SpoErrorText $web
    Write-Bad "Cannot read the site at all: $($report.webError)"
    Write-Info 'Everything below depends on site access; report written with what we have.'
    $report | ConvertTo-Json -Depth 8 | Set-Content $ReportPath
    Write-Host "`nReport written to $ReportPath" -ForegroundColor Yellow
    return
}
Write-Good "Site reachable: '$($web.Data.Title)'"

$me = Invoke-Spo '/_api/web/currentuser?$select=Id,Title,LoginName,Email,IsSiteAdmin'
$report.currentUser = $me.Data
Write-Info "Current user:  $($me.Data.Title) <$($me.Data.Email)>  ($($me.Data.LoginName))"
if ($me.Data.IsSiteAdmin) { Write-Good 'IsSiteAdmin = TRUE (site collection administrator)' }
else                      { Write-Info 'IsSiteAdmin = false' }

$myGroups = Invoke-Spo '/_api/web/currentuser/groups?$select=Id,Title'
$myGroupList = @()
if ($myGroups.Ok -and $myGroups.Data.PSObject.Properties['value']) { $myGroupList = @($myGroups.Data.value) }
$report.currentUserGroups = $myGroupList
Write-Info ("Member of {0} site group(s): {1}" -f $myGroupList.Count, (($myGroupList | ForEach-Object { $_.Title }) -join ', '))

# ---------------------------------------------------------------------------
# 3. Effective permissions under THIS (scope-capped) token
# ---------------------------------------------------------------------------
Write-Step 'Effective base permissions on the web (as capped by the token scope)'
$perms = Invoke-Spo '/_api/web/effectiveBasePermissions'
$permReport = [ordered]@{}
if ($perms.Ok) {
    $high = [uint64]$perms.Data.High
    $low = [uint64]$perms.Data.Low
    $mask = ($high -shl 32) -bor $low
    $permReport.rawHigh = "$high"
    $permReport.rawLow = "$low"
    foreach ($name in $permissionBits.Keys) {
        $has = (($mask -shr $permissionBits[$name]) -band 1) -eq 1
        $permReport[$name] = $has
        if ($name -in 'ManagePermissions', 'CreateGroups', 'BrowseUserInfo', 'ManageWeb') {
            if ($has) { Write-Good "$name = true" } else { Write-Info "$name = false" }
        }
    }
    if (-not $permReport.ManagePermissions) {
        Write-Info '(ManagePermissions=false is EXPECTED under AllSites.Manage - the scope caps it. Group edits must then come via group OWNERSHIP.)'
    }
}
else {
    $permReport.error = Get-SpoErrorText $perms
    Write-Bad "Could not read effectiveBasePermissions: $($permReport.error)"
}
$report.effectivePermissions = $permReport

# ---------------------------------------------------------------------------
# 4. The target group and SharePoint's own membership-edit verdict
# ---------------------------------------------------------------------------
Write-Step "Target group: '$GroupName'"
$encodedGroup = [Uri]::EscapeDataString($GroupName.Replace("'", "''"))
# The expanded Owner is only returned when its fields are also projected in $select.
$group = Invoke-Spo ("/_api/web/sitegroups/getbyname('$encodedGroup')" +
    '?$select=Id,Title,OwnerTitle,AllowMembersEditMembership,OnlyAllowMembersViewMembership,' +
    'CanCurrentUserEditMembership,CanCurrentUserManageGroup,CanCurrentUserViewMembership,' +
    'Owner/Id,Owner/LoginName,Owner/PrincipalType&$expand=Owner')
if (-not $group.Ok) {
    $report.groupError = Get-SpoErrorText $group
    Write-Bad "Group lookup failed: $($report.groupError)"
}
else {
    $g = $group.Data
    $report.group = $g
    Write-Info "AllowMembersEditMembership = $($g.AllowMembersEditMembership)"

    # Owner detection is best-effort evidence only; CanCurrentUserEditMembership
    # below is the authoritative, server-computed answer.
    $iAmOwner = $false
    $ownerNote = $null
    if ($g.PSObject.Properties['Owner'] -and $null -ne $g.Owner) {
        Write-Info "Group Id $($g.Id), Owner: $($g.OwnerTitle) (PrincipalType $($g.Owner.PrincipalType); 1=user, 4=Entra security group, 8=SharePoint group)"
        switch ($g.Owner.PrincipalType) {
            1 { $iAmOwner = ($g.Owner.LoginName -eq $me.Data.LoginName) }
            8 {
                $iAmOwner = $null -ne ($myGroupList | Where-Object { $_.Id -eq $g.Owner.Id })
                if (-not $iAmOwner) { $ownerNote = 'You are not a DIRECT member of the owner group; membership via a nested Entra security group is invisible to this check.' }
            }
            default { $ownerNote = 'The owner is an Entra security group (or another principal type); direct membership cannot be checked from here.' }
        }
    }
    else {
        Write-Info "Group Id $($g.Id), Owner: $($g.OwnerTitle) (owner details not returned by the server)"
        $ownerNote = 'Owner details were not returned; relying on CanCurrentUserEditMembership below.'
    }
    $report.currentUserIsGroupOwner = $iAmOwner
    if ($ownerNote) { $report.ownerCheckNote = $ownerNote }
    if ($iAmOwner) { Write-Good 'You are the group owner (directly or via the owner group).' }
    else {
        Write-Info 'You are NOT (detectably) the group owner.'
        if ($ownerNote) { Write-Info "($ownerNote)" }
    }

    if ($g.CanCurrentUserEditMembership) {
        Write-Good 'CanCurrentUserEditMembership = TRUE - SharePoint says this token CAN edit the membership.'
    }
    else {
        Write-Bad 'CanCurrentUserEditMembership = FALSE - SharePoint says this token CANNOT edit the membership. This is the decisive flag.'
    }
}

# ---------------------------------------------------------------------------
# 5. EnsureUser probe (the add-in's first server call per recipient)
# ---------------------------------------------------------------------------
$probeEmail = $TestEmail
if (-not $probeEmail) { $probeEmail = $me.Data.Email }
if ($probeEmail) {
    Write-Step "EnsureUser probe: $probeEmail"
    $ensure = Invoke-Spo '/_api/web/ensureuser' -Method POST -Body @{ logonName = $probeEmail }
    if ($ensure.Ok) {
        $report.ensureUser = @{ email = $probeEmail; ok = $true; loginName = $ensure.Data.LoginName }
        Write-Good "EnsureUser resolved to: $($ensure.Data.LoginName)"
    }
    else {
        $report.ensureUser = @{ email = $probeEmail; ok = $false; error = (Get-SpoErrorText $ensure) }
        Write-Bad "EnsureUser failed: $($report.ensureUser.error)"
        Write-Info '(If this is the failing step, the problem is user resolution - e.g. external sharing settings for guests - not group permissions.)'
    }
}

# ---------------------------------------------------------------------------
# 6. Optional: the real add call
# ---------------------------------------------------------------------------
if ($AttemptAdd -and $TestEmail -and $group.Ok) {
    Write-Step "Attempting the REAL add: $TestEmail -> '$GroupName'"
    $login = if ($report.ensureUser.ok) { $report.ensureUser.loginName } else { $TestEmail }
    $add = Invoke-Spo "/_api/web/sitegroups/getbyid($($group.Data.Id))/users" -Method POST -VerboseOData `
        -Body @{ '__metadata' = @{ type = 'SP.User' }; LoginName = $login }
    if ($add.Ok) {
        $report.addAttempt = @{ ok = $true }
        Write-Good "Add SUCCEEDED - $TestEmail is now in the group (this was a real change)."
    }
    else {
        $report.addAttempt = @{ ok = $false; error = (Get-SpoErrorText $add) }
        Write-Bad "Add failed: $($report.addAttempt.error)"
    }
}

# ---------------------------------------------------------------------------
# Verdict
# ---------------------------------------------------------------------------
Write-Step 'Verdict'
$verdict = [System.Collections.Generic.List[string]]::new()

if ($consentFixedByRun) {
    $verdict.Add("CAUSE FOUND (and fixed by this run): the `.default` token - the exact kind the add-in uses - carried NO AllSites scope, because consent predated the permission change on the app registration. Every SharePoint write was unauthorized regardless of your rights. Sign-in 2 established consent. FIX: restart Outlook so the add-in discards its cached scope-less token and acquires a fresh one, then retry.")
}

if ($scopes -notmatch 'AllSites\.(Manage|FullControl)') {
    $verdict.Add('CAUSE: the token has no AllSites scope, so every SharePoint write is unauthorized regardless of your rights. FIX: consent could not be established even with an explicit request - if the error above mentions AADSTS65001/admin approval, your tenant requires an ADMIN to consent even for AllSites.Manage: have an admin open the app registration > API permissions > "Grant admin consent". Then restart Outlook.')
}
elseif (-not $group.Ok) {
    $verdict.Add("CAUSE: the group '$GroupName' could not be read on this site ($($report.groupError)) - the add-in fails the same way before permissions even matter. FIX: check the group name configured in the add-in against Site settings > People and groups; it must match the group Title exactly (groups are per-site, so also confirm the configured SITE is the right one).")
}
elseif (-not $group.Data.CanCurrentUserEditMembership) {
    $ownerFix = "FIX: in SharePoint (Site settings > People and groups > '$GroupName' > Settings) set the group OWNER to you or to a group you belong to (e.g. the site Owners group), or set 'Who can edit the membership of the group?' to Members and join the group. Alternatively switch the app registration to AllSites.FullControl (admin consent) so site-admin rights flow through."
    if ($me.Data.IsSiteAdmin -and ($report.effectivePermissions['ManagePermissions'] -eq $false)) {
        $verdict.Add("CAUSE: the AllSites.Manage cap strips site-level rights (ManagePermissions), and you are not the group's owner - so SharePoint denies membership edits even though you could do them in the browser as site admin. $ownerFix")
    }
    else {
        $verdict.Add("CAUSE: this account cannot edit the group's membership (not the owner, and no un-capped site-level rights). $ownerFix")
    }
}
elseif ($report.Contains('ensureUser') -and -not $report.ensureUser.ok) {
    $verdict.Add('CAUSE: group permissions are fine, but EnsureUser fails for the test address - user resolution is the failing step (typical for guests when external sharing is restricted, or for addresses not in the directory). The add-in reports these per-recipient in its results dialog.')
}
elseif (-not $consentFixedByRun) {
    $verdict.Add('No blocker found: the token has the right scope and SharePoint confirms membership-edit rights. If the add-in still fails, compare the signed-in ACCOUNT (this script showed which account the browser picked - a different cached account in the add-in would explain it), and rerun with -AttemptAdd and the exact failing recipient to capture the precise server error.')
}

foreach ($v in $verdict) { Write-Host "`n$v" -ForegroundColor Yellow }
$report.verdict = $verdict

$report | ConvertTo-Json -Depth 8 | Set-Content $ReportPath
Write-Host "`nFull report written to: $ReportPath" -ForegroundColor Yellow
Write-Host 'Share that file (it contains claims and permission data, no tokens).'

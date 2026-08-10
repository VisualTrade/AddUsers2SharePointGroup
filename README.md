# Add Users — Outlook add-in for SharePoint group membership

A VSTO add-in for classic Outlook on Windows. Select an email, click **Add Users** on the
ribbon, and the people on that message are added to a SharePoint Online site group of your
choice. A results dialog shows what happened for each address. The **Configure** button
stores the connection settings.

The add-in signs in interactively as the current user through your own Microsoft Entra app
registration (delegated permissions) and talks to SharePoint via
[PnP.Framework](https://github.com/pnp/pnpframework). It is deployed with ClickOnce from
Azure Blob Storage and checks for updates every time Outlook loads it.

- Product name: **Add Users**
- Publisher: **VisualTrade**
- Target: .NET Framework 4.7.2, classic Outlook desktop (not "new Outlook" / OWA)

## Repository layout

```
AddUsers2SharePointGroup.sln
AddUsers\                     VSTO add-in project (.NET Framework 4.7.2, C# 7.3)
  Ribbon\                     Ribbon XML and callbacks (Add Users / Configure buttons)
  Forms\                      Configuration and results dialogs
  Services\                   Outlook message parsing and SharePoint group operations
  Configuration\              Settings model and per-user persistence
  AddUsers_TemporaryKey.pfx   Signing key — NOT in the repo; created per dev machine (section 8)
build\
  publish.ps1                 Build + ClickOnce publish + upload to Azure Blob Storage
  setup-azure.ps1             One-time creation of the Azure hosting resources
  trust-cert.ps1              Trusts the signing cert on a machine (run elevated; see section 6)
  AddUsersTemporaryKey.cer    Public half of the signing cert, for distribution to machines
  version.txt                 Last published version (created/managed by publish.ps1)
```

## Prerequisites

| Task | You need |
|------|----------|
| Building | Visual Studio 2019 or 2022 with the **Office/SharePoint development** workload (installs the VSTO SDK and MSBuild targets) and the .NET Framework 4.7.2 developer pack |
| Publishing | [Azure CLI](https://aka.ms/installazurecli) and an Azure subscription |
| Entra setup | Rights to create app registrations and grant admin consent in your Microsoft 365 tenant |
| Running | Windows, classic Outlook (Microsoft 365 / 2016 or later), the [VSTO Runtime](https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-install-the-visual-studio-tools-for-office-runtime) (`setup.exe` installs it), and permission to manage the target SharePoint group |

## 1. Register a Microsoft Entra application (one time per tenant)

The add-in needs an Entra application identity to sign users in. PnP-based tools used to
fall back to the shared multi-tenant **PnP Management Shell** application, but Microsoft
retired and deleted that app registration in September 2024. Since then, every tenant must
register its **own** application and pass its client ID to PnP.Framework — there is no
shared default anymore. The upside: your admins get normal control over consent,
Conditional Access, and sign-in logs for this add-in.

1. Open the [Microsoft Entra admin center](https://entra.microsoft.com) and go to
   **Identity > Applications > App registrations > New registration**.
2. Name it, e.g. `Add Users Outlook Add-in`. Under *Supported account types* choose
   **Accounts in this organizational directory only** (single tenant). Click **Register**.
3. Go to **Authentication > Add a platform > Mobile and desktop applications** and add the
   custom redirect URI:

   ```
   http://localhost
   ```

   The interactive sign-in opens the system browser and completes on a loopback port, so
   this exact redirect URI is required.
4. Still under **Authentication**, set **Allow public client flows** to **Yes** and save.
   The add-in runs on user desktops and cannot keep a client secret, so it must be a
   public client.
5. Go to **API permissions > Add a permission > SharePoint > Delegated permissions** and
   add **AllSites.FullControl**. This is a hard requirement, not a preference: group
   membership changes are *security operations*, and SharePoint requires the app's own
   grant to cover them. With any lesser scope (`AllSites.Manage`/`Write`) membership
   edits are denied **even for users who own the group and can edit it in the browser**
   — verified empirically against a production tenant.
6. Click **Grant admin consent for \<your tenant\>** (FullControl always needs admin
   consent). Talking points for your admin: the permission is *delegated only* — the
   add-in has no standing access, every call runs as the signed-in user, and the token
   can never do more than that user can already do in the SharePoint UI.
7. From the **Overview** page copy the **Application (client) ID** and the
   **Directory (tenant) ID** — you will enter both in the add-in's Configure dialog.

Note on delegated permissions: the app never gets standing access of its own. The scope
is only the ceiling of what a token may do; every call still runs as the signed-in user
and is limited by that user's actual SharePoint permissions. A user who cannot edit a
group's membership in the browser cannot do it through the add-in either.

Why won't `AllSites.Manage` do, when group owners can edit membership in the browser
without any site-level rights? Because app-token authorization is two-sided: the *user*
must be allowed (owner status covers that), **and the app's grant must cover the
operation class**. Group membership is a security operation, which only the FullControl
scope covers — the owner special-case relaxes the user-side check, never the app-side
one. In practice: with `AllSites.Manage`, `CanCurrentUserEditMembership` comes back
false and every add fails with an authorization error, even for the group's owner.
(`build\debug-permissions.ps1` demonstrates this from any affected machine.)

## 2. Build

```powershell
git clone <this repo>
```

Open `AddUsers2SharePointGroup.sln` in Visual Studio and build, or from a Developer
PowerShell:

```powershell
msbuild AddUsers\AddUsers.csproj /t:Restore /p:Configuration=Release
msbuild AddUsers\AddUsers.csproj /t:Build   /p:Configuration=Release
```

Pressing **F5** in Visual Studio registers the add-in for debugging and launches Outlook
with it loaded.

If the build fails with *"Cannot find manifest signing certificate in the certificate
store"*, open the project's **Properties > Signing** tab and either re-select
`AddUsers_TemporaryKey.pfx` via **Select from File** or click **Create Test Certificate**
(both update the pinned certificate thumbprint in the project file).

## 3. Configure the add-in

In Outlook, click **Configure** in the add-in's ribbon group:

1. Enter the **Base URL** of your SharePoint tenant — e.g. `https://contoso.sharepoint.com`
   (not a specific site), the **Client ID** from step 1, and optionally the **Tenant ID**
   (leave empty to sign in against the multi-tenant `organizations` authority).
2. Click **Load sites** — a browser window opens for sign-in, then the **Site** dropdown
   fills with every site collection you can access (found via SharePoint search).
3. Pick a site; the **Group** dropdown fills with that site's SharePoint groups.
4. Pick the group and click **Save**.

Settings are stored per user in `%APPDATA%\AddUsers\settings.json`. Sign-ins reuse the
cached token for the rest of the Outlook session.

## 4. Create the Azure hosting (one time)

```powershell
az login
.\build\setup-azure.ps1 -StorageAccountName <globally-unique-name>
```

This creates:

- resource group `rg-addusers` (override with `-ResourceGroup`)
- a `StorageV2` storage account, `Standard_LRS`, TLS 1.2 minimum (`-Location`, default `eastus`)
- blob container `addusers` (`-Container`) with **anonymous read access for blobs**

Public blob read is required because the VSTO runtime on end-user machines downloads the
deployment anonymously. Only blob *reads* are public — nobody can list the container or
write to it. If your organization's policy forbids public blob access, host the same
files behind any HTTPS endpoint (Azure Static Website + Front Door, IIS, a file share)
and pass that URL scheme through the scripts instead.

## 5. Publish a version

```powershell
.\build\publish.ps1 -StorageAccountName <name>
```

The script:

1. Increments the version in `build\version.txt` (creates it at `1.0.0.0` on first run).
   Use `-Version 2.0.0.0` to set an explicit version instead.
2. Finds MSBuild via `vswhere.exe`, preferring a Visual Studio instance with the Office
   workload installed.
3. Runs `msbuild /t:Restore`, then `msbuild /t:Publish` with:

   | MSBuild property | Value | Purpose |
   |---|---|---|
   | `PublishUrl` | same as `InstallUrl` | Required by the VSTO publish targets on the command line |
   | `InstallUrl` | `https://<account>.blob.core.windows.net/<container>/` | Where installed clients look for updates |
   | `IsWebBootstrapper` | `true` | Makes the publish targets generate `setup.exe` (VS sets this silently in its wizard) |
   | `UpdateEnabled` | `true` | Enables automatic update checks |
   | `UpdateInterval` | `0` | Check for updates on **every** add-in load |
   | `MapFileExtensions` | `true` | Payload files get a `.deploy` suffix |
   | `ApplicationVersion` | from `build\version.txt` | ClickOnce deployment version |
   | `ProductName` | `Add Users` | Name shown in the install prompt and Apps list |
   | `PublisherName` | `VisualTrade` | Publisher shown in the install prompt |

4. Uploads the new `Application Files/AddUsers_<version>/...` payload **first** and fixes
   its content types (`.manifest` → `application/x-ms-manifest`, `.deploy` →
   `application/octet-stream` — blob storage guesses these wrong, and the VSTO runtime
   refuses mis-typed downloads).
5. Only then uploads the two root pointers, `setup.exe` and `AddUsers.vsto`
   (`application/x-ms-vsto`), with their content types set at upload time. Pointer-last
   ordering means a client that checks for updates *while* a publish is running still
   sees the old, fully intact version instead of a half-uploaded new one.
6. Prints the install URL.

Each publish adds a new `Application Files/AddUsers_<version>/` folder to the container
and re-points the root `AddUsers.vsto` and `setup.exe` at it. Older version folders stay
in the container; to roll back, re-run `publish.ps1 -Version <higher-version>` from the
older source, since clients only ever move to a *newer* deployment version.

## 6. Install on user machines

> **Trust the signing certificate first.** With the repo's self-signed test certificate,
> installing from the Azure URL fails outright with *"the certificate used to sign the
> deployment manifest ... or its location is not trusted"* — for Internet-zone URLs,
> Windows only offers the ClickOnce install prompt when the certificate chains to a
> trusted root. On each machine, run `build\trust-cert.ps1` from an **elevated**
> PowerShell — the script is self-contained (the public certificate is embedded in
> it), so it is the only file you need to copy. Alternatively push
> `build\AddUsersTemporaryKey.cer` to *Trusted Root Certification Authorities* +
> *Trusted Publishers* via Intune/GPO. A CA-issued code-signing certificate
> (section 8) removes this step.

Send users the URLs the script prints:

- **First-time install:** `https://<account>.blob.core.windows.net/<container>/setup.exe` —
  the bootstrapper installs .NET Framework 4.7.2 and the VSTO Runtime if missing, then the
  add-in.
- **Machines that already have the prerequisites:** open
  `https://<account>.blob.core.windows.net/<container>/AddUsers.vsto` directly.

The add-in itself installs per-user (no admin rights needed). The *prerequisites* that
`setup.exe` installs when missing — .NET Framework 4.7.2 and the VSTO Runtime — are
machine-wide and **do** require administrator rights: for non-admin users, have IT
pre-deploy the VSTO Runtime and send them the `AddUsers.vsto` link instead. Restart
Outlook after installing. Uninstall from **Windows Settings > Apps > Add Users**.

## 7. How auto-update works

With `UpdateEnabled=true` and `UpdateInterval=0`, the VSTO runtime contacts the
`InstallUrl` **every time Outlook loads the add-in**, compares the deployment manifest's
version to the installed one, and downloads and installs a newer version automatically.
Rolling out an update is therefore just:

```powershell
.\build\publish.ps1 -StorageAccountName <name>
```

Users get the new version on their next Outlook start — no reinstall, no notification
emails. If the update server is unreachable, the installed version keeps working.

`MapFileExtensions=true` renames payload files (`.dll`, `.exe`, ...) to `*.deploy` in the
deployment, so a static blob container can serve everything without executable-blocking
proxies or content-sniffing interfering; the runtime strips the suffix during install per
the manifest. That is also why the handful of files that keep their real extension
(`.vsto`, `.manifest`, `setup.exe`) must carry the exact content types the publish script
sets.

## 8. Signing certificate and trust

ClickOnce/VSTO deployments must have signed manifests (`SignManifests` is enabled in the
project). The private key is **not** in the repo: each dev machine needs its own signing
certificate. Easiest is the project's **Properties > Signing** tab > **Create Test
Certificate** (this regenerates `AddUsers_TemporaryKey.pfx` locally and updates the
pinned `ManifestCertificateThumbprint` in the `.csproj`). A test certificate is fine for
development, not for rollout: users will see an "Unknown publisher" warning, and stricter
ClickOnce trust policies can block the install outright. Note that machines only trust
deployments signed by the certificate they imported via `build\trust-cert.ps1`, so
whoever publishes must sign with the certificate matching `build\AddUsersTemporaryKey.cer`
(or redistribute a new `.cer` after changing keys).

For a real rollout, do one of the following:

- **Public code-signing certificate** (OV/EV) from a commercial CA. Best experience:
  the install prompt shows a verified publisher. Select it on the project's
  **Properties > Signing** tab (this updates `ManifestKeyFile` /
  `ManifestCertificateThumbprint` in the `.csproj`). Keep the private key in the build
  machine's certificate store, not in the repo.
- **Internal PKI or self-signed certificate distributed by Group Policy.** Export the
  public certificate (`.cer`) and push it to user machines into:
  - **Trusted Root Certification Authorities** — required if the cert does not chain to a
    trusted root, and
  - **Trusted Publishers** — this is what suppresses the ClickOnce trust prompt entirely
    and allows silent installs/updates.

When you change the signing certificate, publish a new version signed with the new cert
and make sure the new cert is trusted on client machines *before* they pick up that
update.

## 9. Troubleshooting

**Build**

- *vswhere/MSBuild not found, or "Office targets not found"* — install the
  **Office/SharePoint development** workload in Visual Studio Installer.
- *"Cannot find manifest signing certificate in the certificate store"* — see the note in
  section 2.

**Publish**

- *`az` not logged in* — run `az login` (and `az account set -s <subscription>` if you
  have several).
- *Storage account name already taken* — names are global across all of Azure; pick
  another and re-run `setup-azure.ps1`.
- *`PublicAccessNotPermitted`* — either the new account's public-access flag is still
  propagating (the script retries; re-run after a minute) or an Azure Policy in your
  subscription forbids public blob access — see the hosting alternatives in section 4.

**Install / update on client machines**

- *Browser downloads `AddUsers.vsto` instead of installing* — the VSTO Runtime is missing
  (run `setup.exe` first — requires admin rights), or the blob content types were not set (re-run `publish.ps1`,
  which always fixes them).
- *"Application download did not succeed"* — usually wrong content types or a proxy
  mangling downloads; verify the `.vsto` URL opens over HTTPS and re-run `publish.ps1`.
- *Stale or corrupted install* — clear the ClickOnce online cache with
  `rundll32 dfshim.dll CleanOnlineAppCache`, then reinstall from the install URL.
- *Update not arriving* — confirm the publish succeeded (check `build\version.txt` and
  the blob container), then fully restart Outlook; the check runs at add-in load.

**Add-in not visible in Outlook**

- **File > Options > Add-ins > Manage: COM Add-ins > Go** — re-check *Add Users* if it is
  unchecked, and look under *Disabled Items* / slow-start disabling (Outlook disables
  add-ins it blames for slow startup).
- Verify `LoadBehavior` is `3` under
  `HKCU\Software\Microsoft\Office\Outlook\Addins\AddUsers`.

**Sign-in errors (AADSTS codes shown in the error dialog)**

- `700016` — wrong Client ID or Tenant ID in the Configure dialog.
- `65001` — consent not granted; an admin must grant consent (step 1.6).
- `7000218` — **Allow public client flows** is not enabled on the app registration.
- `50011` — the `http://localhost` redirect URI is missing from the app registration.

**SharePoint errors**

- *Authorization errors when adding users* — run `build\debug-permissions.ps1` (PowerShell 7)
  on the affected machine; it signs in with the same app registration, checks the token's
  actual scopes, the group's settings, and SharePoint's own membership-edit verdict, then
  prints the likely cause and writes a shareable JSON report:
  ```powershell
  pwsh -File .\debug-permissions.ps1 -SiteUrl <configured site> -ClientId <app id> `
      -GroupName '<configured group>' -TestEmail <a recipient that failed>
  ```
- *403 / "Access denied" when adding users* — two possibilities, in order of likelihood:
  (1) the app registration's SharePoint delegated permission is not `AllSites.FullControl`
  — lesser scopes cannot perform membership changes at all, regardless of the user's own
  rights (see section 1); (2) the signed-in user cannot edit that group's membership even
  in the browser — they need to be the group's owner, a member when *"Who can edit the
  membership of the group?"* is set to Members, or a site admin. Delegated scopes never
  grant more than the user already has.

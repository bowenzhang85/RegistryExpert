# Registry Expert Installer

Inno Setup script that produces `RegistryExpert-installer-v<version>.exe`
(e.g. `RegistryExpert-installer-v2.3.0.exe`) — a per-user, no-UAC,
silent-default installer used by both end users (interactive install from
GitHub releases) and the in-app auto-updater (silent in-place upgrade).

## Build locally

```powershell
# 1. Build & publish the WPF app
dotnet publish RegistryExpert.Wpf/RegistryExpert.Wpf.csproj `
  -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true -o publish-wpf

# 2. Build the installer (requires Inno Setup 6+ installed)
$env:APP_VERSION = "2.2.1"   # match the version in RegistryExpert.Wpf.csproj
& "C:\Program Files (x86)\Inno Setup 6\iscc.exe" installer\registry-expert.iss

# 3. Output: publish-installer\RegistryExpert-installer-v2.2.1.exe (~74 MB)
#    Uses the app icon (Assets\registry_fixed.ico) so the installer .exe
#    looks identical to RegistryExpert.exe in File Explorer.
```

## Install / upgrade flags

| Command | Behavior |
|---------|----------|
| `RegistryExpert-installer-v2.3.0.exe` | Interactive wizard (minimal — most pages skipped). |
| `RegistryExpert-installer-v2.3.0.exe /VERYSILENT /SUPPRESSMSGBOXES /CLOSEAPPLICATIONS /NORESTART /fromversion=2.2.1` | Fully silent in-place upgrade with relaunch. Used by the in-app auto-updater. |

## What gets installed

- Install root: `%LOCALAPPDATA%\Programs\RegistryExpert\RegistryExpert.exe`
- Start Menu shortcut: `Start ▸ Registry Expert`
- Optional desktop shortcut (unchecked by default in the wizard)
- Apps & Features entry for clean uninstall (Publisher: Microsoft Corporation)
- No HKLM writes — fully per-user, no admin / UAC required

## Asset naming convention

The installer is shipped on GitHub releases as
`RegistryExpert-installer-v<version>.exe` (e.g.
`RegistryExpert-installer-v2.3.0.exe`). The in-app `UpdateChecker` matches
any asset whose name starts with `RegistryExpert-installer-v` and ends with
`.exe` (case-insensitive) — see `IsInstallerAssetName` in `UpdateChecker.cs`.

## Auto-update flow

After install, the running exe lives in a per-user writable folder, so the
existing `AutoUpdater` can either:

- (Preferred) download the installer from a newer release and re-run it
  silently with `/CLOSEAPPLICATIONS` to upgrade in place.
- (Legacy) download `RegistryExpert.exe` (portable) and swap it in via a
  small batch script (used for releases that don't ship the installer, or
  legacy v2.2.1 clients that pre-date the installer feature).

The choice is driven by `UpdateInfo.DownloadKind` from `UpdateChecker`.

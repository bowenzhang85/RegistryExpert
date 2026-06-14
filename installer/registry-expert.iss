; Registry Expert — Inno Setup installer script
;
; Builds RegistryExpert-Setup.exe — a single-file, per-user, no-UAC, silent-default
; installer that can be invoked either interactively (full wizard) or completely
; silently (used by the in-app auto-update flow).
;
; Build locally:
;   1. dotnet publish RegistryExpert.Wpf/RegistryExpert.Wpf.csproj -c Release -r win-x64 ^
;        --self-contained true -p:PublishSingleFile=true -o publish-wpf
;   2. iscc installer\registry-expert.iss
;   Output: publish-installer\RegistryExpert-Setup.exe
;
; Silent in-place upgrade invocation (used by the in-app updater):
;   RegistryExpert-Setup.exe /VERYSILENT /SUPPRESSMSGBOXES /CLOSEAPPLICATIONS /NORESTART /fromversion=2.2.1
;
; CI passes APP_VERSION via environment so AppVersion stays in sync with the .csproj.

#define MyAppName "Registry Expert"
#define MyAppPublisher "Microsoft Corporation"
#define MyAppURL "https://github.com/bowenzhang85/RegistryExpert"
#define MyAppExeName "RegistryExpert.exe"
#define MyAppVersion GetEnv("APP_VERSION")
#if MyAppVersion == ""
  #define MyAppVersion "0.0.0"
#endif

[Setup]
; Stable, unique AppId — never change this across versions; Inno Setup uses it
; to detect existing installations for in-place upgrades and uninstall lookup.
AppId={{B7C9D2E4-4A41-4F61-8B7C-9D2E44A41F61}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
VersionInfoVersion={#MyAppVersion}

; Per-user install: no admin/UAC required. Installs into %LOCALAPPDATA%\Programs\.
DefaultDirName={localappdata}\Programs\RegistryExpert
DefaultGroupName={#MyAppName}
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=

; Streamlined wizard — most pages disabled so a typical install is 1-2 clicks.
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=yes
DisableFinishedPage=no
ShowLanguageDialog=no

; CloseApplications=yes makes /CLOSEAPPLICATIONS the default behavior, so the
; in-app updater can run the installer silently without flags.
CloseApplications=yes
CloseApplicationsFilter=*.exe
RestartApplications=no

; Refresh Explorer's file-association cache (SHChangeNotify) after install/uninstall
; so the hive associations and the right-click verb appear immediately.
ChangesAssociations=yes

; Modern wizard style + branded uninstall icon
WizardStyle=modern
SetupIconFile=..\Assets\registry_fixed.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

; Output: publish-installer\RegistryExpert-installer-vX.Y.Z.exe (versioned filename)
OutputDir=..\publish-installer
OutputBaseFilename=RegistryExpert-installer-v{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; Minimum Windows 10 (1809 / build 17763) so all features used by the app work
MinVersion=10.0.17763

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "..\publish-wpf\RegistryExpert.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Registry]
; ── Right-click "Open with Registry Expert" verb for ALL hive filenames ──
; The AppliesTo filter limits the menu item to known hive names so it does NOT
; pollute every file's context menu. This list MUST stay in sync with
; HiveBundleScanner.KnownHiveNames + KnownHivePrefixes, and the key/command
; format MUST match ShellIntegrationService (the app's stale-cleanup pass parses
; this exact key via GetRegisteredExePath).
; NOTE: on Windows 11 this verb appears under "Show more options" (legacy menu).
Root: HKCU; Subkey: "Software\Classes\*\shell\OpenWithRegistryExpert"; ValueType: string; ValueData: "Open with Registry Expert"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\*\shell\OpenWithRegistryExpert"; ValueType: string; ValueName: "Icon"; ValueData: "{app}\{#MyAppExeName},0"
Root: HKCU; Subkey: "Software\Classes\*\shell\OpenWithRegistryExpert"; ValueType: string; ValueName: "AppliesTo"; ValueData: "System.FileName:=""SYSTEM"" OR System.FileName:=""SOFTWARE"" OR System.FileName:=""SAM"" OR System.FileName:=""SECURITY"" OR System.FileName:=""DEFAULT"" OR System.FileName:=""BCD"" OR System.FileName:=""COMPONENTS"" OR System.FileName:~<""NTUSER"" OR System.FileName:~<""USRCLASS"" OR System.FileName:~<""AMCACHE"""
Root: HKCU; Subkey: "Software\Classes\*\shell\OpenWithRegistryExpert\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" ""%1"""

; ── ProgID for true double-click default on .hiv / .hve ──
Root: HKCU; Subkey: "Software\Classes\RegistryExpert.HiveFile"; ValueType: string; ValueData: "Registry Hive File"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\RegistryExpert.HiveFile\DefaultIcon"; ValueType: string; ValueData: "{app}\{#MyAppExeName},0"
Root: HKCU; Subkey: "Software\Classes\RegistryExpert.HiveFile\shell\open\command"; ValueType: string; ValueData: """{app}\{#MyAppExeName}"" ""%1"""

; ── Associate .hiv and .hve with the ProgID (default action + Open-with list) ──
Root: HKCU; Subkey: "Software\Classes\.hiv"; ValueType: string; ValueData: "RegistryExpert.HiveFile"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\.hiv\OpenWithProgIds"; ValueType: string; ValueName: "RegistryExpert.HiveFile"; ValueData: ""; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\.hve"; ValueType: string; ValueData: "RegistryExpert.HiveFile"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Classes\.hve\OpenWithProgIds"; ValueType: string; ValueName: "RegistryExpert.HiveFile"; ValueData: ""; Flags: uninsdeletevalue

[Run]
; Always relaunch the just-installed exe at end of install (works in both
; interactive and silent modes). Passes --just-updated <prev> so the post-update
; banner shows after auto-update. /fromversion= is read via {param:...}.
;
; Flags:
;   nowait                   — don't block the installer waiting for the app
;   postinstall              — run as the post-install action
;   skipifsilent             — actually we DO want this to run in silent mode too,
;                              so we omit skipifsilent and use it unconditionally
Filename: "{app}\{#MyAppExeName}"; Parameters: "--just-updated {param:fromversion|portable}"; Flags: nowait postinstall

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

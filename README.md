# Registry Expert - Offline Registry Viewer

A lightweight Windows application for viewing and analyzing offline Windows registry hive files.

![.NET](https://img.shields.io/badge/.NET-8.0-blue) ![Windows](https://img.shields.io/badge/platform-Windows%20x64-lightgrey)

## Download

**[Download Latest Release](../../releases/latest)**

No installation required - just download `RegistryExpert.exe` and run it.

## Features

- **Open Offline Registry Hives**: Load and browse SAM, SECURITY, SOFTWARE, SYSTEM, NTUSER.DAT, USRCLASS.DAT, Amcache.hve, and other hive files
- **Tree-based Navigation**: Browse registry keys in a familiar tree structure
- **Value Viewer**: View registry values with type information and hex dump
- **Search Functionality**: Search for keys, values, and data across the entire hive
- **Information Extraction**:
  - System information (OS version, computer name, timezone, hardware)
  - User information (local accounts, profiles, recent documents, typed paths)
  - Software information (installed programs, startup items, services)
  - Network information (interfaces, profiles, wireless networks, shares)
  - USB device history
  - MRU lists
- **Export**: Export keys and values to text files
- **Drag & Drop**: Simply drag a hive file onto the application to open it

## System Requirements

- Windows 10/11 (64-bit)
- No additional runtime required (self-contained executable)

## Building from Source

### Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

### Build Steps

1. Open a command prompt in the project directory
2. Run the build script:
   ```
   build.bat
   ```
3. The executable will be created in the `.\publish` folder

### Manual Build

```bash
# Restore packages
dotnet restore

# Build release
dotnet build -c Release

# Publish self-contained executable
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=true -o .\publish
```

## Usage

### Opening a Hive File

1. Launch `RegistryExpert.exe`
2. Either:
   - Use **File → Open Hive** (Ctrl+O)
   - Drag and drop a hive file onto the application window

### Common Registry Hive Locations

| Hive | Location |
|------|----------|
| SAM | `C:\Windows\System32\config\SAM` |
| SECURITY | `C:\Windows\System32\config\SECURITY` |
| SOFTWARE | `C:\Windows\System32\config\SOFTWARE` |
| SYSTEM | `C:\Windows\System32\config\SYSTEM` |
| DEFAULT | `C:\Windows\System32\config\DEFAULT` |
| NTUSER.DAT | `C:\Users\<username>\NTUSER.DAT` |
| USRCLASS.DAT | `C:\Users\<username>\AppData\Local\Microsoft\Windows\USRCLASS.DAT` |
| Amcache.hve | `C:\Windows\AppCompat\Programs\Amcache.hve` |

> **Note**: Live system registry files are locked. To analyze them, either:
> - Boot from a different OS/recovery environment
> - Use volume shadow copies
> - Use forensic imaging tools

### Extracting Information

Use the **Tools** menu to extract various information:

- **Extract System Info**: OS version, computer name, timezone, hardware details
- **Extract User Info**: User accounts, profiles, recent activity
- **Extract Software Info**: Installed programs, startup items, services
- **Extract Network Info**: Network interfaces, WiFi profiles, shares
- **Full Analysis Report** (F5): Comprehensive report combining all categories

### Searching

1. Press **Ctrl+F** or use **Tools → Search**
2. Enter your search term
3. Double-click a result to navigate to that key

## Supported Hive Types

| Type | Description |
|------|-------------|
| SAM | Security Account Manager - User accounts |
| SECURITY | Security policies and credentials |
| SOFTWARE | Installed software and system settings |
| SYSTEM | Hardware configuration and services |
| NTUSER | Per-user settings and preferences |
| USRCLASS | Per-user file associations |
| AMCACHE | Application execution history |
| DEFAULT | Default user profile template |
| BCD | Boot Configuration Data |



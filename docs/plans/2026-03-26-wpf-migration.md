# WPF Migration Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Migrate RegistryExpert from Windows Forms to WPF with MVVM to permanently solve DPI scaling issues across multi-monitor setups.

**Architecture:** Incremental side-by-side migration. Extract shared business logic into a `RegistryExpert.Core` class library, then build a new `RegistryExpert.Wpf` application consuming it. The existing WinForms app continues to work during migration. MVVM pattern with no framework dependency (raw `INotifyPropertyChanged` + `ICommand`).

**Tech Stack:** .NET 8.0, WPF, XAML, C#, System.Text.Json, vendored Lib/Registry parser

---

## Phase 0: Foundation

Phase 0 establishes the solution structure, extracts the shared Core library, and builds the WPF project shell with theme support. At the end of Phase 0, the WPF app launches with an empty themed window and the WinForms app still builds and runs unchanged.

---

### Task 1: Create Solution Structure

**Files:**
- Create: `RegistryExpert.Core/RegistryExpert.Core.csproj`
- Create: `RegistryExpert.Wpf/RegistryExpert.Wpf.csproj`
- Modify: `RegistryExpert.sln` (add two new projects)
- Rename: `RegistryExpert.csproj` to stay as-is (existing WinForms project)

**Step 1: Create the Core class library project**

```bash
dotnet new classlib -n RegistryExpert.Core -f net8.0-windows -o RegistryExpert.Core
```

The `-f net8.0-windows` is needed because the vendored Lib/Registry uses Windows-specific APIs.

**Step 2: Delete the auto-generated Class1.cs**

```bash
del RegistryExpert.Core\Class1.cs
```

**Step 3: Create the WPF application project**

```bash
dotnet new wpf -n RegistryExpert.Wpf -f net8.0-windows -o RegistryExpert.Wpf
```

**Step 4: Add both projects to the solution**

```bash
dotnet sln RegistryExpert.sln add RegistryExpert.Core\RegistryExpert.Core.csproj
dotnet sln RegistryExpert.sln add RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
```

**Step 5: Add project references**

```bash
dotnet add RegistryExpert.Wpf\RegistryExpert.Wpf.csproj reference RegistryExpert.Core\RegistryExpert.Core.csproj
```

**Step 6: Verify solution builds**

```bash
dotnet build RegistryExpert.sln
```

Expected: 3 projects build successfully (0 errors).

**Step 7: Commit**

```bash
git add -A && git commit -m "feat: create solution structure for WPF migration (Core + Wpf projects)"
```

---

### Task 2: Configure Core Library Project

The Core library will contain all business logic, data models, and the vendored Lib/Registry parser. It must NOT reference any UI framework (no WinForms, no WPF).

**Files:**
- Modify: `RegistryExpert.Core/RegistryExpert.Core.csproj`

**Step 1: Edit the Core csproj**

Replace the entire content of `RegistryExpert.Core/RegistryExpert.Core.csproj` with:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net8.0-windows</TargetFramework>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <RootNamespace>RegistryExpert.Core</RootNamespace>
    <AssemblyName>RegistryExpert.Core</AssemblyName>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="System.Text.Encoding.CodePages" Version="9.0.0" />
  </ItemGroup>

</Project>
```

Key decisions:
- `net8.0-windows` because Lib/Registry uses some Windows-specific APIs
- No `<UseWindowsForms>` or `<UseWPF>` -- this library is UI-agnostic
- `System.Text.Encoding.CodePages` is needed by the registry parser

**Step 2: Verify the project builds (it will be empty at this point)**

```bash
dotnet build RegistryExpert.Core\RegistryExpert.Core.csproj
```

Expected: Build succeeded, 0 errors.

**Step 3: Commit**

```bash
git add -A && git commit -m "feat: configure Core library csproj with correct framework and dependencies"
```

---

### Task 3: Move Vendored Lib/Registry to Core

The vendored `Lib/Registry` parser library is the foundation that both UI projects depend on.

**Files:**
- Move: `Lib/Registry/**` -> `RegistryExpert.Core/Lib/Registry/**`
- Modify: `RegistryExpert.csproj` (update path or add project reference)
- Modify: `RegistryExpert.Core/RegistryExpert.Core.csproj` (include Lib/Registry)

**Step 1: Copy the vendored library to Core**

```bash
xcopy /E /I Lib\Registry RegistryExpert.Core\Lib\Registry
```

**Step 2: Verify files were copied**

```bash
dir RegistryExpert.Core\Lib\Registry
```

Expected: All directories (Abstractions/, Cells/, Lists/, Other/) and .cs files present.

**Step 3: Build the Core project to verify Lib/Registry compiles**

```bash
dotnet build RegistryExpert.Core\RegistryExpert.Core.csproj
```

Expected: Build succeeded. If there are namespace issues, they should resolve because the Lib/Registry files use their own namespace (`RegistryParser`).

**Step 4: Verify the existing WinForms project still builds**

The WinForms project still has its own copy of Lib/Registry. Both projects will have their own copy during the migration period. This is intentional -- the WinForms project should remain completely untouched.

```bash
dotnet build RegistryExpert.csproj
```

Expected: Build succeeded, 0 errors.

**Step 5: Commit**

```bash
git add -A && git commit -m "feat: copy vendored Lib/Registry parser into Core library"
```

---

### Task 4: Move Services to Core

Move all UI-agnostic service classes to the Core library.

**Files:**
- Copy: `Services/OfflineRegistryParser.cs` -> `RegistryExpert.Core/Services/OfflineRegistryParser.cs`
- Copy: `Services/RegistryInfoExtractor.cs` -> `RegistryExpert.Core/Services/RegistryInfoExtractor.cs`
- Copy: `Services/TransactionLogAnalyzer.cs` -> `RegistryExpert.Core/Services/TransactionLogAnalyzer.cs`
- Copy: `Services/UpdateChecker.cs` -> `RegistryExpert.Core/Services/UpdateChecker.cs`
- Copy: `Services/AppSettings.cs` -> `RegistryExpert.Core/Services/AppSettings.cs`

**Step 1: Create the Services directory in Core**

```bash
mkdir RegistryExpert.Core\Services
```

**Step 2: Copy all service files**

```bash
copy Services\OfflineRegistryParser.cs RegistryExpert.Core\Services\
copy Services\RegistryInfoExtractor.cs RegistryExpert.Core\Services\
copy Services\TransactionLogAnalyzer.cs RegistryExpert.Core\Services\
copy Services\UpdateChecker.cs RegistryExpert.Core\Services\
copy Services\AppSettings.cs RegistryExpert.Core\Services\
```

**Step 3: Update namespaces in Core copies**

All copied files currently use `namespace RegistryExpert`. Update them to `namespace RegistryExpert.Core` in every file.

For each file in `RegistryExpert.Core/Services/`:
- Change `namespace RegistryExpert` to `namespace RegistryExpert.Core`
- Keep all `using` statements as-is

**Important:** The `RegistryInfoExtractor.cs` file uses `System.Security.Cryptography.X509Certificates` -- verify this is available in the Core library without WinForms.

**Step 4: Build Core**

```bash
dotnet build RegistryExpert.Core\RegistryExpert.Core.csproj
```

Expected: Build succeeded. If there are errors about missing types, they're likely from WinForms-specific types that leaked into service code (there shouldn't be any based on our analysis).

**Step 5: Verify WinForms project still builds unchanged**

```bash
dotnet build RegistryExpert.csproj
```

Expected: Build succeeded, 0 errors (WinForms project is completely unchanged).

**Step 6: Commit**

```bash
git add -A && git commit -m "feat: copy service layer into Core library with updated namespaces"
```

---

### Task 5: Extract Data Models to Core

The data model classes are currently embedded at the bottom of `MainForm.cs` (lines 10178-10337). Copy them to standalone files in Core.

**Files:**
- Create: `RegistryExpert.Core/Models/AnalysisSection.cs`
- Create: `RegistryExpert.Core/Models/AnalysisItem.cs`
- Create: `RegistryExpert.Core/Models/NetworkAdapterItem.cs`
- Create: `RegistryExpert.Core/Models/NetworkPropertyItem.cs`
- Create: `RegistryExpert.Core/Models/DeviceClassItem.cs`
- Create: `RegistryExpert.Core/Models/DeviceItem.cs`
- Create: `RegistryExpert.Core/Models/DevicePropertyItem.cs`
- Create: `RegistryExpert.Core/Models/RoleFeatureItem.cs`
- Create: `RegistryExpert.Core/Models/MountedDeviceEntry.cs`
- Create: `RegistryExpert.Core/Models/PhysicalDiskEntry.cs`

**Step 1: Create the Models directory**

```bash
mkdir RegistryExpert.Core\Models
```

**Step 2: Create each model file**

Copy each class from `MainForm.cs` lines 10178-10337 into its own file under `RegistryExpert.Core/Models/`, using namespace `RegistryExpert.Core.Models`.

Example for `AnalysisSection.cs`:

```csharp
namespace RegistryExpert.Core.Models
{
    public class AnalysisSection
    {
        public string Title { get; set; } = "";
        public List<AnalysisItem> Items { get; set; } = new();
        public object? Tag { get; set; }
    }
}
```

Example for `AnalysisItem.cs`:

```csharp
namespace RegistryExpert.Core.Models
{
    public class AnalysisItem
    {
        public string Name { get; set; } = "";
        public string Value { get; set; } = "";
        public string RegistryPath { get; set; } = "";
        public string RegistryValue { get; set; } = "";
        public bool IsSubSection { get; set; } = false;
        public bool IsWarning { get; set; } = false;
        public List<AnalysisItem>? SubItems { get; set; }
    }
}
```

Do the same for ALL model classes: `NetworkAdapterItem`, `NetworkPropertyItem`, `DeviceClassItem`, `DeviceItem`, `DevicePropertyItem`, `RoleFeatureItem`, `MountedDeviceEntry`, `PhysicalDiskEntry`.

Also extract `SearchMatch`, `HiveStatistics`, `UpdateInfo`, `TransactionLogDiff`, `ValueChange`, and `TransactionLogChangeType` from their respective service files into Core/Models if they don't already live there.

**Step 3: Update Core service files to use the new model namespace**

Add `using RegistryExpert.Core.Models;` to any Core service file that references these model classes.

**Step 4: Build Core**

```bash
dotnet build RegistryExpert.Core\RegistryExpert.Core.csproj
```

Expected: Build succeeded, 0 errors.

**Step 5: Verify WinForms still builds (unchanged)**

```bash
dotnet build RegistryExpert.csproj
```

Expected: Build succeeded (WinForms project still uses its own inline copies).

**Step 6: Commit**

```bash
git add -A && git commit -m "feat: extract data models into Core library"
```

---

### Task 6: Configure WPF Project

Set up the WPF project with proper configuration, references, and assets.

**Files:**
- Modify: `RegistryExpert.Wpf/RegistryExpert.Wpf.csproj`
- Copy: `Assets/**` -> `RegistryExpert.Wpf/Assets/**`

**Step 1: Edit the WPF csproj**

Replace the entire content of `RegistryExpert.Wpf/RegistryExpert.Wpf.csproj` with:

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>WinExe</OutputType>
    <TargetFramework>net8.0-windows</TargetFramework>
    <UseWPF>true</UseWPF>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
    <RootNamespace>RegistryExpert.Wpf</RootNamespace>
    <AssemblyName>RegistryExpert</AssemblyName>
    <Version>1.4.0</Version>
    <Company>RegistryExpert</Company>
    <Product>Registry Expert - Offline Registry Viewer</Product>
    <Description>A lightweight tool for viewing and analyzing offline Windows registry hive files</Description>
    <ApplicationIcon>..\Assets\registry_fixed.ico</ApplicationIcon>
    <PublishSingleFile>true</PublishSingleFile>
    <SelfContained>true</SelfContained>
    <IncludeNativeLibrariesForSelfExtract>true</IncludeNativeLibrariesForSelfExtract>
    <EnableCompressionInSingleFile>true</EnableCompressionInSingleFile>
    <IncludeAllContentForSelfExtract>true</IncludeAllContentForSelfExtract>
  </PropertyGroup>

  <ItemGroup>
    <ProjectReference Include="..\RegistryExpert.Core\RegistryExpert.Core.csproj" />
  </ItemGroup>

  <!-- Asset images as WPF Resources -->
  <ItemGroup>
    <Resource Include="..\Assets\registry.png" Link="Assets\registry.png" />
    <Resource Include="..\Assets\registry_fixed.ico" Link="Assets\registry_fixed.ico" />
    <Resource Include="..\Assets\registry_icon_system.png" Link="Assets\registry_icon_system.png" />
    <Resource Include="..\Assets\registry_icon_profiles.png" Link="Assets\registry_icon_profiles.png" />
    <Resource Include="..\Assets\registry_icon_services.png" Link="Assets\registry_icon_services.png" />
    <Resource Include="..\Assets\registry_icon_software.png" Link="Assets\registry_icon_software.png" />
    <Resource Include="..\Assets\registry_icon_storage.png" Link="Assets\registry_icon_storage.png" />
    <Resource Include="..\Assets\registry_icon_network.png" Link="Assets\registry_icon_network.png" />
    <Resource Include="..\Assets\registry_icon_rdp.png" Link="Assets\registry_icon_rdp.png" />
    <Resource Include="..\Assets\registry_icon_update.png" Link="Assets\registry_icon_update.png" />
    <Resource Include="..\Assets\registry_icon_open.png" Link="Assets\registry_icon_open.png" />
    <Resource Include="..\Assets\registry_icon_search.png" Link="Assets\registry_icon_search.png" />
    <Resource Include="..\Assets\registry_icon_analyze.png" Link="Assets\registry_icon_analyze.png" />
    <Resource Include="..\Assets\registry_icon_statistics.png" Link="Assets\registry_icon_statistics.png" />
    <Resource Include="..\Assets\registry_icon_compare.png" Link="Assets\registry_icon_compare.png" />
    <Resource Include="..\Assets\registry_icon_timeline.png" Link="Assets\registry_icon_timeline.png" />
    <Resource Include="..\Assets\registry_icon_load-hive.png" Link="Assets\registry_icon_load-hive.png" />
    <Resource Include="..\Assets\registry_icon_health.png" Link="Assets\registry_icon_health.png" />
    <Resource Include="..\Assets\docs_images_reg_bin.png" Link="Assets\reg_bin.png" />
    <Resource Include="..\docs\images\reg_bin.png" Link="Assets\reg_bin.png" />
    <Resource Include="..\docs\images\reg_num.png" Link="Assets\reg_num.png" />
    <Resource Include="..\docs\images\reg_str.png" Link="Assets\reg_str.png" />
  </ItemGroup>

</Project>
```

Key differences from WinForms csproj:
- `<UseWPF>true</UseWPF>` instead of `<UseWindowsForms>true</UseWindowsForms>`
- NO `<ApplicationHighDpiMode>` -- WPF handles DPI natively
- Assets use `<Resource>` build action (WPF's resource system) instead of `<EmbeddedResource>`
- Project reference to Core library

**Step 2: Build**

```bash
dotnet build RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
```

Expected: Build succeeded (default WPF template + our config).

**Step 3: Commit**

```bash
git add -A && git commit -m "feat: configure WPF project with assets and Core reference"
```

---

### Task 7: Create MVVM Infrastructure

Create the base classes for MVVM (ViewModelBase, RelayCommand). No external framework needed.

**Files:**
- Create: `RegistryExpert.Wpf/ViewModels/ViewModelBase.cs`
- Create: `RegistryExpert.Wpf/ViewModels/RelayCommand.cs`

**Step 1: Create ViewModels directory**

```bash
mkdir RegistryExpert.Wpf\ViewModels
```

**Step 2: Create ViewModelBase.cs**

```csharp
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// Base class for all ViewModels providing INotifyPropertyChanged implementation.
    /// </summary>
    public abstract class ViewModelBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}
```

**Step 3: Create RelayCommand.cs**

```csharp
using System.Windows.Input;

namespace RegistryExpert.Wpf.ViewModels
{
    /// <summary>
    /// Basic ICommand implementation for MVVM command binding.
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Func<object?, bool>? _canExecute;

        public RelayCommand(Action<object?> execute, Func<object?, bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public RelayCommand(Action execute, Func<bool>? canExecute = null)
            : this(_ => execute(), canExecute != null ? _ => canExecute() : null)
        {
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter) => _canExecute?.Invoke(parameter) ?? true;

        public void Execute(object? parameter) => _execute(parameter);

        public void RaiseCanExecuteChanged() => CommandManager.InvalidateRequerySuggested();
    }

    /// <summary>
    /// Async variant of RelayCommand for async operations.
    /// </summary>
    public class AsyncRelayCommand : ICommand
    {
        private readonly Func<object?, Task> _execute;
        private readonly Func<object?, bool>? _canExecute;
        private bool _isExecuting;

        public AsyncRelayCommand(Func<object?, Task> execute, Func<object?, bool>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public AsyncRelayCommand(Func<Task> execute, Func<bool>? canExecute = null)
            : this(_ => execute(), canExecute != null ? _ => canExecute() : null)
        {
        }

        public event EventHandler? CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }

        public bool CanExecute(object? parameter) => !_isExecuting && (_canExecute?.Invoke(parameter) ?? true);

        public async void Execute(object? parameter)
        {
            if (_isExecuting) return;

            _isExecuting = true;
            CommandManager.InvalidateRequerySuggested();

            try
            {
                await _execute(parameter);
            }
            finally
            {
                _isExecuting = false;
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }
}
```

**Step 4: Build**

```bash
dotnet build RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
```

Expected: Build succeeded, 0 errors.

**Step 5: Commit**

```bash
git add -A && git commit -m "feat: add MVVM infrastructure (ViewModelBase, RelayCommand)"
```

---

### Task 8: Create Theme System (XAML ResourceDictionaries)

Port the ModernTheme color palette to WPF ResourceDictionaries. This replaces `ModernTheme.cs` (1,132 lines) with clean XAML.

**Files:**
- Create: `RegistryExpert.Wpf/Themes/DarkTheme.xaml`
- Create: `RegistryExpert.Wpf/Themes/LightTheme.xaml`
- Create: `RegistryExpert.Wpf/Themes/SharedStyles.xaml`
- Modify: `RegistryExpert.Wpf/App.xaml`

**Step 1: Create Themes directory**

```bash
mkdir RegistryExpert.Wpf\Themes
```

**Step 2: Create DarkTheme.xaml**

Port all dark theme colors from `ModernTheme.cs` (lines 91-241). Each color becomes a `SolidColorBrush` resource with a consistent naming convention.

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

    <!-- Background / Surface -->
    <SolidColorBrush x:Key="BackgroundBrush" Color="#FF1E1E22" />
    <SolidColorBrush x:Key="SurfaceBrush" Color="#FF27272C" />
    <SolidColorBrush x:Key="SurfaceLightBrush" Color="#FF323238" />
    <SolidColorBrush x:Key="SurfaceHoverBrush" Color="#FF3C3C44" />
    <SolidColorBrush x:Key="BorderBrush" Color="#FF37373E" />

    <!-- Text -->
    <SolidColorBrush x:Key="TextPrimaryBrush" Color="#FFE6E6EB" />
    <SolidColorBrush x:Key="TextSecondaryBrush" Color="#FFA0A0AA" />
    <SolidColorBrush x:Key="TextDisabledBrush" Color="#FF464650" />

    <!-- Accent -->
    <SolidColorBrush x:Key="AccentBrush" Color="#FF38A1DF" />
    <SolidColorBrush x:Key="AccentHoverBrush" Color="#FF3CAFEB" />
    <SolidColorBrush x:Key="AccentDarkBrush" Color="#FF1E82BE" />

    <!-- Status -->
    <SolidColorBrush x:Key="SuccessBrush" Color="#FF10B981" />
    <SolidColorBrush x:Key="WarningBrush" Color="#FFF59E0B" />
    <SolidColorBrush x:Key="ErrorBrush" Color="#FFEF4444" />
    <SolidColorBrush x:Key="InfoBrush" Color="#FF6366F1" />

    <!-- Block/Error Row -->
    <SolidColorBrush x:Key="BlockRowBackgroundBrush" Color="#FF3C2828" />
    <SolidColorBrush x:Key="BlockRowForegroundBrush" Color="#FFFFB4B4" />

    <!-- Diff -->
    <SolidColorBrush x:Key="DiffAddedBrush" Color="#FF2EA043" />
    <SolidColorBrush x:Key="DiffRemovedBrush" Color="#FFFF6E6E" />

    <!-- Health -->
    <SolidColorBrush x:Key="HealthyTextBrush" Color="#FF4CAF50" />
    <SolidColorBrush x:Key="WarningTextBrush" Color="#FFFF6464" />

    <!-- TreeView / ListView -->
    <SolidColorBrush x:Key="TreeViewBackBrush" Color="#FF1E1E22" />
    <SolidColorBrush x:Key="ListViewBackBrush" Color="#FF1E1E22" />
    <SolidColorBrush x:Key="ListViewAltRowBrush" Color="#FF24242A" />
    <SolidColorBrush x:Key="SelectionBrush" Color="#FF2D3746" />
    <SolidColorBrush x:Key="SelectionActiveBrush" Color="#FF2D9CDB" />

    <!-- Headers -->
    <SolidColorBrush x:Key="GradientStartBrush" Color="#FF27272C" />
    <SolidColorBrush x:Key="GradientEndBrush" Color="#FF27272C" />

    <!-- Window Chrome -->
    <Color x:Key="WindowBackgroundColor">#FF1E1E22</Color>

</ResourceDictionary>
```

**Step 3: Create LightTheme.xaml**

Same keys, light theme colors:

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

    <!-- Background / Surface -->
    <SolidColorBrush x:Key="BackgroundBrush" Color="#FFF9F9FB" />
    <SolidColorBrush x:Key="SurfaceBrush" Color="#FFFFFFFF" />
    <SolidColorBrush x:Key="SurfaceLightBrush" Color="#FFF5F5F8" />
    <SolidColorBrush x:Key="SurfaceHoverBrush" Color="#FFEBEBF0" />
    <SolidColorBrush x:Key="BorderBrush" Color="#FFE1E1E6" />

    <!-- Text -->
    <SolidColorBrush x:Key="TextPrimaryBrush" Color="#FF1C1C23" />
    <SolidColorBrush x:Key="TextSecondaryBrush" Color="#FF5A5A64" />
    <SolidColorBrush x:Key="TextDisabledBrush" Color="#FFAAAAB4" />

    <!-- Accent (same for both themes) -->
    <SolidColorBrush x:Key="AccentBrush" Color="#FF38A1DF" />
    <SolidColorBrush x:Key="AccentHoverBrush" Color="#FF3CAFEB" />
    <SolidColorBrush x:Key="AccentDarkBrush" Color="#FF1E82BE" />

    <!-- Status -->
    <SolidColorBrush x:Key="SuccessBrush" Color="#FF10B981" />
    <SolidColorBrush x:Key="WarningBrush" Color="#FFF59E0B" />
    <SolidColorBrush x:Key="ErrorBrush" Color="#FFEF4444" />
    <SolidColorBrush x:Key="InfoBrush" Color="#FF6366F1" />

    <!-- Block/Error Row -->
    <SolidColorBrush x:Key="BlockRowBackgroundBrush" Color="#FFFFEBEB" />
    <SolidColorBrush x:Key="BlockRowForegroundBrush" Color="#FFB42828" />

    <!-- Diff -->
    <SolidColorBrush x:Key="DiffAddedBrush" Color="#FF238636" />
    <SolidColorBrush x:Key="DiffRemovedBrush" Color="#FFCF222E" />

    <!-- Health -->
    <SolidColorBrush x:Key="HealthyTextBrush" Color="#FF2E7D32" />
    <SolidColorBrush x:Key="WarningTextBrush" Color="#FFFF0000" />

    <!-- TreeView / ListView -->
    <SolidColorBrush x:Key="TreeViewBackBrush" Color="#FFFCFCFE" />
    <SolidColorBrush x:Key="ListViewBackBrush" Color="#FFFCFCFE" />
    <SolidColorBrush x:Key="ListViewAltRowBrush" Color="#FFF7F7FA" />
    <SolidColorBrush x:Key="SelectionBrush" Color="#FFE6F2FF" />
    <SolidColorBrush x:Key="SelectionActiveBrush" Color="#FF2D9CDB" />

    <!-- Headers -->
    <SolidColorBrush x:Key="GradientStartBrush" Color="#FFFAFAFC" />
    <SolidColorBrush x:Key="GradientEndBrush" Color="#FFFAFAFC" />

    <!-- Window Chrome -->
    <Color x:Key="WindowBackgroundColor">#FFF9F9FB</Color>

</ResourceDictionary>
```

**Step 4: Create SharedStyles.xaml**

Shared styles that reference DynamicResource theme brushes. These work with either theme.

```xml
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">

    <!-- Fonts (WPF handles point-to-device scaling automatically) -->
    <FontFamily x:Key="DefaultFontFamily">Segoe UI</FontFamily>
    <FontFamily x:Key="MonoFontFamily">Consolas</FontFamily>

    <!-- Base Window Style -->
    <Style x:Key="ModernWindowStyle" TargetType="Window">
        <Setter Property="Background" Value="{DynamicResource BackgroundBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="FontFamily" Value="{StaticResource DefaultFontFamily}" />
        <Setter Property="FontSize" Value="12" />
    </Style>

    <!-- TreeView Styles -->
    <Style TargetType="TreeView">
        <Setter Property="Background" Value="{DynamicResource TreeViewBackBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="BorderThickness" Value="0" />
        <Setter Property="FontFamily" Value="{StaticResource DefaultFontFamily}" />
        <Setter Property="FontSize" Value="12" />
        <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True" />
        <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling" />
    </Style>

    <Style TargetType="TreeViewItem">
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="Padding" Value="4,2" />
        <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="{DynamicResource SelectionActiveBrush}" />
                <Setter Property="Foreground" Value="White" />
            </Trigger>
            <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource SurfaceHoverBrush}" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <!-- DataGrid Styles -->
    <Style TargetType="DataGrid">
        <Setter Property="Background" Value="{DynamicResource SurfaceBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="BorderThickness" Value="0" />
        <Setter Property="GridLinesVisibility" Value="Horizontal" />
        <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource BorderBrush}" />
        <Setter Property="RowHeaderWidth" Value="0" />
        <Setter Property="AutoGenerateColumns" Value="False" />
        <Setter Property="IsReadOnly" Value="True" />
        <Setter Property="SelectionMode" Value="Single" />
        <Setter Property="CanUserAddRows" Value="False" />
        <Setter Property="CanUserDeleteRows" Value="False" />
        <Setter Property="FontFamily" Value="{StaticResource DefaultFontFamily}" />
        <Setter Property="FontSize" Value="13.33" />
        <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True" />
        <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling" />
        <Setter Property="EnableRowVirtualization" Value="True" />
        <Setter Property="EnableColumnVirtualization" Value="True" />
    </Style>

    <Style TargetType="DataGridColumnHeader">
        <Setter Property="Background" Value="{DynamicResource TreeViewBackBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextSecondaryBrush}" />
        <Setter Property="FontWeight" Value="SemiBold" />
        <Setter Property="Padding" Value="8,4" />
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}" />
        <Setter Property="BorderThickness" Value="0,0,0,1" />
    </Style>

    <Style TargetType="DataGridRow">
        <Setter Property="Background" Value="{DynamicResource SurfaceBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="{DynamicResource SelectionActiveBrush}" />
                <Setter Property="Foreground" Value="White" />
            </Trigger>
            <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{DynamicResource SurfaceHoverBrush}" />
            </Trigger>
            <Trigger Property="AlternationIndex" Value="1">
                <Setter Property="Background" Value="{DynamicResource ListViewAltRowBrush}" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style TargetType="DataGridCell">
        <Setter Property="Padding" Value="8,4" />
        <Setter Property="BorderThickness" Value="0" />
        <Setter Property="FocusVisualStyle" Value="{x:Null}" />
        <Style.Triggers>
            <Trigger Property="IsSelected" Value="True">
                <Setter Property="Background" Value="Transparent" />
                <Setter Property="Foreground" Value="White" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <!-- Button Styles -->
    <Style x:Key="AccentButtonStyle" TargetType="Button">
        <Setter Property="Background" Value="{DynamicResource AccentBrush}" />
        <Setter Property="Foreground" Value="White" />
        <Setter Property="FontWeight" Value="SemiBold" />
        <Setter Property="Padding" Value="16,8" />
        <Setter Property="Cursor" Value="Hand" />
        <Setter Property="BorderThickness" Value="0" />
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}"
                            CornerRadius="4"
                            Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="Center"
                                          VerticalAlignment="Center" />
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="{DynamicResource AccentHoverBrush}" />
                        </Trigger>
                        <Trigger Property="IsPressed" Value="True">
                            <Setter Property="Background" Value="{DynamicResource AccentDarkBrush}" />
                        </Trigger>
                        <Trigger Property="IsEnabled" Value="False">
                            <Setter Property="Opacity" Value="0.5" />
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <Style x:Key="SecondaryButtonStyle" TargetType="Button">
        <Setter Property="Background" Value="Transparent" />
        <Setter Property="Foreground" Value="{DynamicResource AccentBrush}" />
        <Setter Property="FontWeight" Value="SemiBold" />
        <Setter Property="Padding" Value="16,8" />
        <Setter Property="Cursor" Value="Hand" />
        <Setter Property="BorderBrush" Value="{DynamicResource AccentBrush}" />
        <Setter Property="BorderThickness" Value="1" />
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="4"
                            Padding="{TemplateBinding Padding}">
                        <ContentPresenter HorizontalAlignment="Center"
                                          VerticalAlignment="Center" />
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="{DynamicResource SelectionBrush}" />
                        </Trigger>
                        <Trigger Property="IsPressed" Value="True">
                            <Setter Property="Background" Value="{DynamicResource SurfaceLightBrush}" />
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <!-- TextBox Style -->
    <Style TargetType="TextBox">
        <Setter Property="Background" Value="{DynamicResource SurfaceBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="CaretBrush" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}" />
        <Setter Property="BorderThickness" Value="1" />
        <Setter Property="Padding" Value="6,4" />
        <Setter Property="FontFamily" Value="{StaticResource DefaultFontFamily}" />
    </Style>

    <!-- Label Style -->
    <Style TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="FontFamily" Value="{StaticResource DefaultFontFamily}" />
    </Style>

    <!-- Menu Styles -->
    <Style TargetType="Menu">
        <Setter Property="Background" Value="{DynamicResource SurfaceBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
    </Style>

    <Style TargetType="MenuItem">
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
    </Style>

    <!-- StatusBar Style -->
    <Style TargetType="StatusBar">
        <Setter Property="Background" Value="{DynamicResource SurfaceBrush}" />
        <Setter Property="Foreground" Value="{DynamicResource TextSecondaryBrush}" />
    </Style>

    <!-- GridSplitter Style -->
    <Style TargetType="GridSplitter">
        <Setter Property="Background" Value="{DynamicResource BorderBrush}" />
    </Style>

    <!-- Section Header Style -->
    <Style x:Key="SectionHeaderStyle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource TextSecondaryBrush}" />
        <Setter Property="FontWeight" Value="SemiBold" />
        <Setter Property="FontSize" Value="12" />
        <Setter Property="Padding" Value="10,6" />
    </Style>

    <!-- Title Style -->
    <Style x:Key="TitleStyle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="FontSize" Value="18.67" />
        <Setter Property="FontWeight" Value="Light" />
    </Style>

    <!-- Header Style -->
    <Style x:Key="HeaderStyle" TargetType="TextBlock">
        <Setter Property="Foreground" Value="{DynamicResource TextPrimaryBrush}" />
        <Setter Property="FontSize" Value="16" />
        <Setter Property="FontWeight" Value="SemiBold" />
    </Style>

</ResourceDictionary>
```

**Step 5: Create ThemeManager helper class**

Create `RegistryExpert.Wpf/Helpers/ThemeManager.cs`:

```csharp
using System.Windows;

namespace RegistryExpert.Wpf.Helpers
{
    /// <summary>
    /// Manages runtime theme switching by swapping ResourceDictionaries.
    /// </summary>
    public static class ThemeManager
    {
        public enum Theme { Dark, Light }

        private static Theme _currentTheme = Theme.Dark;
        public static Theme CurrentTheme => _currentTheme;

        private static readonly Uri DarkThemeUri = new("Themes/DarkTheme.xaml", UriKind.Relative);
        private static readonly Uri LightThemeUri = new("Themes/LightTheme.xaml", UriKind.Relative);

        public static event EventHandler? ThemeChanged;

        /// <summary>
        /// Switch the application theme at runtime.
        /// </summary>
        public static void SetTheme(Theme theme)
        {
            if (_currentTheme == theme) return;
            _currentTheme = theme;

            var mergedDicts = Application.Current.Resources.MergedDictionaries;

            // Remove existing theme dictionary (it's always at index 0)
            if (mergedDicts.Count > 0)
            {
                var themeUri = theme == Theme.Dark ? DarkThemeUri : LightThemeUri;
                // Check if first dictionary is a theme dictionary
                var existing = mergedDicts[0];
                if (existing.Source == DarkThemeUri || existing.Source == LightThemeUri)
                {
                    mergedDicts.RemoveAt(0);
                }
            }

            // Insert new theme at position 0
            var newTheme = new ResourceDictionary
            {
                Source = theme == Theme.Dark ? DarkThemeUri : LightThemeUri
            };
            mergedDicts.Insert(0, newTheme);

            // Apply dark title bar via DWM
            foreach (Window window in Application.Current.Windows)
            {
                ApplyWindowChrome(window);
            }

            ThemeChanged?.Invoke(null, EventArgs.Empty);
        }

        /// <summary>
        /// Apply dark title bar and rounded corners on Windows 11.
        /// </summary>
        public static void ApplyWindowChrome(Window window)
        {
            if (Environment.OSVersion.Version.Build >= 22000)
            {
                var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                if (hwnd == IntPtr.Zero) return;

                int darkMode = _currentTheme == Theme.Dark ? 1 : 0;
                DwmSetWindowAttribute(hwnd, 20, ref darkMode, sizeof(int)); // DWMWA_USE_IMMERSIVE_DARK_MODE
                int cornerPref = 2; // DWMWCP_ROUND
                DwmSetWindowAttribute(hwnd, 33, ref cornerPref, sizeof(int)); // DWMWA_WINDOW_CORNER_PREFERENCE
            }
        }

        [System.Runtime.InteropServices.DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    }
}
```

**Step 6: Update App.xaml to load theme dictionaries**

```xml
<Application x:Class="RegistryExpert.Wpf.App"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             StartupUri="Views/MainWindow.xaml">
    <Application.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <!-- Theme (index 0 -- swapped at runtime by ThemeManager) -->
                <ResourceDictionary Source="Themes/DarkTheme.xaml" />
                <!-- Shared styles (uses DynamicResource to pick up theme colors) -->
                <ResourceDictionary Source="Themes/SharedStyles.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Application.Resources>
</Application>
```

**Step 7: Build**

```bash
dotnet build RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
```

Expected: Build succeeded. May need to create the Views directory and MainWindow.xaml first (next task).

**Step 8: Commit**

```bash
git add -A && git commit -m "feat: add WPF theme system with dark/light ResourceDictionaries and ThemeManager"
```

---

### Task 9: Create MainWindow Shell

Create the initial MainWindow with basic layout matching the WinForms app structure: menu bar, toolbar area, status bar, and a centered placeholder.

**Files:**
- Create: `RegistryExpert.Wpf/Views/MainWindow.xaml`
- Create: `RegistryExpert.Wpf/Views/MainWindow.xaml.cs`
- Create: `RegistryExpert.Wpf/ViewModels/MainViewModel.cs`
- Modify: `RegistryExpert.Wpf/App.xaml` (StartupUri)
- Delete: `RegistryExpert.Wpf/MainWindow.xaml` and `.cs` (the auto-generated ones)

**Step 1: Create Views directory**

```bash
mkdir RegistryExpert.Wpf\Views
```

**Step 2: Delete auto-generated MainWindow**

```bash
del RegistryExpert.Wpf\MainWindow.xaml
del RegistryExpert.Wpf\MainWindow.xaml.cs
```

**Step 3: Create MainWindow.xaml**

```xml
<Window x:Class="RegistryExpert.Wpf.Views.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:vm="clr-namespace:RegistryExpert.Wpf.ViewModels"
        Title="{Binding WindowTitle}"
        Width="1400" Height="900"
        MinWidth="800" MinHeight="600"
        Style="{DynamicResource ModernWindowStyle}"
        WindowStartupLocation="CenterScreen"
        AllowDrop="True">

    <Window.DataContext>
        <vm:MainViewModel />
    </Window.DataContext>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />  <!-- Menu -->
            <RowDefinition Height="Auto" />  <!-- Toolbar -->
            <RowDefinition Height="*" />     <!-- Content -->
            <RowDefinition Height="Auto" />  <!-- Status bar -->
        </Grid.RowDefinitions>

        <!-- Menu Bar -->
        <Menu Grid.Row="0">
            <MenuItem Header="_File">
                <MenuItem Header="_Open Hive..." InputGestureText="Ctrl+O" />
                <MenuItem Header="_Close Hive" />
                <Separator />
                <MenuItem Header="E_xit" InputGestureText="Alt+F4" />
            </MenuItem>
            <MenuItem Header="_View">
                <MenuItem Header="Switch Theme" />
            </MenuItem>
            <MenuItem Header="_Tools">
                <MenuItem Header="_Search..." InputGestureText="Ctrl+F" />
                <MenuItem Header="_Analyze" />
                <MenuItem Header="S_tatistics" />
                <MenuItem Header="_Compare Hives" />
                <MenuItem Header="_Timeline" />
            </MenuItem>
            <MenuItem Header="_Help">
                <MenuItem Header="Check for _Updates" />
                <MenuItem Header="_About" />
            </MenuItem>
        </Menu>

        <!-- Toolbar Placeholder -->
        <Border Grid.Row="1"
                Background="{DynamicResource SurfaceBrush}"
                BorderBrush="{DynamicResource BorderBrush}"
                BorderThickness="0,0,0,1"
                Padding="8,4">
            <TextBlock Text="Toolbar placeholder"
                       Foreground="{DynamicResource TextSecondaryBrush}" />
        </Border>

        <!-- Main Content Area (placeholder for now) -->
        <Grid Grid.Row="2" Background="{DynamicResource BackgroundBrush}">
            <StackPanel VerticalAlignment="Center"
                        HorizontalAlignment="Center">
                <TextBlock Text="Registry Expert"
                           Style="{DynamicResource TitleStyle}"
                           HorizontalAlignment="Center"
                           Margin="0,0,0,8" />
                <TextBlock Text="WPF Edition - Drag and drop a registry hive file to get started"
                           Foreground="{DynamicResource TextSecondaryBrush}"
                           HorizontalAlignment="Center" />
            </StackPanel>
        </Grid>

        <!-- Status Bar -->
        <StatusBar Grid.Row="3">
            <StatusBarItem>
                <TextBlock Text="{Binding StatusText}" />
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
```

**Step 4: Create MainWindow.xaml.cs**

```csharp
using System.Windows;
using RegistryExpert.Wpf.Helpers;

namespace RegistryExpert.Wpf.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Apply dark title bar after handle is available
            Loaded += (s, e) => ThemeManager.ApplyWindowChrome(this);
        }
    }
}
```

**Step 5: Create MainViewModel.cs (minimal shell)**

```csharp
namespace RegistryExpert.Wpf.ViewModels
{
    public class MainViewModel : ViewModelBase
    {
        private string _windowTitle = "Registry Expert";
        public string WindowTitle
        {
            get => _windowTitle;
            set => SetProperty(ref _windowTitle, value);
        }

        private string _statusText = "Ready";
        public string StatusText
        {
            get => _statusText;
            set => SetProperty(ref _statusText, value);
        }
    }
}
```

**Step 6: Update App.xaml StartupUri**

Make sure `StartupUri="Views/MainWindow.xaml"` is set (should be from Task 8).

**Step 7: Update App.xaml.cs**

```csharp
using System.Windows;

namespace RegistryExpert.Wpf
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Register code page encoding support (required by Lib/Registry parser)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        }
    }
}
```

**Step 8: Build and run**

```bash
dotnet build RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
dotnet run --project RegistryExpert.Wpf\RegistryExpert.Wpf.csproj
```

Expected: A themed dark window appears with menu bar, toolbar placeholder, centered welcome text, and status bar. **The window should render correctly on any DPI display and when moved between monitors with different DPIs.**

**Step 9: Verify WinForms project still builds**

```bash
dotnet build RegistryExpert.csproj
```

Expected: Build succeeded, 0 errors.

**Step 10: Commit**

```bash
git add -A && git commit -m "feat: add MainWindow shell with menu, toolbar, status bar, and MVVM binding"
```

---

### Task 10: Verify DPI Behavior

This is a manual verification step, not a code change. Move the WPF window between monitors with different DPI settings.

**Verification checklist:**
1. Launch `dotnet run --project RegistryExpert.Wpf\RegistryExpert.Wpf.csproj`
2. Drag the window to a monitor with different DPI (e.g., 100% -> 150%)
3. Verify: text scales smoothly, no layout breaking, no clipping
4. Verify: menu items, status bar, placeholder text all scale correctly
5. Verify: dark title bar appears on Windows 11

**If DPI works correctly**, Phase 0 is complete. The foundation is solid and ready for Phase 1 (Core Browser implementation).

---

## Phase 0 Summary

At the end of Phase 0, you have:

| Component | Status |
|-----------|--------|
| `RegistryExpert.Core` library | Contains Lib/Registry, Services, Models |
| `RegistryExpert.Wpf` app | Launches with themed window, menu, status bar |
| Theme system | Dark/Light switching via `ThemeManager.SetTheme()` |
| MVVM infrastructure | `ViewModelBase`, `RelayCommand`, `AsyncRelayCommand` |
| WinForms app | Still builds and runs unchanged |
| DPI behavior | Automatic, no manual scaling needed |

---

## Future Phases (not detailed here)

### Phase 1: Core Browser
- Registry tree view with lazy loading
- Value list view
- Detail pane
- Drag-drop hive loading
- Multi-hive tab support
- Bookmark sidebar

### Phase 2: Search
- SearchWindow with async search
- Result navigation (click result -> navigate tree + select value)

### Phase 3: Analyze Dialog
- Category list, subcategory buttons
- Content grid with detail pane
- Network, Firewall, Device Manager, Roles, Certs, Disk views

### Phase 4: Compare & Timeline
- Side-by-side hive comparison
- Transaction log timeline

### Phase 5: Polish
- About, Update, Statistics dialogs
- Settings persistence
- Remove WinForms project

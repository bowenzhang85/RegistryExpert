# AGENTS.md - RegistryExpert

## Project Overview

RegistryExpert is a .NET 8.0 Windows Forms application for viewing and analyzing offline Windows registry hive files.

## Build/Run Commands

```bash
dotnet restore              # Restore dependencies
dotnet build                # Build (debug)
dotnet run                  # Run the application
dotnet build -c Release     # Build (release)

# Publish self-contained executable
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o .\publish
```

## CI/CD

GitHub Actions (`.github/workflows/release.yml`) auto-builds releases on version tags:
```bash
git tag v1.0.0 && git push origin v1.0.0
```

### Release Checklist

**IMPORTANT: Before pushing a new release tag, always verify:**

1. **Update version in `RegistryExpert.csproj`** - The `<Version>` element must match your tag:
   ```xml
   <Version>1.0.1</Version>
   ```
2. **Commit the version change** before creating the tag
3. **Verify the About dialog** shows the correct version by running locally:
   ```bash
   dotnet publish -c Release -r win-x64 --self-contained true -o publish
   # Run publish\RegistryExpert.exe and check Help > About
   ```
4. **Create and push the tag** only after confirming the version is correct

## Testing

**No automated tests.** Manual testing: load SYSTEM/SOFTWARE/SAM/SECURITY/NTUSER.DAT hives, search, theme switching, export, compare features.

## Project Structure

```
Program.cs               - Entry point
MainForm.cs              - Main window - tree/list view, menus, toolbar
ModernTheme.cs           - Theme system (dark/light) and UI factory methods
DpiHelper.cs             - DPI scaling utilities for High DPI support
OfflineRegistryParser.cs - Registry hive loading wrapper
RegistryInfoExtractor.cs - Extracts system/user/network info from hives
RegistryComparer.cs      - Diff logic for comparing two hives
CompareForm.cs           - Side-by-side hive comparison UI
SearchForm.cs            - Search dialog with cancellation support
TimelineForm.cs          - Timeline view of registry modifications
UpdateChecker.cs         - GitHub API integration for update checking
AppSettings.cs           - Persisted settings (JSON in %LOCALAPPDATA%)
```

## Dependencies

- **Registry** (v1.3.4): Offline registry hive parser
- **Superpower** (v3.1.0): Parser combinator (Registry dependency)

## Code Style

### Naming Conventions

| Element | Style | Example |
|---------|-------|---------|
| Classes | PascalCase | `OfflineRegistryParser` |
| Methods | PascalCase | `LoadHive`, `GetRootKey` |
| Private fields | `_camelCase` | `_parser`, `_hiveType` |
| Local variables | camelCase | `computerName` |
| UI controls | `_descriptiveName` | `_treeView`, `_searchButton` |

### Nullability

**Nullable reference types enabled.** Handle nulls explicitly:

```csharp
private RegistryHive? _hive;
public RegistryKey? GetRootKey() => _hive?.Root;

// IMPORTANT: Use null-conditional before method calls
switch (value.ValueType?.ToUpperInvariant() ?? "")  // NOT: value.ValueType.ToUpperInvariant()
```

### Error Handling

```csharp
try {
    // risky operation
} catch (Exception ex) {
    System.Diagnostics.Debug.WriteLine($"Error: {ex.Message}");
}
```

### Async & Cancellation

```csharp
private CancellationTokenSource? _cts;

private async Task DoWorkAsync()
{
    _cts = new CancellationTokenSource();
    await Task.Run(() => Work(_cts.Token), _cts.Token).ConfigureAwait(true);
    UpdateUI();  // Safe - ConfigureAwait(true) returns to UI thread
}

private void CancelButton_Click(object? s, EventArgs e) => _cts?.Cancel();
```

### Resource Disposal

**Dispose images explicitly** - PictureBox.Image is not auto-disposed:

```csharp
if (ctrl is PictureBox pb && pb.Image != null)
{
    pb.Image.Dispose();
    pb.Image = null;
}
```

### Windows Forms Patterns

```csharp
// UI controls - use null! for InitializeComponent-initialized controls
private TreeView _treeView = null!;

// Always apply theme
ModernTheme.ApplyTo(this);

// Subscribe to theme changes, unsubscribe in FormClosing
private EventHandler? _themeHandler;
_themeHandler = (s, e) => ApplyTheme();
ModernTheme.ThemeChanged += _themeHandler;
FormClosing += (s, e) => ModernTheme.ThemeChanged -= _themeHandler;
```

### Data Validation

Validate external data (e.g., registry values) before using:

```csharp
if (year < 1 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31)
    return "Unknown";
```

### Common Patterns

```csharp
// Pattern matching
if (node.Tag is RegistryKey key) { ... }

// Switch expressions
return startType switch { "0" => "Boot", "1" => "System", _ => startType };

// LINQ
keys.OrderBy(k => k.KeyName).Take(50).ToList();
```

## Theme System

Access via `ModernTheme` static class:
- Colors: `ModernTheme.Background`, `ModernTheme.TextPrimary`, `ModernTheme.Accent`
- Apply: `ModernTheme.ApplyTo(control)`
- Switch: `ModernTheme.SetTheme(ThemeType.Dark)` or `ModernTheme.ToggleTheme()`
- Events: `ModernTheme.ThemeChanged`

## DPI Awareness (REQUIRED)

**All UI code must be DPI-aware.** This application supports High DPI displays (125%, 150%, 200%, etc.) via PerMonitorV2 mode.

### DpiHelper Class

Use `DpiHelper` static class for all hardcoded pixel values:

```csharp
// Scale individual values
int height = DpiHelper.Scale(32);           // 32px at 100% → 48px at 150%

// Scale sizes and positions
Size = DpiHelper.ScaleSize(300, 200);       // Form/control sizes
Location = DpiHelper.ScalePoint(25, 30);    // Control positions
Padding = DpiHelper.ScalePadding(10);       // Margins and padding
Padding = DpiHelper.ScalePadding(10, 5, 10, 5);  // Left, Top, Right, Bottom
```

### What to Scale

| Scale with DpiHelper | Do NOT Scale |
|---------------------|--------------|
| Form sizes (`Size`, `MinimumSize`) | Font point sizes (already DPI-aware) |
| Control sizes and locations | `AutoScaleMode = AutoScaleMode.Dpi` forms |
| Padding and margins | TreeView `ItemHeight` with default drawing |
| DataGridView row/column heights | |
| Custom drawing positions | |
| Icon/image dimensions | |

### New Form DPI Pattern

For new Form classes, override `OnDpiChanged` to handle multi-monitor scenarios:

```csharp
protected override void OnDpiChanged(DpiChangedEventArgs e)
{
    // Reset cached scale factor so it recalculates for new DPI
    DpiHelper.ResetScaleFactor();
    
    base.OnDpiChanged(e);
    
    // Update controls that need manual DPI adjustment
    _myGrid.RowTemplate.Height = DpiHelper.Scale(28);
    _myGrid.ColumnHeadersHeight = DpiHelper.Scale(32);
}
```

### Inline Dialog DPI Pattern

For dialogs created inline (not separate Form classes), scale all values at creation:

```csharp
using var dialog = new Form
{
    Size = DpiHelper.ScaleSize(400, 300),
    // ...
};

var label = new Label
{
    Location = DpiHelper.ScalePoint(20, 15),
    Size = DpiHelper.ScaleSize(360, 25),
    // Font sizes in points - do NOT scale
    Font = new Font("Segoe UI", 12F),
};
```

## Adding New Features

### New Analysis Category
1. Add extraction method in `RegistryInfoExtractor.cs`
2. Create `AnalysisSection` with `AnalysisItem` entries
3. Add to `GetFullAnalysis()`

### New Form
1. Inherit from `Form`, call `ModernTheme.ApplyTo(this)`
2. Subscribe to `ThemeChanged`, unsubscribe in `FormClosing`
3. Use `CancellationTokenSource` for long-running operations
4. Track in `MainForm` for disposal
5. **Override `OnDpiChanged`** - call `DpiHelper.ResetScaleFactor()` and update scaled controls
6. **Use `DpiHelper`** for all hardcoded pixel values (sizes, positions, padding)

### New Inline Dialog
1. Call `ModernTheme.ApplyTo(dialog)`
2. **Use `DpiHelper.ScaleSize()`** for dialog size
3. **Use `DpiHelper.ScalePoint()`** for control positions
4. **Use `DpiHelper.ScalePadding()`** for margins
5. Do NOT scale font point sizes (already DPI-aware)

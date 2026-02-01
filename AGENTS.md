# AGENTS.md - RegistryExpert

## Project Overview

RegistryExpert is a .NET 8.0 Windows Forms application for viewing and analyzing offline Windows registry hive files. Uses Eric Zimmerman's Registry library for parsing.

## Build/Run Commands

```bash
dotnet restore              # Restore dependencies
dotnet build                # Build (debug)
dotnet build -c Release     # Build (release)
dotnet run                  # Run the application

# Publish self-contained executable
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o .\publish
```

## Testing

**No automated test suite exists.** Manual testing should verify:
- Loading hive types: SYSTEM, SOFTWARE, SAM, SECURITY, NTUSER.DAT, USRCLASS.DAT, Amcache.hve
- Search across keys, values, and data
- Theme switching (dark/light)
- Export and Compare features

## Project Structure

```
Program.cs              - Entry point
MainForm.cs             - Main window (tree/list view, menus, toolbar)
ModernTheme.cs          - Theme system (dark/light) and UI styling
OfflineRegistryParser.cs - Registry hive loading and parsing
RegistryInfoExtractor.cs - Extracts system/user/network info from hives
CompareForm.cs          - Side-by-side hive comparison UI
SearchForm.cs           - Search dialog
TimelineForm.cs         - Timeline view of registry modifications
AppSettings.cs          - Persisted user settings (JSON)
```

## Dependencies

- **Registry** (v1.3.4): Offline registry hive parser
- **Superpower** (v3.1.0): Parser combinator (Registry dependency)

## Code Style Guidelines

### Naming Conventions

| Element | Style | Example |
|---------|-------|---------|
| Classes | PascalCase | `OfflineRegistryParser` |
| Methods | PascalCase | `LoadHive`, `GetRootKey` |
| Private fields | `_camelCase` | `_parser`, `_hiveType` |
| Local variables | camelCase | `computerName`, `shutdownTime` |
| Properties | PascalCase | `IsLoaded`, `CurrentHiveType` |
| UI controls | `_descriptiveSuffix` | `_treeView`, `_statusPanel` |

### Imports

Order: System namespaces → Third-party → Project namespaces

```csharp
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Registry;
using Registry.Abstractions;

namespace RegistryExpert
{
```

### Nullability

**Nullable reference types are enabled.** Always handle nullability explicitly:

```csharp
private RegistryHive? _hive;
private string? _filePath;

public RegistryKey? GetRootKey() => _hive?.Root;

// Use null-conditional and null-coalescing operators
var name = value?.ToString() ?? "Unknown";
```

### Error Handling

- Try-catch for file I/O and parsing operations
- Log to `System.Diagnostics.Debug.WriteLine()` for debugging
- Wrap exceptions with context in public APIs
- Silently catch in recursive operations to avoid breaking traversal

```csharp
catch (Exception ex)
{
    System.Diagnostics.Debug.WriteLine($"Error: {ex.Message}");
    throw new Exception($"Failed to load hive: {ex.Message}", ex);
}
```

### Dispose Pattern

Implement full `IDisposable` for classes holding resources:

```csharp
public class MyClass : IDisposable
{
    private bool _disposed;

    public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }
    
    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed) { /* cleanup */ _disposed = true; }
    }
    
    ~MyClass() { Dispose(false); }
}
```

### Windows Forms Patterns

```csharp
// UI control declarations - use null! for InitializeComponent-initialized controls
private TreeView _treeView = null!;
private Panel _statusPanel = null!;

// Always apply theme
ModernTheme.ApplyTo(this);
ModernTheme.ApplyTo(_treeView);

// Theme change subscription (unsubscribe in FormClosing)
private EventHandler? _themeChangedHandler;
_themeChangedHandler = (s, e) => ApplyTheme();
ModernTheme.ThemeChanged += _themeChangedHandler;

// Unsubscribe in FormClosing to prevent memory leaks
ModernTheme.ThemeChanged -= _themeChangedHandler;
```

### Async Patterns

```csharp
// Use Task.Run for CPU work, ConfigureAwait(true) to return to UI thread
private async Task LoadHiveFileAsync(string filePath)
{
    await Task.Run(() => _parser.LoadHive(filePath)).ConfigureAwait(true);
    PopulateTreeView(); // Safe - back on UI thread
}
```

### Common Patterns

```csharp
// Pattern matching
if (node.Tag is RegistryKey key) { ... }

// Switch expressions
return startType switch { "0" => "Boot", "1" => "System", _ => startType };

// LINQ
keys.OrderBy(k => k.KeyName).Take(50).ToList();

// StringBuilder for loops
var sb = new StringBuilder();
foreach (var item in items) sb.AppendLine(item);
```

### Performance

- `HashSet<T>` for O(1) lookups in search
- `ConcurrentDictionary` for thread-safe collections  
- Cache frequently accessed registry keys
- Limit recursion depth (`MaxSearchDepth = 100`)

## Theme System

Access via `ModernTheme` static class:
- Colors: `ModernTheme.Background`, `ModernTheme.TextPrimary`, `ModernTheme.Accent`
- Apply: `ModernTheme.ApplyTo(control)` - overloads for Form, TreeView, ListView, etc.
- Switch: `ModernTheme.SetTheme(ThemeType.Dark)` or `ModernTheme.ToggleTheme()`
- Events: Subscribe to `ModernTheme.ThemeChanged` for dynamic updates

## Adding New Features

### New Analysis Category
1. Add extraction method in `RegistryInfoExtractor.cs`
2. Create `AnalysisSection` with `AnalysisItem` entries
3. Include registry paths and value names
4. Add to `GetFullAnalysis()` or category method

### New Form
1. Inherit from `Form`, call `ModernTheme.ApplyTo(this)`
2. Track in `MainForm` for disposal (e.g., `_myForm?.Dispose()`)
3. Subscribe to `ThemeChanged`, unsubscribe in `FormClosing`
4. Implement dispose pattern if holding resources

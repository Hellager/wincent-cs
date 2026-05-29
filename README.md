# Wincent

Wincent is a .NET Framework library for managing Windows Explorer Quick Access items. It is a C# implementation of [wincent-rs](https://github.com/Hellager/wincent-rs), with APIs for recent files and frequent folders.

## Features

- Query Quick Access items
- Manage recent files
- Manage frequent folders
- Pin and unpin folders
- Clear Quick Access sections

## Requirements

- .NET Framework 4.8
- Windows with PowerShell

## Usage

```csharp
using Wincent;

var manager = new QuickAccessManager();

var allItems = manager.GetItems(QuickAccess.All);
var recentFiles = manager.GetItems(QuickAccess.RecentFiles);
var frequentFolders = manager.GetItems(QuickAccess.FrequentFolders);

bool containsKeyword = manager.ContainsItem("report", QuickAccess.RecentFiles);
bool containsExactPath = manager.ContainsItemExact(@"C:\path\to\file.txt", QuickAccess.RecentFiles);
```

### Add And Remove Items

```csharp
var manager = new QuickAccessManager();

manager.AddItem(@"C:\path\to\file.txt", QuickAccess.RecentFiles);
manager.AddItem(
    @"C:\path\to\folder",
    QuickAccess.FrequentFolders);

manager.RemoveItem(@"C:\path\to\file.txt", QuickAccess.RecentFiles);
manager.RemoveItem(
    @"C:\path\to\folder",
    QuickAccess.FrequentFolders,
    new RemoveOptions { DeepCleanRecentLinks = false });
```

### Clear Items

```csharp
var manager = new QuickAccessManager();

manager.ClearItems(QuickAccess.RecentFiles);
manager.ClearItems(
    QuickAccess.FrequentFolders,
    new ClearOptions
    {
        RemovePinnedFolders = true,
        RefreshExplorer = true
    });
```

### Batch Operations

```csharp
var manager = new QuickAccessManager();

BatchResult result = manager.AddItems(
    new[]
    {
        QuickAccessItem.RecentFile(@"C:\path\to\file.txt"),
        QuickAccessItem.FrequentFolder(@"C:\path\to\folder")
    },
    new BatchOptions { RefreshRecentFiles = true });

foreach (BatchFailure failure in result.Failed)
{
    Console.WriteLine($"{failure.Item.Path}: {failure.Error.Message}");
}
```

## Breaking Changes In The 0.2 Migration

- Public async APIs were removed.
- `IQuickAccessManager`, `ExecutionFeasibilityStatus`, and `QuickAccessManager.ClearCache()` were removed from the public surface.
- Operation failures are reported with exceptions instead of `bool` return values.
- Boolean behavior switches were replaced by options classes.

## License

MIT License

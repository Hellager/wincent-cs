# Wincent

Wincent is a .NET Framework library that provides a clean interface for managing Windows Quick Access items. It's a C# implementation of [wincent-rs](https://github.com/Hellager/wincent-rs), allowing you to programmatically interact with Windows Explorer's Quick Access feature, including recent files and frequent folders.

## Features

- Query Quick Access items
- Manage recent files
- Handle frequent folders
- Pin/Unpin folders to Quick Access
- UTF-8 encoding support

## Installation

Install via NuGet Package Manager:

```powershell
dotnet add package Wincent --version 0.1.4
```

## Usage

### Basic Setup

```csharp
using Wincent;

// Initialize QuickAccessManager
var quickAccessManager = new QuickAccessManager();

// Check system compatibility，Not Necessary
var (queryFeasible, handleFeasible) = await quickAccessManager.CheckFeasibleAsync();
if (!queryFeasible)
{
    Console.WriteLine("System compatibility check failed");
    return;
}
```

### Query Quick Access Items

```csharp
// Get all items
var allItems = await quickAccessManager.GetItemsAsync(QuickAccess.All);

// Get recent files
var recentFiles = await quickAccessManager.GetItemsAsync(QuickAccess.RecentFiles);

// Get frequent folders
var frequentFolders = await quickAccessManager.GetItemsAsync(QuickAccess.FrequentFolders);

// Display items
foreach (var item in recentFiles)
{
    Console.WriteLine(item);
}
```

### Add Items to Quick Access

```csharp
// Add a file to recent files
await quickAccessManager.AddItemAsync(@"C:\path\to\file.txt", QuickAccess.RecentFiles);

// Add a folder to frequent folders
await quickAccessManager.AddItemAsync(@"C:\path\to\folder", QuickAccess.FrequentFolders);
```

### Remove Items from Quick Access

```csharp
// Remove a file from recent files
await quickAccessManager.RemoveItemAsync(@"C:\path\to\file.txt", QuickAccess.RecentFiles);

// Remove a folder from frequent folders
await quickAccessManager.RemoveItemAsync(@"C:\path\to\folder", QuickAccess.FrequentFolders);
```

### Batch Operations

```csharp
// Clear all recent files
await quickAccessManager.EmptyItemsAsync(QuickAccess.RecentFiles);

// Clear all frequent folders
await quickAccessManager.EmptyItemsAsync(QuickAccess.FrequentFolders);

// Clear all items
await quickAccessManager.EmptyItemsAsync(QuickAccess.All);

// Clear with system defaults
await quickAccessManager.EmptyItemsAsync(QuickAccess.All, alsoSystemDefault: true);
```

### Check Item Existence

```csharp
// Check if a file exists in recent files
bool exists = await quickAccessManager.CheckItemAsync(@"C:\path\to\file.txt", QuickAccess.RecentFiles);
```

## Requirements

- .NET Framework 4.8
- Windows OS with PowerShell
- Administrative privileges for certain operations

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Credits

This project is a C# port of [wincent-rs](https://github.com/Hellager/wincent-rs).

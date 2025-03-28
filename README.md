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
dotnet add package Wincent --version 0.1.3
```

## Usage

### Query Recent Files

```csharp
using Wincent;

static void PrintList(IEnumerable<string> items)
{
    if (items == null) return;

    var index = 1;
    foreach (var item in items)
    {
        Console.WriteLine($"  {index++}. {item}");
    }
    Console.WriteLine($"Found {items.Count()} items");
}

var files = await QuickAccessQuery.GetRecentFilesAsync();
Console.WriteLine("\nRecent Files:");
PrintList(files);
``` 

### Query Frequent Folders

```csharp
using Wincent;

static void PrintList(IEnumerable<string> items)
{
    if (items == null) return;

    var index = 1;
    foreach (var item in items)
    {
        Console.WriteLine($"  {index++}. {item}");
    }
    Console.WriteLine($"Found {items.Count()} items");
}

var files = await QuickAccessQuery.GetFrequentFoldersAsync();
Console.WriteLine("\nFrequent Folders:");
PrintList(files);
```

### Pin Folder to Quick Access

```csharp
using Wincent;

await QuickAccessManager.AddItemAsync(@"C:\YourFolder", QuickAccessItemType.Directory);
``` 

### Remove Recent File

```csharp
using Wincent;

await QuickAccessManager.RemoveItemAsync(@"C:\YourFile.txt", QuickAccessItemType.File);
```     

## Requirements

- .NET Framework 4.8
- Windows OS with PowerShell

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Credits

This project is a C# port of [wincent-rs](https://github.com/Hellager/wincent-rs).

# Wincent

Read this in other languages: [English](README.md) | [中文](README.cn.md)

## Overview

Wincent is a .NET Framework library for managing Windows Explorer Quick Access. It is a C# implementation inspired by
[wincent-rs](https://github.com/Hellager/wincent-rs), with APIs for recent files, frequent folders, visibility, backing
file locks, and DestList metadata.

Wincent uses native Windows Shell APIs where available and PowerShell fallbacks where needed.

## Features

- Query recent files and frequent folders
- Add and remove items with duplicate detection
- Clear Quick Access sections
- Check item existence by exact path or keyword
- Batch add/remove with per-item error collection
- Timeout and retry controls for shell operations
- Visibility control for Quick Access sections and Start Recommended recent documents
- Backing file locks with Windows Recent shortcut snapshots
- DestList metadata parsing, including visible-entry helpers, hostname, DROID GUIDs, and file DROID MAC extraction
- Experimental rebuild-based DestList removal

## Installation

Reference the `Wincent` project from this solution, or install the package from your NuGet feed when published:

```powershell
Install-Package Wincent
```

## Quick Start

```csharp
using System;
using Wincent;

using (var manager = new QuickAccessManager())
{
    manager.AddItem(
        @"C:\Projects\report.docx",
        QuickAccess.RecentFiles,
        new AddOptions { RefreshRecentFiles = true });

    manager.AddItem(@"C:\Projects", QuickAccess.FrequentFolders);

    var recent = manager.GetItems(QuickAccess.RecentFiles);
    Console.WriteLine($"Recent files: {recent.Count}");

    bool exact = manager.ContainsItemExact(
        @"C:\Projects\report.docx",
        QuickAccess.RecentFiles);
    Console.WriteLine($"report.docx in Recent Files: {exact}");

    bool anyMatch = manager.ContainsItem("Projects", QuickAccess.All);
    Console.WriteLine($"Any Quick Access item contains 'Projects': {anyMatch}");

    manager.RemoveItem(
        @"C:\Projects\report.docx",
        QuickAccess.RecentFiles,
        new RemoveOptions { DeepCleanRecentLinks = true });
}
```

## System Requirements and Limitations

- **OS**: Windows 10 or Windows 11.
- **Framework**: .NET Framework 4.8.
- **PowerShell**: Required for fallback operations.
- **Explorer session**: Some Shell namespace and refresh operations require an interactive Explorer desktop session.
- **COM STA workers**: Native Shell COM operations run on background STA workers. A timeout stops waiting for the
  result but cannot cancel the Shell call; a timed-out worker may continue until Explorer returns, with at most four
  active STA workers.
- **Consistency**: Quick Access state is maintained by Windows Explorer. Results may lag behind mutations, and Explorer
  may rebuild backing files asynchronously.
- **User state**: Clear and experimental rebuild operations can remove or reset Quick Access state. Frequent Folders
  rebuild can restore the default pinned folders: Desktop, Downloads, Documents, and Pictures.

## Development

```powershell
dotnet build Wincent.sln
dotnet test Wincent.sln
```

Integration tests are marked with `TestCategory("Integration")` and skipped by default. They require an interactive
Explorer desktop session and may read or refresh the current user's real Explorer state.

## Disclaimer

This library interacts with Windows Explorer state for the current user. Review the behavior of clear and experimental
operations before using them on important user profiles.

## Acknowledgements

- [wincent-rs](https://github.com/Hellager/wincent-rs)
- [Castorix31](https://learn.microsoft.com/en-us/answers/questions/1087928/how-to-get-recent-docs-list-and-delete-some-of-the)
- [Yohan Ney](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel)
- [libyal](https://github.com/libyal/dtformats/blob/main/documentation/Jump%20lists%20format.asciidoc)
- [Eric Zimmerman](https://github.com/EricZimmerman/JumpList)
- [kacos2000](https://github.com/kacos2000/Jumplist-Browser)

## License

MIT License. See [LICENSE](LICENSE) for details.

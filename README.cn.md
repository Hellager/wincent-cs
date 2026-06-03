# Wincent

其他语言版本：[English](README.md) | [中文](README.cn.md)

## 概览

Wincent 是一个用于管理 Windows Explorer 快速访问的 .NET Framework 库。它是受
[wincent-rs](https://github.com/Hellager/wincent-rs) 启发的 C# 实现，提供最近文件、常用文件夹、可见性、backing
file 锁和 DestList 元数据相关 API。

Wincent 会优先使用原生 Windows Shell API，并在需要时回退到 PowerShell。

## 功能

- 查询最近文件和常用文件夹
- 添加和删除项目，并检测重复项
- 清空快速访问分区
- 按精确路径或关键字检查项目是否存在
- 批量添加/删除，并逐项收集错误
- 为 Shell 操作配置超时和重试
- 控制快速访问分区可见性
- 锁定 backing file，并记录 Windows Recent 快捷方式快照
- 解析 DestList 元数据
- 实验性通过重建删除 DestList 项

## 安装

可以在当前解决方案中引用 `Wincent` 项目；如果已发布到你的 NuGet 源，也可以安装包：

```powershell
Install-Package Wincent
```

## 快速开始

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

## 系统要求与限制

- **操作系统**：Windows 10 或 Windows 11。
- **框架**：.NET Framework 4.8。
- **PowerShell**：回退操作需要 PowerShell。
- **Explorer 会话**：部分 Shell namespace 和刷新操作需要交互式 Explorer 桌面会话。
- **一致性**：快速访问状态由 Windows Explorer 维护。修改后查询结果可能短暂滞后，Explorer 也可能异步重建 backing file。
- **用户状态**：清空和实验性重建操作可能移除或重置快速访问状态。常用文件夹重建可能恢复默认固定文件夹：桌面、下载、文档和图片。

## 开发

```powershell
dotnet build Wincent.sln
dotnet test Wincent.sln
```

部分集成测试需要交互式 Explorer 桌面会话，默认会跳过。

## 免责声明

本库会操作当前用户的 Windows Explorer 状态。在重要用户配置上使用清空或实验性操作前，请先确认相关行为。

## 致谢

- [wincent-rs](https://github.com/Hellager/wincent-rs)
- [Castorix31](https://learn.microsoft.com/en-us/answers/questions/1087928/how-to-get-recent-docs-list-and-delete-some-of-the)
- [Yohan Ney](https://stackoverflow.com/questions/30051634/is-it-possible-programmatically-add-folders-to-the-windows-10-quick-access-panel)
- [libyal](https://github.com/libyal/dtformats/blob/main/documentation/Jump%20lists%20format.asciidoc)
- [Eric Zimmerman](https://github.com/EricZimmerman/JumpList)
- [kacos2000](https://github.com/kacos2000/Jumplist-Browser)

## 许可证

MIT License。详情见 [LICENSE](LICENSE)。

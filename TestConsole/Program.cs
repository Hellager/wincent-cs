using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Text;
using Wincent;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();
        }

        static async Task MainAsync(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            try
            {
                // 初始化QuickAccessManager并检查系统兼容性
                var quickAccessManager = new QuickAccessManager();

                // 显示欢迎信息
                DisplayWelcomeMessage();

                // 检查系统健康状态
                await CheckSystemHealth(quickAccessManager);

                // 进入主菜单循环
                await RunMainMenuLoop(quickAccessManager);
            }
            catch (Exception ex)
            {
                DisplayError($"程序运行时发生错误: {ex.Message}");
                Console.WriteLine("按任意键退出...");
                Console.ReadKey();
            }
        }

        static void DisplayWelcomeMessage()
        {
            Console.Clear();
            Console.WriteLine("==========================================");
            Console.WriteLine("      Windows 快速访问管理工具 v1.0       ");
            Console.WriteLine("==========================================");
            Console.WriteLine("本工具帮助您管理 Windows 的快速访问项目");
            Console.WriteLine();
        }

        static async Task CheckSystemHealth(IQuickAccessManager quickAccessManager)
        {
            Console.WriteLine("正在检查系统兼容性...");

            var (queryFeasible, handleFeasible) = await quickAccessManager.CheckFeasibleAsync();

            if (queryFeasible)
            {
                Console.WriteLine("✅ 查询功能正常");
            }
            else
            {
                Console.WriteLine("❌ 查询功能不可用，请确保系统权限设置正确");
            }

            if (handleFeasible)
            {
                Console.WriteLine("✅ 操作功能正常");
            }
            else
            {
                Console.WriteLine("⚠️ 操作功能不可用，某些功能可能无法使用");
            }

            Console.WriteLine();

            if (!queryFeasible)
            {
                Console.WriteLine("系统检查失败，程序无法继续执行。按任意键退出...");
                Console.ReadKey();
                Environment.Exit(1);
            }
        }

        static async Task RunMainMenuLoop(IQuickAccessManager quickAccessManager)
        {
            bool exit = false;

            while (!exit)
            {
                DisplayMainMenu();

                var choice = Console.ReadKey(true).KeyChar;
                Console.WriteLine();

                switch (choice)
                {
                    case '1':
                        await ViewQuickAccessItems(quickAccessManager);
                        break;
                    case '2':
                        await AddNewItem(quickAccessManager);
                        break;
                    case '3':
                        await RemoveItem(quickAccessManager);
                        break;
                    case '4':
                        await PerformBatchOperations(quickAccessManager);
                        break;
                    case '5':
                        await SearchItems(quickAccessManager);
                        break;
                    case '0':
                        exit = true;
                        break;
                    default:
                        Console.WriteLine("无效的选择，请重试");
                        WaitForKeyPress();
                        break;
                }
            }

            Console.WriteLine("感谢使用 Windows 快速访问管理工具，再见！");
        }

        static void DisplayMainMenu()
        {
            Console.Clear();
            Console.WriteLine("=== 主菜单 ===");
            Console.WriteLine("1. 查看快速访问项");
            Console.WriteLine("2. 添加新项目");
            Console.WriteLine("3. 移除项目");
            Console.WriteLine("4. 批量操作");
            Console.WriteLine("5. 关键字搜索");
            Console.WriteLine("0. 退出程序");
            Console.Write("请选择: ");
        }

        static async Task ViewQuickAccessItems(IQuickAccessManager quickAccessManager)
        {
            Console.Clear();
            Console.WriteLine("=== 查看快速访问项 ===");
            Console.WriteLine("1. 查看所有项目");
            Console.WriteLine("2. 查看最近使用的文件");
            Console.WriteLine("3. 查看常用文件夹");
            Console.WriteLine("0. 返回主菜单");
            Console.Write("请选择: ");

            var choice = Console.ReadKey(true).KeyChar;
            Console.WriteLine();

            QuickAccess qaType;
            string title;

            switch (choice)
            {
                case '1':
                    qaType = QuickAccess.All;
                    title = "所有项目";
                    break;
                case '2':
                    qaType = QuickAccess.RecentFiles;
                    title = "最近使用的文件";
                    break;
                case '3':
                    qaType = QuickAccess.FrequentFolders;
                    title = "常用文件夹";
                    break;
                case '0':
                    return;
                default:
                    Console.WriteLine("无效的选择，请重试");
                    WaitForKeyPress();
                    return;
            }

            await DisplayItems(quickAccessManager, qaType, title);
        }

        static async Task DisplayItems(IQuickAccessManager quickAccessManager, QuickAccess qaType, string title)
        {
            Console.Clear();
            Console.WriteLine($"=== {title} ===");

            try
            {
                Console.WriteLine("正在获取数据...");
                var items = await quickAccessManager.GetItemsAsync(qaType);

                Console.Clear();
                Console.WriteLine($"=== {title} ===");

                if (items.Count == 0)
                {
                    Console.WriteLine("没有找到任何项目");
                }
                else
                {
                    for (int i = 0; i < items.Count; i++)
                    {
                        string itemType = IsDirectory(items[i]) ? "[目录]" : "[文件]";
                        Console.WriteLine($"{i + 1}. {itemType} {items[i]}");
                    }
                    Console.WriteLine($"共 {items.Count} 个项目");
                }
            }
            catch (Exception ex)
            {
                DisplayError($"获取项目列表时出错: {ex.Message}");
            }

            WaitForKeyPress();
        }

        static async Task AddNewItem(IQuickAccessManager quickAccessManager)
        {
            Console.Clear();
            Console.WriteLine("=== 添加新项目 ===");
            Console.WriteLine("1. 添加最近使用的文件");
            Console.WriteLine("2. 添加常用文件夹");
            Console.WriteLine("0. 返回主菜单");
            Console.Write("请选择: ");

            var choice = Console.ReadKey(true).KeyChar;
            Console.WriteLine();

            QuickAccess qaType;
            PathType pathType;

            switch (choice)
            {
                case '1':
                    qaType = QuickAccess.RecentFiles;
                    pathType = PathType.File;
                    break;
                case '2':
                    qaType = QuickAccess.FrequentFolders;
                    pathType = PathType.Directory;
                    break;
                case '0':
                    return;
                default:
                    Console.WriteLine("无效的选择，请重试");
                    WaitForKeyPress();
                    return;
            }

            Console.Write($"请输入要添加的{(pathType == PathType.File ? "文件" : "文件夹")}完整路径: ");
            string path = Console.ReadLine().Trim();

            if (string.IsNullOrEmpty(path))
            {
                DisplayError("路径不能为空");
                WaitForKeyPress();
                return;
            }

            // 验证路径是否有效
            if ((pathType == PathType.File && !File.Exists(path)) ||
                (pathType == PathType.Directory && !Directory.Exists(path)))
            {
                DisplayError($"{(pathType == PathType.File ? "文件" : "文件夹")}不存在");
                WaitForKeyPress();
                return;
            }

            try
            {
                Console.WriteLine("正在添加项目...");
                bool result = await quickAccessManager.AddItemAsync(path, qaType, true);

                if (result)
                {
                    Console.WriteLine("✅ 项目添加成功");
                }
                else
                {
                    Console.WriteLine("⚠️ 项目添加失败");
                }
            }
            catch (Exception ex)
            {
                DisplayError($"添加项目时出错: {ex.Message}");
            }

            WaitForKeyPress();
        }

        static async Task RemoveItem(IQuickAccessManager quickAccessManager)
        {
            Console.Clear();
            Console.WriteLine("=== 移除项目 ===");
            Console.WriteLine("1. 移除最近使用的文件");
            Console.WriteLine("2. 移除常用文件夹");
            Console.WriteLine("0. 返回主菜单");
            Console.Write("请选择: ");

            var choice = Console.ReadKey(true).KeyChar;
            Console.WriteLine();

            QuickAccess qaType;
            string title;

            switch (choice)
            {
                case '1':
                    qaType = QuickAccess.RecentFiles;
                    title = "最近使用的文件";
                    break;
                case '2':
                    qaType = QuickAccess.FrequentFolders;
                    title = "常用文件夹";
                    break;
                case '0':
                    return;
                default:
                    Console.WriteLine("无效的选择，请重试");
                    WaitForKeyPress();
                    return;
            }

            try
            {
                Console.WriteLine("正在获取数据...");
                var items = await quickAccessManager.GetItemsAsync(qaType);

                Console.Clear();
                Console.WriteLine($"=== 移除{title} ===");

                if (items.Count == 0)
                {
                    Console.WriteLine("没有找到任何项目");
                    WaitForKeyPress();
                    return;
                }

                for (int i = 0; i < items.Count; i++)
                {
                    string itemType = IsDirectory(items[i]) ? "[目录]" : "[文件]";
                    Console.WriteLine($"{i + 1}. {itemType} {items[i]}");
                }

                Console.WriteLine($"\n请选择要移除的项目编号 (1-{items.Count})，或输入0返回: ");
                if (!int.TryParse(Console.ReadLine(), out int itemIndex) || itemIndex < 0 || itemIndex > items.Count)
                {
                    Console.WriteLine("无效的选择");
                    WaitForKeyPress();
                    return;
                }

                if (itemIndex == 0)
                {
                    return;
                }

                string selectedPath = items[itemIndex - 1];

                Console.WriteLine($"您确定要移除 \"{selectedPath}\" 吗? (Y/N)");
                if (Console.ReadKey(true).Key != ConsoleKey.Y)
                {
                    Console.WriteLine("\n已取消操作");
                    WaitForKeyPress();
                    return;
                }

                Console.WriteLine("\n正在移除项目...");
                bool result = await quickAccessManager.RemoveItemAsync(selectedPath, qaType);

                if (result)
                {
                    Console.WriteLine("✅ 项目移除成功");
                }
                else
                {
                    Console.WriteLine("⚠️ 项目移除失败");
                }
            }
            catch (Exception ex)
            {
                DisplayError($"移除项目时出错: {ex.Message}");
            }

            WaitForKeyPress();
        }

        static async Task PerformBatchOperations(IQuickAccessManager quickAccessManager)
        {
            Console.Clear();
            Console.WriteLine("=== 批量操作 ===");
            Console.WriteLine("1. 清空最近使用的文件");
            Console.WriteLine("2. 清空常用文件夹");
            Console.WriteLine("3. 清空所有快速访问项目");
            Console.WriteLine("0. 返回主菜单");
            Console.Write("请选择: ");

            var choice = Console.ReadKey(true).KeyChar;
            Console.WriteLine();

            QuickAccess qaType;
            string operationName;

            switch (choice)
            {
                case '1':
                    qaType = QuickAccess.RecentFiles;
                    operationName = "最近使用的文件";
                    break;
                case '2':
                    qaType = QuickAccess.FrequentFolders;
                    operationName = "常用文件夹";
                    break;
                case '3':
                    qaType = QuickAccess.All;
                    operationName = "所有快速访问项目";
                    break;
                case '0':
                    return;
                default:
                    Console.WriteLine("无效的选择，请重试");
                    WaitForKeyPress();
                    return;
            }

            Console.WriteLine($"⚠️ 警告：此操作将清空{operationName}，且不可恢复。");
            Console.WriteLine("是否继续? (Y/N)");

            if (Console.ReadKey(true).Key != ConsoleKey.Y)
            {
                Console.WriteLine("\n已取消操作");
                WaitForKeyPress();
                return;
            }

            Console.WriteLine("\n是否同时清除系统默认项目? (Y/N)");
            bool alsoSystemDefault = Console.ReadKey(true).Key == ConsoleKey.Y;

            try
            {
                Console.WriteLine("\n正在执行操作...");
                bool result = await quickAccessManager.EmptyItemsAsync(qaType, true, alsoSystemDefault);

                if (result)
                {
                    Console.WriteLine($"✅ 已成功清空{operationName}");
                }
                else
                {
                    Console.WriteLine("⚠️ 操作失败");
                }
            }
            catch (Exception ex)
            {
                DisplayError($"执行批量操作时出错: {ex.Message}");
            }

            WaitForKeyPress();
        }

        static async Task SearchItems(IQuickAccessManager quickAccessManager)
        {
            Console.Clear();
            Console.WriteLine("=== 关键字搜索 ===");
            Console.WriteLine("1. 搜索所有项目");
            Console.WriteLine("2. 搜索最近使用的文件");
            Console.WriteLine("3. 搜索常用文件夹");
            Console.WriteLine("0. 返回主菜单");
            Console.Write("请选择: ");

            var choice = Console.ReadKey(true).KeyChar;
            Console.WriteLine();

            QuickAccess qaType;
            string title;

            switch (choice)
            {
                case '1':
                    qaType = QuickAccess.All;
                    title = "所有项目";
                    break;
                case '2':
                    qaType = QuickAccess.RecentFiles;
                    title = "最近使用的文件";
                    break;
                case '3':
                    qaType = QuickAccess.FrequentFolders;
                    title = "常用文件夹";
                    break;
                case '0':
                    return;
                default:
                    Console.WriteLine("无效的选择，请重试");
                    WaitForKeyPress();
                    return;
            }

            Console.Write("请输入搜索关键字: ");
            string keyword = Console.ReadLine().Trim().ToLower();

            if (string.IsNullOrEmpty(keyword))
            {
                DisplayError("搜索关键字不能为空");
                WaitForKeyPress();
                return;
            }

            Console.Clear();
            Console.WriteLine($"=== 搜索结果 - {title} - 关键字: \"{keyword}\" ===");

            try
            {
                Console.WriteLine("正在搜索...");
                var items = await quickAccessManager.GetItemsAsync(qaType);
                var results = items.Where(item => item.ToLower().Contains(keyword)).ToList();

                Console.Clear();
                Console.WriteLine($"=== 搜索结果 - {title} - 关键字: \"{keyword}\" ===");

                if (results.Count == 0)
                {
                    Console.WriteLine("没有找到匹配的项目");
                }
                else
                {
                    for (int i = 0; i < results.Count; i++)
                    {
                        string itemType = IsDirectory(results[i]) ? "[目录]" : "[文件]";
                        Console.WriteLine($"{i + 1}. {itemType} {results[i]}");
                    }
                    Console.WriteLine($"共找到 {results.Count} 个匹配项");

                    Console.WriteLine("\n您想对找到的项目执行什么操作？");
                    Console.WriteLine("1. 移除选定项目");
                    Console.WriteLine("0. 返回主菜单");
                    Console.Write("请选择: ");

                    var operationChoice = Console.ReadKey(true).KeyChar;
                    Console.WriteLine();

                    switch (operationChoice)
                    {
                        case '1':
                            await RemoveSearchResult(quickAccessManager, qaType, results);
                            break;
                        case '0':
                            return;
                        default:
                            Console.WriteLine("无效的选择，返回主菜单");
                            WaitForKeyPress();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                DisplayError($"搜索项目时出错: {ex.Message}");
            }

            WaitForKeyPress();
        }

        static async Task RemoveSearchResult(IQuickAccessManager quickAccessManager, QuickAccess qaType, List<string> items)
        {
            Console.WriteLine($"\n请选择要移除的项目编号 (1-{items.Count})，或输入0返回: ");
            if (!int.TryParse(Console.ReadLine(), out int itemIndex) || itemIndex < 0 || itemIndex > items.Count)
            {
                Console.WriteLine("无效的选择");
                return;
            }

            if (itemIndex == 0)
            {
                return;
            }

            string selectedPath = items[itemIndex - 1];

            Console.WriteLine($"您确定要移除 \"{selectedPath}\" 吗? (Y/N)");
            if (Console.ReadKey(true).Key != ConsoleKey.Y)
            {
                Console.WriteLine("\n已取消操作");
                return;
            }

            Console.WriteLine("\n正在移除项目...");
            bool result = await quickAccessManager.RemoveItemAsync(selectedPath, qaType);

            if (result)
            {
                Console.WriteLine("✅ 项目移除成功");
            }
            else
            {
                Console.WriteLine("⚠️ 项目移除失败");
            }
        }

        static bool IsDirectory(string path)
        {
            try
            {
                return Directory.Exists(path);
            }
            catch
            {
                return false;
            }
        }

        static void DisplayError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"错误: {message}");
            Console.ResetColor();
        }

        static void WaitForKeyPress()
        {
            Console.WriteLine("\n按任意键继续...");
            Console.ReadKey(true);
        }
    }
}
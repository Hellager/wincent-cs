using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Threading.Tasks;
using Wincent;

namespace TestConsole
{
    class Program
    {
        private static readonly Dictionary<int, string> MenuOptions = new Dictionary<int, string>()
        {
            {1, "检查执行策略"},
            {2, "修复执行策略"}, // 新增选项
            {3, "添加文件到最近访问"},
            {4, "固定文件夹到常用位置"},
            {5, "列出最近访问文件"},
            {6, "列出常用文件夹"},
            {7, "清空所有快速访问项"},
            {8, "测试路径保护检测"}, // 原7改为8
            {0, "退出程序"}

        };

        static async Task Main(string[] args)
        {
            Console.Title = "Wincent 类库测试工具";
            ShowWelcome();

            while (true)
            {
                ShowMenu();
                var input = GetUserInput("请选择操作编号");

                if (!int.TryParse(input, out int choice) || !MenuOptions.ContainsKey(choice))
                {
                    ShowError("无效的选项，请重新输入");
                    continue;
                }

                if (choice == 0) break;

                try
                {
                    Console.Clear();
                    ShowSectionTitle(MenuOptions[choice]);

                    switch (choice)
                    {
                        case 1:
                            CheckExecutionPolicy();
                            break;
                        case 2: // 新增case
                            FixExecutionPolicy();
                            break;
                        case 3:
                            await HandleAddFile();
                            break;
                        case 4:
                            await HandlePinFolder();
                            break;
                        case 5:
                            await ListRecentFiles();
                            break;
                        case 6:
                            await ListFrequentFolders();
                            break;
                        case 7:
                            HandleClearQuickAccess();
                            break;
                        case 8:
                            TestProtectedPaths();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    ShowError($"操作失败: {ex.Message}", pause: true);
                }

                Console.WriteLine("\n按任意键返回主菜单...");
                Console.ReadKey();
                Console.Clear();
            }

            ShowFarewell();
        }

        static void ShowWelcome()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(@"
    _       ___                       __ 
| |     / (_)___  ________  ____  / /_
| | /| / / / __ \/ ___/ _ \/ __ \/ __/
| |/ |/ / / / / / /__/  __/ / / / /_  
|__/|__/_/_/ /_/\___/\___/_/ /_/\__/ 
");
            Console.ResetColor();
            Console.WriteLine("欢迎使用 Wincent 类库交互测试工具\n");
        }

        static void ShowMenu()
        {
            Console.WriteLine("════════════ 主菜单 ════════════");
            foreach (var option in MenuOptions)
            {
                Console.WriteLine($"  [{option.Key}] {option.Value}");
            }
            Console.WriteLine("════════════════════════════════");
        }

        static async Task HandleAddFile()
        {
            var path = GetUserInput("请输入文件完整路径");
            Console.WriteLine();
            await ShowLoading("添加文件中", async () =>
            {
                await QuickAccessManager.AddItemAsync(path, QuickAccessItemType.File);
                return true;
            });
            ShowSuccess("文件已成功添加到最近访问！");
        }

        static async Task HandlePinFolder()
        {
            var path = GetUserInput("请输入文件夹完整路径");
            Console.WriteLine();
            await ShowLoading("固定文件夹中", async () =>
            {
                await QuickAccessManager.AddItemAsync(path, QuickAccessItemType.Directory);
                return true;
            });
            ShowSuccess("文件夹已成功固定到常用位置！");
        }

        static void CheckExecutionPolicy()
        {
            Console.WriteLine("当前执行策略状态：");
            Console.WriteLine($"  ▪ 策略名称: {ExecutionFeasible.GetExecutionPolicy()}");
            Console.WriteLine($"  ▪ 允许执行: {ExecutionFeasible.CheckScriptFeasible().ToYesNo()}");
            Console.WriteLine($"  ▪ 管理员权限: {ExecutionFeasible.IsAdministrator().ToYesNo()}");
        }

        static void FixExecutionPolicy()
        {

            var confirm = GetUserInput("确定要将执行策略设置为RemoteSigned吗？(y/n)").ToLower();
            if (confirm != "y") return;

            try
            {
                ExecutionFeasible.FixExecutionPolicy();

                // 二次验证
                var newPolicy = ExecutionFeasible.GetExecutionPolicy();
                if (newPolicy.Equals("RemoteSigned", StringComparison.OrdinalIgnoreCase))
                {
                    ShowSuccess("执行策略已成功设置为RemoteSigned！");
                }
                else
                {
                    ShowError($"策略设置异常，当前策略：{newPolicy}", pause: true);
                }
            }
            catch (SecurityException ex)
            {
                ShowError($"权限不足：{ex.Message}", pause: true);
            }
            catch (Exception ex)
            {
                ShowError($"策略修改失败：{ex.Message}", pause: true);
            }
        }


        static async Task ListRecentFiles()
        {
            var files = await ShowLoading("获取最近文件", async () =>
                await QuickAccessQuery.GetRecentFilesAsync());

            Console.WriteLine("\n最近访问文件列表：");
            PrintList(files);
        }

        static async Task ListFrequentFolders()
        {
            var folders = await ShowLoading("获取常用文件夹", async () =>
                await QuickAccessQuery.GetFrequentFoldersAsync());

            Console.WriteLine("\n常用文件夹列表：");
            PrintList(folders);
        }

        static void HandleClearQuickAccess()
        {
            var confirm = GetUserInput("确定要清空所有快速访问项吗？(y/n)").ToLower();
            if (confirm != "y") return;

            Console.WriteLine();
            ShowLoading("正在清空", () => Task.Run(() =>
            {
                QuickAccessManager.EmptyQuickAccess();
                return true;
            })).Wait();
            ShowSuccess("已成功清空所有快速访问项！");
        }

        static void TestProtectedPaths()
        {
            var testCases = new Dictionary<string, QuickAccessItemType>
            {
                { @"C:\Windows\System32\drivers", QuickAccessItemType.Directory },
                { @"C:\Program Files\Windows NT", QuickAccessItemType.Directory }
            };

            Console.WriteLine("保护路径测试结果：");
            foreach (var testCase in testCases)
            {
                try
                {
                    QuickAccessManager.ValidatePathSecurity(testCase.Key, testCase.Value);
                    Console.WriteLine($"  ✓ {testCase.Key}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  ✕ {testCase.Key} ({ex.Message})");
                }
            }
        }

        static string GetUserInput(string prompt)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"{prompt} => ");
            Console.ResetColor();
            return Console.ReadLine().Trim();
        }

        static void PrintList(IEnumerable<string> items)
        {
            if (items == null) return;

            var index = 1;
            foreach (var item in items)
            {
                Console.WriteLine($"  {index++}. {item}");
            }
            Console.WriteLine($"共发现 {items.Count()} 个项目");
        }

        // 新增非泛型版本
        static async Task ShowLoading(string text, Func<Task> action)
        {
            Console.Write($"{text} ");
            var dots = new ConsoleSpinner();
            var task = action.Invoke();

            while (!task.IsCompleted)
            {
                dots.Update();
                await Task.Delay(100);
            }

            Console.WriteLine("✓");
        }

        // 保留原有泛型版本
        static async Task<T> ShowLoading<T>(string text, Func<Task<T>> action)
        {
            Console.Write($"{text} ");
            var dots = new ConsoleSpinner();
            var task = action.Invoke();

            while (!task.IsCompleted)
            {
                dots.Update();
                await Task.Delay(100);
            }

            Console.WriteLine("✓");
            return await task;
        }

        static void ShowSectionTitle(string title)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"【{title}】\n");
            Console.ResetColor();
        }

        static void ShowSuccess(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"✔ {message}");
            Console.ResetColor();
        }

        static void ShowError(string message, bool pause = false)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"✘ {message}");
            Console.ResetColor();
            if (pause)
            {
                Console.WriteLine("按任意键继续...");
                Console.ReadKey();
            }
        }

        static void ShowFarewell()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("感谢使用，再见！");
            Console.ResetColor();
            Task.Delay(1000).Wait();
        }
    }

    public static class Extensions
    {
        public static string ToYesNo(this bool value) => value ? "是" : "否";
    }

    public class ConsoleSpinner
    {
        private int _counter;

        public void Update()
        {
            _counter++;
            var spinChars = new[] { '⣾', '⣽', '⣻', '⢿', '⡿', '⣟', '⣯', '⣷' };
            Console.Write($"\b{spinChars[_counter % spinChars.Length]}");
        }
    }
}

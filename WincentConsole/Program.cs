using Wincent;

class Program
{
    static async Task Main(string[] args)
    {
        while (true)
        {
            Console.Clear();
            Console.WriteLine("=== Wincent 快速访问管理工具 ===");
            Console.WriteLine("1. 添加文件到最近访问");
            Console.WriteLine("2. 添加文件夹到快速访问");
            Console.WriteLine("3. 从最近访问移除文件");
            Console.WriteLine("4. 从快速访问移除文件夹");
            Console.WriteLine("5. 清空最近访问的文件");
            Console.WriteLine("6. 清空常用文件夹");
            Console.WriteLine("7. 清空所有快速访问项");
            Console.WriteLine("8. 查看最近访问的文件");
            Console.WriteLine("9. 查看常用文件夹");
            Console.WriteLine("0. 退出");
            Console.Write("\n请选择操作 (0-9): ");

            if (!int.TryParse(Console.ReadLine(), out int choice))
            {
                Console.WriteLine("\n无效的输入，请按任意键继续...");
                Console.ReadKey();
                continue;
            }

            try
            {
                switch (choice)
                {
                    case 0:
                        return;

                    case 1:
                        Console.Write("\n请输入文件路径: ");
                        string filePath = Console.ReadLine() ?? "";
                        await QuickAccessManager.AddItemAsync(filePath, QuickAccessItemType.File);
                        Console.WriteLine("文件已添加到最近访问");
                        break;

                    case 2:
                        Console.Write("\n请输入文件夹路径: ");
                        string folderPath = Console.ReadLine() ?? "";
                        await QuickAccessManager.AddItemAsync(folderPath, QuickAccessItemType.Directory);
                        Console.WriteLine("文件夹已添加到快速访问");
                        break;

                    case 3:
                        Console.Write("\n请输入要移除的文件路径: ");
                        string removeFilePath = Console.ReadLine() ?? "";
                        await QuickAccessManager.RemoveItemAsync(removeFilePath, QuickAccessItemType.File);
                        Console.WriteLine("文件已从最近访问移除");
                        break;

                    case 4:
                        Console.Write("\n请输入要移除的文件夹路径: ");
                        string removeFolderPath = Console.ReadLine() ?? "";
                        await QuickAccessManager.RemoveItemAsync(removeFolderPath, QuickAccessItemType.Directory);
                        Console.WriteLine("文件夹已从快速访问移除");
                        break;

                    case 5:
                        QuickAccessManager.EmptyRecentFiles();
                        Console.WriteLine("已清空最近访问的文件");
                        break;

                    case 6:
                        QuickAccessManager.EmptyFrequentFolders();
                        Console.WriteLine("已清空常用文件夹");
                        break;

                    case 7:
                        QuickAccessManager.EmptyQuickAccess();
                        Console.WriteLine("已清空所有快速访问项");
                        break;

                    case 8:
                        var recentFiles = await QuickAccessQuery.GetRecentFilesAsync();
                        Console.WriteLine("\n最近访问的文件:");
                        foreach (var file in recentFiles)
                        {
                            Console.WriteLine(file);
                        }
                        break;

                    case 9:
                        var frequentFolders = await QuickAccessQuery.GetFrequentFoldersAsync();
                        Console.WriteLine("\n常用文件夹:");
                        foreach (var folder in frequentFolders)
                        {
                            Console.WriteLine(folder);
                        }
                        break;

                    default:
                        Console.WriteLine("\n无效的选择");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\n操作失败: {ex.Message}");
            }

            Console.WriteLine("\n按任意键继续...");
            Console.ReadKey();
        }
    }
}

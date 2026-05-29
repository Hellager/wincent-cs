using System;
using Wincent;

namespace TestConsole
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            var manager = new QuickAccessManager();

            Console.WriteLine("Wincent Quick Access sample");
            Console.WriteLine();

            foreach (var item in manager.GetItems(QuickAccess.All))
            {
                Console.WriteLine(item);
            }
        }
    }
}

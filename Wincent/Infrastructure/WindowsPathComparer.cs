using System;
using System.IO;

namespace Wincent
{
    internal static class WindowsPathComparer
    {
        public static bool Equals(string left, string right)
        {
            return string.Equals(Normalize(left), Normalize(right), StringComparison.OrdinalIgnoreCase);
        }

        internal static string Normalize(string path)
        {
            if (path == null)
                return null;

            string normalized = path.Replace('/', '\\').TrimEnd('\\');
            try
            {
                normalized = Path.GetFullPath(normalized).TrimEnd('\\');
            }
            catch (Exception)
            {
                // Invalid paths still participate in lightweight Windows-style comparison.
            }

            return normalized;
        }
    }
}

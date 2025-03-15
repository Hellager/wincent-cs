using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Wincent
{
    /// <summary>
    /// Wrapper class for temporary files with automatic cleanup, supports custom file extensions
    /// Usage example:
    /// using var tempFile = TempFile.Create("Hello World", "txt");
    /// var content = File.ReadAllText(tempFile.FullPath);
    /// </summary>
    public sealed class TempFile : IDisposable
    {
        /// <summary>
        /// Gets the full path of the temporary file
        /// </summary>
        public string FullPath { get; }

        /// <summary>
        /// Gets the filename (including extension)
        /// </summary>
        public string FileName => Path.GetFileName(FullPath);

        /// <summary>
        /// Gets the directory name where temporary files are created
        /// </summary>
        public static string DirName { get; set; } = "WincentTemp";

        private bool _disposed;

        /// <summary>
        /// Creates a temporary file and writes binary content
        /// </summary>
        /// <param name="content">Binary file content</param>
        /// <param name="extension">File extension (with or without leading dot)</param>
        public static TempFile Create(byte[] content, string extension = ".tmp", Encoding encoding = null)
        {
            if (encoding == null) encoding = Encoding.UTF8;
            _ = content ?? throw new ArgumentNullException(nameof(content));

            ValidateExtension(ref extension);

            string fullPath = GenerateFilePath(extension);

            try
            {
                File.WriteAllBytes(fullPath, content);
                return new TempFile(fullPath);
            }
            catch
            {
                SafeDelete(fullPath);
                throw;
            }
        }

        /// <summary>
        /// Creates a temporary file and writes text content
        /// </summary>
        /// <param name="content">File content</param>
        /// <param name="extension">File extension (with or without leading dot)</param>
        public static TempFile Create(string content, string extension = ".tmp")
        {
            _ = content ?? throw new ArgumentNullException(nameof(content));

            ValidateExtension(ref extension);

            string fullPath = GenerateFilePath(extension);

            try
            {
                File.WriteAllText(fullPath, content);
                return new TempFile(fullPath);
            }
            catch
            {
                SafeDelete(fullPath);
                throw;
            }
        }

        #region Helper Methods
        private static void ValidateExtension(ref string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
                extension = ".tmp";

            if (!extension.StartsWith("."))
                extension = "." + extension;
        }

        private static string GenerateFilePath(string extension)
        {
            string baseTempDir = Path.GetTempPath();
            string tempDir = Path.Combine(baseTempDir, DirName);

            // Create the custom temp directory if it doesn't exist
            if (!Directory.Exists(tempDir))
            {
                try
                {
                    Directory.CreateDirectory(tempDir);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to create temp directory: {ex.Message}");
                    // Fall back to system temp directory if custom directory creation fails
                    tempDir = baseTempDir;
                }
            }

            string fileName = $"{Guid.NewGuid():N}{extension}";
            return Path.Combine(tempDir, fileName);
        }
        #endregion

        private TempFile(string fullPath)
        {
            FullPath = fullPath ?? throw new ArgumentNullException(nameof(fullPath));
        }

        /// <summary>
        /// Gets file stream (read mode by default)
        /// </summary>
        public FileStream OpenRead()
        {
            return _disposed ? throw new ObjectDisposedException(nameof(TempFile)) : File.OpenRead(FullPath);
        }

        /// <summary>
        /// Reads all text content from the file
        /// </summary>
        public string ReadAllText() => ReadAllText(Encoding.UTF8);

        public string ReadAllText(Encoding encoding)
        {
            return _disposed ? throw new ObjectDisposedException(nameof(TempFile)) : File.ReadAllText(FullPath, encoding);
        }

        public void Dispose()
        {
            if (_disposed) return;

            SafeDelete(FullPath);
            _disposed = true;
            GC.SuppressFinalize(this);
        }

        private static void SafeDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"TempFile deletion failed: {path}\nError: {ex.Message}");
            }
        }

        ~TempFile() => Dispose();
    }
}

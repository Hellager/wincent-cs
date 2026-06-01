using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;

namespace Wincent
{
    /// <summary>
    /// Wrapper class for temporary files with automatic cleanup, supports custom file extensions
    /// Usage example:
    /// using var tempFile = TempFile.Create("Hello World", "txt");
    /// var content = File.ReadAllText(tempFile.FullPath);
    /// </summary>
    internal sealed class TempFile : IDisposable
    {
        public const string DefaultDirectoryName = "WincentTemp";

        private static readonly Encoding DefaultTextEncoding = new UTF8Encoding(false);

        /// <summary>
        /// Gets the full path of the temporary file
        /// </summary>
        public string FullPath { get; }

        /// <summary>
        /// Gets the filename (including extension)
        /// </summary>
        public string FileName => Path.GetFileName(FullPath);

        private int _disposed;

        /// <summary>
        /// Creates a temporary file and writes binary content
        /// </summary>
        /// <param name="content">Binary file content</param>
        /// <param name="extension">File extension (with or without leading dot)</param>
        /// <param name="directoryName">Temp subdirectory name.</param>
        public static TempFile Create(
            byte[] content,
            string extension = ".ps1",
            string directoryName = DefaultDirectoryName)
        {
            _ = content ?? throw new ArgumentNullException(nameof(content));

            ValidateExtension(ref extension);

            string fullPath = GenerateFilePath(extension, directoryName);

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
        /// Creates a temporary file and writes text content with the specified encoding
        /// </summary>
        /// <param name="content">File content</param>
        /// <param name="extension">File extension (with or without leading dot)</param>
        /// <param name="encoding">Text encoding used to write the file. Uses UTF-8 without BOM when null.</param>
        /// <param name="directoryName">Temp subdirectory name.</param>
        public static TempFile Create(
            string content,
            string extension = ".tmp",
            Encoding encoding = null,
            string directoryName = DefaultDirectoryName)
        {
            _ = content ?? throw new ArgumentNullException(nameof(content));

            ValidateExtension(ref extension);

            string fullPath = GenerateFilePath(extension, directoryName);

            try
            {
                File.WriteAllText(fullPath, content, encoding ?? DefaultTextEncoding);
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
                extension = ".ps1";

            if (!extension.StartsWith("."))
                extension = "." + extension;
        }

        private static string GenerateFilePath(string extension, string directoryName)
        {
            string baseTempDir = Path.GetTempPath();
            string tempDir = baseTempDir;
            string normalizedDirectoryName = string.IsNullOrWhiteSpace(directoryName)
                ? DefaultDirectoryName
                : directoryName;

            try
            {
                string customTempDir = Path.Combine(baseTempDir, normalizedDirectoryName);

                // Create the custom temp directory if it doesn't exist
                if (!Directory.Exists(customTempDir))
                {
                    Directory.CreateDirectory(customTempDir);
                }

                tempDir = customTempDir;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to create temp directory: {ex.Message}");
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
            ThrowIfDisposed();
            return File.OpenRead(FullPath);
        }

        /// <summary>
        /// Reads all text content from the file
        /// </summary>
        public string ReadAllText() => ReadAllText(DefaultTextEncoding);

        public string ReadAllText(Encoding encoding)
        {
            ThrowIfDisposed();
            return File.ReadAllText(FullPath, encoding);
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0)
                return;

            SafeDelete(FullPath);
        }

        private void ThrowIfDisposed()
        {
            if (Volatile.Read(ref _disposed) != 0)
                throw new ObjectDisposedException(nameof(TempFile));
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
    }
}

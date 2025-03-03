using System;
using System.IO;

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

        private bool _disposed;

        /// <summary>
        /// Creates a temporary file and writes binary content
        /// </summary>
        /// <param name="content">Binary file content</param>
        /// <param name="extension">File extension (with or without leading dot)</param>
        public static TempFile Create(byte[]? content, string? extension = ".tmp")
        {
            ArgumentNullException.ThrowIfNull(content);

            ValidateExtension(ref extension!);

            string fullPath = GenerateFilePath(extension);

            try
            {
                File.WriteAllBytes(fullPath, content);
            }
            catch
            {
                SafeDelete(fullPath);
                throw;
            }

            return new TempFile(fullPath);
        }

        /// <summary>
        /// Creates a temporary file and writes text content
        /// </summary>
        /// <param name="content">File content</param>
        /// <param name="extension">File extension (with or without leading dot)</param>
        public static TempFile Create(string? content, string? extension = ".tmp")
        {
            ArgumentNullException.ThrowIfNull(content);

            ValidateExtension(ref extension!);

            string fullPath = GenerateFilePath(extension);

            try
            {
                File.WriteAllText(fullPath, content);
            }
            catch
            {
                SafeDelete(fullPath);
                throw;
            }

            return new TempFile(fullPath);
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
            string tempDir = Path.GetTempPath();
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
        public FileStream OpenRead() => File.OpenRead(FullPath);

        /// <summary>
        /// Reads all text content from the file
        /// </summary>
        public string ReadAllText() => File.ReadAllText(FullPath);

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
            catch
            {
                // Logging or other error handling logic
            }
        }

        ~TempFile() => Dispose();
    }
}

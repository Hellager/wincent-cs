using System;
using System.IO;

namespace Wincent
{
    /// <summary>
    /// File system operations interface for dependency injection and unit testing
    /// </summary>
    public interface IFileSystem
    {
        bool FileExists(string path);
        void DeleteFile(string path);
        DateTime GetLastWriteTime(string path);
    }

    /// <summary>
    /// Implementation of real file system operations
    /// </summary>
    public class RealFileSystem : IFileSystem
    {
        public bool FileExists(string path) => File.Exists(path);
        public void DeleteFile(string path) => File.Delete(path);
        public DateTime GetLastWriteTime(string path) => File.GetLastWriteTime(path);
    }

    /// <summary>
    /// Windows Quick Access data file management and information
    /// </summary>
    public class QuickAccessDataFiles
    {
        private readonly string _recentFilesPath;
        private readonly string _frequentFoldersPath;
        private readonly IFileSystem _fileSystem;

        /// <summary>
        /// Creates Quick Access data file manager using real file system
        /// </summary>
        /// <exception cref="InvalidPathException">Thrown when unable to retrieve Windows Recent folder path</exception>
        public QuickAccessDataFiles() : this(new RealFileSystem()) { }

        /// <summary>
        /// Creates Quick Access data file manager with specified file system
        /// </summary>
        /// <param name="fileSystem">File system operations interface</param>
        /// <exception cref="InvalidPathException">Thrown when unable to retrieve Windows Recent folder path</exception>
        /// <exception cref="ArgumentNullException">Thrown when fileSystem parameter is null</exception>
        public QuickAccessDataFiles(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));

            string recentFolder = GetWindowsRecentFolder();
            string automaticDestDir = Path.Combine(recentFolder, "AutomaticDestinations");

            _recentFilesPath = Path.Combine(automaticDestDir, "5f7b5f1e01b83767.automaticDestinations-ms");
            _frequentFoldersPath = Path.Combine(automaticDestDir, "f01b4d95cf55d32a.automaticDestinations-ms");
        }

        /// <summary>
        /// Retrieves Windows Recent folder path
        /// </summary>
        /// <returns>Path to Recent folder</returns>
        /// <exception cref="InvalidPathException">Thrown when path retrieval fails</exception>
        private string GetWindowsRecentFolder()
        {
            string recentFolder = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            return string.IsNullOrEmpty(recentFolder)
                ? throw new InvalidPathException("Windows Recent Folder", "Failed to retrieve Windows Recent folder path")
                : recentFolder;
        }

        /// <summary>
        /// Unified file deletion logic with error handling
        /// </summary>
        /// <param name="path">File path to delete</param>
        /// <exception cref="IOException">Thrown when file deletion fails</exception>
        private void RemoveFile(string path)
        {
            try
            {
                if (_fileSystem.FileExists(path))
                    _fileSystem.DeleteFile(path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                throw new IOException($"Failed to delete file: {path}", ex);
            }
        }

        /// <summary>
        /// Deletes Recent Files data file (ignores if file doesn't exist)
        /// </summary>
        /// <exception cref="IOException">Thrown when file deletion fails</exception>
        public void RemoveRecentFile() => RemoveFile(_recentFilesPath);

        /// <summary>
        /// Gets last modified time of Recent Files data
        /// </summary>
        /// <returns>Last modified timestamp</returns>
        public DateTime GetRecentFilesModifiedTime() => GetModifiedTime(_recentFilesPath);

        /// <summary>
        /// Gets last modified time of Frequent Folders data
        /// </summary>
        /// <returns>Last modified timestamp</returns>
        public DateTime GetFrequentFoldersModifiedTime() => GetModifiedTime(_frequentFoldersPath);

        /// <summary>
        /// Gets latest modified time across Quick Access data files
        /// </summary>
        /// <returns>Newest timestamp from both data files</returns>
        public DateTime GetQuickAccessModifiedTime()
        {
            DateTime recentTime = GetRecentFilesModifiedTime();
            DateTime frequentTime = GetFrequentFoldersModifiedTime();
            return recentTime > frequentTime ? recentTime : frequentTime;
        }

        /// <summary>
        /// Gets modified time based on script type
        /// </summary>
        /// <param name="scriptType">Script type identifier</param>
        /// <returns>Corresponding data file modified time</returns>
        public DateTime GetModifiedTimeForScript(PSScript scriptType)
        {
            switch (scriptType)
            {
                case PSScript.QueryRecentFile:
                    return GetRecentFilesModifiedTime();
                case PSScript.QueryFrequentFolder:
                    return GetFrequentFoldersModifiedTime();
                case PSScript.QueryQuickAccess:
                    return GetQuickAccessModifiedTime();
                default:
                    return DateTime.Now; // Use current time for non-query scripts
            }
        }

        private DateTime GetModifiedTime(string path)
        {
            try
            {
                return _fileSystem.FileExists(path)
                    ? _fileSystem.GetLastWriteTime(path)
                    : DateTime.Now;
            }
            catch
            {
                return DateTime.Now;
            }
        }

        /// <summary>
        /// Path to Recent Files data file
        /// </summary>
        public string RecentFilesPath => _recentFilesPath;

        /// <summary>
        /// Path to Frequent Folders data file
        /// </summary>
        public string FrequentFoldersPath => _frequentFoldersPath;
    }
}

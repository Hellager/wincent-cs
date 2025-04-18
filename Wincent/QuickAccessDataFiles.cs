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
        public bool FileExists(string path)
        {
            return File.Exists(path);
        }

        public void DeleteFile(string path)
        {
            File.Delete(path);
        }

        public DateTime GetLastWriteTime(string path)
        {
            return File.GetLastWriteTime(path);
        }
    }

    /// <summary>
    /// Quick Access data files management interface
    /// </summary>
    public interface IQuickAccessDataFiles
    {
        /// <summary>
        /// Gets modification time for script-related data file
        /// </summary>
        /// <param name="scriptType">Script type</param>
        /// <returns>Modification timestamp</returns>
        DateTime GetModifiedTimeForScript(PSScript scriptType);

        /// <summary>
        /// Removes recent files data file
        /// </summary>
        void RemoveRecentFile();

        /// <summary>
        /// Gets path to recent files data in Quick Access
        /// </summary>
        string RecentFilesPath { get; }

        /// <summary>
        /// Gets path to frequent folders data in Quick Access
        /// </summary>
        string FrequentFoldersPath { get; }
    }

    /// <summary>
    /// Windows Quick Access data file management
    /// </summary>
    public class QuickAccessDataFiles : IQuickAccessDataFiles
    {
        private readonly string _recentFilesPath;
        private readonly string _frequentFoldersPath;
        private readonly IFileSystem _fileSystem;

        /// <summary>
        /// Initializes new instance of <see cref="QuickAccessDataFiles"/>
        /// </summary>
        public QuickAccessDataFiles()
            : this(new RealFileSystem())
        {
        }

        /// <summary>
        /// Initializes new instance of <see cref="QuickAccessDataFiles"/>
        /// </summary>
        /// <param name="fileSystem">File system interface</param>
        public QuickAccessDataFiles(IFileSystem fileSystem)
        {
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));

            string recentFolder = GetWindowsRecentFolder();
            string automaticDestDir = Path.Combine(recentFolder, "AutomaticDestinations");

            _recentFilesPath = Path.Combine(automaticDestDir, "5f7b5f1e01b83767.automaticDestinations-ms");
            _frequentFoldersPath = Path.Combine(automaticDestDir, "f01b4d95cf55d32a.automaticDestinations-ms");
        }

        /// <summary>
        /// Gets path to Windows Recent folder
        /// </summary>
        /// <returns>Recent folder path</returns>
        /// <exception cref="InvalidPathException">Thrown when path retrieval fails</exception>
        private string GetWindowsRecentFolder()
        {
            string recentFolder = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            if (string.IsNullOrEmpty(recentFolder))
            {
                throw new InvalidPathException("Windows Recent Folder", "Unable to retrieve Windows Recent folder path");
            }
            return recentFolder;
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
                {
                    _fileSystem.DeleteFile(path);
                }
                // If file doesn't exist, no exception is thrown
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                throw new IOException($"Failed to delete file: {path}", ex);
            }
        }

        /// <summary>
        /// Removes recent files data file (ignores if file doesn't exist)
        /// </summary>
        /// <exception cref="IOException">Thrown when deletion fails</exception>
        public void RemoveRecentFile()
        {
            RemoveFile(_recentFilesPath);
        }

        /// <summary>
        /// Gets modification time of recent files data
        /// </summary>
        /// <returns>Modification timestamp</returns>
        /// <exception cref="FileNotFoundException">Thrown when file doesn't exist</exception>
        public DateTime GetRecentFilesModifiedTime()
        {
            if (_fileSystem.FileExists(_recentFilesPath))
            {
                try
                {
                    return _fileSystem.GetLastWriteTime(_recentFilesPath);
                }
                catch (Exception ex)
                {
                    // Capture and rethrow access errors
                    throw new IOException($"Failed to get file timestamp: {_recentFilesPath}", ex);
                }
            }
            else
            {
                throw new FileNotFoundException($"Recent files data file not found: {_recentFilesPath}", _recentFilesPath);
            }
        }

        /// <summary>
        /// Gets modification time of frequent folders data
        /// </summary>
        /// <returns>Modification timestamp</returns>
        /// <exception cref="FileNotFoundException">Thrown when file doesn't exist</exception>
        public DateTime GetFrequentFoldersModifiedTime()
        {
            if (_fileSystem.FileExists(_frequentFoldersPath))
            {
                try
                {
                    return _fileSystem.GetLastWriteTime(_frequentFoldersPath);
                }
                catch (Exception ex)
                {
                    throw new IOException($"Failed to get file timestamp: {_frequentFoldersPath}", ex);
                }
            }
            else
            {
                throw new FileNotFoundException($"Frequent folders data file not found: {_frequentFoldersPath}", _frequentFoldersPath);
            }
        }

        /// <summary>
        /// Gets modification time of Quick Access data
        /// </summary>
        /// <returns>Modification timestamp</returns>
        /// <exception cref="FileNotFoundException">Thrown when no data files exist</exception>
        public DateTime GetQuickAccessModifiedTime()
        {
            DateTime recentTime = DateTime.MinValue;
            DateTime frequentTime = DateTime.MinValue;
            bool anyFileExists = false;

            try
            {
                recentTime = GetRecentFilesModifiedTime();
                anyFileExists = true;
            }
            catch (FileNotFoundException)
            {
                // Ignore missing recent files
            }

            try
            {
                frequentTime = GetFrequentFoldersModifiedTime();
                anyFileExists = true;
            }
            catch (FileNotFoundException)
            {
                // Ignore missing frequent folders
            }

            if (!anyFileExists)
            {
                throw new FileNotFoundException("No available Quick Access data files");
            }

            return recentTime > frequentTime ? recentTime : frequentTime;
        }

        /// <summary>
        /// Gets modification time for specific script type
        /// </summary>
        /// <param name="scriptType">Script type</param>
        /// <returns>Related data file modification time</returns>
        /// <exception cref="FileNotFoundException">Thrown when target data file is missing</exception>
        public DateTime GetModifiedTimeForScript(PSScript scriptType)
        {
            try
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
            catch (FileNotFoundException) when (scriptType != PSScript.QueryRecentFile &&
                                               scriptType != PSScript.QueryFrequentFolder &&
                                               scriptType != PSScript.QueryQuickAccess)
            {
                // Return current time for non-specific query scripts
                return DateTime.Now;
            }
        }

        /// <summary>
        /// Gets path to recent files data
        /// </summary>
        public string RecentFilesPath => _recentFilesPath;

        /// <summary>
        /// Gets path to frequent folders data
        /// </summary>
        public string FrequentFoldersPath => _frequentFoldersPath;
    }
}

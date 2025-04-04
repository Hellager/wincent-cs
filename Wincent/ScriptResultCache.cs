using System;
using System.Collections.Concurrent;
using System.IO;

namespace Wincent
{
    internal class ScriptResultCache
    {
        private static readonly string RecentFileTrackerPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            @"Microsoft\Windows\Recent\AutomaticDestinations\5f7b5f1e01b83767.automaticDestinations-ms");

        private static readonly string FrequentFolderTrackerPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            @"Microsoft\Windows\Recent\AutomaticDestinations\f01b4d95cf55d32a.automaticDestinations-ms");

        private class CacheEntry
        {
            public ScriptResult Result { get; }
            public DateTime LastModifiedTime { get; }

            public CacheEntry(ScriptResult result, DateTime lastModifiedTime)
            {
                Result = result;
                LastModifiedTime = lastModifiedTime;
            }
        }

        private readonly ConcurrentDictionary<PSScript, CacheEntry> _cache = new ConcurrentDictionary<PSScript, CacheEntry>();

        public ScriptResult GetCachedResult(PSScript script)
        {
            if (!ShouldCache(script))
                return null;

            if (_cache.TryGetValue(script, out var entry))
            {
                var currentModTime = GetTrackerLastModifiedTime(script);
                if (currentModTime <= entry.LastModifiedTime)
                {
                    return entry.Result;
                }
            }
            return null;
        }

        public void UpdateCache(PSScript script, ScriptResult result)
        {
            if (!ShouldCache(script))
                return;

            var entry = new CacheEntry(result, GetTrackerLastModifiedTime(script));
            _cache.AddOrUpdate(script, entry, (_, __) => entry);
        }

        private static bool ShouldCache(PSScript script)
        {
            return script == PSScript.QueryQuickAccess ||
                   script == PSScript.QueryRecentFile ||
                   script == PSScript.QueryFrequentFolder;
        }

        private static DateTime GetTrackerLastModifiedTime(PSScript script)
        {
            try
            {
                switch (script)
                {
                    case PSScript.QueryRecentFile:
                        return File.GetLastWriteTime(RecentFileTrackerPath);
                    
                    case PSScript.QueryFrequentFolder:
                        return File.GetLastWriteTime(FrequentFolderTrackerPath);
                    
                    case PSScript.QueryQuickAccess:
                        var recentTime = File.GetLastWriteTime(RecentFileTrackerPath);
                        var frequentTime = File.GetLastWriteTime(FrequentFolderTrackerPath);
                        return recentTime > frequentTime ? recentTime : frequentTime;
                    
                    default:
                        return DateTime.MinValue;
                }
            }
            catch
            {
                return DateTime.MinValue;
            }
        }
    }
} 
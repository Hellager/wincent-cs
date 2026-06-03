using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Wincent;

namespace TestConsole
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                    InteractiveLoop();
                else
                    Run(args);
            }
            catch (Exception ex)
            {
                PrintError(ex);
                Environment.Exit(1);
            }
        }

        #region Interactive loop

        private static void InteractiveLoop()
        {
            Console.WriteLine("wincent interactive example CLI");
            Console.WriteLine("Type `help` to list commands, `exit` or `quit` to leave.");

            while (true)
            {
                Console.Write("wincent> ");
                var line = Console.ReadLine();
                if (line == null)
                    break;

                line = line.Trim();
                if (line.Length == 0)
                    continue;
                if (line == "exit" || line == "quit")
                    break;

                var cmdArgs = SplitCommandLine(line);
                if (cmdArgs.Length == 0)
                    continue;

                try
                {
                    Run(cmdArgs);
                }
                catch (Exception ex)
                {
                    PrintError(ex);
                }
            }
        }

        #endregion

        #region Command dispatch

        private static void Run(string[] args)
        {
            if (args.Length == 0 || args[0] == "help" || args[0] == "--help" || args[0] == "-h")
            {
                PrintHelp();
                return;
            }

            var (timeout, remaining) = ParseGlobalOptions(args);
            if (remaining.Length == 0)
            {
                PrintHelp();
                return;
            }

            // Re-check help after global options have been stripped.
            if (remaining[0] == "help" || remaining[0] == "--help" || remaining[0] == "-h")
            {
                PrintHelp();
                return;
            }

            var manager = new QuickAccessManager(new QuickAccessManagerOptions { Timeout = timeout });

            using (manager)
            {
                switch (remaining[0])
                {
                    case "features":      CmdFeatures(); break;
                    case "list":          CmdList(manager, remaining); break;
                    case "list-paths":    CmdListPaths(manager, remaining); break;
                    case "check":         CmdCheck(manager, remaining); break;
                    case "contains":      CmdContains(manager, remaining); break;
                    case "add":           CmdAdd(manager, remaining); break;
                    case "remove":        CmdRemove(manager, remaining); break;
                    case "batch-add":     CmdBatchAdd(manager, remaining); break;
                    case "batch-remove":  CmdBatchRemove(manager, remaining); break;
                    case "lock":          CmdLock(manager, remaining); break;
                    case "empty":         CmdEmpty(manager, remaining); break;
                    case "clear-cache":   CmdClearCache(); break;
                    case "retry":         CmdRetry(remaining); break;
                    case "classify":      CmdClassify(remaining); break;
                    case "invalid-path":  CmdInvalidPath(remaining); break;
                    case "visible":       CmdVisible(manager, remaining); break;
                    case "dest":          CmdDest(manager, remaining); break;
                    default:
                        throw new ArgumentException($"unknown command: {remaining[0]}");
                }
            }
        }

        #endregion

        #region Commands — core

        private static void CmdFeatures()
        {
            Console.WriteLine("visible: supported");
            Console.WriteLine("destlist: supported");
        }

        private static void CmdList(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 2, "list <recent|frequent|all> [--paths]");
            var qa = ParseQuickAccess(args[1], allowAll: true);

            if (args.Skip(2).Any(a => a == "--paths"))
            {
                CmdListPathsCore(manager, qa);
                return;
            }

            var items = manager.GetItems(qa);
            PrintStringList("items", items);
        }

        private static void CmdListPaths(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 2, "list-paths <recent|frequent|all>");
            var qa = ParseQuickAccess(args[1], allowAll: true);
            CmdListPathsCore(manager, qa);
        }

        private static void CmdListPathsCore(QuickAccessManager manager, QuickAccess qa)
        {
            var items = manager.GetItems(qa);
            Console.WriteLine("paths: {0}", items.Count);
            foreach (var path in items)
                Console.WriteLine(path);
        }

        private static void CmdCheck(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "check <recent|frequent|all> <path>");
            var qa = ParseQuickAccess(args[1], allowAll: true);
            var exists = manager.ContainsItemExact(args[2], qa);
            Console.WriteLine(exists);
        }

        private static void CmdContains(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "contains <recent|frequent|all> <keyword>");
            var qa = ParseQuickAccess(args[1], allowAll: true);
            var exists = manager.ContainsItem(args[2], qa);
            Console.WriteLine(exists);
        }

        private static void CmdAdd(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "add <recent|frequent> <path> [--refresh]");
            var qa = ParseQuickAccess(args[1], allowAll: false);
            var refresh = args.Skip(3).Any(a => a == "--refresh");

            manager.AddItem(args[2], qa, new AddOptions { RefreshRecentFiles = refresh });
            Console.WriteLine("added {0}", args[2]);
        }

        private static void CmdRemove(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "remove <recent|frequent> <path> [--deep-clean]");
            var qa = ParseQuickAccess(args[1], allowAll: false);
            var deepClean = args.Skip(3).Any(a => a == "--deep-clean");

            manager.RemoveItem(args[2], qa, new RemoveOptions { DeepCleanRecentLinks = deepClean });
            Console.WriteLine("removed {0}", args[2]);
        }

        private static void CmdBatchAdd(QuickAccessManager manager, string[] args)
        {
            var (refresh, itemArgs) = SplitFlag(args, 1, "--refresh");
            if (itemArgs.Length == 0)
                throw new ArgumentException("batch-add requires at least one item");

            var items = ParseBatchItems(itemArgs);
            var options = new BatchOptions { RefreshRecentFiles = refresh };
            var result = manager.AddItems(items, options);
            PrintBatchResult(result);
        }

        private static void CmdBatchRemove(QuickAccessManager manager, string[] args)
        {
            var (deepClean, itemArgs) = SplitFlag(args, 1, "--deep-clean");
            if (itemArgs.Length == 0)
                throw new ArgumentException("batch-remove requires at least one item");

            var items = ParseBatchItems(itemArgs);
            var options = new RemoveOptions { DeepCleanRecentLinks = deepClean };
            var result = manager.RemoveItems(items, options);
            PrintBatchResult(result);
        }

        private static void CmdLock(QuickAccessManager manager, string[] args)
        {
            var cleanup = args.Skip(1).Any(a => a == "--cleanup-new-links");
            var targetArg = args.Skip(1).FirstOrDefault(a => a != "--cleanup-new-links") ?? "all";
            var target = ParseLockTarget(targetArg);

            QuickAccessLock @lock;
            switch (target)
            {
                case QuickAccessLockTarget.RecentFiles:
                    @lock = manager.LockRecentFiles();
                    break;
                case QuickAccessLockTarget.FrequentFolders:
                    @lock = manager.LockFrequentFolders();
                    break;
                default:
                    @lock = manager.LockQuickAccess();
                    break;
            }

            using (@lock)
            {
                Console.WriteLine("locked Quick Access backing files");
                Console.WriteLine("target: {0}", @lock.Target);
                Console.WriteLine("recent_folder: {0}", @lock.RecentFolder);
                Console.WriteLine("initial_lnk_paths: {0}", @lock.InitialShortcutPaths.Count);
                Console.WriteLine("locked_file_count: {0}", @lock.LockedFileCount);
                Console.WriteLine("press Enter to unlock");

                Console.ReadLine();

                var unlockOptions = new QuickAccessUnlockOptions { CleanupNewRecentLinks = cleanup };
                var report = @lock.Unlock(unlockOptions);
                Console.WriteLine("current_lnk_paths: {0}", report.CurrentShortcutPaths.Count);
                Console.WriteLine("new_lnk_paths: {0}", report.NewShortcutPaths.Count);
                Console.WriteLine("deleted_lnk_paths: {0}", report.DeletedShortcutPaths.Count);
                Console.WriteLine("failed_lnk_deletions: {0}", report.FailedShortcutDeletions.Count);
                foreach (var p in report.DeletedShortcutPaths)
                    Console.WriteLine("  deleted {0}", p);
                foreach (var f in report.FailedShortcutDeletions)
                    Console.WriteLine("  failed {0}: {1}", f.Path, f.Error.Message);
            }
        }

        private static void CmdEmpty(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 2, "empty <recent|frequent|all> [--pinned] [--refresh]");
            var qa = ParseQuickAccess(args[1], allowAll: true);

            bool pinned = false, refresh = false;
            for (int i = 2; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--pinned":  pinned = true; break;
                    case "--refresh": refresh = true; break;
                    default:
                        throw new ArgumentException($"unknown empty option: {args[i]}");
                }
            }

            manager.ClearItems(qa, new ClearOptions
            {
                RemovePinnedFolders = pinned,
                RefreshExplorer = refresh
            });
            Console.WriteLine("cleared {0}", CategoryName(qa));
        }

        private static void CmdClearCache()
        {
            Console.WriteLine("cache cleared (currently a no-op)");
        }

        #endregion

        #region Commands — utility

        private static void CmdRetry(string[] args)
        {
            RequireMinArgs(args, 2, "retry <policy> [--attempt N] [--max-attempts N] [--initial-ms N] [--max-ms N] [--factor N] [--jitter true|false]");
            var policyName = args[1];

            int? maxAttempts = null;
            int? initialMs = null;
            int? maxMs = null;
            double? factor = null;
            bool? jitter = null;
            int attempt = 0;

            for (int i = 2; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--attempt":
                        attempt = ParseInt(RequireNextArg(args, ref i, "attempt"), "attempt");
                        break;
                    case "--max-attempts":
                        maxAttempts = ParseInt(RequireNextArg(args, ref i, "max-attempts"), "max-attempts");
                        break;
                    case "--initial-ms":
                        initialMs = ParseInt(RequireNextArg(args, ref i, "initial-ms"), "initial-ms");
                        break;
                    case "--max-ms":
                        maxMs = ParseInt(RequireNextArg(args, ref i, "max-ms"), "max-ms");
                        break;
                    case "--factor":
                        factor = ParseDouble(RequireNextArg(args, ref i, "factor"), "factor");
                        break;
                    case "--jitter":
                        jitter = ParseBool(RequireNextArg(args, ref i, "jitter"));
                        break;
                    default:
                        throw new ArgumentException($"unknown retry option: {args[i]}");
                }
            }

            RetryPolicy policy;
            switch (policyName)
            {
                case "none": case "no-retry":    policy = RetryPolicy.None; break;
                case "fast":                     policy = RetryPolicy.Fast; break;
                case "default": case "standard": policy = RetryPolicy.Standard; break;
                case "aggressive":               policy = RetryPolicy.Aggressive; break;
                case "custom":
                    policy = new RetryPolicy(
                        maxAttempts ?? 3,
                        TimeSpan.FromMilliseconds(initialMs ?? 100),
                        TimeSpan.FromMilliseconds(maxMs ?? 5000),
                        factor ?? 2.0,
                        jitter ?? true);
                    Console.WriteLine("max_attempts: {0}", policy.MaxRetryCount);
                    Console.WriteLine("initial_delay_ms: {0}", policy.InitialDelay.TotalMilliseconds);
                    Console.WriteLine("max_delay_ms: {0}", policy.MaxDelay.TotalMilliseconds);
                    Console.WriteLine("backoff_factor: {0}", policy.BackoffFactor);
                    Console.WriteLine("jitter: {0}", policy.UseJitter);
                    if (policy.MaxRetryCount > 0)
                        Console.WriteLine("delay_at_{0}_ms: {1}", attempt, policy.GetDelay(attempt).TotalMilliseconds);
                    return;
                default:
                    throw new ArgumentException($"unknown retry policy: {policyName}");
            }

            Console.WriteLine("max_attempts: {0}", policy.MaxRetryCount);
            Console.WriteLine("initial_delay_ms: {0}", policy.InitialDelay.TotalMilliseconds);
            Console.WriteLine("max_delay_ms: {0}", policy.MaxDelay.TotalMilliseconds);
            Console.WriteLine("backoff_factor: {0}", policy.BackoffFactor);
            Console.WriteLine("jitter: {0}", policy.UseJitter);
            if (policy.MaxRetryCount > 0)
                Console.WriteLine("delay_at_{0}_ms: {1}", attempt, policy.GetDelay(attempt).TotalMilliseconds);
        }

        private static void CmdClassify(string[] args)
        {
            RequireMinArgs(args, 2, "classify <stderr text>");
            var stderr = string.Join(" ", args.Skip(1));
            var inferred = InferKindFromStderr(stderr);
            var classified = InferKindCustom(stderr);
            Console.WriteLine("infer_kind_from_stderr: {0}", inferred);
            Console.WriteLine("classify_with: {0}", classified);
        }

        private static void CmdInvalidPath(string[] args)
        {
            RequireMinArgs(args, 2, "invalid-path <reason> [path]");
            var reason = args[1];
            var path = args.Length > 2 ? args[2] : null;

            Console.WriteLine("reason: {0}", reason);
            if (path != null)
                Console.WriteLine("path: {0}", path);
            else
                Console.WriteLine("path: <none>");
            Console.WriteLine("display: InvalidPathError {{ reason: \"{0}\"{1} }}",
                reason, path != null ? $", path: \"{path}\"" : "");
        }

        #endregion

        #region Commands — visible

        private static void CmdVisible(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 2, "visible <command>");

            switch (args[1])
            {
                case "get":
                    RequireMinArgs(args, 3, "visible get <recent|frequent|all>");
                    Console.WriteLine(manager.IsVisible(ParseQuickAccess(args[2], allowAll: true)));
                    break;
                case "set":
                    RequireMinArgs(args, 4, "visible set <recent|frequent|all> <true|false> [--refresh]");
                    var setTarget = ParseQuickAccess(args[2], allowAll: true);
                    var setOptions = ParseVisibilityOptions(args, 4);
                    manager.SetVisible(
                        setTarget,
                        ParseBool(args[3]),
                        setOptions);
                    PrintVisibilityMutationResult("updated visibility", setTarget, setOptions);
                    break;
                case "show":
                    RequireMinArgs(args, 3, "visible show <recent|frequent|all> [--refresh]");
                    var showTarget = ParseQuickAccess(args[2], allowAll: true);
                    var showOptions = ParseVisibilityOptions(args, 3);
                    manager.ShowSection(showTarget, showOptions);
                    PrintVisibilityMutationResult("shown", showTarget, showOptions);
                    break;
                case "hide":
                    RequireMinArgs(args, 3, "visible hide <recent|frequent|all> [--refresh]");
                    var hideTarget = ParseQuickAccess(args[2], allowAll: true);
                    var hideOptions = ParseVisibilityOptions(args, 3);
                    manager.HideSection(hideTarget, hideOptions);
                    PrintVisibilityMutationResult("hidden", hideTarget, hideOptions);
                    break;
                case "get-recent":
                    Console.WriteLine(manager.IsVisible(QuickAccess.RecentFiles));
                    break;
                case "get-frequent":
                    Console.WriteLine(manager.IsVisible(QuickAccess.FrequentFolders));
                    break;
                case "set-recent":
                    RequireMinArgs(args, 3, "visible set-recent <true|false> [--refresh]");
                    var recentOptions = ParseVisibilityOptions(args, 3);
                    manager.SetVisible(QuickAccess.RecentFiles, ParseBool(args[2]), recentOptions);
                    PrintVisibilityMutationResult("updated recent visibility", QuickAccess.RecentFiles, recentOptions);
                    break;
                case "set-frequent":
                    RequireMinArgs(args, 3, "visible set-frequent <true|false> [--refresh]");
                    var frequentOptions = ParseVisibilityOptions(args, 3);
                    manager.SetVisible(QuickAccess.FrequentFolders, ParseBool(args[2]), frequentOptions);
                    PrintVisibilityMutationResult("updated frequent visibility", QuickAccess.FrequentFolders, frequentOptions);
                    break;
                default:
                    throw new ArgumentException($"unknown visible command: {args[1]}");
            }
        }

        #endregion

        #region Commands — destlist

        private static void CmdDest(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 2, "dest <command>");

            switch (args[1])
            {
                case "parse":
                    CmdDestParse(manager, args);
                    break;
                case "parse-bytes":
                    CmdDestParseBytes(args);
                    break;
                case "manager":
                    CmdDestManager(manager, args);
                    break;
                case "filetime":
                    RequireMinArgs(args, 3, "dest filetime <value>");
                    PrintFileTime(ParseUInt64(args[2], "filetime"));
                    break;
                case "remove":
                    CmdDestRemove(useEntries: false, manager, args);
                    break;
                case "remove-entries":
                    CmdDestRemove(useEntries: true, manager, args);
                    break;
                default:
                    throw new ArgumentException($"unknown dest command: {args[1]}");
            }
        }

        private static void CmdDestParse(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "dest parse <recent|frequent> [--limit N]");
            var limit = ParseLimit(args, 2, 20);

            IReadOnlyList<DestListEntry> entries;
            switch (args[2])
            {
                case "recent":
                    entries = manager.GetRecentFilesMetadata();
                    break;
                case "frequent":
                    entries = manager.GetFrequentFoldersMetadata();
                    break;
                default:
                    throw new ArgumentException($"dest parse target must be recent or frequent (file/parse-bytes via public API not available): {args[2]}");
            }

            PrintDestEntries(entries, limit);
        }

        private static void CmdDestParseBytes(string[] args)
        {
            throw new NotSupportedException(
                "dest parse-bytes is not available via the public Wincent API. The low-level DestList parser is internal.");
        }

        private static void CmdDestManager(QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "dest manager <recent|frequent> [--limit N]");
            var limit = ParseLimit(args, 2, 20);

            IReadOnlyList<DestListEntry> entries;
            switch (args[2])
            {
                case "recent":
                    entries = manager.GetRecentFilesMetadata();
                    break;
                case "frequent":
                    entries = manager.GetFrequentFoldersMetadata();
                    break;
                default:
                    throw new ArgumentException($"unknown metadata target: {args[2]}");
            }

            PrintDestEntries(entries, limit);
        }

        private static void CmdDestRemove(bool useEntries, QuickAccessManager manager, string[] args)
        {
            RequireMinArgs(args, 3, "dest remove|remove-entries <recent|frequent> [--delay-ms N] <path>...");
            var kind = ParseDestKind(args[2]);
            var delay = TimeSpan.FromMilliseconds(500);
            var paths = new List<string>();
            for (int i = 3; i < args.Length; i++)
            {
                if (args[i] == "--delay-ms")
                    delay = TimeSpan.FromMilliseconds(ParseInt(RequireNextArg(args, ref i, "delay-ms"), "delay-ms"));
                else
                    paths.Add(args[i]);
            }

            if (paths.Count == 0)
                throw new ArgumentException("dest remove requires at least one target path");

            var options = new ExperimentalRemoveOptions { RebuildDelay = delay };
            ExperimentalRemoveReport report;
            if (useEntries)
            {
                IReadOnlyList<DestListEntry> entries = kind == AutomaticDestinationsKind.RecentFiles
                    ? manager.GetRecentFilesMetadata()
                    : manager.GetFrequentFoldersMetadata();
                var matching = entries
                    .Where(e => paths.Any(p => string.Equals(e.Path, p, StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                report = ExperimentalDestListRemoval.RemoveEntriesByRebuild(kind, matching, options);
            }
            else
            {
                report = ExperimentalDestListRemoval.RemoveEntryPathsByRebuild(kind, paths, options);
            }

            PrintRemoveReport(report);
        }

        private static void PrintDestEntries(IReadOnlyList<DestListEntry> entries, int limit)
        {
            Console.WriteLine("entries: {0}", entries.Count);
            foreach (var e in entries.Take(limit))
            {
                Console.WriteLine(
                    "entry offset={0} len={1} id={2} number={3} number_unknown={4} " +
                    "stream={5} pinned={6} pin_status={7} pin_order={8} rank={9} recent_rank={10} " +
                    "count={11} access_count={12} score={13} last_access={14} " +
                    "last_interaction={15} sps_size={16} raw_path={17} path={18}",
                    e.EntryOffset, e.EntryLength, e.EntryId, e.EntryNumber,
                    e.EntryNumberReserved, e.StreamName, e.IsPinned, e.PinStatus,
                    e.PinOrder, e.Rank, e.RecentRank, 0, e.AccessCount, e.Score,
                    e.LastAccessTime, e.LastInteractionTime,
                    e.SerializedPropertyStoreSize,
                    e.RawPath, e.Path);
            }
        }

        private static void PrintRemoveReport(ExperimentalRemoveReport report)
        {
            Console.WriteLine("kind: {0}", report.Kind);
            Console.WriteLine("recent_folder: {0}", report.RecentFolder);
            Console.WriteLine("dest_path: {0}", report.DestinationPath);
            PrintRebuildObservation(report.Kind);
            PrintStringList("requested_paths", report.RequestedPaths);
            PrintStringList("matching_paths_before", report.MatchingPathsBefore);
            Console.WriteLine("deleted_lnk_paths: {0}", report.DeletedShortcutPaths.Count);
            foreach (var p in report.DeletedShortcutPaths)
                Console.WriteLine(p);
            PrintStringList("missing_lnk_target_paths", report.MissingShortcutTargetPaths);
            Console.WriteLine("dest_deleted: {0}", report.DestinationDeleted);
            Console.WriteLine("rebuilt: {0}", report.Rebuilt);
            Console.WriteLine("rebuild_parse_elapsed: {0}", report.RebuildParseElapsed);
            Console.WriteLine("rebuild_parse_error: {0}", report.RebuildParseError);
            PrintStringList("remaining_paths_after_rebuild", report.RemainingPathsAfterRebuild);
            Console.WriteLine("success: {0}", report.Success);
        }

        private static void PrintRebuildObservation(AutomaticDestinationsKind kind)
        {
            switch (kind)
            {
                case AutomaticDestinationsKind.FrequentFolders:
                    Console.WriteLine(
                        "rebuild_note: frequent backing file rebuild can reset folder pins to Desktop, Downloads, " +
                        "Documents, and Pictures; Explorer leaves Windows Recent folder .lnk files in place.");
                    break;
                case AutomaticDestinationsKind.RecentFiles:
                    Console.WriteLine(
                        "rebuild_note: recent backing file rebuild uses Windows Recent file .lnk files, " +
                        "so stale shortcuts can make files reappear.");
                    break;
            }
        }

        private static void PrintFileTime(ulong filetime)
        {
            try
            {
                var dt = DateTime.FromFileTimeUtc((long)filetime);
                Console.WriteLine(dt.ToString("o"));
            }
            catch (ArgumentOutOfRangeException)
            {
                Console.WriteLine("<before unix epoch or out of range>");
            }
        }

        private static int ParseLimit(string[] args, int startIndex, int defaultValue)
        {
            for (int i = startIndex; i < args.Length; i++)
            {
                if (args[i] == "--limit" && i + 1 < args.Length)
                    return ParseInt(args[i + 1], "limit");
            }
            return defaultValue;
        }

        #endregion

        #region Help

        private static void PrintHelp()
        {
            Console.WriteLine(@"wincent example CLI

Usage:
  TestConsole.exe
  TestConsole.exe [--timeout-ms N] [--timeout-secs N] <command> [args]

Interactive:
  help
  exit
  quit

Core:
  features
  list <recent|frequent|all> [--paths]
  list-paths <recent|frequent|all>
  check <recent|frequent|all> <path>
  contains <recent|frequent|all> <keyword>
  add <recent|frequent> <path> [--refresh]
  remove <recent|frequent> <path> [--deep-clean]
  batch-add [--refresh] <recent:path|frequent:path>...
  batch-remove [--deep-clean] <recent:path|frequent:path>...
  lock [recent|frequent|all] [--cleanup-new-links]
  empty <recent|frequent|all> [--pinned] [--refresh]
  clear-cache

Utility APIs:
  retry <default|none|fast|standard|aggressive|custom> [--attempt N] [custom options]
  classify <stderr text>
  invalid-path <reason> [path]

Visible:
  visible get <recent|frequent|all>
  visible set <recent|frequent|all> <true|false> [--refresh]
  visible show <recent|frequent|all> [--refresh]
  visible hide <recent|frequent|all> [--refresh]
  visible get-recent | get-frequent
  visible set-recent <true|false> [--refresh]
  visible set-frequent <true|false> [--refresh]

  --refresh refreshes Explorer windows after a visible set/show/hide command.
  Frequent visibility affects only unpinned frequent folders; pinned folders stay pinned.

Destlist:
  dest parse <recent|frequent> [--limit N]
  dest manager <recent|frequent> [--limit N]
  dest filetime <value>
  dest remove <recent|frequent> [--delay-ms N] <path>...
  dest remove-entries <recent|frequent> [--delay-ms N] <path>...
");
        }

        #endregion

        #region Argument parsing

        private static (TimeSpan timeout, string[] remaining) ParseGlobalOptions(string[] args)
        {
            var timeout = TimeSpan.FromSeconds(10);
            var rest = new List<string>();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--timeout-ms":
                        timeout = TimeSpan.FromMilliseconds(ParseInt(RequireNextArg(args, ref i, "timeout-ms"), "timeout-ms"));
                        break;
                    case "--timeout-secs":
                        timeout = TimeSpan.FromSeconds(ParseInt(RequireNextArg(args, ref i, "timeout-secs"), "timeout-secs"));
                        break;
                    default:
                        rest.Add(args[i]);
                        break;
                }
            }

            return (timeout, rest.ToArray());
        }

        private static string[] SplitCommandLine(string line)
        {
            var args = new List<string>();
            var current = new System.Text.StringBuilder();
            char? quote = null;

            for (int i = 0; i < line.Length; i++)
            {
                var ch = line[i];

                if (quote.HasValue)
                {
                    if (ch == quote.Value)
                    {
                        quote = null;
                    }
                    else if (ch == '\\' && i + 1 < line.Length)
                    {
                        // Windows-style: inside double quotes, only \" is an escape.
                        // Inside single quotes, \ escapes any next char (shell style).
                        var next = line[i + 1];
                        if (quote.Value == '"' && (next == '"' || next == '\\'))
                        {
                            i++;
                            current.Append(next);
                        }
                        else if (quote.Value == '\'')
                        {
                            i++;
                            current.Append(next);
                        }
                        else
                        {
                            // Inside double quotes with non-special next char:
                            // keep both the backslash and the next character.
                            current.Append(ch);
                        }
                    }
                    else
                    {
                        current.Append(ch);
                    }
                }
                else if (ch == '"' || ch == '\'')
                {
                    quote = ch;
                }
                else if (char.IsWhiteSpace(ch))
                {
                    if (current.Length > 0)
                    {
                        args.Add(current.ToString());
                        current.Clear();
                    }
                }
                else
                {
                    current.Append(ch);
                }
            }

            if (quote.HasValue)
                throw new ArgumentException($"unterminated quote: {quote.Value}");

            if (current.Length > 0)
                args.Add(current.ToString());

            return args.ToArray();
        }

        private static string RequireNextArg(string[] args, ref int index, string name)
        {
            index++;
            if (index >= args.Length)
                throw new ArgumentException($"{name} requires a value");
            return args[index];
        }

        private static QuickAccess ParseQuickAccess(string value, bool allowAll)
        {
            switch (value)
            {
                case "recent": case "recent-files": case "files": return QuickAccess.RecentFiles;
                case "frequent": case "frequent-folders": case "folders": return QuickAccess.FrequentFolders;
                case "all" when allowAll: return QuickAccess.All;
                case "all":
                    throw new ArgumentException("QuickAccess.All is not valid for this operation");
                default:
                    throw new ArgumentException($"unknown Quick Access category: {value}");
            }
        }

        private static QuickAccessLockTarget ParseLockTarget(string value)
        {
            switch (value)
            {
                case "recent": case "recent-files": case "files": return QuickAccessLockTarget.RecentFiles;
                case "frequent": case "frequent-folders": case "folders": return QuickAccessLockTarget.FrequentFolders;
                case "all": return QuickAccessLockTarget.All;
                default: throw new ArgumentException($"unknown lock target: {value}");
            }
        }

        private static AutomaticDestinationsKind ParseDestKind(string value)
        {
            switch (value)
            {
                case "recent": case "recent-files": case "files": return AutomaticDestinationsKind.RecentFiles;
                case "frequent": case "frequent-folders": case "folders": return AutomaticDestinationsKind.FrequentFolders;
                default: throw new ArgumentException($"unknown AutomaticDestinations kind: {value}");
            }
        }

        private static List<QuickAccessItem> ParseBatchItems(string[] itemArgs)
        {
            var items = new List<QuickAccessItem>();
            foreach (var item in itemArgs)
            {
                var parts = item.Split(new[] { ':' }, 2);
                if (parts.Length != 2)
                    throw new ArgumentException($"batch item must be recent:path or frequent:path: {item}");

                switch (parts[0])
                {
                    case "recent": case "file": case "files":
                        items.Add(QuickAccessItem.RecentFile(parts[1]));
                        break;
                    case "frequent": case "folder": case "folders":
                        items.Add(QuickAccessItem.FrequentFolder(parts[1]));
                        break;
                    default:
                        throw new ArgumentException($"unknown batch item type: {parts[0]}");
                }
            }
            return items;
        }

        private static (bool flag, string[] remaining) SplitFlag(string[] args, int startIndex, string flag)
        {
            var flagValue = args.Skip(startIndex).Any(a => a == flag);
            var rest = args.Where((a, i) => i < startIndex || a != flag).ToArray();
            return (flagValue, rest);
        }

        private static VisibilityOptions ParseVisibilityOptions(string[] args, int startIndex)
        {
            var options = new VisibilityOptions();
            for (int i = startIndex; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--refresh":
                        options.RefreshExplorer = true;
                        break;
                    default:
                        throw new ArgumentException($"unknown visible option: {args[i]}");
                }
            }

            return options;
        }

        private static void PrintVisibilityMutationResult(
            string message,
            QuickAccess target,
            VisibilityOptions options)
        {
            Console.WriteLine(message);
            Console.WriteLine("explorer_refreshed: {0}", options.RefreshExplorer);
            if (target == QuickAccess.FrequentFolders || target == QuickAccess.All)
                Console.WriteLine("note: pinned Quick Access folders are not affected by frequent visibility");
        }

        private static bool ParseBool(string value)
        {
            switch (value)
            {
                case "true": case "1": case "yes": case "on": case "show": return true;
                case "false": case "0": case "no": case "off": case "hide": return false;
                default: throw new ArgumentException($"invalid bool value: {value}");
            }
        }

        private static int ParseInt(string value, string name)
        {
            if (int.TryParse(value, out var result))
                return result;
            throw new ArgumentException($"invalid {name}: {value}");
        }

        private static double ParseDouble(string value, string name)
        {
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
                return result;
            throw new ArgumentException($"invalid {name}: {value}");
        }

        private static ulong ParseUInt64(string value, string name)
        {
            if (ulong.TryParse(value, out var result))
                return result;
            throw new ArgumentException($"invalid {name}: {value}");
        }

        private static void RequireMinArgs(string[] args, int min, string usage)
        {
            if (args.Length < min)
                throw new ArgumentException($"usage: {usage}");
        }

        #endregion

        #region Output helpers

        private static void PrintStringList(string label, IEnumerable<string> values)
        {
            Console.WriteLine("{0}: {1}", label, values.Count());
            foreach (var value in values)
                Console.WriteLine(value);
        }

        private static void PrintBatchResult(BatchResult result)
        {
            Console.WriteLine("total: {0}", result.Total);
            Console.WriteLine("succeeded: {0}", result.Succeeded.Count);
            Console.WriteLine("failed: {0}", result.Failed.Count);
            Console.WriteLine("success_rate: {0:F1}%", result.SuccessRate * 100.0);
            Console.WriteLine("complete_success: {0}", result.IsCompleteSuccess);
            Console.WriteLine("partial_success: {0}", result.HasPartialSuccess);

            foreach (var item in result.Succeeded)
                Console.WriteLine("ok: {0}", item.Path);
            foreach (var failure in result.Failed)
                Console.WriteLine("failed: {0}: {1}", failure.Item.Path, failure.Error.Message);
        }

        private static string CategoryName(QuickAccess qa)
        {
            switch (qa)
            {
                case QuickAccess.RecentFiles: return "Recent Files";
                case QuickAccess.FrequentFolders: return "Frequent Folders";
                case QuickAccess.All: return "All";
                default: return "Unknown";
            }
        }

        private static void PrintError(Exception error)
        {
            Console.Error.WriteLine("error: {0}", error.Message);

            if (error is PowerShellExecutionException psee)
            {
                Console.Error.WriteLine("powershell.kind: {0}", psee.Kind);
                Console.Error.WriteLine("powershell.operation: {0}", psee.Operation);
                Console.Error.WriteLine("powershell.exit_code: {0}", psee.ExitCode);
                Console.Error.WriteLine("powershell.script_path: {0}", psee.ScriptPath);
                Console.Error.WriteLine("powershell.parameters: {0}", psee.Parameters);
                Console.Error.WriteLine("powershell.duration: {0}", psee.Duration);
                Console.Error.WriteLine("powershell.native_error_code: {0}", psee.NativeErrorCode);
                Console.Error.WriteLine("powershell.stdout: {0}", psee.StandardOutput);
                Console.Error.WriteLine("powershell.stderr: {0}", psee.StandardError);
            }
            else if (error is QuickAccessItemAlreadyExistsException already)
            {
                Console.Error.WriteLine("already_exists.path: {0}", already.Path);
                Console.Error.WriteLine("already_exists.qa_type: {0}", already.Target);
            }
            else if (error is QuickAccessItemNotFoundException notFound)
            {
                Console.Error.WriteLine("not_found.path: {0}", notFound.Path);
                Console.Error.WriteLine("not_found.qa_type: {0}", notFound.Target);
            }
            else if (error is PartialClearException partial)
            {
                Console.Error.WriteLine("partial_clear.recent: {0}", partial.RecentFilesCleared);
                Console.Error.WriteLine("partial_clear.frequent: {0}", partial.FrequentFoldersCleared);
                if (partial.InnerException != null)
                    Console.Error.WriteLine("partial_clear.source: {0}", partial.InnerException.Message);
            }
            else if (error is UnsupportedQuickAccessOperationException unsupported)
            {
                Console.Error.WriteLine("unsupported.target: {0}", unsupported.Target);
                Console.Error.WriteLine("unsupported.operation: {0}", unsupported.Operation);
            }
            else if (error is DestListUnsupportedVersionException versionEx)
            {
                Console.Error.WriteLine("destlist.version: {0}", versionEx.Version);
                Console.Error.WriteLine("destlist.file: {0}", versionEx.FilePath);
                Console.Error.WriteLine("destlist.offset: {0}", versionEx.Offset);
                Console.Error.WriteLine("destlist.details: {0}", versionEx.Details);
            }
            else if (error is DestListParseException destEx)
            {
                Console.Error.WriteLine("destlist.file: {0}", destEx.FilePath);
                Console.Error.WriteLine("destlist.offset: {0}", destEx.Offset);
                Console.Error.WriteLine("destlist.details: {0}", destEx.Details);
            }
        }

        private static PowerShellErrorKind InferKindFromStderr(string stderr)
        {
            var lower = stderr.ToLowerInvariant();
            if (lower.Contains("access denied") || lower.Contains("拒绝访问"))
                return PowerShellErrorKind.AccessDenied;
            if (lower.Contains("timed out") || lower.Contains("timeout"))
                return PowerShellErrorKind.Timeout;
            if (lower.Contains("execution policy") || lower.Contains("executionpolicy"))
                return PowerShellErrorKind.ExecutionPolicy;
            if (lower.Contains("not recognized") || lower.Contains("not found"))
                return PowerShellErrorKind.CmdletNotFound;
            return PowerShellErrorKind.ProcessFailed;
        }

        private static PowerShellErrorKind InferKindCustom(string stderr)
        {
            var lower = stderr.ToLowerInvariant();
            if (lower.Contains("access denied") || lower.Contains("拒绝访问"))
                return PowerShellErrorKind.AccessDenied;
            if (lower.Contains("timed out") || lower.Contains("timeout"))
                return PowerShellErrorKind.Timeout;
            return PowerShellErrorKind.ProcessFailed;
        }

        #endregion
    }
}

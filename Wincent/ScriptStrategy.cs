using System;
using System.Collections.Generic;

namespace Wincent
{
    internal enum PSScript
    {
        RefreshExplorer,
        QueryQuickAccess,
        QueryRecentFile,
        QueryFrequentFolder,
        AddRecentFile,
        RemoveRecentFile,
        PinToFrequentFolder,
        UnpinFromFrequentFolder,
        EmptyPinnedFolders,
    }

    internal static class PSScriptExtensions
    {
        public static PowerShellOperation ToPowerShellOperation(this PSScript script)
        {
            switch (script)
            {
                case PSScript.RefreshExplorer:
                    return PowerShellOperation.RefreshExplorer;
                case PSScript.QueryQuickAccess:
                    return PowerShellOperation.QueryQuickAccess;
                case PSScript.QueryRecentFile:
                    return PowerShellOperation.QueryRecentFiles;
                case PSScript.QueryFrequentFolder:
                    return PowerShellOperation.QueryFrequentFolders;
                case PSScript.AddRecentFile:
                    return PowerShellOperation.AddRecentFile;
                case PSScript.RemoveRecentFile:
                    return PowerShellOperation.RemoveRecentFile;
                case PSScript.PinToFrequentFolder:
                    return PowerShellOperation.PinFrequentFolder;
                case PSScript.UnpinFromFrequentFolder:
                    return PowerShellOperation.UnpinFrequentFolder;
                case PSScript.EmptyPinnedFolders:
                    return PowerShellOperation.ClearPinnedFolders;
                default:
                    throw new NotSupportedException($"Unsupported script type: {script}");
            }
        }
    }

    internal interface IPSScriptStrategy
    {
        string GenerateScript(string parameter);
    }

    internal abstract class PSScriptStrategyBase : IPSScriptStrategy
    {
        internal const string AlreadyExistsSentinel = "WINCENT_ALREADY_EXISTS";
        internal const string NotInQuickAccessSentinel = "WINCENT_NOT_IN_QUICK_ACCESS";

        protected const string EncodingSetup = @"
            $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        ";

        protected const string ShellApplicationSetup = @"
            $shellApplication = New-Object -ComObject Shell.Application;
        ";

        protected const string NormalizePathFunction = @"
            function Normalize-WincentPath($p) {
                if ($null -eq $p) {
                    return ''
                }

                $n = $p.ToLower().Replace('/', '\')
                if ($n.Length -gt 3 -and $n.EndsWith('\')) {
                    $n = $n.TrimEnd('\')
                }

                return $n
            }
        ";

        protected static string EscapePowerShellString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            return input.Replace("'", "''");
        }

        public abstract string GenerateScript(string parameter);
    }

    internal static class ShellNamespaces
    {
        public const string QuickAccess = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
        public const string FrequentFolders = "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";
    }

    internal class RefreshExplorerStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $windows = $shellApplication.Windows();
            $windows | ForEach-Object {{ $_.Refresh() }}
        ";
    }

    internal class QueryRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                where {{ $_.IsFolder -eq $false }} | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    internal class QueryFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    internal class QueryQuickAccessStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    internal class AddRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            parameter = EscapePowerShellString(parameter);

            // Recent-file additions use the native SHAddToRecentDocs path in QuickAccessManager.
            // This legacy script shape is retained for existing script-generation tests.
            return $@"
                {EncodingSetup}
                {ShellApplicationSetup}
                Write-Output '{parameter}'
            ";
        }
    }

    internal class RemoveRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            parameter = EscapePowerShellString(parameter);

            return $@"
                {EncodingSetup}
                {ShellApplicationSetup}
                {NormalizePathFunction}
                $requestedPath = '{parameter}';
                $files = $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                          where {{ $_.IsFolder -eq $false }};
                $target = $files |
                    Where-Object {{ (Normalize-WincentPath $_.Path) -eq (Normalize-WincentPath $requestedPath) }} |
                    Select-Object -First 1;
                if ($null -eq $target) {{
                    Write-Output '{NotInQuickAccessSentinel}';
                    exit 1;
                }}

                $target.InvokeVerb('remove');
            ";
        }
    }

    internal class PinToFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            parameter = EscapePowerShellString(parameter);

            return $@"
                {EncodingSetup}
                {ShellApplicationSetup}
                {NormalizePathFunction}
                $requestedPath = '{parameter}';
                $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items() |
                    where {{ $_.IsFolder -eq $true }};
                $target = $folders |
                    Where-Object {{ (Normalize-WincentPath $_.Path) -eq (Normalize-WincentPath $requestedPath) }} |
                    Select-Object -First 1;
                if ($null -ne $target) {{
                    Write-Output '{AlreadyExistsSentinel}';
                    exit 1;
                }}

                $shellApplication.Namespace($requestedPath).Self.InvokeVerb('pintohome');
            ";
        }
    }

    internal class UnpinFromFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            parameter = EscapePowerShellString(parameter);

            return $@"
                {EncodingSetup}
                {ShellApplicationSetup}
                {NormalizePathFunction}
                $requestedPath = '{parameter}';

                function Find-WincentFrequentFolder {{
                    $finderShell = New-Object -ComObject Shell.Application;
                    $folder = $finderShell.Namespace('{ShellNamespaces.FrequentFolders}');
                    if ($null -eq $folder) {{
                        throw 'Failed to open Frequent Folders namespace'
                    }}

                    $folder.Items() |
                        Where-Object {{ (Normalize-WincentPath $_.Path) -eq (Normalize-WincentPath $requestedPath) }} |
                        Select-Object -First 1
                }}

                function Wait-WincentFrequentFolderPresence($expected) {{
                    $deadline = (Get-Date).AddMilliseconds(1000)
                    while ((Get-Date) -lt $deadline) {{
                        $exists = $null -ne (Find-WincentFrequentFolder)
                        if ($exists -eq $expected) {{
                            return $true
                        }}

                        $remaining = [int][Math]::Ceiling(($deadline - (Get-Date)).TotalMilliseconds)
                        if ($remaining -le 0) {{
                            break
                        }}
                        Start-Sleep -Milliseconds ([Math]::Min(100, $remaining))
                    }}

                    $exists = $null -ne (Find-WincentFrequentFolder)
                    return ($exists -eq $expected)
                }}

                function Invoke-WincentUnpinFromHome($item) {{
                    try {{
                        $item.InvokeVerb('unpinfromhome');
                    }} catch {{}}
                }}

                function Invoke-WincentPinToHomeToggle {{
                    $shellApplication.Namespace($requestedPath).Self.InvokeVerb('pintohome')
                }}

                $target = Find-WincentFrequentFolder;
                if ($null -eq $target) {{
                    Write-Output '{NotInQuickAccessSentinel}';
                    exit 1;
                }}

                Invoke-WincentUnpinFromHome $target;
                if (Wait-WincentFrequentFolderPresence $false) {{
                    return
                }}

                Invoke-WincentPinToHomeToggle;
                if (Wait-WincentFrequentFolderPresence $false) {{
                    return
                }}

                $target = Find-WincentFrequentFolder;
                if ($null -eq $target) {{
                    return
                }}

                $isWin11 = (Get-CimInstance -Class Win32_OperatingSystem).Caption -Match 'Windows 11'
                if ($isWin11)
                {{
                    Invoke-WincentPinToHomeToggle
                }}
                else
                {{
                    Invoke-WincentUnpinFromHome $target;
                }}

                if (Wait-WincentFrequentFolderPresence $false) {{
                    return
                }}

                throw ""Failed to remove frequent folder: $requestedPath""
            ";
        }
    }

    internal class EmptyPinnedFoldersStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
            $folders | ForEach-Object {{ $_.InvokeVerb('unpinfromhome') }}
        ";
    }

    internal interface IPSScriptStrategyFactory
    {
        IPSScriptStrategy GetStrategy(PSScript method);
    }

    internal class DefaultPSScriptStrategyFactory : IPSScriptStrategyFactory
    {
        private static readonly IReadOnlyDictionary<PSScript, Func<IPSScriptStrategy>> _strategyMap =
            new Dictionary<PSScript, Func<IPSScriptStrategy>>
            {
                [PSScript.RefreshExplorer] = () => new RefreshExplorerStrategy(),
                [PSScript.QueryRecentFile] = () => new QueryRecentFileStrategy(),
                [PSScript.QueryFrequentFolder] = () => new QueryFrequentFolderStrategy(),
                [PSScript.QueryQuickAccess] = () => new QueryQuickAccessStrategy(),
                [PSScript.AddRecentFile] = () => new AddRecentFileStrategy(),
                [PSScript.RemoveRecentFile] = () => new RemoveRecentFileStrategy(),
                [PSScript.PinToFrequentFolder] = () => new PinToFrequentFolderStrategy(),
                [PSScript.UnpinFromFrequentFolder] = () => new UnpinFromFrequentFolderStrategy(),
                [PSScript.EmptyPinnedFolders] = () => new EmptyPinnedFoldersStrategy(),
            };

        public IPSScriptStrategy GetStrategy(PSScript method)
        {
            if (_strategyMap.TryGetValue(method, out var strategyCreator))
                return strategyCreator();

            throw new NotSupportedException($"Unsupported script type: {method}");
        }
    }

    internal class ScriptGenerationException : Exception
    {
        public ScriptGenerationException(string message, Exception inner)
            : base(message, inner) { }
    }
}

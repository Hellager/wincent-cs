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
        protected const string EncodingSetup = @"
            $OutputEncoding = [Console]::OutputEncoding = [System.Text.Encoding]::UTF8;
        ";

        protected const string ShellApplicationSetup = @"
            $shellApplication = New-Object -ComObject Shell.Application;
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
                $files = $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                         where {{ $_.IsFolder -eq $false }};
                $target = $files | where {{ $_.Path -eq '{parameter}' }};
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
                $shellApplication.Namespace('{parameter}').Self.InvokeVerb('pintohome');
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
                $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                $target = $folders | where {{ $_.Path -eq '{parameter}' }};
                if ($null -eq $target) {{ throw 'Target path not found in Frequent Folders namespace: {parameter}' }}

                $target.InvokeVerb('unpinfromhome');
                Start-Sleep -Milliseconds 1000;

                $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                $target = $folders | where {{ $_.Path -eq '{parameter}' }};
                if ($null -eq $target) {{ return }}

                $shellApplication.Namespace('{parameter}').Self.InvokeVerb('pintohome');
                Start-Sleep -Milliseconds 1000;

                $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                $target = $folders | where {{ $_.Path -eq '{parameter}' }};
                if ($null -eq $target) {{ return }}

                $isWin11 = (Get-CimInstance -Class Win32_OperatingSystem).Caption -Match 'Windows 11'
                if ($isWin11)
                {{
                    $shellApplication.Namespace('{parameter}').Self.InvokeVerb('pintohome')
                }}
                else
                {{
                    $target.InvokeVerb('unpinfromhome');
                }}

                Start-Sleep -Milliseconds 1000;

                $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                $target = $folders | where {{ $_.Path -eq '{parameter}' }};
                if ($null -ne $target) {{ throw 'Failed to remove frequent folder: {parameter}' }}
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

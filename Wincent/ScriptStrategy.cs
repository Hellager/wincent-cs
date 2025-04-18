using System;
using System.IO;
using System.Collections.Generic;

namespace Wincent
{
    public enum PSScript
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
        CheckQueryFeasible,
        CheckPinUnpinFeasible,
    }

    public interface IPSScriptStrategy
    {
        string GenerateScript(string parameter);
    }

    public abstract class PSScriptStrategyBase : IPSScriptStrategy
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

        protected static string CreateTimeoutWrapper(string scriptBlock, int timeoutSeconds = 5)
        {
            return $@"
            $timeout = {timeoutSeconds}

            $scriptBlock = {{
                {scriptBlock}
            }}.ToString()

            $arguments = ""-Command & {{$scriptBlock}}""
            $process = Start-Process powershell -ArgumentList $arguments -NoNewWindow -PassThru

            if (-not $process.WaitForExit($timeout * 1000)) {{
                try {{
                    $process.Kill()
                    Write-Error ""Process execution timed out (${{timeout}}s), forcefully terminated""
                    exit 1
                }}
                catch {{
                    Write-Error ""Error occurred while terminating process: $_""
                    exit 1
                }}
            }}
            ";
        }

        protected static string CreateTimeoutWrapperWithParam(string scriptBlock, string paramName, string paramValue, int timeoutSeconds = 5)
        {
            return $@"
            $timeout = {timeoutSeconds}

            $scriptBlock = {{
                param(${paramName})
                {scriptBlock}
            }}.ToString()

            $arguments = ""-Command & {{$scriptBlock}} -{paramName} '{EscapePowerShellString(paramValue)}'""
            $process = Start-Process powershell -ArgumentList $arguments -NoNewWindow -PassThru

            if (-not $process.WaitForExit($timeout * 1000)) {{
                try {{
                    $process.Kill()
                    Write-Error ""Process execution timed out (${{timeout}}s), forcefully terminated""
                    exit 1
                }}
                catch {{
                    Write-Error ""Error occurred while terminating process: $_""
                    exit 1
                }}
            }}
            ";
        }

        public abstract string GenerateScript(string parameter);
    }

    public static class ShellNamespaces
    {
        public const string QuickAccess = "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}";
        public const string FrequentFolders = "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}";
    }

    public class RefreshExplorerStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $windows = $shellApplication.Windows();
            $windows | ForEach-Object {{ $_.Refresh() }}
        ";
    }

    public class QueryRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                where {{ $_.IsFolder -eq $false }} | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class QueryFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class QueryQuickAccessStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class CheckQueryFeasibleStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {CreateTimeoutWrapper($@"
                {ShellApplicationSetup}
                $shellApplication.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                    ForEach-Object {{ $_.Path }}
            ")}
        ";
    }

    public class CheckPinUnpinFeasibleStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            $currentPath = $PSScriptRoot
            {CreateTimeoutWrapperWithParam($@"
                {ShellApplicationSetup}
                $shellApplication.Namespace($scriptPath).Self.InvokeVerb('pintohome')

                Start-Sleep -Seconds 3

                $isWin11 = (Get-CimInstance -Class Win32_OperatingSystem).Caption -Match ""Windows 11""
                if ($isWin11) 
                {{
                    $shellApplication.Namespace($scriptPath).Self.InvokeVerb('pintohome')
                }}
                else
                {{
                    $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                    $target = $folders | Where-Object {{$_.Path -eq $scriptPath}};
                    $target.InvokeVerb('unpinfromhome');
                }}
            ", "scriptPath", "$currentPath", 10)}
        ";
    }

    public class AddRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            parameter = EscapePowerShellString(parameter);

            return $@"
                {EncodingSetup}
                {ShellApplicationSetup}
                Write-Output '{parameter}'
            ";
        }
    }

    public class RemoveRecentFileStrategy : PSScriptStrategyBase
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

    public class PinToFrequentFolderStrategy : PSScriptStrategyBase
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

    public class UnpinFromFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            parameter = EscapePowerShellString(parameter);

            return $@"
                {EncodingSetup}
                {ShellApplicationSetup}
                $shellApplication.Namespace($scriptPath).Self.InvokeVerb('pintohome')                

                $isWin11 = (Get-CimInstance -Class Win32_OperatingSystem).Caption -Match ""Windows 11""
                if ($isWin11)
                {{
                    $shellApplication.Namespace('{parameter}').Self.InvokeVerb('pintohome')
                }}
                else
                {{
                    $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                    $target = $folders | where {{ $_.Path -eq '{parameter}' }};
                    $target.InvokeVerb('unpinfromhome');
                }}
            ";
        }
    }   

    public class EmptyPinnedFoldersStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            {ShellApplicationSetup}
            $folders = $shellApplication.Namespace('{ShellNamespaces.FrequentFolders}').Items();
            $folders | ForEach-Object {{ $_.InvokeVerb('unpinfromhome') }}
        ";
    }

    public interface IPSScriptStrategyFactory
    {
        IPSScriptStrategy GetStrategy(PSScript method);
    }


    public class DefaultPSScriptStrategyFactory : IPSScriptStrategyFactory
    {
        private static readonly IReadOnlyDictionary<PSScript, Func<IPSScriptStrategy>> _strategyMap =
            new Dictionary<PSScript, Func<IPSScriptStrategy>>
            {
                [PSScript.RefreshExplorer] = () => new RefreshExplorerStrategy(),
                [PSScript.QueryRecentFile] = () => new QueryRecentFileStrategy(),
                [PSScript.QueryFrequentFolder] = () => new QueryFrequentFolderStrategy(),
                [PSScript.QueryQuickAccess] = () => new QueryQuickAccessStrategy(),
                [PSScript.CheckQueryFeasible] = () => new CheckQueryFeasibleStrategy(),
                [PSScript.CheckPinUnpinFeasible] = () => new CheckPinUnpinFeasibleStrategy(),
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

    public class ScriptGenerationException : Exception
    {
        public ScriptGenerationException(string message, Exception inner)
            : base(message, inner) { }
    }
}

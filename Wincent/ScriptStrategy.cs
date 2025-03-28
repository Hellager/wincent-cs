﻿using System;
using System.Collections.Generic;

namespace Wincent
{
    public enum PSScript
    {
        RefreshExplorer,
        QueryQuickAccess,
        QueryRecentFile,
        QueryFrequentFolder,
        RemoveRecentFile,
        PinToFrequentFolder,
        UnpinFromFrequentFolder,
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
            $shellApplication = New-Object -ComObject Shell.Application;
            $windows = $shellApplication.Windows();
            $windows | ForEach-Object {{ $_.Refresh() }}
        ";
    }

    public class QueryRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            $shell = New-Object -ComObject Shell.Application;
            $shell.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                where {{ $_.IsFolder -eq $false }} | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class QueryFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            $shell = New-Object -ComObject Shell.Application;
            $shell.Namespace('{ShellNamespaces.FrequentFolders}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class QueryQuickAccessStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            $shell = New-Object -ComObject Shell.Application;
            $shell.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class CheckQueryFeasibleStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            $shell = New-Object -ComObject Shell.Application
            $shell.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
                ForEach-Object {{ $_.Path }}
        ";
    }

    public class CheckPinUnpinFeasibleStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter) => $@"
            {EncodingSetup}
            $shell = New-Object -ComObject Shell.Application
            $shell.Namespace($PSScriptRoot).Self.InvokeVerb('pintohome')

            $folders = $shell.Namespace('{ShellNamespaces.FrequentFolders}').Items();
            $target = $folders | where {{ $_.Path -eq $PSScriptRoot }};
            $target.InvokeVerb('unpinfromhome');
        ";
    }

    public class RemoveRecentFileStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            return $@"
                {EncodingSetup}
                $shell = New-Object -ComObject Shell.Application;
                $files = $shell.Namespace('{ShellNamespaces.QuickAccess}').Items() | 
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

            return $@"
                {EncodingSetup}
                $shell = New-Object -ComObject Shell.Application;
                $shell.Namespace('{parameter}').Self.InvokeVerb('pintohome');
            ";
        }
    }

    public class UnpinFromFrequentFolderStrategy : PSScriptStrategyBase
    {
        public override string GenerateScript(string parameter)
        {
            if (string.IsNullOrWhiteSpace(parameter))
                throw new ArgumentException("Valid file path parameter required");

            return $@"
                {EncodingSetup}
                $shell = New-Object -ComObject Shell.Application;
                $folders = $shell.Namespace('{ShellNamespaces.FrequentFolders}').Items();
                $target = $folders | where {{ $_.Path -eq '{parameter}' }};
                $target.InvokeVerb('unpinfromhome');
            ";
        }
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
                [PSScript.RemoveRecentFile] = () => new RemoveRecentFileStrategy(),
                [PSScript.PinToFrequentFolder] = () => new PinToFrequentFolderStrategy(),
                [PSScript.UnpinFromFrequentFolder] = () => new UnpinFromFrequentFolderStrategy(),
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

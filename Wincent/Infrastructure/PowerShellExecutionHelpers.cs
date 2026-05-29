using System;
using System.ComponentModel;
using System.IO;
using System.Text;

namespace Wincent
{
    internal static class PowerShellErrorClassifier
    {
        public static PowerShellErrorKind InferKind(string stderr)
        {
            string value = (stderr ?? string.Empty).ToLowerInvariant();

            if ((value.Contains("access") && value.Contains("denied")) ||
                value.Contains("unauthorizedaccessexception"))
                return PowerShellErrorKind.AccessDenied;

            if (value.Contains("execution policy") || value.Contains("executionpolicy"))
                return PowerShellErrorKind.ExecutionPolicy;

            if (value.Contains("not recognized as") || value.Contains("commandnotfoundexception"))
                return PowerShellErrorKind.CmdletNotFound;

            if (value.Contains("timeout") || value.Contains("timed out"))
                return PowerShellErrorKind.Timeout;

            return PowerShellErrorKind.ProcessFailed;
        }

        public static PowerShellErrorKind InferKind(Exception exception, string stderr)
        {
            if (IsAccessDenied(exception))
                return PowerShellErrorKind.AccessDenied;

            return InferKind(stderr);
        }

        public static int? GetNativeErrorCode(Exception exception)
        {
            var win32Exception = exception as Win32Exception;
            if (win32Exception != null)
                return win32Exception.NativeErrorCode;

            int hResult = exception?.HResult ?? 0;
            int facility = (hResult >> 16) & 0x0FFF;
            if (facility != 7)
                return null;

            int win32Code = hResult & 0xFFFF;
            return win32Code == 0 ? (int?)null : win32Code;
        }

        private static bool IsAccessDenied(Exception exception)
        {
            if (exception is UnauthorizedAccessException)
                return true;

            var win32Exception = exception as Win32Exception;
            if (win32Exception != null && win32Exception.NativeErrorCode == 5)
                return true;

            return GetNativeErrorCode(exception) == 5;
        }
    }

    internal static class PowerShellCommandLine
    {
        public static string BuildArguments(string scriptPath, string parameter)
        {
            var arguments = new StringBuilder();
            arguments.Append("-NoProfile -NonInteractive -ExecutionPolicy Bypass -File ");
            arguments.Append(QuoteArgument(scriptPath));

            if (!string.IsNullOrEmpty(parameter))
            {
                arguments.Append(' ');
                arguments.Append(QuoteArgument(parameter));
            }

            return arguments.ToString();
        }

        public static string QuoteArgument(string argument)
        {
            if (argument == null)
                return "\"\"";

            var result = new StringBuilder();
            result.Append('"');

            int backslashes = 0;
            foreach (char c in argument)
            {
                if (c == '\\')
                {
                    backslashes++;
                    continue;
                }

                if (c == '"')
                {
                    result.Append('\\', backslashes * 2 + 1);
                    result.Append(c);
                    backslashes = 0;
                    continue;
                }

                result.Append('\\', backslashes);
                backslashes = 0;
                result.Append(c);
            }

            result.Append('\\', backslashes * 2);
            result.Append('"');
            return result.ToString();
        }
    }
}

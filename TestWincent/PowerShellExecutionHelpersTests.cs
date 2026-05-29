using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.ComponentModel;
using Wincent;

namespace TestWincent
{
    [TestClass]
    public class PowerShellExecutionHelpersTests
    {
        [TestMethod]
        public void InferKind_ClassifiesCommonStderrPatterns()
        {
            Assert.AreEqual(PowerShellErrorKind.AccessDenied, PowerShellErrorClassifier.InferKind("Access is denied"));
            Assert.AreEqual(PowerShellErrorKind.AccessDenied, PowerShellErrorClassifier.InferKind("UnauthorizedAccessException"));
            Assert.AreEqual(PowerShellErrorKind.ExecutionPolicy, PowerShellErrorClassifier.InferKind("running scripts is disabled by execution policy"));
            Assert.AreEqual(PowerShellErrorKind.CmdletNotFound, PowerShellErrorClassifier.InferKind("foo is not recognized as the name of a cmdlet"));
            Assert.AreEqual(PowerShellErrorKind.CmdletNotFound, PowerShellErrorClassifier.InferKind("CommandNotFoundException"));
            Assert.AreEqual(PowerShellErrorKind.Timeout, PowerShellErrorClassifier.InferKind("operation timed out"));
            Assert.AreEqual(PowerShellErrorKind.ProcessFailed, PowerShellErrorClassifier.InferKind("unknown failure"));
        }

        [TestMethod]
        public void InferKind_SystemEvidenceTakesPrecedenceOverStderr()
        {
            var exception = new Win32Exception(5);

            Assert.AreEqual(
                PowerShellErrorKind.AccessDenied,
                PowerShellErrorClassifier.InferKind(exception, "operation timed out"));
        }

        [TestMethod]
        public void InferKind_NonWin32HResultLowWordFive_DoesNotClassifyAsAccessDenied()
        {
            var exception = new NonWin32LowWordFiveException();

            Assert.AreEqual(
                PowerShellErrorKind.Timeout,
                PowerShellErrorClassifier.InferKind(exception, "operation timed out"));
        }

        [TestMethod]
        public void GetNativeErrorCode_OnlyExtractsFacilityWin32Codes()
        {
            Assert.AreEqual(5, PowerShellErrorClassifier.GetNativeErrorCode(new UnauthorizedAccessException()));
            Assert.IsNull(PowerShellErrorClassifier.GetNativeErrorCode(new InvalidOperationException()));
        }

        [TestMethod]
        public void QuoteArgument_EscapesQuotesAndTrailingBackslashes()
        {
            Assert.AreEqual(@"""C:\Path With Spaces\file.ps1""", PowerShellCommandLine.QuoteArgument(@"C:\Path With Spaces\file.ps1"));
            Assert.AreEqual(@"""C:\Path \""Quoted\""\\""", PowerShellCommandLine.QuoteArgument(@"C:\Path ""Quoted""\"));
            Assert.AreEqual(@"""C:\O'Brien\file.ps1""", PowerShellCommandLine.QuoteArgument(@"C:\O'Brien\file.ps1"));
        }

        [TestMethod]
        public void BuildArguments_QuotesScriptPathAndParameter()
        {
            var arguments = PowerShellCommandLine.BuildArguments(
                @"C:\Scripts\Test Script.ps1",
                @"C:\Input Path\O'Brien.txt");

            StringAssert.Contains(arguments, @"-File ""C:\Scripts\Test Script.ps1""");
            StringAssert.Contains(arguments, @"""C:\Input Path\O'Brien.txt""");
        }

        private sealed class NonWin32LowWordFiveException : Exception
        {
            public NonWin32LowWordFiveException()
            {
                HResult = unchecked((int)0x80040005);
            }
        }
    }
}

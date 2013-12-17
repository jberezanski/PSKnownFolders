using System;
using System.Linq;
using System.Management.Automation.Runspaces;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#if FX40
namespace BlaSoft.PowerShell.KnownFolders.Tests.Fx40
#else
namespace BlaSoft.PowerShell.KnownFolders.Tests
#endif
{
    using System.Management.Automation;

    [TestClass]
    public class BasicPowerShell : PowerShellTestBase
    {
        [TestMethod]
        public void Is_expected_version()
        {
            AssertPowerShellIsExpectedVersion();
        }

        [TestMethod]
        public void Can_run_single_command()
        {
            using (var psh = PowerShell.Create())
            {
                var output = psh.AddCommand("Get-Location").Invoke();
                Assert.IsNotNull(output);
                Assert.AreEqual(1, output.Count);
                var psObj = output.First();
                Assert.IsNotNull(psObj);
                var pathProp = psObj.Properties["Path"];
                Assert.IsNotNull(pathProp);
                var pathObj = pathProp.Value;
                Assert.IsNotNull(pathObj);
                Assert.IsInstanceOfType(pathObj, typeof(string));
                var pathStr = (string)pathObj;
                Assert.AreEqual(Environment.CurrentDirectory, pathStr);
            }
        }

        [TestMethod]
        [ExpectedException(typeof(CommandNotFoundException))]
        public void Cannot_run_standard_command_with_empty_session_state()
        {
            var sessionState = InitialSessionState.Create();
            using (var runspace = RunspaceFactory.CreateRunspace(sessionState))
            using (var psh = PowerShell.Create())
            {
                psh.Runspace = runspace;
                runspace.Open();
                psh.AddCommand("Get-Location").Invoke();
                Assert.Fail();
            }
        }
    }
}

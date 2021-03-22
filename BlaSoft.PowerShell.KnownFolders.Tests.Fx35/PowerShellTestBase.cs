using System;
using System.Collections;
using System.Linq;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#if FX40
namespace BlaSoft.PowerShell.KnownFolders.Tests.Fx40
#else
namespace BlaSoft.PowerShell.KnownFolders.Tests.Fx35
#endif
{
    using System.Management.Automation;

    public abstract class PowerShellTestBase
    {
        protected PowerShellTestBase()
        {
        }

        protected static void AssertPowerShellIsExpectedVersion()
        {
            using (var psh = PowerShell.Create())
            {
                var output = psh.AddScript("$PSVersionTable").Invoke();
                Assert.IsNotNull(output);
                Assert.AreEqual(1, output.Count);
                var psObj = output.First();
                Assert.IsNotNull(psObj);
                Assert.IsInstanceOfType(psObj.BaseObject, typeof(Hashtable));
                var table = (Hashtable)psObj.BaseObject;
                Assert.IsTrue(table.ContainsKey("PSVersion"));
                var psVersion = (Version)table["PSVersion"];
                Assert.IsTrue(table.ContainsKey("CLRVersion"));
                var clrVersion = (Version)table["CLRVersion"];
#if FX40
                Assert.AreEqual(4, clrVersion.Major, "Unexpected CLR version.");
                Assert.AreNotEqual(2, psVersion.Major, "Unexpected PowerShell version.");
                Assert.AreNotEqual(1, psVersion.Major, "Unexpected PowerShell version.");
#else
                Assert.AreEqual(2, clrVersion.Major, "This test project requires CLR 2.0. Older Visual Studio will use CLR 4 if running tests from .NET 3.5 and 4.0 projects together; to work around this, run tests from this project separately from other test projects. Modern Visual Studio does not support running tests under .NET 3.5 at all (https://github.com/Microsoft/vstest/issues/1896), sorry.");
                Assert.AreEqual(2, psVersion.Major, "Unexpected PowerShell version.");
#endif
            }
        }
    }
}

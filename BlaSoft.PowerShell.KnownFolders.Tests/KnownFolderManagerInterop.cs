using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders.Tests
{
    [TestClass]
    public class KnownFolderManagerInterop
    {
        [TestMethod]
        public void CanCreateKnownFolderManager()
        {
            KnownFolderManager obj = new KnownFolderManager();
            Assert.IsNotNull(obj);
        }

        [TestMethod]
        public void KnownFolderManager_implements_IKnownFolderManager()
        {
            KnownFolderManager obj = new KnownFolderManager();
            IKnownFolderManager knownFolderManager = (IKnownFolderManager)obj;
            Assert.IsNotNull(knownFolderManager);
        }
    }
}

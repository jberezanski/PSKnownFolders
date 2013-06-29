using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders.Tests
{
    [TestClass]
    public class KnownFolderManagerInterop
    {
        [TestMethod]
        public void Can_create_KnownFolderManager()
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

        [TestMethod]
        public void Can_obtain_documents_folder()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            var id = KnownFolderIds.FOLDERID_Documents;
            knownFolderManager.GetFolder(ref id, out knownFolder);
            Assert.IsNotNull(knownFolder);
        }
    }
}

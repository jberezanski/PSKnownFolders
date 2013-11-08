using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BlaSoft.PowerShell.KnownFolders.Win32;
using System.Runtime.CompilerServices;

namespace BlaSoft.PowerShell.KnownFolders.Tests
{
    [TestClass]
    public class KnownFolderInterop
    {
        private IKnownFolder documentsFolder;

        [TestInitialize]
        public void SetUp()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            var id = KnownFolderIds.FOLDERID_Documents;
            knownFolderManager.GetFolder(ref id, out documentsFolder);
        }

        [TestMethod]
        public void Can_obtain_path()
        {
            string path;
            this.documentsFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out path);
            Assert.IsNotNull(path);
        }

        [TestMethod]
        public void Can_obtain_redirection_capabilities()
        {
            KF_REDIRECTION_CAPABILITIES caps;
            this.documentsFolder.GetRedirectionCapabilities(out caps);
        }

        [TestMethod]
        public void Can_obtain_documents_folder_definition()
        {
            var def = KNOWNFOLDER_DEFINITION.FromKnownFolder(this.documentsFolder);
            Assert.AreEqual("Personal", def.name);
        }
    }
}

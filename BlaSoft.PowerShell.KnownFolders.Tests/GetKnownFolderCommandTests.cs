using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace BlaSoft.PowerShell.KnownFolders.Tests
{
    using System.Management.Automation;

    [TestClass]
    public class GetKnownFolderCommandTests
    {

        [TestInitialize]
        public void SetUp()
        {
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder()
        {
            TestDocumentsFolderLowLevel(output =>
            {
                Assert.IsNotNull(output);
                Assert.AreEqual(1, output.Count);
                var psObj = output.First();
                Assert.IsNotNull(psObj);
                Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
            });
        }

        [TestMethod]
        public void Can_inspect_properties_of_user_documents_folder()
        {
            TestDocumentsFolder(psObj =>
            {
                var nameProp = psObj.Properties["Name"];
                Assert.IsNotNull(nameProp);

                var folderIdProp = psObj.Properties["FolderId"];
                Assert.IsNotNull(folderIdProp);

                var pathProp = psObj.Properties["Path"];
                Assert.IsNotNull(pathProp);

                var categoryProp = psObj.Properties["Category"];
                Assert.IsNotNull(categoryProp);

                var canRedirectProp = psObj.Properties["CanRedirect"];
                Assert.IsNotNull(canRedirectProp);

                var folderTypeIdProp = psObj.Properties["FolderTypeId"];
                Assert.IsNotNull(folderTypeIdProp);
            });
        }

        [TestMethod]
        public void Property_Name_of_user_documents_folder_is_correct()
        {
            TestDocumentsFolder(psObj =>
            {
                var nameProp = psObj.Properties["Name"];
                var nameVal = nameProp.Value;
                Assert.IsNotNull(nameVal);
                Assert.IsInstanceOfType(nameVal, typeof(string));
                Assert.AreEqual("unknown", nameVal);
            });
        }

        [TestMethod]
        public void Property_FolderId_of_user_documents_folder_is_correct()
        {
            TestDocumentsFolder(psObj =>
            {
                var folderIdProp = psObj.Properties["FolderId"];
                var folderIdVal = folderIdProp.Value;
                Assert.IsNotNull(folderIdVal);
                Assert.IsInstanceOfType(folderIdVal, typeof(Guid));
                Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, folderIdVal);
            });
        }

        [TestMethod]
        public void Property_Path_of_user_documents_folder_is_correct()
        {
            TestDocumentsFolder(psObj =>
            {
                var pathProp = psObj.Properties["Path"];
                var pathVal = pathProp.Value;
                Assert.IsNotNull(pathVal);
                Assert.IsInstanceOfType(pathVal, typeof(string));
                Assert.AreEqual(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), pathVal);
            });
        }

        [TestMethod]
        public void Property_Category_of_user_documents_folder_is_correct()
        {
            TestDocumentsFolder(psObj =>
            {
                var categoryProp = psObj.Properties["Category"];
                var categoryVal = categoryProp.Value;
                Assert.IsNotNull(categoryVal);
                Assert.IsInstanceOfType(categoryVal, typeof(string));
                Assert.AreEqual(KF_CATEGORY.KF_CATEGORY_PERUSER.ToString(), categoryVal);
            });
        }

        [TestMethod]
        public void Property_CanRedirect_of_user_documents_folder_is_correct()
        {
            TestDocumentsFolder(psObj =>
            {
                var canRedirectProp = psObj.Properties["CanRedirect"];
                var canRedirectVal = canRedirectProp.Value;
                Assert.IsNotNull(canRedirectVal);
                Assert.IsInstanceOfType(canRedirectVal, typeof(bool));
                Assert.AreEqual(true, canRedirectVal);
            });
        }

        [TestMethod]
        public void Cannot_obtain_property_FolderTypeId_of_user_documents_folder()
        {
            TestDocumentsFolder(psObj =>
            {
                var folderTypeIdProp = psObj.Properties["FolderTypeId"];
                Assert.IsNotNull(folderTypeIdProp);
                try
                {
                    var folderTypeIdVal = folderTypeIdProp.Value;
                    Assert.Fail();
                }
                catch (GetValueInvocationException x)
                {
                    Assert.IsNotNull(x.InnerException);
                    Assert.IsInstanceOfType(x.InnerException, typeof(COMException));
                    Assert.AreEqual(unchecked((int)0x80004005), ((COMException)x.InnerException).ErrorCode);
                }
            });
        }

        private static void TestDocumentsFolderLowLevel(Action<Collection<PSObject>> test)
        {
            var sessionState = InitialSessionState.Create();
            sessionState.Commands.Add(new SessionStateCmdletEntry("Get-KnownFolder", typeof(GetKnownFolderCommand), null));
            using (var runspace = RunspaceFactory.CreateRunspace(sessionState))
            using (var psh = PowerShell.Create())
            {
                psh.Runspace = runspace;
                runspace.Open();
                var output = psh
                    .AddCommand("Get-KnownFolder")
                    .AddParameter("FolderId", KnownFolderIds.FOLDERID_Documents.value)
                    .Invoke();
                test(output);
            }
        }

        private static void TestDocumentsFolder(Action<PSObject> test)
        {
            TestDocumentsFolderLowLevel(output => test(output.First()));
        }
    }
}

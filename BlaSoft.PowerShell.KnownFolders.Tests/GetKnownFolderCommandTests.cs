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
        public void Can_obtain_user_documents_folder_by_id()
        {
            TestDocumentsFolderLowLevel(output =>
            {
                Assert.IsNotNull(output);
                Assert.AreEqual(1, output.Count);
                var psObj = output.First();
                Assert.IsNotNull(psObj);
                Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, ((KnownFolder)psObj.BaseObject).FolderId);
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

        [TestMethod]
        public void Can_obtain_all_folders()
        {
            TestCommand(
                psh => psh.AddParameter("All"),
                output =>
                {
                    Assert.IsNotNull(output);
                    Assert.IsTrue(1 < output.Count);
                    CollectionAssert.AllItemsAreNotNull(output);
                    CollectionAssert.AllItemsAreUnique(output);
                    CollectionAssert.AllItemsAreInstancesOfType(output.Select(pso => pso.BaseObject).ToArray(), typeof(KnownFolder));
                });
        }

        [TestMethod]
        public void All_is_implicit()
        {
            TestCommand(
                _ => { },
                output =>
                {
                    Assert.IsNotNull(output);
                    Assert.IsTrue(1 < output.Count);
                    CollectionAssert.AllItemsAreNotNull(output);
                    CollectionAssert.AllItemsAreUnique(output);
                    CollectionAssert.AllItemsAreInstancesOfType(output.Select(pso => pso.BaseObject).ToArray(), typeof(KnownFolder));
                });
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_name()
        {
            TestCommand(
               psh => psh.AddParameter("Name", "Personal"),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(1, output.Count);
                   var psObj = output.First();
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_SpecialFolder_using_canonical_name()
        {
            TestCommand(
               psh => psh.AddParameter("SpecialFolder", Environment.SpecialFolder.Personal),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(1, output.Count);
                   var psObj = output.First();
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_SpecialFolder_using_alternate_name()
        {
            TestCommand(
               psh => psh.AddParameter("SpecialFolder", Environment.SpecialFolder.MyDocuments),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(1, output.Count);
                   var psObj = output.First();
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_user_desktop_folder_by_SpecialFolder()
        {
            TestCommand(
               psh => psh.AddParameter("SpecialFolder", Environment.SpecialFolder.Desktop),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(1, output.Count);
                   var psObj = output.First();
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Desktop.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_common_documents_folder_by_id()
        {
            TestCommand(
               psh => psh.AddParameter("FolderId", KnownFolderIds.FOLDERID_PublicDocuments.value),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(1, output.Count);
                   var psObj = output.First();
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_PublicDocuments.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_multiple_user_folders_by_name()
        {
            TestCommand(
               psh => psh.AddParameter("Name", new[] { "Personal", "Desktop" }),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(2, output.Count);
                   PSObject psObj;
                   psObj = output[0];
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, ((KnownFolder)psObj.BaseObject).FolderId);
                   psObj = output[1];
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Desktop.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_multiple_user_folders_by_SpecialFolder()
        {
            TestCommand(
               psh => psh.AddParameter("SpecialFolder", new[] { Environment.SpecialFolder.Desktop, Environment.SpecialFolder.Favorites }),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(2, output.Count);
                   PSObject psObj;
                   psObj = output[0];
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Desktop.value, ((KnownFolder)psObj.BaseObject).FolderId);
                   psObj = output[1];
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Favorites.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        [TestMethod]
        public void Can_obtain_multiple_user_folders_by_id()
        {
            TestCommand(
               psh => psh.AddParameter("FolderId", new[] { KnownFolderIds.FOLDERID_Documents.value, KnownFolderIds.FOLDERID_Downloads.value }),
               output =>
               {
                   Assert.IsNotNull(output);
                   Assert.AreEqual(2, output.Count);
                   PSObject psObj;
                   psObj = output[0];
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Documents.value, ((KnownFolder)psObj.BaseObject).FolderId);
                   psObj = output[1];
                   Assert.IsNotNull(psObj);
                   Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
                   Assert.AreEqual(KnownFolderIds.FOLDERID_Downloads.value, ((KnownFolder)psObj.BaseObject).FolderId);
               });
        }

        private static void TestCommand(Action<PowerShell> setupCommand, Action<Collection<PSObject>> verify)
        {
            var sessionState = InitialSessionState.Create();
            sessionState.Commands.Add(new SessionStateCmdletEntry("Get-KnownFolder", typeof(GetKnownFolderCommand), null));
            using (var runspace = RunspaceFactory.CreateRunspace(sessionState))
            using (var psh = PowerShell.Create())
            {
                psh.Runspace = runspace;
                runspace.Open();
                psh.AddCommand("Get-KnownFolder");
                setupCommand(psh);
                var output = psh.Invoke();
                verify(output);
            }
        }

        private static void TestDocumentsFolderLowLevel(Action<Collection<PSObject>> verify)
        {
            TestCommand(psh => psh.AddParameter("FolderId", KnownFolderIds.FOLDERID_Documents.value), verify);
        }

        private static void TestDocumentsFolder(Action<PSObject> test)
        {
            TestDocumentsFolderLowLevel(output => test(output.First()));
        }
    }
}

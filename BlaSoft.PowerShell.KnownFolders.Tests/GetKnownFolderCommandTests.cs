using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;
using Microsoft.VisualStudio.TestTools.UnitTesting;

#if FX40
namespace BlaSoft.PowerShell.KnownFolders.Tests.Fx40
#else
namespace BlaSoft.PowerShell.KnownFolders.Tests
#endif
{
    using System.Management.Automation;

    [TestClass]
    public class GetKnownFolderCommandTests : PowerShellTestBase
    {
        private static readonly HashSet<Guid> UserFolders = new HashSet<Guid>()
        {
            KnownFolderIds.FOLDERID_Documents.value,
            KnownFolderIds.FOLDERID_Contacts.value,
            KnownFolderIds.FOLDERID_Links.value,
            KnownFolderIds.FOLDERID_Music.value,
            KnownFolderIds.FOLDERID_Pictures.value,
            KnownFolderIds.FOLDERID_Downloads.value,
            KnownFolderIds.FOLDERID_Desktop.value,
            KnownFolderIds.FOLDERID_Favorites.value,
            KnownFolderIds.FOLDERID_Videos.value,
            KnownFolderIds.FOLDERID_SavedSearches.value,
            KnownFolderIds.FOLDERID_SavedGames.value,
        };

        private static readonly HashSet<Guid> PublicFolders = new HashSet<Guid>()
        {
            KnownFolderIds.FOLDERID_PublicDocuments.value,
            KnownFolderIds.FOLDERID_PublicMusic.value,
            KnownFolderIds.FOLDERID_PublicPictures.value,
            KnownFolderIds.FOLDERID_PublicDownloads.value,
            KnownFolderIds.FOLDERID_PublicVideos.value,
        };

        [ClassInitialize]
        public static void SetUp(TestContext context)
        {
            AssertPowerShellIsExpectedVersion();
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_id()
        {
            TestDocumentsFolderLowLevel(
                output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Documents, output));
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
                Assert.AreEqual("Personal", nameVal);
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
                Assert.IsInstanceOfType(categoryVal, typeof(KnownFolderCategory));
                Assert.AreEqual(KnownFolderCategory.PerUser, categoryVal);
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
        public void Property_FolderTypeId_of_user_documents_folder_is_null()
        {
            TestDocumentsFolder(psObj =>
            {
                var folderTypeIdProp = psObj.Properties["FolderTypeId"];
                Assert.IsNotNull(folderTypeIdProp);
                var folderTypeIdVal = folderTypeIdProp.Value;
                Assert.IsNull(folderTypeIdVal);
            });
        }

        [TestMethod]
        public void Can_obtain_all_folders()
        {
            TestGetKnownFolderCommand(
                psh => psh.AddParameter("All"),
                output =>
                {
                    AssertMultipleFoldersOutputAtLeast(20, output);
                });
        }

        [TestMethod]
        public void Can_obtain_all_per_user_folders()
        {
            TestGetKnownFolderCommand(
                psh => psh.AddParameter("PerUser"),
                output =>
                {
                    AssertMultipleFoldersOutput(UserFolders, output);
                });
        }

        [TestMethod]
        public void Can_obtain_all_public_folders()
        {
            TestGetKnownFolderCommand(
                psh => psh.AddParameter("Public"),
                output =>
                {
                    AssertMultipleFoldersOutput(PublicFolders, output);
                });
        }

        [TestMethod]
        public void PerUser_is_implicit()
        {
            TestGetKnownFolderCommand(
                _ => { },
                output =>
                {
                    AssertMultipleFoldersOutput(UserFolders, output);
                });
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_name()
        {
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("Name", "Personal"),
               output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Documents, output));
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_SpecialFolder_using_canonical_name()
        {
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("SpecialFolder", Environment.SpecialFolder.Personal),
               output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Documents, output));

            TestScript(
                "Get-KnownFolder -SpecialFolder Personal",
                output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Documents, output));
        }

        [TestMethod]
        public void Can_obtain_user_documents_folder_by_SpecialFolder_using_alternate_name()
        {
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("SpecialFolder", Environment.SpecialFolder.MyDocuments),
               output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Documents, output));

            TestScript(
                "Get-KnownFolder -SpecialFolder MyDocuments",
                output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Documents, output));
        }

        [TestMethod]
        public void Can_obtain_user_desktop_folder_by_SpecialFolder()
        {
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("SpecialFolder", Environment.SpecialFolder.Desktop),
               output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_Desktop, output));
        }

        [TestMethod]
        public void Can_obtain_common_documents_folder_by_id()
        {
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("FolderId", KnownFolderIds.FOLDERID_PublicDocuments.value),
               output => AssertSingleFolderOutput(KnownFolderIds.FOLDERID_PublicDocuments, output));
        }

        [TestMethod]
        public void Can_obtain_multiple_user_folders_by_name()
        {
            var folderIds = new[] { KnownFolderIds.FOLDERID_Desktop.value, KnownFolderIds.FOLDERID_Documents.value };
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("Name", new[] { "Personal", "Desktop" }),
               output =>
               {
                   AssertMultipleFoldersOutput(folderIds, output);
               });
        }

        [TestMethod]
        public void Can_obtain_multiple_user_folders_by_SpecialFolder()
        {
            var folderIds = new[] { KnownFolderIds.FOLDERID_Desktop.value, KnownFolderIds.FOLDERID_Favorites.value };
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("SpecialFolder", new[] { Environment.SpecialFolder.Desktop, Environment.SpecialFolder.Favorites }),
               output =>
               {
                   AssertMultipleFoldersOutput(folderIds, output);
               });
        }

        [TestMethod]
        public void Can_obtain_multiple_user_folders_by_id()
        {
            var folderIds = new[] { KnownFolderIds.FOLDERID_Documents.value, KnownFolderIds.FOLDERID_Downloads.value };
            TestGetKnownFolderCommand(
               psh => psh.AddParameter("FolderId", folderIds),
               output =>
               {
                   AssertMultipleFoldersOutput(folderIds, output);
               });
        }

        private static void AssertMultipleFoldersOutputAtLeast(int atLeast, Collection<PSObject> output)
        {
            Assert.IsNotNull(output);
            Assert.IsTrue(atLeast <= output.Count);
            CollectionAssert.AllItemsAreNotNull(output);
            CollectionAssert.AllItemsAreUnique(output);
            CollectionAssert.AllItemsAreInstancesOfType(output.Select(pso => pso.BaseObject).ToArray(), typeof(KnownFolder));
            var outputIds = output.Select(pso => pso.BaseObject).Cast<KnownFolder>().Select(kf => kf.FolderId).ToArray();
            CollectionAssert.AllItemsAreUnique(outputIds);
        }

        private static void AssertMultipleFoldersOutput(ICollection<Guid> expected, Collection<PSObject> output)
        {
            Assert.IsNotNull(output);
            ////Assert.AreEqual(expected.Count, output.Count);
            CollectionAssert.AllItemsAreNotNull(output);
            CollectionAssert.AllItemsAreUnique(output);
            CollectionAssert.AllItemsAreInstancesOfType(output.Select(pso => pso.BaseObject).ToArray(), typeof(KnownFolder));
            var outputIds = output.Select(pso => pso.BaseObject).Cast<KnownFolder>().Select(kf => kf.FolderId).ToArray();
            CollectionAssert.AllItemsAreUnique(outputIds);
            CollectionAssert.AreEquivalent(expected.ToArray(), outputIds);
        }

        private static void AssertSingleFolderOutput(KNOWNFOLDERID expectedId, Collection<PSObject> output)
        {
            Assert.IsNotNull(output);
            Assert.AreEqual(1, output.Count);
            var psObj = output.First();
            Assert.IsNotNull(psObj);
            Assert.IsInstanceOfType(psObj.BaseObject, typeof(KnownFolder));
            Assert.AreEqual(expectedId.value, ((KnownFolder)psObj.BaseObject).FolderId);
        }

        private static void TestGetKnownFolderCommand(Action<PowerShell> setupCommand, Action<Collection<PSObject>> verify)
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
                if (psh.Streams.Error.Any())
                {
                    Assert.Fail("Command execution failed, first error: {0}", psh.Streams.Error.First());
                }

                verify(output);
            }
        }

        private static void TestScript(string script, Action<Collection<PSObject>> verify)
        {
            var sessionState = InitialSessionState.Create();
            sessionState.LanguageMode = PSLanguageMode.FullLanguage;
            sessionState.Commands.Add(new SessionStateCmdletEntry("Get-KnownFolder", typeof(GetKnownFolderCommand), null));
            using (var runspace = RunspaceFactory.CreateRunspace(sessionState))
            using (var psh = PowerShell.Create())
            {
                psh.Runspace = runspace;
                runspace.Open();
                psh.AddScript(script);
                var output = psh.Invoke();
                if (psh.Streams.Error.Any())
                {
                    Assert.Fail("Script execution failed, first error: {0}", psh.Streams.Error.First());
                }

                verify(output);
            }
        }

        private static void TestDocumentsFolderLowLevel(Action<Collection<PSObject>> verify)
        {
            TestGetKnownFolderCommand(psh => psh.AddParameter("FolderId", KnownFolderIds.FOLDERID_Documents.value), verify);
        }

        private static void TestDocumentsFolder(Action<PSObject> test)
        {
            TestDocumentsFolderLowLevel(output => test(output.First()));
        }
    }
}

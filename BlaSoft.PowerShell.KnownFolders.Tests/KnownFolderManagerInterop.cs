using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
        public void Can_obtain_documents_folder_by_id()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            var id = KnownFolderIds.FOLDERID_Documents;
            knownFolderManager.GetFolder(ref id, out knownFolder);
            Assert.IsNotNull(knownFolder);
            knownFolder.GetId(out id);
            Assert.AreEqual(KnownFolderIds.FOLDERID_Documents, id);
        }

        [TestMethod]
        public void Can_obtain_documents_folder_by_name()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            knownFolderManager.GetFolderByName("Personal", out knownFolder);
            Assert.IsNotNull(knownFolder);
            KNOWNFOLDERID id;
            knownFolder.GetId(out id);
            Assert.AreEqual(KnownFolderIds.FOLDERID_Documents, id);
        }

        [TestMethod]
        public void Can_obtain_all_folder_ids()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            uint count = 0;
            object ppIds = IntPtr.Zero;
            var h = GCHandle.Alloc(ppIds, GCHandleType.Pinned);
            try
            {
                knownFolderManager.GetFolderIds(h.AddrOfPinnedObject(), ref count);
                IntPtr pIds = (IntPtr)ppIds;
                if (IntPtr.Zero == pIds)
                {
                    throw new InvalidOperationException("GetFolderIds returned NULL");
                }

                try
                {
                    var ids = new KNOWNFOLDERID[count];
                    var ptr = pIds.ToInt64();
                    for (uint u = 0; u < count; ++u)
                    {
                        ids[u] = (KNOWNFOLDERID)Marshal.PtrToStructure((IntPtr)ptr, typeof(KNOWNFOLDERID));
                        ptr += Marshal.SizeOf(typeof(KNOWNFOLDERID));
                    }

                    foreach (var id in ids)
                    {
                        Assert.AreNotEqual(Guid.Empty, id.value);
                    }
                }
                finally
                {
                    Marshal.FreeCoTaskMem(pIds);
                }
            }
            finally
            {
                h.Free();
            }
        }

        [TestMethod]
        public void Cannot_redirect_Links_to_invalid_path()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            var id = KnownFolderIds.FOLDERID_Links;
            string error;
            var hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_NONE,
                "g:e:f*&#%^NIY",
                0,
                null,
                out error);
            Assert.AreEqual(HResult.WIN32_E_PATHNOTFOUND, hr);
        }

        [TestMethod]
        public void Can_check_if_Links_is_redirected_to_itself()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            var id = KnownFolderIds.FOLDERID_Links;
            knownFolderManager.GetFolder(ref id, out knownFolder);
            string currentPath;
            knownFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out currentPath);
            string error;
            var hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_CHECK_ONLY,
                currentPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
        }

        [TestMethod]
        public void Can_check_if_Links_is_not_already_redirected()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            var id = KnownFolderIds.FOLDERID_Links;
            string newPath = GetTempRedirectionPath();
            string error;
            var hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_CHECK_ONLY,
                newPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_FALSE, hr);
        }

        [TestMethod]
        public void Can_redirect_Links_without_moving_existing_data()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            var id = KnownFolderIds.FOLDERID_Links;
            knownFolderManager.GetFolder(ref id, out knownFolder);
            string currentPath;
            knownFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out currentPath);
            string newPath = GetTempRedirectionPath();
            string error;
            HResult hr;
            hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_NONE,
                newPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
            hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_NONE,
                currentPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
            DeleteTempRedirectionDirectory(newPath);
        }

        [TestMethod]
        public void Can_redirect_Links_with_copying_existing_data()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            var id = KnownFolderIds.FOLDERID_Links;
            knownFolderManager.GetFolder(ref id, out knownFolder);
            string currentPath;
            knownFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out currentPath);
            string newPath = GetTempRedirectionPath();
            string error;
            HResult hr;
            hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_COPY_CONTENTS,
                newPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
            hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_NONE,
                currentPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
            DeleteTempRedirectionDirectory(newPath);
        }

        [TestMethod]
        public void Can_redirect_Links_with_moving_existing_data()
        {
            var knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
            IKnownFolder knownFolder;
            var id = KnownFolderIds.FOLDERID_Links;
            knownFolderManager.GetFolder(ref id, out knownFolder);
            string currentPath;
            knownFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out currentPath);
            string newPath = GetTempRedirectionPath();
            string error;
            HResult hr;
            hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_COPY_CONTENTS | KF_REDIRECT_FLAGS.KF_REDIRECT_DEL_SOURCE_CONTENTS,
                newPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
            hr = knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                KF_REDIRECT_FLAGS.KF_REDIRECT_COPY_CONTENTS | KF_REDIRECT_FLAGS.KF_REDIRECT_DEL_SOURCE_CONTENTS,
                currentPath,
                0,
                null,
                out error);
            Assert.AreEqual(HResult.S_OK, hr);
            DeleteTempRedirectionDirectory(newPath);
        }

        private static void DeleteTempRedirectionDirectory(string newPath)
        {
            if (Directory.Exists(newPath))
            {
                var attrs = File.GetAttributes(newPath);
                if ((attrs & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    File.SetAttributes(newPath, attrs & (~FileAttributes.ReadOnly));
                }

                Directory.Delete(newPath, true);
            }
        }

        private static string GetTempRedirectionPath()
        {
            string newPath = Path.Combine(Path.Combine(Path.GetTempPath(), Assembly.GetExecutingAssembly().GetName().Name), Guid.NewGuid().ToString("N"));
            return newPath;
        }
    }
}

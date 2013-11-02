using System;
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
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    [Cmdlet(VerbsCommon.Get, "KnownFolder")]
    public class GetKnownFolderCommand : Cmdlet
    {
        private IKnownFolderManager knownFolderManager;

        [Parameter(ParameterSetName = "ByName", Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; }

        [Parameter(ParameterSetName = "ByFolderId", Mandatory = true, Position = 0)]
        //[ValidateScript(ScriptBlock.Create("$_ -ne [Guid]::Empty"))]
        public Guid FolderId { get; set; }

        [Parameter(ParameterSetName = "All", Mandatory = true, Position = 0)]
        public SwitchParameter All { get; set; }

        protected override void BeginProcessing()
        {
            this.knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
        }

        protected override void ProcessRecord()
        {
            IKnownFolder nativeKnownFolder;
            IEnumerable<IKnownFolder> result;
            if (this.FolderId != Guid.Empty)
            {
                var knownFolderId = new KNOWNFOLDERID(this.FolderId.ToString());
                nativeKnownFolder = this.GetKnownFolderById(knownFolderId);
                result = new[] { nativeKnownFolder };
            }
            else if (this.Name != null)
            {
                this.knownFolderManager.GetFolderByName(this.Name, out nativeKnownFolder);
                result = new[] { nativeKnownFolder };
            }
            else if (this.All)
            {
                uint count = 0;
                object boxpIds = IntPtr.Zero;
                var h = GCHandle.Alloc(boxpIds, GCHandleType.Pinned);
                try
                {
                    this.knownFolderManager.GetFolderIds(h.AddrOfPinnedObject(), ref count);
                    IntPtr pIds = (IntPtr)boxpIds;
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

                        result = ids.Select(kfi => this.GetKnownFolderById(kfi));
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
            else
            {
                throw new ArgumentException("Unknown parameter set");
            }

            foreach (var kf in result)
            {
                this.WriteObject(new KnownFolder(kf, "unknown"));
            }
        }

        private IKnownFolder GetKnownFolderById(KNOWNFOLDERID knownFolderId)
        {
            IKnownFolder nativeKnownFolder;
            this.knownFolderManager.GetFolder(ref knownFolderId, out nativeKnownFolder);
            return nativeKnownFolder;
        }
    }
}

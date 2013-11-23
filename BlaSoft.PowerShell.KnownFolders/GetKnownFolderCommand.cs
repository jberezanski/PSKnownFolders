using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    [Cmdlet(VerbsCommon.Get, "KnownFolder", DefaultParameterSetName = "All")]
    [OutputType(typeof(KnownFolder))]
    public sealed class GetKnownFolderCommand : PSCmdlet
    {
        private IKnownFolderManager knownFolderManager;

        [Parameter(ParameterSetName = "ByName", Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty]
        public string Name { get; set; }

        [Parameter(ParameterSetName = "ByFolderId", Mandatory = true, Position = 0)]
        public Guid FolderId { get; set; }

        [Parameter(ParameterSetName = "BySpecialFolder", Mandatory = true, Position = 0)]
        public Environment.SpecialFolder SpecialFolder { get; set; }

        [Parameter(ParameterSetName = "All", Position = 0)]
        public SwitchParameter All { get; set; }

        protected override void BeginProcessing()
        {
            this.knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
        }

        protected override void ProcessRecord()
        {
            IEnumerable<IKnownFolder> result;
            switch (this.ParameterSetName)
            {
                case "ByName":
                    result = GetByName(this.Name);
                    break;
                case "BySpecialFolder":
                    if (!Enum.IsDefined(typeof(Environment.SpecialFolder), this.SpecialFolder))
                    {
                        throw new ArgumentException("Invalid SpecialFolder value");
                    }

                    result = GetByName(this.SpecialFolder.ToString());
                    break;
                case "ByFolderId":
                    result = GetById(this.FolderId);
                    break;
                case "All":
                    result = GetAll();
                    break;
                default:
                    throw new ArgumentException("Unsupported parameter set name");
            }

            foreach (var kf in result)
            {
                this.WriteObject(new KnownFolder(kf, "unknown"));
            }
        }

        private IEnumerable<IKnownFolder> GetAll()
        {
            IEnumerable<IKnownFolder> result;
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

            return result;
        }

        private IEnumerable<IKnownFolder> GetByName(string name)
        {
            IKnownFolder nativeKnownFolder;
            this.knownFolderManager.GetFolderByName(name, out nativeKnownFolder);
            var result = new[] { nativeKnownFolder };
            return result;
        }

        private IEnumerable<IKnownFolder> GetById(Guid folderId)
        {
            IKnownFolder nativeKnownFolder;
            var knownFolderId = new KNOWNFOLDERID(folderId.ToString());
            nativeKnownFolder = this.GetKnownFolderById(knownFolderId);
            var result = new[] { nativeKnownFolder };
            return result;
        }

        private IKnownFolder GetKnownFolderById(KNOWNFOLDERID knownFolderId)
        {
            IKnownFolder nativeKnownFolder;
            this.knownFolderManager.GetFolder(ref knownFolderId, out nativeKnownFolder);
            return nativeKnownFolder;
        }
    }
}

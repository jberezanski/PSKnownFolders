using System;
using System.Collections.Generic;
using System.IO;
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
        public string[] Name { get; set; }

        [Parameter(ParameterSetName = "ByFolderId", Mandatory = true, Position = 0)]
        public Guid[] FolderId { get; set; }

        [Parameter(ParameterSetName = "BySpecialFolder", Mandatory = true, Position = 0)]
        public Environment.SpecialFolder[] SpecialFolder { get; set; }

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
                    Validate(this.SpecialFolder);
                    result = GetByName(this.SpecialFolder.Select(sf => sf == Environment.SpecialFolder.Personal ? "Personal" : sf.ToString()));
                    break;
                case "ByFolderId":
                    result = GetById(this.FolderId);
                    break;
                case "All":
                    result = GetAll();
                    break;
                default:
                    throw new ArgumentException("Unsupported parameter set name: " + this.ParameterSetName);
            }

            foreach (var kf in result)
            {
                this.WriteObject(new KnownFolder(kf, "unknown"));
            }
        }

        private static void Validate(IEnumerable<Environment.SpecialFolder> specialFolders)
        {
            foreach (var specialFolder in specialFolders)
            {
                if (!Enum.IsDefined(typeof(Environment.SpecialFolder), specialFolder))
                {
                    throw new ArgumentException("Invalid SpecialFolder value: " + specialFolder);
                }
            }
        }

        private IEnumerable<IKnownFolder> GetAll()
        {
            KNOWNFOLDERID[] ids;

            object boxpIds = IntPtr.Zero;
            var h = GCHandle.Alloc(boxpIds, GCHandleType.Pinned);
            try
            {
                uint count = 0;
                this.knownFolderManager.GetFolderIds(h.AddrOfPinnedObject(), ref count);
                IntPtr pIds = (IntPtr)boxpIds;
                if (IntPtr.Zero == pIds)
                {
                    throw new InvalidOperationException("GetFolderIds returned NULL");
                }

                try
                {
                    ids = new KNOWNFOLDERID[count];
                    var ptr = pIds.ToInt64();
                    for (uint u = 0; u < count; ++u)
                    {
                        ids[u] = (KNOWNFOLDERID)Marshal.PtrToStructure((IntPtr)ptr, typeof(KNOWNFOLDERID));
                        ptr += Marshal.SizeOf(typeof(KNOWNFOLDERID));
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

            var result = ids.Select(kfi => this.GetKnownFolderById(kfi));

            return result;
        }

        private IEnumerable<IKnownFolder> GetByName(IEnumerable<string> names)
        {
            return names.Select(name => this.GetKnownFolderByName(name));
        }

        private IEnumerable<IKnownFolder> GetById(IEnumerable<Guid> folderIds)
        {
            return folderIds.Select(folderId => this.GetKnownFolderById(new KNOWNFOLDERID(folderId.ToString())));
        }

        private IKnownFolder GetKnownFolderById(KNOWNFOLDERID knownFolderId)
        {
            IKnownFolder nativeKnownFolder;
            try
            {
                this.knownFolderManager.GetFolder(ref knownFolderId, out nativeKnownFolder);
            }
            catch (FileNotFoundException x)
            {
                throw new FileNotFoundException(string.Format("Known folder not found: {0}", knownFolderId.value), x);
            }

            return nativeKnownFolder;
        }

        private IKnownFolder GetKnownFolderByName(string name)
        {
            IKnownFolder nativeKnownFolder;
            try
            {
                this.knownFolderManager.GetFolderByName(name, out nativeKnownFolder);
            }
            catch (FileNotFoundException x)
            {
                throw new FileNotFoundException(string.Format("Known folder not found: '{0}'", name), x);
            }

            return nativeKnownFolder;
        }
    }
}

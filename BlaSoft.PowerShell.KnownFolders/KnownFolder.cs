using System;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    public class KnownFolder
    {
        private IKnownFolder nativeKnownFolder;

        internal KnownFolder(IKnownFolder nativeKnownFolder, string name)
        {
            if (nativeKnownFolder == null)
            {
                throw new ArgumentNullException("nativeKnownFolder");
            }

            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("name is null or empty", "name");
            }

            this.nativeKnownFolder = nativeKnownFolder;
            this.Name = name;
        }

        public string Name { get; private set; }

        public string Path
        {
            get
            {
                string path;
                this.nativeKnownFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out path);
                return path;
            }
        }

        public bool CanRedirect
        {
            get
            {
                KF_REDIRECTION_CAPABILITIES caps;
                this.nativeKnownFolder.GetRedirectionCapabilities(out caps);
                return caps.HasMask(KF_REDIRECTION_CAPABILITIES.KF_REDIRECTION_CAPABILITIES_ALLOW_ALL)
                    && !caps.HasMask(KF_REDIRECTION_CAPABILITIES.KF_REDIRECTION_CAPABILITIES_DENY_ALL);
            }
        }

        public string Category
        {
            get
            {
                KF_CATEGORY cat;
                this.nativeKnownFolder.GetCategory(out cat);
                return cat.ToString();
            }
        }

        public Guid FolderId
        {
            get
            {
                KNOWNFOLDERID id;
                this.nativeKnownFolder.GetId(out id);
                return id.value;
            }
        }

        public Guid FolderTypeId
        {
            get
            {
                FOLDERTYPEID typeid;
                this.nativeKnownFolder.GetFolderType(out typeid);
                return typeid.value;
            }
        }
    }
}

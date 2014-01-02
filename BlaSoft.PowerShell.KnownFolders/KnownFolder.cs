using System;
using System.IO;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    public sealed class KnownFolder
    {
        private readonly IKnownFolder nativeKnownFolder;

        private KnownFolderDefinition definition;

        internal KnownFolder(IKnownFolder nativeKnownFolder)
        {
            if (nativeKnownFolder == null)
            {
                throw new ArgumentNullException("nativeKnownFolder");
            }

            this.nativeKnownFolder = nativeKnownFolder;
        }

        public string Name
        {
            get
            {
                return this.Definition.Name;
            }
        }

        public string Path
        {
            get
            {
                string path;
                try
                {
                    this.nativeKnownFolder.GetPath(KNOWN_FOLDER_FLAG.KF_FLAG_NONE, out path);
                }
                catch (DirectoryNotFoundException)
                {
                    return null;
                }
                catch (FileNotFoundException)
                {
                    return null;
                }
                catch (COMException)
                {
                    return null;
                }

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

        public KnownFolderCategory Category
        {
            get
            {
                KF_CATEGORY cat;
                this.nativeKnownFolder.GetCategory(out cat);
                return (KnownFolderCategory)cat;
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

        public Guid? FolderTypeId
        {
            get
            {
                FOLDERTYPEID typeid;
                try
                {
                    this.nativeKnownFolder.GetFolderType(out typeid);
                }
                catch (COMException)
                {
                    return null;
                }

                return typeid.value;
            }
        }

        public KnownFolderDefinition Definition
        {
            get
            {
                if (this.definition == null)
                {
                    this.definition = new KnownFolderDefinition(this.nativeKnownFolder);
                }

                return this.definition;
            }
        }
    }
}

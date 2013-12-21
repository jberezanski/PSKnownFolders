using System;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    public sealed class KnownFolderDefinition
    {
        private readonly KNOWNFOLDER_DEFINITION nativeDefinition;

        internal KnownFolderDefinition(IKnownFolder nativeKnownFolder)
        {
            if (nativeKnownFolder == null)
            {
                throw new ArgumentNullException("nativeKnownFolder");
            }

            this.nativeDefinition = KNOWNFOLDER_DEFINITION.FromKnownFolder(nativeKnownFolder);
        }

        public string Name
        {
            get
            {
                return this.nativeDefinition.name;
            }
        }

        public string ParsingName
        {
            get
            {
                return this.nativeDefinition.parsingName;
            }
        }

        public string LocalizedName
        {
            get
            {
                return this.nativeDefinition.localizedName;
            }
        }

        public string Description
        {
            get
            {
                return this.nativeDefinition.description;
            }
        }

        public string ToolTip
        {
            get
            {
                return this.nativeDefinition.toolTip;
            }
        }

        public string Icon
        {
            get
            {
                return this.nativeDefinition.icon;
            }
        }

        public string Security
        {
            get
            {
                return this.nativeDefinition.security;
            }
        }

        public string RelativePath
        {
            get
            {
                return this.nativeDefinition.relativePath;
            }
        }

        public Guid? FolderTypeId
        {
            get
            {
                return this.nativeDefinition.ftidType != Guid.Empty ? (Guid?)this.nativeDefinition.ftidType : null;
            }
        }

        public Guid? ParentId
        {
            get
            {
                return this.nativeDefinition.fidParent != Guid.Empty ? (Guid?)this.nativeDefinition.fidParent : null;
            }
        }

        public string Category
        {
            get
            {
                return this.nativeDefinition.category.ToString();
            }
        }

        public string Flags
        {
            get
            {
                return this.nativeDefinition.kfdFlags.ToString();
            }
        }

        public uint Attributes
        {
            get
            {
                return this.nativeDefinition.dwAttributes;
            }
        }
    }
}

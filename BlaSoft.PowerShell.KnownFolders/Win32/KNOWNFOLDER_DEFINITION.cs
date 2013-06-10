using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 4)]
    struct KNOWNFOLDER_DEFINITION
    {
        public KF_CATEGORY category;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszName;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszCreator;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszDescription;

        public Guid fidParent;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszRelativePath;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszParsingName;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszToolTip;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszLocalizedName;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszIcon;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string pszSecurity;

        public uint dwAttributes;

        public KF_DEFINITION_FLAGS kfdFlags;

        public Guid ftidType;
    }
}

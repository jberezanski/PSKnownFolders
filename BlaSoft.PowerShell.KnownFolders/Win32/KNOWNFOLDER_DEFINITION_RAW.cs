using System;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct KNOWNFOLDER_DEFINITION_RAW
    {
        public KF_CATEGORY category;

        public IntPtr pszName;

        public IntPtr pszDescription;

        public Guid fidParent;

        public IntPtr pszRelativePath;

        public IntPtr pszParsingName;

        public IntPtr pszToolTip;

        public IntPtr pszLocalizedName;

        public IntPtr pszIcon;

        public IntPtr pszSecurity;

        public uint dwAttributes;

        public KF_DEFINITION_FLAGS kfdFlags;

        public Guid ftidType;

        public void FreeKnownFolderDefinitionFields()
        {
            /*
http://www.winehq.org/pipermail/wine-cvs/2011-June/078517.html
+cpp_quote("    CoTaskMemFree(pKFD->pszName);")
+cpp_quote("    CoTaskMemFree(pKFD->pszDescription);")
+cpp_quote("    CoTaskMemFree(pKFD->pszRelativePath);")
+cpp_quote("    CoTaskMemFree(pKFD->pszParsingName);")
+cpp_quote("    CoTaskMemFree(pKFD->pszTooltip);")
+cpp_quote("    CoTaskMemFree(pKFD->pszLocalizedName);")
+cpp_quote("    CoTaskMemFree(pKFD->pszIcon);")
+cpp_quote("    CoTaskMemFree(pKFD->pszSecurity);")
             * * */
            Marshal.FreeCoTaskMem(this.pszName);
            Marshal.FreeCoTaskMem(this.pszDescription);
            Marshal.FreeCoTaskMem(this.pszRelativePath);
            Marshal.FreeCoTaskMem(this.pszParsingName);
            Marshal.FreeCoTaskMem(this.pszToolTip);
            Marshal.FreeCoTaskMem(this.pszLocalizedName);
            Marshal.FreeCoTaskMem(this.pszIcon);
            Marshal.FreeCoTaskMem(this.pszSecurity);
        }
    }
}

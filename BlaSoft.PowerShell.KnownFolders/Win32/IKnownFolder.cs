using System;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [ComImport]
    [Guid("3AA7AF7E-9B36-420c-A8E3-F77D4674A488")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IKnownFolder
    {
        void GetId(out KNOWNFOLDERID pkfid);

        void GetCategory(out KF_CATEGORY pCategory);

        // Can return IShellItem or IShellItem2, based on passed riid
        void GetShellItem([In] KNOWN_FOLDER_FLAG dwFlags, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);

        void GetPath([In] KNOWN_FOLDER_FLAG dwFlags, [MarshalAs(UnmanagedType.LPWStr)] out string ppszPath);

        void SetPath([In] KNOWN_FOLDER_FLAG dwFlags, [In, MarshalAs(UnmanagedType.LPWStr)] string pszPath);

        void GetIDList([In] KNOWN_FOLDER_FLAG dwFlags, [Out, ComAliasName("ShellObjects.wirePIDL")] out SafeCoTaskMemHandle ppidl);

        void GetFolderType(out FOLDERTYPEID pftid);

        void GetRedirectionCapabilities(out KF_REDIRECTION_CAPABILITIES pCapabilities);

        // Should be cleaned up by calling FreeKnownFolderDefinitionFields
        void GetFolderDefinition(out KNOWNFOLDER_DEFINITION_RAW pKFD);
    }
}

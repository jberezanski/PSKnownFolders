using System;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [ComImport]
    [Guid("8BE2D872-86AA-4d47-B776-32CCA40C7018")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IKnownFolderManager
    {
        void FolderIdFromCsidl([In] int nCsidl, out Guid pfid);

        void FolderIdToCsidl([In] ref KNOWNFOLDERID rfid, out int pnCsidl);

        void GetFolderIds([Out] IntPtr ppKFId, [In, Out] ref uint pCount);

        void GetFolder([In] ref KNOWNFOLDERID rfid, [MarshalAs(UnmanagedType.Interface)] out IKnownFolder ppkf);

        void GetFolderByName(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName,
            [MarshalAs(UnmanagedType.Interface)] out IKnownFolder ppkf);

        void RegisterFolder([In] ref KNOWNFOLDERID rfid, [In] ref KNOWNFOLDER_DEFINITION pKFD);

        void UnregisterFolder([In] ref KNOWNFOLDERID rfid);

        void FindFolderFromPath(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            [In] FFFP_MODE mode,
            [MarshalAs(UnmanagedType.Interface)] out IKnownFolder ppkf);

        void FindFolderFromIDList([In] IntPtr pidl, [MarshalAs(UnmanagedType.Interface)] out IKnownFolder ppkf);

        void Redirect(
            [In] ref KNOWNFOLDERID rfid, 
            [In, Optional] IntPtr hwnd, 
            [In] KF_REDIRECT_FLAGS Flags,
            [In, Optional, MarshalAs(UnmanagedType.LPWStr)] string pszTargetPath, 
            [In] uint cFolders,
            [In] KNOWNFOLDERID[] pExclusion, 
            [MarshalAs(UnmanagedType.LPWStr)] out string ppszError);
    }
}

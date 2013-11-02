using System;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct FOLDERTYPEID
    {
        public Guid value;
    }
}

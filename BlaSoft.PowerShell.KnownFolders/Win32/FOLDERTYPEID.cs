using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    struct FOLDERTYPEID
    {
        public Guid value;
    }
}

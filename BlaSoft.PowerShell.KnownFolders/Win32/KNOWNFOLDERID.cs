using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [StructLayout(LayoutKind.Sequential)]
    struct KNOWNFOLDERID
    {
        public Guid value;

        public KNOWNFOLDERID(string id)
        {
            this.value = new Guid(id);
        }

        public KNOWNFOLDERID(uint a, ushort b, ushort c, byte d, byte e, byte f, byte g, byte h, byte i, byte j, byte k)
        {
            this.value = new Guid(a, b, c, d, e, f, g, h, i, j, k);
        }
    }
}

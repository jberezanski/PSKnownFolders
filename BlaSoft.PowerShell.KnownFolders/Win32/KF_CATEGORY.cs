using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    enum KF_CATEGORY
    {
        KF_CATEGORY_INVALID = 0,
        KF_CATEGORY_VIRTUAL = 0x00000001,
        KF_CATEGORY_FIXED = 0x00000002,
        KF_CATEGORY_COMMON = 0x00000003,
        KF_CATEGORY_PERUSER = 0x00000004
    }
}

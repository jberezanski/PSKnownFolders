using System;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [Flags]
    internal enum KF_DEFINITION_FLAGS : uint
    {
        KFDF_NONE = 0,
        KFDF_PERSONALIZE = 0x00000001,
        KFDF_LOCAL_REDIRECT_ONLY = 0x00000002,
        KFDF_ROAMABLE = 0x00000004,
    }
}

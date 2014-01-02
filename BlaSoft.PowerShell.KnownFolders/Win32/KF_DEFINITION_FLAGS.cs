using System;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [Flags]
    internal enum KF_DEFINITION_FLAGS : uint
    {
        KFDF_NONE = 0,
        [Obsolete("Not documented in Windows 8.1 SDK")]
        KFDF_PERSONALIZE = 0x00000001,
        KFDF_LOCAL_REDIRECT_ONLY = 0x00000002,
        KFDF_ROAMABLE = 0x00000004,
        KFDF_PRECREATE = 0x00000008,
        KFDF_STREAM = 0x00000010,
        KFDF_PUBLISHEXPANDEDPATH = 0x00000020,
        KFDF_NO_REDIRECT_UI = 0x00000040,
    }
}

using System;

namespace BlaSoft.PowerShell.KnownFolders
{
    [Flags]
    public enum KnownFolderDefinitionFlags
    {
        None = 0,
        [Obsolete("Not documented in Windows 8.1 SDK")]
        Personalize = 0x00000001,
        LocalRedirectOnly = 0x00000002,
        Roamable = 0x00000004,
        PreCreate = 0x00000008,
        Stream = 0x00000010,
        PublishExpandedPath = 0x00000020,
        NoRedirectUI = 0x00000040,
    }
}

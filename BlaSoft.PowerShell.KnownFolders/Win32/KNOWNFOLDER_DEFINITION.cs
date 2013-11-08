using System;
using System.Runtime.CompilerServices;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;

namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    internal struct KNOWNFOLDER_DEFINITION
    {
        public KF_CATEGORY category;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string name;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string description;

        public Guid fidParent;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string relativePath;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string parsingName;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string toolTip;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string localizedName;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string icon;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string security;

        public uint dwAttributes;

        public KF_DEFINITION_FLAGS kfdFlags;

        public Guid ftidType;

        public static KNOWNFOLDER_DEFINITION FromKnownFolder(IKnownFolder knownFolder)
        {
            if (knownFolder == null)
            {
                throw new ArgumentNullException("knownFolder");
            }

            KNOWNFOLDER_DEFINITION def;

            using (var handles = new KnownFolderDefinitionHandles())
            {
                KNOWNFOLDER_DEFINITION_RAW rawDef;

                RuntimeHelpers.PrepareConstrainedRegions();
                try
                {
                }
                finally
                {
                    knownFolder.GetFolderDefinition(out rawDef);
                    try
                    {
                        handles.SecureHandles(ref rawDef);
                    }
                    finally
                    {
                        // in case SecureHandles throws
                        rawDef.FreeKnownFolderDefinitionFields();
                    }
                }

                def = new KNOWNFOLDER_DEFINITION(ref rawDef, handles);
            }

            return def;
        }

        private KNOWNFOLDER_DEFINITION(ref KNOWNFOLDER_DEFINITION_RAW d, KnownFolderDefinitionHandles h)
        {
            this.category = d.category;
            this.fidParent = d.fidParent;
            this.dwAttributes = d.dwAttributes;
            this.kfdFlags = d.kfdFlags;
            this.ftidType = d.ftidType;

            this.name = UnmarshalString(h.hName, "hName");
            this.description = UnmarshalString(h.hDescription, "hDescription");
            this.relativePath = UnmarshalString(h.hRelativePath, "hRelativePath");
            this.parsingName = UnmarshalString(h.hParsingName, "hParsingName");
            this.toolTip = UnmarshalString(h.hToolTip, "hToolTip");
            this.localizedName = UnmarshalString(h.hLocalizedName, "hLocalizedName");
            this.icon = UnmarshalString(h.hIcon, "hIcon");
            this.security = UnmarshalString(h.hSecurity, "hSecurity");
        }

        private static string UnmarshalString(SafeCoTaskMemHandle hValue, string handleName)
        {
            string value;
            bool success;

            RuntimeHelpers.PrepareConstrainedRegions();
            try
            {
            }
            finally
            {
                success = false;
                hValue.DangerousAddRef(ref success);
                if (success)
                {
                    value = Marshal.PtrToStringUni(hValue.DangerousGetHandle());
                    hValue.DangerousRelease();
                }
                else
                {
                    value = null;
                }
            }

            if (!success)
            {
                throw new InvalidOperationException("Failed to AddRef on " + handleName);
            }

            return value;
        }

        private sealed class KnownFolderDefinitionHandles : IDisposable
        {
            public readonly SafeCoTaskMemHandle hName = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hDescription = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hRelativePath = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hParsingName = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hToolTip = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hLocalizedName = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hIcon = new SafeCoTaskMemHandle();

            public readonly SafeCoTaskMemHandle hSecurity = new SafeCoTaskMemHandle();

            [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
            public void SecureHandles(ref KNOWNFOLDER_DEFINITION_RAW d)
            {
                SecureHandle(this.hName, ref d.pszName);
                SecureHandle(this.hDescription, ref d.pszDescription);
                SecureHandle(this.hRelativePath, ref d.pszRelativePath);
                SecureHandle(this.hParsingName, ref d.pszParsingName);
                SecureHandle(this.hToolTip, ref d.pszToolTip);
                SecureHandle(this.hLocalizedName, ref d.pszLocalizedName);
                SecureHandle(this.hIcon, ref d.pszIcon);
                SecureHandle(this.hSecurity, ref d.pszSecurity);
            }

            [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
            private static void SecureHandle(SafeCoTaskMemHandle handle, ref IntPtr ptr)
            {
                handle.SetHandle(ptr);
                ptr = IntPtr.Zero;
            }

            public void Dispose()
            {
                this.hName.Dispose();
                this.hDescription.Dispose();
                this.hRelativePath.Dispose();
                this.hParsingName.Dispose();
                this.hToolTip.Dispose();
                this.hLocalizedName.Dispose();
                this.hIcon.Dispose();
                this.hSecurity.Dispose();
                GC.SuppressFinalize(this);
            }
        }
    }
}

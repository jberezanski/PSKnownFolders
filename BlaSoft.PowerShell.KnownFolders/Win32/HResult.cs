namespace BlaSoft.PowerShell.KnownFolders.Win32
{
    /// <summary>
    /// Common <c>HRESULT</c> values obtained from <see href="http://msdn.microsoft.com/en-us/windows/desktop/aa378137(v=vs.85).aspx" />.
    /// </summary>
    internal enum HResult : uint
    {
        /// <summary>
        /// Operation successful
        /// </summary>
        S_OK = 0x00000000,

        /// <summary>
        /// Some methods use S_FALSE to mean, roughly, a negative condition that is not a failure. 
        /// It can also indicate a "no-op"—the method succeeded, but had no effect.
        /// </summary>
        S_FALSE = 0x00000001,

        /// <summary>
        /// Operation aborted 
        /// </summary>
        E_ABORT = 0x80004004,

        /// <summary>
        /// General access denied error 
        /// </summary>
        E_ACCESSDENIED = 0x80070005,

        /// <summary>
        /// Unspecified failure 
        /// </summary>
        E_FAIL = 0x80004005,

        /// <summary>
        /// Handle that is not valid 
        /// </summary>
        E_HANDLE = 0x80070006,

        /// <summary>
        /// One or more arguments are not valid 
        /// </summary>
        E_INVALIDARG = 0x80070057,

        /// <summary>
        /// No such interface supported 
        /// </summary>
        E_NOINTERFACE = 0x80004002,

        /// <summary>
        /// Not implemented 
        /// </summary>
        E_NOTIMPL = 0x80004001,

        /// <summary>
        /// Failed to allocate necessary memory 
        /// </summary>
        E_OUTOFMEMORY = 0x8007000E,

        /// <summary>
        /// Pointer that is not valid 
        /// </summary>
        E_POINTER = 0x80004003,

        /// <summary>
        /// Unexpected failure 
        /// </summary>
        E_UNEXPECTED = 0x8000FFFF,

        WIN32_E_PATHNOTFOUND = 0x80070003,
    }
}

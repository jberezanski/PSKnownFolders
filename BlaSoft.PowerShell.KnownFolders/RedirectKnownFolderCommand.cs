using System;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    [Cmdlet(VerbsCommon.Move, "KnownFolder", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [OutputType(typeof(KnownFolder))]
    public sealed class RedirectKnownFolderCommand : PSCmdlet
    {
        private IKnownFolderManager knownFolderManager;

        private bool yesToAll;

        private bool noToAll;

        [Parameter(ParameterSetName = "KnownFolder", Mandatory = true, ValueFromPipeline = true)]
        [ValidateNotNull]
        public KnownFolder Folder { get; set; }

        [Parameter(ParameterSetName = "KnownFolder", Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty]
        public string Destination { get; set; }

        [Parameter]
        public SwitchParameter Force { get; set; }

        [Parameter]
        public SwitchParameter CheckOnly { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        protected override void BeginProcessing()
        {
            this.knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
        }

        protected override void ProcessRecord()
        {
            if (!this.Folder.CanRedirect)
            {
                throw new InvalidOperationException("Folder cannot be redirected.");
            }

            string currentPath = this.Folder.Path;
            var newPath = this.Destination;
            if (!this.ShouldProcess(this.Folder.Name, string.Format("Redirect from {0} to {1}", currentPath, newPath)))
            {
                return;
            }

            if (!this.Force && !this.ShouldContinue("Do it?", "Folder redirection", ref yesToAll, ref noToAll))
            {
                return;
            }

            var id = new KNOWNFOLDERID(this.Folder.FolderId);
            KF_REDIRECT_FLAGS flags = KF_REDIRECT_FLAGS.KF_REDIRECT_NONE;
            if (this.CheckOnly)
            {
                flags |= KF_REDIRECT_FLAGS.KF_REDIRECT_CHECK_ONLY;
            }
            else
            {
                flags |= KF_REDIRECT_FLAGS.KF_REDIRECT_COPY_CONTENTS | KF_REDIRECT_FLAGS.KF_REDIRECT_DEL_SOURCE_CONTENTS;
            }

            string error;
            HResult hr;
            hr = this.knownFolderManager.Redirect(
                ref id,
                IntPtr.Zero,
                flags,
                newPath,
                0,
                null,
                out error);
            if (hr != HResult.S_OK)
            {
                if (string.IsNullOrEmpty(error))
                {
                    Marshal.ThrowExceptionForHR(unchecked((int)hr));
                    this.WriteWarning(string.Format("Redirection returned success code other than S_OK ({0})", hr));
                }
                else
                {
                    throw new TargetInvocationException(
                        error,
                        Marshal.GetExceptionForHR(unchecked((int)hr)));
                }
            }

            if (this.PassThru)
            {
                this.WriteObject(this.Folder);
            }
        }
    }
}

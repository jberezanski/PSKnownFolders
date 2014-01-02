using System;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using System.Runtime.InteropServices;
using BlaSoft.PowerShell.KnownFolders.Win32;

namespace BlaSoft.PowerShell.KnownFolders
{
    [Cmdlet(VerbsCommon.Move, "KnownFolder", DefaultParameterSetName = "SingleFolder", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.High)]
    [OutputType(typeof(KnownFolder))]
    public sealed class RedirectKnownFolderCommand : PSCmdlet
    {
        private IKnownFolderManager knownFolderManager;

        private bool yesToAll;

        private bool noToAll;

        [Parameter(ParameterSetName = "SingleFolder", Mandatory = true, Position = 0)]
        public KnownFolder SingleFolder { get; set; }

        [Parameter(ParameterSetName = "SingleFolder", Mandatory = true, Position = 1)]
        public string NewPath { get; set; }

        [Parameter(ParameterSetName = "MultipleFolders", Mandatory = true, ValueFromPipeline = true)]
        public KnownFolder Folder { get; set; }

        [Parameter(ParameterSetName = "MultipleFolders", Mandatory = true, Position = 0)]
        public string Destination { get; set; }

        [Parameter]
        public SwitchParameter Force { get; set; }

        [Parameter]
        public SwitchParameter CheckOnly { get; set; }

        [Parameter]
        public SwitchParameter PassThru { get; set; }

        [Parameter]
        public SwitchParameter DontMoveExistingData { get; set; }

        protected override void BeginProcessing()
        {
            this.knownFolderManager = (IKnownFolderManager)new KnownFolderManager();
        }

        protected override void ProcessRecord()
        {
            switch (this.ParameterSetName)
            {
                case "SingleFolder":
                    this.RedirectOneFolder(this.SingleFolder, this.NewPath);
                    break;
                case "MultipleFolders":
                    var newPath = DetermineNewPath(this.Folder, this.Destination);
                    this.RedirectOneFolder(this.Folder, newPath);
                    break;
                default:
                    throw new ArgumentException("Unsupported parameter set name: " + this.ParameterSetName);
            }
        }

        private string DetermineNewPath(KnownFolder knownFolder, string destinationPath)
        {
            var existingPath = knownFolder.Path;
            if (string.IsNullOrEmpty(existingPath))
            {
                throw new InvalidOperationException(string.Format("Unable to determine new path of folder '{0}': existing path is empty.", knownFolder.Name));
            }

            var existingDirectoryName = Path.GetFileName(existingPath);
            if (string.IsNullOrEmpty(existingDirectoryName))
            {
                throw new InvalidOperationException(string.Format("Unable to determine new path of folder '{0}': cannot determine existing directory name.", knownFolder.Name));
            }

            return Path.Combine(destinationPath, existingDirectoryName);
        }

        private void RedirectOneFolder(KnownFolder folder, string newPath)
        {
            if (!folder.CanRedirect)
            {
                throw new InvalidOperationException(string.Format("Folder '{0}' cannot be redirected.", folder.Name));
            }

            string currentPath = folder.Path;
            if (!this.ShouldProcess(folder.Name, string.Format("Redirect from {0} to {1}", currentPath, newPath)))
            {
                return;
            }

            if (!this.Force && !this.ShouldContinue("Do it?", "Folder redirection", ref yesToAll, ref noToAll))
            {
                return;
            }

            var id = new KNOWNFOLDERID(folder.FolderId);
            KF_REDIRECT_FLAGS flags = KF_REDIRECT_FLAGS.KF_REDIRECT_NONE;
            if (this.CheckOnly)
            {
                flags |= KF_REDIRECT_FLAGS.KF_REDIRECT_CHECK_ONLY;
            }

            if (!this.DontMoveExistingData)
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
                this.WriteObject(folder);
            }
        }
    }
}

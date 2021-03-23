[CmdletBinding(SupportsShouldProcess = $true)]
Param
(
    [switch] $Production,
    [PSCredential] $NuGetApiKeyCredential,
    [string] $ModulePath = "$PSScriptRoot\Output\PSModules\KnownFolders"
)

Set-StrictMode -Version 5
$ErrorActionPreference = 'Stop'

$verboseValue = $VerbosePreference -ne 'SilentlyContinue'

if ($Production)
{
    $repository = 'PSGallery'
}
else
{
    $repository = 'JBMyGet'
    $repositorySourceUrl = 'https://www.myget.org/F/jberezanski-psmodules/api/v2'
    $repositoryPublishUrl = 'https://www.myget.org/F/jberezanski-psmodules/api/v2/package'
    if ($null -eq (Get-PSRepository -Name $repository -ErrorAction SilentlyContinue))
    {
        if ($PSCmdlet.ShouldProcess("PSRepository $repository", 'Register'))
        {
            Register-PSRepository -Name $repository -SourceLocation $repositorySourceUrl -PublishLocation $repositoryPublishUrl -Verbose:$verboseValue
        }
    }
}

Resolve-Path -Path $ModulePath | Out-Null

if ($null -eq $NuGetApiKeyCredential)
{
    $NuGetApiKeyCredential = Get-Credential -Credential 'API key'
}

$plainApiKey = $NuGetApiKeyCredential.GetNetworkCredential().Password

$whatIfValue = $true
if ($PSCmdlet.ShouldProcess($ModulePath, "Publish to repository $repository"))
{
    $whatIfValue = $false
}

Publish-Module -Path $ModulePath -NuGetApiKey $plainApiKey -Repository $repository -Verbose:$verboseValue -WhatIf:$whatIfValue

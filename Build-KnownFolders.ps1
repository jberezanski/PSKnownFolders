Remove-Item -Path .\Output -Recurse -Force
nuget restore .\KnownFolders.sln
msbuild /m /v:m .\KnownFolders.sln /p:Configuration=Release
robocopy .\BlaSoft.PowerShell.KnownFolders\bin\Release .\Output\PSModules\KnownFolders /E /XF System.Management.Automation.dll

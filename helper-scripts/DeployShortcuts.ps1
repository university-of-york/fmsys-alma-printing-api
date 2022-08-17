#Requires -RunAsAdministrator
#Requires -Version 4.0
<#
  .SYNOPSIS
  A script to deploy shortcuts to 'shell:common startup' and 'shell:common desktop' locations.
  Also creates a new system environment variable to hold the powershell -Command value.
#>
$envVariableName = "ALMA_PRINTING_CMD"
$shortcutFilename = read-host "Please type a shortcut filename (note the Powershell window title will carry this name too)"
$startup=[Environment]::GetFolderPath("CommonStartup")
$desktop=[Environment]::GetFolderPath("CommonDesktop")
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$desktop\$shortcutFilename.lnk")
$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$Shortcut.Arguments  = "-NoLogo -NoProfile -Command %$envVariableName%"
$Shortcut.WorkingDirectory  = $(Get-Item $PSCommandPath ).DirectoryName | Split-Path
$Shortcut.WindowStyle = 7
$Shortcut.Save()
Copy-Item "$desktop\$shortcutFilename.lnk" -Destination "$startup"
"The next step is to create the system environment variable $envVariableName"
$envVariableValue = read-host "Paste in your Powershell -Command value (including the surrounding double-quotes) to populate $envVariableName"
[Environment]::SetEnvironmentVariable($envVariableName, $envVariableValue, "Machine")

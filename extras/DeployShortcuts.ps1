#Requires -RunAsAdministrator
#Requires -Version 4.0
$envVariableName = "ALMA_PRINTING_CMD"
$shortcutFilename = read-host "Please type a shortcut filename"
$startup=[Environment]::GetFolderPath("CommonStartup")
$desktop=[Environment]::GetFolderPath("CommonDesktop")
$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$desktop\$shortcutFilename.lnk")
$Shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
$Shortcut.Arguments  = "-NoLogo -Command %$envVariableName%"
$Shortcut.WorkingDirectory  = $(Get-Item $PSCommandPath ).DirectoryName | Split-Path
$Shortcut.WindowStyle = 7
$Shortcut.Save()
Copy-Item "$desktop\$shortcutFilename.lnk" -Destination "$startup"
$envVariableValue = read-host "Paste in your params"
[Environment]::SetEnvironmentVariable($envVariableName, $envVariableValue, "Machine")

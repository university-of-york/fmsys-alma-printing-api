#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
  .SYNOPSIS
  A script to deploy shortcuts to 'shell:common startup' and 'shell:common desktop' filesystem locations.

  Usage:
  Run in an elevated Powershell window, e.g.:
  .\DeployShortcuts.ps1 -Deploy -ShortcutFilename 'Alma Interlending Printing' `
  -ShortcutArguments "-NoLogo -NoProfile -Command `"& {Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '19195349880001381' -localPrinterName 'PUSH_ITSPRN0705 [Harry Fairhurst - Information Services LFA/ LFA023](Mobility)' -marginTop '0.3' -jpgBarcode}`""

  For the -ShortcutArguments parameter, you'll need to escape any double-quotes with backticks.

  .PARAMETER Deploy
  Add this switch to copy the shortcut to the Common Startup and Common Desktop directories [switch]

  .PARAMETER ShortcutFilename
  The filename you want to give the shortcut. Default is 'Alma printing' [string]

  .PARAMETER ShortcutArguments
  The arguments that follow the full path to powershell.exe. This ends up appended to the 'Target' field [mandatory] [string]
  Note that when viewing the shortcut properties in Explorer, the Target field display is limited to 260 characters. This does not mean that the field data is incorrect; it's an Explorer display limitation only.

  .PARAMETER ShortcutWorkingDirectory
  The Working Directory a.k.a 'Start in' folder. This defaults to the root directory of this repo [string]

  .PARAMETER VerifyOnly
  Add this switch to check the LNK field content for an existing shortcut [switch]

  .PARAMETER VerifyFilePath
  The full file path to the LNK file that you want to verify [string]

#>
[CmdletBinding(DefaultParameterSetName = "create")]
param (
    # The Powershell window title will carry the $ShortcutFilename too
    [Parameter(ParameterSetName="create")]
    [string]$ShortcutFilename = 'Alma printing',

    [Parameter(Mandatory, ParameterSetName="create")]
    [string]$ShortcutArguments,

    [Parameter(ParameterSetName="create")]
    [string]$ShortcutWorkingDirectory,

    [Parameter(ParameterSetName="create")]
    [switch]$Deploy,

    [Parameter(ParameterSetName="verify")]
    [switch]$VerifyOnly,

    [Parameter(Mandatory, ParameterSetName="verify")]
    [string]$VerifyFilePath
)

$commonStartup = [Environment]::GetFolderPath("CommonStartup")
$commonDesktop = [Environment]::GetFolderPath("CommonDesktop")
$WshShell = New-Object -comObject WScript.Shell

If (-Not($VerifyOnly)) {
    $shortcutPath = "${env:TMP}\tmp$([convert]::tostring((get-random 65535),16).padleft(4,'0')).lnk"
  }
  Else {
    If (Test-Path -Path $VerifyFilePath -PathType Leaf) {
      $shortcutPath = $VerifyFilePath
    } Else {
      Throw '-VerifyFilePath does not exist'
    }
  }

$shortcut = $WshShell.CreateShortcut($shortcutPath)

If (-Not($VerifyOnly)) {
  # Windows auto-strips double-quotes from unspaced file paths, as well as auto-adding the full path to a given executable when just the filename is supplied
  # See https://stackoverflow.com/questions/31815286/creating-quoted-path-for-shortcut-with-arguments-in-powershell
  $shortcut.TargetPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
  $Shortcut.Arguments  = $ShortcutArguments
  If ([string]::IsNullOrEmpty($ShortcutWorkingDirectory)) {
    # Set the working directory to the root directory; this script is in a subfolder
    $shortcut.WorkingDirectory  = $(Get-Item $PSCommandPath).DirectoryName | Split-Path
  } Else {
    # Check that the user-specified directory exists
    If (Test-Path $ShortcutWorkingDirectory -PathType Container) {
      $shortcut.WorkingDirectory  = $ShortcutWorkingDirectory
    } Else {
      Throw 'Working directory does not exist'
    }
  }
  # Launch in a minimised window
  $shortcut.WindowStyle = 7
  $shortcut.Save()
}

If ($Deploy) {
  Copy-Item $shortcutPath -Destination "${commonDesktop}\${shortcutFilename}.lnk" -Verbose
  Copy-Item $shortcutPath -Destination "${commonStartup}\${shortcutFilename}.lnk" -Verbose
}

'For your information, here are the properties of the LNK file:'
"Arguments: $($shortcut.Arguments)"
"Description: $($shortcut.Description)"
"Hotkey: $($shortcut.Hotkey)"
"Icon Location: $($shortcut.IconLocation)"
"Link Path:  $($shortcut.FullName)"
"Target: $(Try {Split-Path $shortcut.TargetPath -Leaf } Catch { 'n/a'})"
"Target Path: $($shortcut.TargetPath)"
"Link: $(Try {Split-Path $shortcut.LinkPath -Leaf } Catch { 'n/a'})"
"Window Style: $($shortcut.WindowStyle)"
"Working Directory: $($shortcut.WorkingDirectory)"
$shortcut = $null
$WshShell = $null

If (-Not($VerifyOnly)) {
  Remove-Item -Path $shortcutPath -Verbose
}

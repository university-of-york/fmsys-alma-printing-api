#Requires -RunAsAdministrator
#Requires -Version 4.0
# Disable the 'Set up Internet Explorer 11' box - code snagged from https://stackoverflow.com/a/52985845/1754517
$keyPath = 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Main'
$valueName = 'DisableFirstRunCustomize'
$valueData = 1
if (!(Test-Path $keyPath)) {
  "Creating subkey(s)"
  $null = New-Item $keyPath -Force
}
"Creating $valueName value"
Set-ItemProperty -Path $keyPath -Name $valueName -Value $valueData
"Done"

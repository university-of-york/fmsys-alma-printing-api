#Requires -RunAsAdministrator
# Disable the 'Set up Internet Explorer 11' box - code snagged from https://stackoverflow.com/a/52985845/1754517
$keyPath = 'Registry::HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Internet Explorer\Main'
if (!(Test-Path $keyPath)) {
    New-Item $keyPath -Force
}
Set-ItemProperty -Path $keyPath -Name "DisableFirstRunCustomize" -Value 1

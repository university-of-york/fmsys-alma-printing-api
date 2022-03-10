#Requires -Version 3.0
#
# The first thing to do is to perform the one-time step of saving your API key to a file for ongoing use.
# Enter:
# Set-Location <script-dir>
# . .\FetchAlmaPrint.ps1;Invoke-Setup

param (
  # North America = na, Europe = eu (default), Asia Pacific = ap, Canada = ca, China = cn ( . .\FetchAlmaPrint.ps1 -apiRegion "na" )
  [string]$apiRegion = "eu"
)

$apiBaseUrl = -join ("https://api-",$apiRegion,".hosted.exlibrisgroup.com")
$printoutsApiUrlPath = "/almaws/v1/task-lists/printouts?"
$tmpPrintoutsPath = "$PSScriptRoot\tmp_printouts"
$apiKeysPath = "$PSScriptRoot\auth"

# . .\FetchAlmaPrint.ps1;Invoke-Setup
function Invoke-Setup {
  "Script setup"
  If ((Test-Path -Path $apiKeysPath) -ne $true) {
      $null = New-Item -Type 'directory' -Path "$apiKeysPath" -Force
  }
  $apikey = read-host "Enter the Ex Libris Alma API key"
  $apikey | export-clixml -Path "$apiKeysPath\apikey.xml" -Force
  If ((Test-Path -Path $tmpPrintoutsPath) -ne $true) {
      $null = New-Item -Type 'directory' -Path $tmpPrintoutsPath -Force
  }
}

# . .\FetchAlmaPrint.ps1;Fetch-Printers
function Fetch-Printers {
  $fetchPrintersApiUrlPath = "/almaws/v1/conf/printers?"
  $fetchPrintersApiUrlParameters = -join ("library=ALL&printout_queue=ALL&name=ALL&code=ALL&limit=100&offset=0")
  $fetchPrintersApiFullUrl = -join ($apiBaseUrl,$fetchPrintersApiUrlPath,$fetchPrintersApiUrlParameters)
  # Use grouping operator per https://education.launchcode.org/azure/chapters/powershell-intro/cmdlet-invoke-restmethod.html#grouping-to-access-fields-of-the-json-response
  (Invoke-RestMethod -Uri $fetchPrintersApiFullUrl -Method Get -Headers (getHeaders)).printer | Format-Table `
  @{N='ID';E={ $_.id };width=20}, `
  @{N='Code';E={ $_.code };width=10}, `
  @{N='Name';E={ $_.name };width=30}, `
  @{N='Description';E={$_.description};width=40}, `
  @{N='Email';E={ $_.email };width=25}, `
  @{N='Queue';E={ $_.printout_queue };width=5}, `
  @{N='Library';E={ $_.library.desc };width=30}
}

# . .\FetchAlmaPrint.ps1;Fetch-Jobs -printerId "848838010001381" -localPrinterName "EPSON TM-T88III Receipt" -printStatuses "ALL"
function Fetch-Jobs(
  [parameter(mandatory)] [string]$printerId,
  [string]$localPrinterName = "EPSON TM-T88III Receipt",
  [int]$checkInterval = 30,
  [string]$marginTop = "0.000000",
  [string]$marginBottom = "0.000000",
  [string]$marginLeft = "0.155560",
  [string]$marginRight = "0.144440",
  [string]$printStatuses = "PENDING") {

  if ($localPrinterName -ne (Get-WmiObject -Class Win32_Printer -Filter "Name='$localPrinterName'").Name) {
    Write-Host "The printer specified was not found" -ForegroundColor red
    return
  }

  if (-not (Test-Path "$apiKeysPath\apikey.xml")) {
    Write-Host "The apikey.xml file doesn't exist" -ForegroundColor red
    return
  }

  # To avoid 'The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)' errors, enable IE Protected Mode for the 'Local Intranet' zone
  # See https://www.reddit.com/r/PowerShell/comments/2rsr8b/automating_ie_getting_strange_behaviour_when/
  $protectedModeKeyPath = 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1'
  $protectedModeValueName = '2500'
  $protectedModeValueData = 0
  $protectedModeValueEval = (Get-ItemProperty -Path $protectedModeKeyPath -ErrorAction SilentlyContinue).$protectedModeValueName
  if ($protectedModeValueEval -ne $protectedModeValueData) {
    Set-ItemProperty -Path $protectedModeKeyPath -Name $protectedModeValueName -Value $protectedModeValueData
  }

  $fetchJobsApiUrlParameters = -join ("letter=ALL&status=",$printStatuses,"&printer_id=",$printerId)
  $fetchJobsApiFullUrl = -join ($apiBaseUrl,$printoutsApiUrlPath,$fetchJobsApiUrlParameters)
  "Beginning at $(Get-Date -UFormat "%A %d/%m/%Y %T")"
  $script:RegPath = "HKCU:\Software\Microsoft\Internet Explorer\PageSetup"
  backupPageSetup

while ($true) {
  # Do Page Setup stuff
  setPageSetup $marginTop $marginBottom $marginLeft $marginRight

  $correctPrinter = $false
  $defaultPrinterChanged = $false
  # If the specified printer is the default printer, set $correctPrinter to $true
  if ($localPrinterName -eq (getDefaultPrinter)) {
    $correctPrinter = $true
  }
  "Working.."
  $letterRequest = (Invoke-RestMethod -Uri $fetchJobsApiFullUrl -Method Get -Headers (getHeaders)).printout
  ForEach($letter in $letterRequest) {
    if ($correctPrinter -eq $false) {
      setDefaultPrinter($localPrinterName)
      $defaultPrinterChanged = $true
    }
    $ie = new-object -com "InternetExplorer.Application"
    $letterId = $letter.id
    $letterHtml = $letter.letter
    $outputFilename = -Join ("document-",$letterId,".html")
    $printOut = -Join ($tmpPrintoutsPath,'\',$outputFilename)
    # Ridiculous hack to get around "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" - see https://stackoverflow.com/a/721519/1754517
    "<!-- saved from url=(0016)http://localhost -->`r`n" + $letterHtml | Out-File -FilePath "$printOut"
    # Begin printing
    $ie.Navigate($printOut)
    Start-Sleep -seconds 3
    "$(Get-Date -UFormat "%A %d/%m/%Y %T") - printing $outputFilename"
    $ie.ExecWB(6,2)
    Start-Sleep -seconds 3
    $ie.quit()
    # Done
    markAsPrinted $letterId
  }

  resetPageSetup
  # Make the original default printer the default again
  if ($defaultPrinterChanged -eq $true) {
    setDefaultPrinter($defaultPrinter)
  }

  "Finished..going to sleep for $checkInterval seconds. Press CTRL+C to quit."
   $i = $checkInterval
   do {
    Write-Host -NoNewline "`rRestarting in $i seconds."
    Start-Sleep -seconds 1
    $i--
  } while ($i -gt 0)
  ""
  }
}

function getHeaders {
  $apikey = Import-Clixml -Path "$apiKeysPath\apikey.xml"
  return $headers = @{
    'Accept' = 'application/json'
    'Authorization' = "apikey $apikey"
    }
}

function markAsPrinted ([string]$letterId){
  $markAsPrintedApiUrlParameters = -join ("letter=ALL&status=ALL&printout_id=",$letterId,"&op=mark_as_printed")
  $markAsPrintedApiFullUrl = -join ($apiBaseUrl,$printoutsApiUrlPath,$markAsPrintedApiUrlParameters)
  Write-Information -MessageData "$(Get-Date -UFormat "%A %d/%m/%Y %T") - marking letter ID $letterId as printed" -InformationAction Continue
  $null = Invoke-RestMethod -Uri $markAsPrintedApiFullUrl -Method Post -Headers (getHeaders)
}

function getDefaultPrinter {
    $script:defaultPrinter = Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default=$true" | select object -ExpandProperty Name
}

function setDefaultPrinter ([string]$printerName){
    $null = (Get-WmiObject -Class Win32_Printer -Filter "Name='$printerName'").SetDefaultPrinter()
}

function backupPageSetup {
  # Get PSObject excluding PS properties (this works when the value names collide)
  Get-Item $RegPath | ForEach-Object {
      $RegKey = $_
      $script:PropertyHash = @{}
      $_.GetValueNames() -replace "^$", "(default)" | ForEach-Object {
          $script:PropertyHash.$_ = $RegKey.GetValue($_)
      }
      $null = New-Object PSObject -Property $script:PropertyHash
  }
}

function setPageSetup ([string]$marginTop, [string]$marginBottom, [string]$marginLeft, [string]$marginRight){
  # Set the values we need for the print job
  # "MarginTop is $marginTop, MarginBottom is $MarginBottom, MarginLeft is $MarginLeft, MarginRight is $MarginRight"
  Set-ItemProperty -Path $RegPath -Name "margin_bottom" -Value $marginBottom -Type "String"
  Set-ItemProperty -Path $RegPath -Name "margin_top" -Value $marginTop -Type "String"
  Set-ItemProperty -Path $RegPath -Name "margin_left" -Value $marginLeft -Type "String"
  Set-ItemProperty -Path $RegPath -Name "margin_right" -Value $marginRight -Type "String"
  Set-ItemProperty -Path $RegPath -Name "header" -Value "" -Type "String"
  Set-ItemProperty -Path $RegPath -Name "footer" -Value "" -Type "String"
}

function resetPageSetup {
  # Delete all properties (as these instances of the properties didn't exist before)
  Remove-ItemProperty -Path $RegPath -Name "margin_bottom"
  Remove-ItemProperty -Path $RegPath -Name "margin_top"
  Remove-ItemProperty -Path $RegPath -Name "margin_left"
  Remove-ItemProperty -Path $RegPath -Name "margin_right"
  Remove-ItemProperty -Path $RegPath -Name "header"
  Remove-ItemProperty -Path $RegPath -Name "footer"

  if ($PropertyHash.Count -gt 0 ) {
    # Restore properties to what they were before
    $PropertyHash.GetEnumerator() | ForEach-Object {
        Set-ItemProperty -Path $RegPath -Name $_.key -Value $_.value -Type "String"
    }
  }
}

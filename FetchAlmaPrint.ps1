#Requires -Version 5.1
<#
  .SYNOPSIS
  Powershell script to facilitate printing of Ex Libris Alma generated printouts

  .DESCRIPTION
  The two main functions of FetchAlmaPrint.ps1 are to (1) retrieve a list of available Alma printers and (2) query, download and print any pending printouts

  .PARAMETER apiRegion
  Specifies the API region code. Available codes are:
    North America = na
    Europe = eu (default)
    Asia Pacific = ap
    Canada = ca
    China = cn

  .EXAMPLE
  PS> . .\FetchAlmaPrint.ps1 -apiRegion "na"
#>

param (
  [string]$apiRegion = 'eu'
)

Add-Type -AssemblyName system.drawing
$apiBaseUrl = -join ('https://api-', $apiRegion, '.hosted.exlibrisgroup.com')
$printoutsApiUrlPath = '/almaws/v1/task-lists/printouts?'
$tmpPrintoutsPath = "$PSScriptRoot\tmp_printouts"
$apiKeysPath = "$PSScriptRoot\auth"

function Invoke-Setup {
  <#
  .SYNOPSIS
  A function to run one time, in order to store the Ex Libris API key to a file for later use, and to ensure the tmp_printouts directory exists for storing the HTML printout files.

  .EXAMPLE
  PS> . .\FetchAlmaPrint.ps1;Invoke-Setup
  #>
  'Script setup'
  If ((Test-Path -Path $apiKeysPath) -ne $true) {
    $null = New-Item -Type 'directory' -Path "$apiKeysPath" -Force
  }
  While ($null -eq $apikey -or $apikey.Length -eq 0) {
    $apikey = Read-Host -AsSecureString -Prompt 'Enter the Ex Libris Alma API key'
  }
  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($apikey)
  $plainApikey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
  $plainApikey | Export-Clixml -Path "$apiKeysPath\apikey.xml" -Force
  If ((Test-Path -Path $tmpPrintoutsPath) -ne $true) {
    $null = New-Item -Type 'directory' -Path $tmpPrintoutsPath -Force
  }
}

function Fetch-Printers {
  <#
  .SYNOPSIS
  Retrieve a list of available Alma printers for use with the Fetch-Jobs function.
  The printer ID should be saved for later, as this is a required named parameter of the Fetch-Jobs function.

  .EXAMPLE
  PS> . .\FetchAlmaPrint.ps1;Fetch-Printers
  #>
  $fetchPrintersApiUrlPath = '/almaws/v1/conf/printers?'
  $fetchPrintersApiUrlParameters = -join ('library=ALL&printout_queue=ALL&name=ALL&code=ALL&limit=100&offset=0')
  $fetchPrintersApiFullUrl = -join ($apiBaseUrl, $fetchPrintersApiUrlPath, $fetchPrintersApiUrlParameters)
  # Use grouping operator per https://education.launchcode.org/azure/chapters/powershell-intro/cmdlet-invoke-restmethod.html#grouping-to-access-fields-of-the-json-response
  (Invoke-RestMethod -Uri $fetchPrintersApiFullUrl -Method Get -Headers (getHeaders)).printer | Format-Table `
  @{N = 'ID'; E = { $_.id }; width = 20 }, `
  @{N = 'Code'; E = { $_.code }; width = 10 }, `
  @{N = 'Name'; E = { $_.name }; width = 30 }, `
  @{N = 'Description'; E = { $_.description }; width = 40 }, `
  @{N = 'Email'; E = { $_.email }; width = 25 }, `
  @{N = 'Queue'; E = { $_.printout_queue }; width = 5 }, `
  @{N = 'Library'; E = { $_.library.desc }; width = 30 }
}

function Fetch-Jobs(
  [parameter(mandatory)] [string[]]$printerId,
  [string]$localPrinterName = 'EPSON TM-T88III Receipt',
  [int]$checkInterval = 30,
  [string]$marginTop = '0.000000',
  [string]$marginBottom = '0.000000',
  [string]$marginLeft = '0.155560',
  [string]$marginRight = '0.144440',
  [string]$printoutsWithStatus = 'PENDING',
  [switch]$jpgBarcode,
  [switch]$multiInstanceOverride) {
  <#
  .SYNOPSIS
  Query, download and print any printouts according to status.

  .PARAMETER printerId
  This is the ID of the Alma printer returned by the Fetch-Printers function. Multiple IDs can be specified, in which case you'll need to comma delimit these.

  .PARAMETER localPrinterName
  This is the name of the physical printer to which the printouts should be sent.

  .PARAMETER checkInterval
  This is the interval in seconds that defines the frequency of checking for pending printouts.

  .PARAMETER marginTop
  This is the marginTop value which sets HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup → margin_top. The value's unit of measurement is inches, which differs from IE's Page Setup dialogue which is in millimetres.

  .PARAMETER marginBottom
  This is the marginBottom value which sets HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup → margin_bottom. The value's unit of measurement is inches, which differs from IE's Page Setup dialogue which is in millimetres.

  .PARAMETER marginLeft
  This is the marginLeft value which sets HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup → margin_left. The value's unit of measurement is inches, which differs from IE's Page Setup dialogue which is in millimetres.

  .PARAMETER marginRight
  This is the marginRight value which sets HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup → margin_right. The value's unit of measurement is inches, which differs from IE's Page Setup dialogue which is in millimetres.

  .PARAMETER printoutsWithStatus
  This determines which of the available statuses of printouts to check for.
  Valid values are:
    Pending (default)
    Printed
    Canceled
    ALL

  .PARAMETER jpgBarcode
  This switch, when used, will replace the base64-encoded PNG barcode data (if found) with base64-encoded JPG data, to ensure it is readable by barcode scanners.

  .PARAMETER multiInstanceOverride
  This switch, when used, will allow multiple instance of the Fetch-Jobs function to run concurrently. Ordinarily this is undesirable because duplicate printouts will be printed.

  .EXAMPLE
  PS> . .\FetchAlmaPrint.ps1;Fetch-Jobs -printerId '14195349480001361' -localPrinterName 'HP Laserjet Pro - Basement' -checkInterval 15 -marginTop '0.3' -marginBottom '0.3' -marginLeft '0.3' -marginRight '0.3' -jpgBarcode
  #>

  if (-not ($multiInstanceOverride)) {
    $scriptName = $(Get-Item $PSCommandPath ).Name
    $scriptInstances = (Get-CimInstance -ClassName Win32_Process | Select-Object CommandLine | Where-Object { $_ -ilike "*${scriptName}*" } | Measure-Object).Count
    if ($scriptInstances -gt 1 ) {
      Throw 'The script is apparently already running in the background'
    }
  }

  if ($localPrinterName -ne (Get-CimInstance -ClassName Win32_Printer -Filter "Name='$localPrinterName'").Name) {
    Throw 'The printer specified was not found'
  }

  if (-not (Test-Path "$apiKeysPath\apikey.xml")) {
    Throw 'The apikey.xml file doesn''t exist'
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

  ForEach ($id in $printerId) {
    $fetchJobsApiUrlParameters = -join ('letter=ALL&status=', $printoutsWithStatus, '&printer_id=', $id)
    [string[]]$fetchJobsApiFullUrl += -join ($apiBaseUrl, $printoutsApiUrlPath, $fetchJobsApiUrlParameters)
  }

  "Beginning at $(Get-Date -UFormat '%A %d/%m/%Y %T')"
  $script:RegPath = 'HKCU:\Software\Microsoft\Internet Explorer\PageSetup'
  backupPageSetup

  while ($true) {
    'Checking Online Queue...'
    ForEach ($url in $fetchJobsApiFullUrl) {
      $letterResponse = (Invoke-RestMethod -Uri $url -Method Get -Headers (getHeaders)).printout

      $defaultPrinterChanged = $false
      if ($null -ne $letterResponse) {
        # If the specified printer isn't the default printer, make it the default printer
        if ($localPrinterName -ne (getDefaultPrinter)) {
          setDefaultPrinter($localPrinterName)
          $defaultPrinterChanged = $true
        }
        # Do Page Setup stuff
        setPageSetup $marginTop $marginBottom $marginLeft $marginRight
      }

      ForEach ($letter in $letterResponse) {
        $ie = New-Object -com 'InternetExplorer.Application'
        $letterId = $letter.id
        if ($jpgBarcode) {
          $letterHtml = base64Png2Jpg $letter.letter
        }
        else {
          $letterHtml = $letter.letter
        }
        $outputFilename = -Join ('document-', $letterId, '.html')
        $printOut = -Join ($tmpPrintoutsPath, '\', $outputFilename)
        # Ridiculous hack to get around "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" - see https://stackoverflow.com/a/721519/1754517
        "<!-- saved from url=(0016)http://localhost -->`r`n" + $letterHtml | Out-File -Encoding utf8 -FilePath "$printOut"
        # Begin printing
        $ie.Navigate($printOut)
        Start-Sleep -Seconds 3
        "$(Get-Date -UFormat '%A %d/%m/%Y %T') - printing $outputFilename"
        $ie.ExecWB(6, 2)
        Start-Sleep -Seconds 3
        $ie.quit()
        # Done
        markAsPrinted $letterId
      }

      # Restore the original default printer
      if ($defaultPrinterChanged -eq $true) {
        setDefaultPrinter($originalDefaultPrinter)
      }

      # Restore the margins, etc
      if ($null -ne $letterResponse) {
        restorePageSetup
      }
    }

    "Finished..going to sleep for ${checkInterval} seconds. Press CTRL+C to quit."
    $i = $checkInterval
    do {
      Write-Host -NoNewline "`rRestarting in ${i} seconds."
      Start-Sleep -Seconds 1
      $i--
    } while ($i -gt 0)
    ''
  }
}

function getHeaders {
  <#
  .SYNOPSIS
   A function to return the required HTTP headers for inclusion in Invoke-RestMethod requests.
  #>
  $apikey = Import-Clixml -Path "$apiKeysPath\apikey.xml"
  return @{
    'Accept'        = 'application/json'
    'Authorization' = "apikey $apikey"
  }
}

function markAsPrinted ([string]$letterId) {
  <#
  .SYNOPSIS
   A function to mark the printed printout with status 'Printed'.
  .PARAMETER letterId
  This is the Alma letterId of the printout to be marked as 'Printed'.
  #>
  $markAsPrintedApiUrlParameters = -join ('letter=ALL&status=ALL&printout_id=', $letterId, '&op=mark_as_printed')
  $markAsPrintedApiFullUrl = -join ($apiBaseUrl, $printoutsApiUrlPath, $markAsPrintedApiUrlParameters)
  Write-Information -MessageData "$(Get-Date -UFormat '%A %d/%m/%Y %T') - marking letter ID $letterId as printed" -InformationAction Continue
  $null = Invoke-RestMethod -Uri $markAsPrintedApiFullUrl -Method Post -Headers (getHeaders)
}

function getDefaultPrinter {
  <#
  .SYNOPSIS
   A function to get the Windows default printer, and populate a script-level variable with its name.
  #>
  $script:originalDefaultPrinter = Get-CimInstance -Query "SELECT * FROM Win32_Printer WHERE Default=$true" | Select-Object -ExpandProperty Name
}

function setDefaultPrinter ([string]$printerName) {
  <#
  .SYNOPSIS
   A function to set the Windows default printer.
   .PARAMETER printerName
   This is the Windows printer name as listed in System settings > Printers & Scanners.
  #>
  $null = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$printerName'" | Invoke-CimMethod -MethodName SetDefaultPrinter
}

function backupPageSetup {
  <#
  .SYNOPSIS
   A function to make a backup of existing Page Setup registry values, for restoration after printing.
  #>
  Get-Item $RegPath | ForEach-Object {
    $RegKey = $_
    $script:PropertyHash = @{}
    $_.GetValueNames() -replace '^$', '(default)' | ForEach-Object {
      $script:PropertyHash.$_ = $RegKey.GetValue($_)
    }
    $null = New-Object PSObject -Property $script:PropertyHash
  }
}

function setPageSetup ([string]$marginTop, [string]$marginBottom, [string]$marginLeft, [string]$marginRight) {
  <#
  .SYNOPSIS
  A function to set the Page Setup values according to what is required for printing in the current environment.
  #>
  Set-ItemProperty -Path $RegPath -Name 'margin_bottom' -Value $marginBottom -Type 'String'
  Set-ItemProperty -Path $RegPath -Name 'margin_top' -Value $marginTop -Type 'String'
  Set-ItemProperty -Path $RegPath -Name 'margin_left' -Value $marginLeft -Type 'String'
  Set-ItemProperty -Path $RegPath -Name 'margin_right' -Value $marginRight -Type 'String'
  Set-ItemProperty -Path $RegPath -Name 'header' -Value '' -Type 'String'
  Set-ItemProperty -Path $RegPath -Name 'footer' -Value '' -Type 'String'
}

function restorePageSetup {
  <#
  .SYNOPSIS
  A function to restore the Page Setup values as they were before printing.
  #>
  # Delete all properties (as these instances of the properties didn't exist before)
  Remove-ItemProperty -Path $RegPath -Name 'margin_bottom'
  Remove-ItemProperty -Path $RegPath -Name 'margin_top'
  Remove-ItemProperty -Path $RegPath -Name 'margin_left'
  Remove-ItemProperty -Path $RegPath -Name 'margin_right'
  Remove-ItemProperty -Path $RegPath -Name 'header'
  Remove-ItemProperty -Path $RegPath -Name 'footer'

  if ($PropertyHash.Count -gt 0 ) {
    # Restore properties to what they were before
    $PropertyHash.GetEnumerator() | ForEach-Object {
      Set-ItemProperty -Path $RegPath -Name $_.key -Value $_.value -Type 'String'
    }
  }
}

function base64Png2Jpg ([string]$html) {
  <#
  .SYNOPSIS
  A function to replace the base64-encoded PNG barcode data with the equivalent base64-encoded JPG data.
  This function was created after finding that the barcodes were not readable when embedded in the HTML as base64-encoded PNG data.

  .PARAMETER html
  This is the HTML data to search and replace.
  #>

  $pattern = '<img src="(data:image/\.png;base64,([A-Za-z0-9-+/]*={0,3}))" alt="Item Barcode">'
  $srcAttribute = Select-String -InputObject $html -Pattern $pattern | ForEach-Object { $_.matches.groups[1] } | Select-Object -ExpandProperty Value
  $base64PngMatchString = Select-String -InputObject $html -Pattern $pattern | ForEach-Object { $_.matches.groups[2] } | Select-Object -ExpandProperty Value

  If ($null -ne $srcAttribute -and $null -ne $base64PngMatchString) {
    $oMemoryStream = New-Object -TypeName System.IO.MemoryStream
    $oImgFormat = [System.Drawing.Imaging.ImageFormat]::Jpeg
    $pngBytes = [System.Convert]::FromBase64String($base64PngMatchString)
    $pngMemoryStream = New-Object -TypeName System.IO.MemoryStream -ArgumentList (,$pngBytes)
    $Image = [System.Drawing.Bitmap]::FromStream($pngMemoryStream)
    $Image.Save($oMemoryStream, $oImgFormat)
    $pngMemoryStream.Dispose()
    $cImgBytes = [Byte[]]($oMemoryStream.ToArray())
    $sBase64 = [System.Convert]::ToBase64String($cImgBytes)
    Return $html.replace($srcAttribute, 'data:image/.jpg;base64,' + $sBase64)
  }
  Else {
    Return $html
  }
}

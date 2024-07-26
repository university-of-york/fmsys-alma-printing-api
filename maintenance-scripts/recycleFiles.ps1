param (
    [string]$FolderPath,
    [int]$DaysOlderThan = 30,
    [string]$FileExtension = "html",
    [switch]$Delete
)

if (-not $FolderPath) {
    Write-Host "-FolderPath is a required parameter"
    exit 1
}

if (-not (Test-Path -Path $FolderPath -PathType Container)) {
    Write-Host "The specified folder does not exist"
    exit 1
}

# Function to send files to recycle bin
function Move-ToRecycleBin {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.Namespace(0)
    $file = $folder.ParseName($Path)
    $file.InvokeVerb("delete")
}

# Get the appropriate delete command
$deleteCommand = if ($Delete) {
    { Remove-Item -Path $_.FullName -Force }
} else {
    { Move-ToRecycleBin -Path $_.FullName }
}

# Find and process files
Get-ChildItem -Path $FolderPath -File -Recurse | Where-Object {
    $_.CreationTime -lt (Get-Date).AddDays(-$DaysOlderThan) -and
    $_.Extension -eq ".$FileExtension"
} | ForEach-Object $deleteCommand

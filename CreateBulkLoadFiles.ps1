## CreateBulkLoadFiles.ps1
##


if ($Args.Length -ne 1) {
  Write-Host "Usage: $($PSCommandPath) filename.json"
  Exit 1
}

$JsonFile = Get-ChildItem $Args[0]
$JsonDir = $JsonFile.DirectoryName

# Read in the JSON file to a PSCustomObject
$Json = Get-Content -Path $JsonFile | ConvertFrom-Json


Write-Host "Creating files for AAD invite bulk loads"

foreach ($channel in $Json.Channels) {
  Write-Host $channel.PartnerName -ForegroundColor Cyan
  foreach ($guest in $channel.Guests) {
    Write-Host $guest.Email
  }
}

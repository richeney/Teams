## CreatePartnerHackTeam.ps1
##
## Creates a Microsoft Team suitable for a partner hack
## Note that a private Team must currently be pre-created.
## Reads from a custom JSON file. See accompanying JSON files
## and the GeneratePartnerHackJson.ps1 script.
##
## PowerShell v7
## Install-Module -Name MicrosoftTeams -Force -AllowClobber
##
## Example:
##
## 1. Create a valid JSON file
## 2. Connect-MicrosoftTeams
## 3. PartnerHackCreateTeam.ps1 AzureArcForServersJan2022.json

if ($Args.Length -ne 1) {
  Write-Host "Usage: $($PSCommandPath) filename.json"
  Exit 1
}

$JsonFile = Get-ChildItem $Args[0]
$JsonDir = $JsonFile.DirectoryName

# Read in the JSON file to a PSCustomObject
$Json = Get-Content -Path $JsonFile | ConvertFrom-Json


$Guests = $Json.Channels.Guests

$Guests | Join-String -Separator "," -Property Email
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


Write-Host "Proctors" -ForegroundColor Cyan
foreach ($owner in $Json.Owners) { Write-Host $($owner.Upn) }

foreach ($channel in $Json.Channels) {
  Write-Host $channel.PartnerName -ForegroundColor Cyan
  $UserInvite = "$JSonDir\$($channel.Code)-UserInvite.csv"
  $GroupImportMembers = "$JSonDir\$($channel.Code)-GroupImportMembers.csv"

  New-Item $UserInvite -ItemType File -Force
  Add-Content $UserInvite -Value "version:v1.0,,,"
  Add-Content $UserInvite -Value "Email address to invite [inviteeEmail] Required,Redirection url [inviteRedirectURL] Required,Send invitation message (true or false) [sendEmail],Customized invitation message [customizedMessageBody]"

  New-Item $GroupImportMembers -ItemType File -Force
  Add-Content $GroupImportMembers -Value "version:v1.0"
  Add-Content $GroupImportMembers -Value "Member object ID or user principal name [memberObjectIdOrUpn] Required"

  foreach ($owner in $Json.Owners) {
    $UserInviteLine = $owner.Upn + ",https://myapplications.microsoft.com" + ",true, Welcome to the Azure Arc Partner Hack!"
    Add-Content $UserInvite -Value $UserInviteLine

    $GroupImportMembersLine = $owner.Upn.Replace('@', '_') + "#EXT#@arc" + $channel.Code +  "outlook.onmicrosoft.com"
    Add-Content $GroupImportMembers -Value $GroupImportMembersLine
  }

  foreach ($guest in $channel.Guests) {
    $UserInviteLine = $guest.Email + ",https://myapplications.microsoft.com" + ",true, Welcome to the Azure Arc Partner Hack!"
    Add-Content $UserInvite -Value $UserInviteLine

    $GroupImportMembersLine = $guest.Email.Replace('@', '_') + "#EXT#@arc" + $channel.Code +  "outlook.onmicrosoft.com"
    Add-Content $GroupImportMembers -Value $GroupImportMembersLine
  }
}

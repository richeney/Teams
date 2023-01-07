## ConfigureTeam.ps1
##
## Creates a Microsoft Team suitable for a partner hack
## Note that a private Team must currently be pre-created.
## Reads from a custom JSON file. See accompanying JSON files
## and the GeneratePartnerHackJson.ps1 script.
##
## PowerShell v7
## Install-Module -Name MicrosoftTeams -Force -AllowClobber
## Import-Module MicrosoftTeams

## Example:
##
## 1. Create a valid JSON file
## 2. Connect-MicrosoftTeams -TenantId 72f988bf-86f1-41af-91ab-2d7cd011db47
## 3. .\ConfigureTeam.ps1 .\2023-01\AzureArcJan2023.json

if ($Args.Length -ne 1) {
  Write-Host "Usage: $($PSCommandPath) filename.json"
  Exit 1
}

$JsonFile = Get-ChildItem $Args[0]
$JsonDir = $JsonFile.DirectoryName

# Read in the JSON file to a PSCustomObject
$Json = Get-Content -Path $JsonFile | ConvertFrom-Json

# Get the GroupId
if ([bool]($Json.PSobject.Properties.name -match "GroupId")) {
  $GroupId = $Json.GroupId
  Write-Host "GroupId found in $($JsonFile.Name): $($GroupId)"
}
else {
  Write-Host -NoNewline "Getting the GroupId for", $Json.DisplayName, ": "
  $GroupId = (Get-Team -DisplayName $Json.DisplayName).GroupId
  $GroupId
}

# $Team = Get-Team -GroupId $GroupId

# Team level changes

Write-Host "Setting the picture:", $Json.Picture
$picture = $JsonDir + '\' + $Json.Picture
Set-TeamPicture -GroupId $GroupId -ImagePath $picture

Write-Host "Setting the description:", $Json.Description
Set-Team -GroupId $GroupId -Description $Json.Description

Write-Host "Setting Owners"

$ExistingUsers = Get-TeamUser -GroupId $GroupId
$ExistingOwners = $ExistingUsers | Where-Object { $_.Role -eq "owner" }
[bool]$OwnersAdded = $false

foreach ($owner in $Json.Owners) {
  if ($ExistingOwners.User.Contains($owner.Upn)) {
    Write-Host "  $($owner.Upn) ($($owner.Name)) ✔️"
  }
  else {
    Write-Host "  $($owner.Upn) ($($owner.Name)) - adding"
    Add-TeamUser -GroupId $GroupId -User $owner.Upn -Role Owner
    $OwnersAdded = $true
  }
}

# Refresh owners if we've added
if ($OwnersAdded) {
  Write-Host "Refreshing list of owners"
  $ExistingUsers = Get-TeamUser -GroupId $GroupId
  $ExistingOwners = $ExistingUsers | Where-Object { $_.Role -eq "owner" }
}

Write-Host "Adding Guests"

$ExistingGuests = $ExistingUsers | Where-Object { $_.Role -eq "guest" }

foreach ($guest in $Json.Channels.Guests) {
  $upn = "$($guest.Email.Replace('@', '_'))#EXT#@microsoft.onmicrosoft.com"
  if ($ExistingGuests.Count -gt 0 -And $ExistingGuests.User.ToLower() -contains $upn) {
    Write-Host "  $($guest.Email) ($($guest.Name)) ✔️"
  }
  else {
    try {
      Add-TeamUser -GroupId $GroupId -User $guest.Email -ErrorAction SilentlyContinue
     }
     catch {
      Write-Host "    $($guest.Email) ($($guest.Name)) ❌ (Manually add email, edit name)"
      Continue
    }
    Write-Host "    $($guest.Email) ($($guest.Name)) ✔️ (added)"
  }
}

Write-Host "Creating Channels"

$ExistingPrivateChannels = (Get-TeamChannel -GroupId $GroupId -MembershipType Private).DisplayName

foreach ($channel in $Json.Channels) {
  if ($ExistingPrivateChannels.Count -gt 0 -And $ExistingPrivateChannels -contains $channel.PartnerName) {
    Write-Host "  $($channel.PartnerName) ✔️"
  }
  else {
    Write-Host "  $($channel.PartnerName) - adding"
    New-TeamChannel -GroupId $GroupId -DisplayName $channel.PartnerName -MembershipType Private
  }

  $ExistingChannelUsers = Get-TeamChannelUser -GroupId $GroupId -DisplayName $channel.PartnerName
  $ExistingChannelOwners = $ExistingChannelUsers | Where-Object {$_.Role -eq "Owner"}
  $ExistingChannelGuests = $ExistingChannelUsers | Where-Object {$_.Role -eq "Guest"}

  foreach ($owner in $ExistingOwners) {
      # Use the objectId as the team uses the alias email whilst the channel uses the long form
      if (-Not ($ExistingChannelOwners.UserId.Contains($owner.UserId))) {
      Write-Host "    $($owner.Name) - adding"
      # Two step - add then promote to owner. Fails if you try to go staight to owner.
      Add-TeamChannelUser -GroupId $GroupId -DisplayName $channel.PartnerName -User $owner.UserId
      Add-TeamChannelUser -GroupId $GroupId -DisplayName $channel.PartnerName -User $owner.UserId -Role "Owner"
    }
  }

  foreach ($guest in $channel.Guests) {
    $ExtEmail = "$($guest.Email.Replace('@', '_'))#EXT#@microsoft.onmicrosoft.com"

    if ($ExistingGuests.User -NotContains $ExtEmail) {
      Write-Host "    $($guest.Email) ($($guest.Name)) ❌ (Needs adding manually to the Team.)"
    } else {
      $UserId = ($ExistingGuests | Where-Object {$_.User -eq $ExtEmail}).UserId

      if ($ExistingChannelGuests.Count -gt 0 -And $ExistingChannelGuests.UserId -Contains $UserId) {
        Write-Host "    $($guest.Email) ($($guest.Name)) ✔️"
      } else {
        try {
          Add-TeamChannelUser -GroupId $GroupId -DisplayName $channel.PartnerName -User $UserId -ErrorVariable err -ErrorAction SilentlyContinue
        }
        catch {
          # Write-Host $_.Exception.Message -ForegroundColor Red
          Write-Host "    $($guest.Email) ($($guest.Name)) ❌"
          Continue
        }
        Write-Host "    $($guest.Email) ($($guest.Name)) ✔️ (added)"
      }
    }
  }
  # exit
}

# Input bindings are passed in via param block.
param($Timer)
$currentUTCtime = (Get-Date).ToUniversalTime()
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"
##########################################
Import-Module Az.Accounts
Import-Module PnP.PowerShell
Write-Host "==========================================================================="

$tenantId = "db05faca-c82a-4b9d-b9c5-0f64b6755421"
$clientId = "30841bf3-964f-4b14-89ff-0f409119f784"
$VaultName = 'spo-spa-87653'
$secretName = 'secret01'
$clientSc = Get-AzKeyVaultSecret -VaultName $VaultName -Name $secretName -AsPlainText
if ($clientSc) {
  Write-host "Got creds from keyvault: " $clientSc.Substring(0,5) "..."
} else {
  write-error "Could not get creds from keyvault"
}
##########################################################################################
# $eap = $ErrorActionPreference; 
# $ErrorActionPreference = "SilentlyContinue"

# $baseUrl = "https://dsistg.sharepoint.com/sites/" # stage
# $intakeUrl = "https://uhgazure.sharepoint.com/sites/request_site/" # prod
$tempLogFile = New-TemporaryFile 
$baseUrl = "https://uhgazure.sharepoint.com/sites/" # stage
$intakeUrl = "https://uhgazure.sharepoint.com/teams/SPODevTools/" # prod
$listRelUrl = "Lists/SiteLookupPoC"

if ($clientId) {
  "Got Client Id: " + $clientId 
} else {
  "Failed to Get Client Id..." 
}
$adminUrl = "https://uhgazure-admin.sharepoint.com/"
$connectionAdmin = Connect-PnPOnline -Url $adminUrl  -ClientId $clientId -ClientSecret $clientSc -ReturnConnection
if ($connectionAdmin.Url -eq $adminUrl) {
  "Authenticated to: " + $connectionAdmin.Url 
  Write-Host "Authenticated to:" $adminUrl -fore Green
} else {
  "Failed to Authenticate to admin Url..." 
  Write-Host "Failed to auth to:" $adminUrl -fore Yellow
}

$connectionIntake = Connect-PnPOnline -Url $intakeUrl -ClientId $clientId -ClientSecret $clientSc -ReturnConnection
if ($connectionIntake.Url -eq $intakeUrl.Trim("/") ) {
  "Connected to: " + $connectionIntake.Url 
  Write-Host "Connected to:" $intakeUrl -fore Green
} else {
  "Failed to connect to intake Url..." 
  Write-Host "Failed to connect to:" $intakeUrl -fore Yellow
}

$list = Get-PnPList -Connection $connectionIntake -Identity $listRelUrl -Includes Fields
if ($?) {
  $message = " Got Intake list"; Write-Host $message; $message 
} else {
  $message = " Failed to get Intake list"; Write-Host $message; $message 
}
# $list | Format-Table -a

$siteRequests = Get-PnPListItem -List $list -Connection $connectionIntake 
"Number of items in the request list: " + $siteRequests.Count.ToString()  
Write-Host "Total number of requests in the list:" $siteRequests.Count
Write-Host "New requests in the list:" $($siteRequests | ?{$_.FieldValues["RequestStatus"] -eq "New"} | measure-object).Count
Write-Host "In Progress requests in the list:" $($siteRequests | ?{$_.FieldValues["RequestStatus"] -eq "In Progress"} | measure-object).Count
# $siteRequest = $siteRequests | select -Last 1
# $siteRequest = $siteRequests[-2]; $siteRequest
# $siteRequest = $siteRequests | ?{$_.FieldValues["RequestStatus"] -eq "Failed"} | ?{$_.FieldValues["Title"] -like "*Enterpr*"}; $siteRequest
foreach($siteRequest in $siteRequests) {
  Write-Host "site:" $siteRequest.FieldValues["Title"]  -NoNewline
  if ($siteRequest.FieldValues["RequestStatus"] -eq "Completed") {
    Write-Host " - This is completed - skipping - " -ForegroundColor Green
    continue
  } else {
    Write-Host " - Let's start working with this NEW request" -ForegroundColor Green
  }
  $siteExists = $null
  $siteExists = Get-PnPTenantSite -Url $siteRequest.FieldValues["Title"] -Connection $connectionAdmin
  if($siteExists) {
    Write-Host " - Site Found: " $siteExists.Title
    Set-PnPListItem -List $list -Identity $siteRequest -Values @{"SiteTitle"=$siteExists.Title} -Connection $connectionIntake
    Set-PnPListItem -List $list -Identity $siteRequest -Values @{"SiteOwners"=$siteExists.OwnerName} -Connection $connectionIntake
  } else {
    Write-Host " - Site not Found:" $siteRequest.FieldValues["Title"]
    Set-PnPListItem -List $list -Identity $siteRequest -Values @{"SiteTitle"="Error: Site Not Found"} -Connection $connectionIntake
  }    
  Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus"="Completed"} -Connection $connectionIntake
}

# $ErrorActionPreference = $eap

return

#       if ($?) {Write-host " = no errors = "} else {$lastErrorMessage = $Error[0].Exception.Message}
  #$siteRequest.FieldValues["Short_Name"]
  Get-PnPWebTemplates
  Get-PnPTimeZoneId -Match Central

  $list = Get-PnPList -Identity $listRelUrl -Connection $connectionIntake -Includes Fields
  $list | Format-Table -a
  
  $list.Fields | ft -a
  $list.Fields | ft -a | clip
  $siteRequest = $siteRequests | select -First 1
  $siteRequest = $siteRequests | select -Last 1
  $siteRequest = $siteRequests | ?{$_.Title -eq 'Broker Bonus Adminstration'}
  $siteRequest = $siteRequests | ?{$_.GUID -eq '7fb07d89-f098-4b50-8df4-4be595754067'}
  $siteRequest = $siteRequests[-4]
  $siteRequest 

  $oneItem = Get-PnPListItem -List $list -Id  $siteRequest.Id -fields Id, Title, RequestStatus -Connection $connectionIntake 
  $oneItem  
  Set-PnPListItem -List $list -Identity $oneItem.Id -Values @{"RequestStatus" = "In Progress"} -Connection $connectionIntake
  Set-PnPListItem -List $list -Identity $oneItem.Id -Values @{"SiteURL" = "https://uhgazure.sharepoint.com/sites/brokerbonusrequest"} -Connection $connectionIntake

#  Administratorâ€™s
#  Administrator's

return
Connect-AzAccount

$siteRequest = $siteRequests | ?{$_.FieldValues["RequestStatus"] -eq "Failed"} | select -Last 1
$siteRequest 
Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "In Progress"} -Connection $connectionIntake
Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "New"} -Connection $connectionIntake
Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Primary_x0020_Admin" = "allison_f_copeland@uhc.com"} -Connection $connectionIntake

$list.Fields





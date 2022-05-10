# Input bindings are passed in via param block.
param($Timer)
$currentUTCtime = (Get-Date).ToUniversalTime()
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

Import-Module Az.Accounts
Import-Module PnP.PowerShell
Get-Module | ft -AutoSize
Write-Host "==========================================================================="

$tenantId = "db05faca-c82a-4b9d-b9c5-0f64b6755421"
$clientId = "30841bf3-964f-4b14-89ff-0f409119f784"
$VaultName = 'spo-spa-87653'
$secretName = 'secret01'
$clientSc = Get-AzKeyVaultSecret -VaultName $VaultName -Name $secretName -AsPlainText

$intakeUrl = "https://uhgazure.sharepoint.com/sites/request_site/" # prod
$listRelUrl = "/Lists/SCA_Requests"

$siteUrlFieldName = "Site_Url"
$requestStatusFieldName = "Request_Status"
$tempAdminFieldName = "Temp_Admin"
$automationCommentFieldName = "Automation_Comment"
$adminUrl = "https://uhgazure-admin.sharepoint.com/"
Connect-PnPOnline -Url $adminUrl  -ClientId $clientId -ClientSecret $clientSc 
Get-PnPSite | ft -a
$connectionAdmin = Get-PnPConnection
Connect-PnPOnline -Url $intakeUrl -ClientId $clientId -ClientSecret $clientSc
Get-PnPSite | ft -a
$connectionIntake = Get-PnPConnection

$list = Get-PnPList -Identity $listRelUrl -Connection $connectionIntake
# $list | Format-Table -a

$scaRequests = @()
$scaRequests = Get-PnPListItem -List $list -Connection $connectionIntake
$scaRequests.Count

$scaRequest = $scaRequests | select -last 1; 
$scaRequest.FieldValues[$siteUrlFieldName]
$scaRequest.FieldValues[$requestStatusFieldName]
foreach($scaRequest in $scaRequests) {
  Write-Host "Site:" $scaRequest.FieldValues[$siteUrlFieldName]  -NoNewline
  switch($scaRequest.FieldValues[$requestStatusFieldName]){
    "Closed" { Write-Host " - request closed " -ForegroundColor Green } # site was provisioned 
    "New" { 
      Write-Host " - Working with the NEW request" -ForegroundColor Green
      $tempAdmin = $null
      $tempAdmin = Get-PnPUser -Identity $scaRequest.FieldValues[$tempAdminFieldName].LookupId
      if ($tempAdmin) {
      } else {
        $message = "Temp Admin must be specified."
        Write-Host $message + "Continue." -ForegroundColor Yellow
        $requestUpdateValues = @{
          $automationCommentFieldName = $message;
          $requestStatusFieldName     = "Failed"
        }
        Set-PnPListItem -List $list -Identity $scaRequest -Values $requestUpdateValues -Connection $connectionIntake
        continue
      }

      $siteExists = $null
      $siteExists = Get-PnPTenantSite -Url $scaRequest.FieldValues[$siteUrlFieldName] -Connection $connectionAdmin 
      if($siteExists) {
      } else {
        $message = "Site could not be Found."
        Write-Host " - $message . Continue.  " -ForegroundColor Yellow
        $requestUpdateValues = @{
          $automationCommentFieldName = $message;
          $requestStatusFieldName     = "Failed"
        }
        Set-PnPListItem -List $list -Identity $scaRequest -Values $requestUpdateValues -Connection $connectionIntake
        continue
      }    

      # check Temp_Admin for eligability

      Set-PnPTenantSite -Url $scaRequest.FieldValues[$siteUrlFieldName] -Owner $tempAdmin.LoginName -Connection $connectionAdmin
      if ($?) {
        $message = "Access has been provided."
        Write-Host " - $message "
        $requestUpdateValues = @{
          $automationCommentFieldName = $message;
          $requestStatusFieldName     = "Completed"
        }
        Set-PnPListItem -List $list -Identity $scaRequest -Values $requestUpdateValues -Connection $connectionIntake
      } else {
        $message = $Error[0].ErrorDetails.Message
        Write-Host $message -ForegroundColor Yellow
        $requestUpdateValues = @{
          $automationCommentFieldName = $message;
          $requestStatusFieldName     = "Failed"
        }
        Set-PnPListItem -List $list -Identity $scaRequest -Values $requestUpdateValues -Connection $connectionIntake
      }


    }
    "Completed" {
      # remove access after some time
    }
    Default {
      Write-Host " - " $scaRequest.FieldValues[$requestStatusFieldName] 
    }
  }
}

#$ErrorActionPreference = $eap

return

#       if ($?) {Write-host " = no errors = "} else {$lastErrorMessage = $Error[0].Exception.Message}
  #$scaRequest.FieldValues["Short_Name"]
  Get-PnPWebTemplates
  Get-PnPTimeZoneId -Match Central


  $list = Get-PnPList -Identity $listRelUrl -Connection $connectionIntake -Includes Fields
  $list | Format-Table -a
  
  $list.Fields | ft -a
  $scaRequest = $scaRequests | select -First 1

$requestUpdateValues = @{
  $automationCommentFieldName = $message;
  $requestStatusFieldName     = "New"
}
Set-PnPListItem -List $list -Identity $scaRequest -Values $requestUpdateValues -Connection $connectionIntake

$scaRequest.FieldValues[$tempAdminFieldName] | fl
$scaRequest.FieldValues[$tempAdminFieldName]

Get-PnPUser -Identity 38




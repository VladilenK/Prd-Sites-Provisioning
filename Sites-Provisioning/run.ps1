# Input bindings are passed in via param block.
param($Timer)
$currentUTCtime = (Get-Date).ToUniversalTime()
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"

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
$outputFileName = $tempLogFile.DirectoryName + "\Sites_Creaton_Log.txt"
"======================  Date: " + $(Get-Date).ToString() | Out-File $outputFileName -Append
$baseUrl = "https://uhgazure.sharepoint.com/sites/" # stage
$intakeUrl = "https://uhgazure.sharepoint.com/sites/CenterofExcellence/" # prod
$listRelUrl = "Lists/NewSiteRequest"

# Configure safe domains list:
$safeDomains = @()
$safeDomains += "cdnapisec.kaltura.com"
$safeDomains += "uhg.video.uhc.com"

if ($clientId) {
  "Got Client Id: " + $clientId | Out-File $outputFileName -Append
} else {
  "Failed to Get Client Id..." | Out-File $outputFileName -Append
}
$adminUrl = "https://uhgazure-admin.sharepoint.com/"
$connectionAdmin = Connect-PnPOnline -Url $adminUrl  -ClientId $clientId -ClientSecret $clientSc -ReturnConnection
if ($connectionAdmin.Url -eq $adminUrl) {
  "Authenticated to: " + $connectionAdmin.Url | Out-File $outputFileName -Append
  Write-Host "Authenticated to:" $adminUrl -fore Green
} else {
  "Failed to Authenticate to admin Url..." | Out-File $outputFileName -Append
  Write-Host "Failed to auth to:" $adminUrl -fore Yellow
}

$connectionIntake = Connect-PnPOnline -Url $intakeUrl -ClientId $clientId -ClientSecret $clientSc -ReturnConnection
if ($connectionIntake.Url -eq $intakeUrl.Trim("/") ) {
  "Connected to: " + $connectionIntake.Url | Out-File $outputFileName -Append
  Write-Host "Connected to:" $intakeUrl -fore Green
} else {
  "Failed to connect to intake Url..." | Out-File $outputFileName -Append
  Write-Host "Failed to connect to:" $intakeUrl -fore Yellow
}

$list = Get-PnPList -Identity $listRelUrl -Connection $connectionIntake
if ($?) {
  $message = " Got Intake list"; Write-Host $message; $message | Out-File $outputFileName -Append
} else {
  $message = " Failed to get Intake list"; Write-Host $message; $message | Out-File $outputFileName -Append
}
# $list | Format-Table -a

$siteRequests = Get-PnPListItem -List $list -Connection $connectionIntake 
"Number of items in the request list: " + $siteRequests.Count.ToString() | Out-File $outputFileName -Append 
Write-Host "Total number of requests in the list:" $siteRequests.Count
Write-Host "New requests in the list:" $($siteRequests | ?{$_.FieldValues["RequestStatus"] -eq "New"} | measure-object).Count
Write-Host "In Progress requests in the list:" $($siteRequests | ?{$_.FieldValues["RequestStatus"] -eq "In Progress"} | measure-object).Count
$siteRequest = $siteRequests | select -Last 1
foreach($siteRequest in $siteRequests) {
  if     ($siteRequest.FieldValues["RequestStatus"] -eq "New") {} 
  elseif ($siteRequest.FieldValues["RequestStatus"] -eq "In Progress" ) {}
  elseif ($siteRequest.FieldValues["RequestStatus"] -eq "Completed" ) {}
  elseif ($siteRequest.FieldValues["RequestStatus"] -eq "Failed" ) {}
  elseif ($siteRequest.FieldValues["RequestStatus"] -eq "Rejected" ) {}
  elseif ($siteRequest.FieldValues["RequestStatus"] -eq "WaitingForManagerApproval" ) {}
  else {
    $message = "Undefined request status:" + $siteRequest.FieldValues["RequestStatus"] + " for site request:" + $siteRequest.FieldValues["ShortName"] 
    $message | Out-File $outputFileName -Append
    Write-Host $message -ForegroundColor Yellow
  }
  if ( $siteRequest.FieldValues["Modified"].AddMinutes(3).ToLocalTime() -gt $(Get-Date) ) {
    # modified less than 3 minutes ago, skip as some workflows migh be running
    continue
  }
  #Write-Host "site:" $siteRequest.FieldValues["Title"]  -NoNewline
  switch($siteRequest.FieldValues["RequestStatus"]){
    "New" { 
      "New:" + $siteRequest.FieldValues["ShortName"]  | Out-File $outputFileName -Append 
      Write-Host "site:" $siteRequest.FieldValues["Title"]  -NoNewline
      Write-Host " - Let's start working with this NEW request" -ForegroundColor Green
      #Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Assigned"} -Connection $connectionIntake
      $template = "STS#3"
      if ($siteRequest.FieldValues["Template"] -eq "Communication") {
        $template = "SITEPAGEPUBLISHING#0"
      }
      $consent1 = $siteRequest.FieldValues["SharePointIsSolution"]
      $consent2 = $siteRequest.FieldValues["I_x0020_attest"]
      if ($consent1 -eq "Yes" -and $consent2 -eq "Yes") {
        # proceed
      } else {
        Write-Host " - User have not attested..." -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"="You have not attested to SharePoint being the right tool and that the admins understand their roles & responsibilities"} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
        continue
      }
      $siteAdmin1 = $siteRequest.FieldValues["Primary_x0020_Admin"]
      $siteAdmin2 = $siteRequest.FieldValues["Secondary_x0020_Admin"]
      if ($siteAdmin1.LookupValue -ne $siteAdmin2.LookupValue) {
        # proceed
      } else {
        Write-Host " - SC Primary and Secondary Administrator s cannot be the same person" -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"="Primary and Secondary Administrator's cannot be the same person"} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
        continue
      }
      $siteAdmin1email = [mailaddress]$siteAdmin1.Email
      $siteAdmin2email = [mailaddress]$siteAdmin2.Email
      if ($siteAdmin1email) {
        # proceed
        # $upProps = $null
        # $upProps = Get-PnPUserProfileProperty -Account $siteAdmin1email -Connection $connectionAdmin
        # if ([mailaddress]$upProps.UserProfileProperties.'SPS-UserPrincipalName') {
        #   $siteAdmin1email = [mailaddress]$upProps.UserProfileProperties.'SPS-UserPrincipalName'
        # }
        Write-host "email 1 is OK"
      } else {
        $message = "Primary Administrator account email propery is empty or incorrect."
        Write-Host $message -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"=$message} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
        continue
      }
      if ($siteAdmin2email) {
        # proceed
        # $upProps = $null
        # Get-PnPAADUser -Identity "vijay.vellabati@optum.com" -Connection $connectionAdmin
        # $upProps = Get-PnPUserProfileProperty -Account $siteAdmin2email -Connection $connectionIntake
        # $upProps
        # if ([mailaddress]$upProps.UserProfileProperties.'SPS-UserPrincipalName') {
        #   $siteAdmin2email = [mailaddress]$upProps.UserProfileProperties.'SPS-UserPrincipalName'
        # }
        Write-host "email 2 is OK"
      } else {
        $message = "Secondary Administrator account email propery is empty or incorrect."
        Write-Host $message -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"=$message} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
        continue
      }

      $shortName = $siteRequest.FieldValues["ShortName"]
      $regex = '[^a-zA-Z0-9-_]'
      $cleanShortName = $shortName -replace $regex, ""
      if ($cleanShortName.Length -gt 0) {
        $siteUrlProposed = $baseUrl + $cleanShortName
        $siteExists = $null
        $siteExists = Get-PnPTenantSite -Url $siteUrlProposed -Connection $connectionAdmin
        $deletedSiteExists = $null
        $deletedSiteExists = Get-PnPTenantRecycleBinItem -Connection $connectionAdmin | Where-Object{$_.Url -eq $siteUrlProposed }
      } else {
        Write-Host " - Shortname is too short " -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"="ShortName contains special characters or too short"} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
        continue
      }
      if($siteExists) {
        Write-Host " - Site Already Exists: " $siteUrlProposed -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"="There is an existing site with the Url:" + $siteExists.Url} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
      } elseif ($deletedSiteExists)  {
        Write-Host " - Deleted Site Exists: " $siteUrlProposed -ForegroundColor Yellow
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"="There is a deleted site with the Url:" + $deletedSiteExists.Url} -Connection $connectionIntake
        Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
      } else {
        Write-Host " - Trying to create a site:" $siteUrlProposed
        New-PnPTenantSite -Url $siteUrlProposed -Owner $siteAdmin1.Email -StorageQuota $(200*1024) -Template $template -Title $siteRequest.FieldValues["Title"] -Connection $connectionAdmin -TimeZone 11 
        if ($?) {
          $message = " New-PnPTenantSite command completed successfully."; Write-Host $message; $message | Out-File $outputFileName -Append
          $message = " Site:" + $siteUrlProposed; Write-Host $message; $message | Out-File $outputFileName -Append
          Set-PnPListItem -List $list -Identity $siteRequest -Values @{"SiteURL" = $siteUrlProposed; "SiteCreated"= Get-Date} -Connection $connectionIntake
          Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"="Site is being created"} -Connection $connectionIntake
          Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "In Progress"} -Connection $connectionIntake
        } else {
          start-sleep -Seconds 3
          $lastError = $Error[0]
          $message = $lastError.Exception.Message    ; Write-Host $message -ForegroundColor Yellow; $message | Out-File $outputFileName -Append
          Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"= $message} -Connection $connectionIntake
          Set-PnPListItem -List $list -Identity $siteRequest -Values @{"RequestStatus" = "Failed"} -Connection $connectionIntake
          Remove-PnPTenantSite -Connection  $connectionAdmin -Url $siteUrlProposed -Force -SkipRecycleBin
        }
      }    
    }
    "In Progress" {
      Write-Host "site:" $siteRequest.FieldValues["Title"]  -NoNewline
      Write-Host " - Request is in progress - let us check if it's done" -ForegroundColor Green
      #Connect-PnPOnline -Url $adminUrl  -ClientId $clientId -ClientSecret $clientSc 
      #$connectionAdmin
      #$context = Get-PnPContext 
      $newPnpTenantSite = $null
      "InProgress:" + $siteRequest.FieldValues["SiteURL"]| Out-File $outputFileName -Append
      $newPnpTenantSite = Get-PnPTenantSite -Url $siteRequest.FieldValues["SiteURL"] -Connection $connectionAdmin -Detailed 
      if ($newPnpTenantSite) {
        "New Site Created:" + $newPnpTenantSite.Url | Out-File $outputFileName -Append  
        $siteAdmin2 = $siteRequest.FieldValues["Secondary_x0020_Admin"]
        Set-PnPTenantSite -Url $newPnpTenantSite.Url -Owners $siteAdmin2.Email  -Connection $connectionAdmin
        $groupLoginName = "c:0t.c|tenant|99d76526-258f-4841-a830-c0a00c2ee945" # "SPO_Support" azure ad security group
        Set-PnPTenantSite -Url $newPnpTenantSite.Url -Owners $groupLoginName -Connection $connectionAdmin

        # disable custom scripts
        Set-PnPTenantSite -Url $newPnpTenantSite.Url -DenyAddAndCustomizePages  -Connection $connectionAdmin
        # $DenyAddAndCustomizePagesStatusEnum = [Microsoft.Online.SharePoint.TenantAdministration.DenyAddAndCustomizePagesStatus]
        # $newPnpTenantSite.DenyAddAndCustomizePages = $DenyAddAndCustomizePagesStatusEnum::Enabled
        # $newPnpTenantSite.Update()
        # Invoke-PnPQuery -Connection $connectionAdmin
        $DenyAddAndCustomizePages = $null
        $DenyAddAndCustomizePages = Get-PnPTenantSite -Url $siteRequest.FieldValues["SiteURL"] -Connection $connectionAdmin -Detailed | Select-Object DenyAddAndCustomizePages -ExpandProperty DenyAddAndCustomizePages
        if ($DenyAddAndCustomizePages -eq "Enabled") {
          $message = " Custom Scripts has been disabled. DenyAddAndCustomizePages = " + $DenyAddAndCustomizePages; Write-Host $message; $message | Out-File $outputFileName -Append
        } else {
          $message = " Failed to disable Custom Scripts. DenyAddAndCustomizePages = " + $DenyAddAndCustomizePages; Write-Host $message; $message | Out-File $outputFileName -Append
        }

        # add sites to HTML Safe domains
        $connectionNewSite = Connect-PnPOnline $newPnpTenantSite.Url  -ClientId $clientId -ClientSecret $clientSc -ReturnConnection
        $pnpSite = Get-PnPSite -Includes CustomScriptSafeDomains, Owner -Connection $connectionNewSite
        $pnpsite | fl Url, Owner, CustomScriptSafeDomains 
        $pnpsite | Select-Object CustomScriptSafeDomains -expandproperty CustomScriptSafeDomains
        foreach ($safeDomain in $safeDomains) 
        {
          #Write-Host " - " $safeDomain "..." -noNewline
          $ssDomain = [Microsoft.SharePoint.Client.ScriptSafeDomainEntityData]::new()
          $ssDomain.DomainName = $safeDomain
          $pnpSite.CustomScriptSafeDomains.Create($ssDomain)
          #$pnpSite
          Invoke-PnPQuery -Connection $connectionNewSite
          if ($?) {
          } else {
            Write-Host "Could not invoke pnp query" -ForegroundColor Yellow
          }
        }
        $pnpsite | Select-Object CustomScriptSafeDomains -expandproperty CustomScriptSafeDomains

        Write-Host " - Success: " $newPnpTenantSite.Url
        $listItemValuesUpdate = @{"RequestStatus" = "Completed"; "Automation_Comment"="Success" }
        $updatedListItem = Set-PnPListItem -List $list -Identity $siteRequest -Values $listItemValuesUpdate -Connection $connectionIntake 
        "Completed:" + $newPnpTenantSite.Url | Out-File $outputFileName -Append
      } else {
        Write-Host " - Still in progress (Something went wrong?)" 
        if ($siteRequest.FieldValues["Automation_Comment"] -match "Still in progress") {
        } else {
          Set-PnPListItem -List $list -Identity $siteRequest -Values @{"Automation_Comment"= $("Still in progress..."   ) } -Connection $connectionIntake
        }
      }
    }
    Default {
      #Write-Host " - " $siteRequest.FieldValues["RequestStatus"]  -ForegroundColor Yellow
    }
  }
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


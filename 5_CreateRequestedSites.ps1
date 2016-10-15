param(
  [string]$staticUrlWithCustomizations = "/"
  ,[bool]$recreateSiteIfExists = $false
  ,[bool]$removeSiteOnly = $false # Overwrites the value of $recreateSiteIfExists
  ,[string]$listUrlDeploymentRequests = "Lists/DeploymentRequests"
  ,[int]$compatibilityLevel = 15
  ,[int]$storageMaximumLevel = 100
  ,[int]$userCodeMaximumLevel = 100 # Server Resource Quota in Admin Center
  ,[int]$timeoutToClearHangingRequestsMinutes = 20
  ,[string]$dateTimeStampFormatLong = "yyyy-MM-dd HH:mm"
  ,[string]$dateTimeStampFormatShort = "HH:mm:ss"
  ,[bool]$deployLegacySolutionsForCustomWebTemplates = $true
  ,[int]$maxSitesToProcessInSingleRun = 10
  ,[string[]]$defaultSiteAdministrators = @("paul", "spa")
  ,[string[]]$defaultVisitors = @("paul", "spa")
)

######################################################## FUNCTIONS #########################################################
function AddSiteAdmin(){
param (
  [Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[string]$siteUrl
  ,[string]$userName
)
  if( $tenant -eq $null -or [string]::IsNullOrEmpty($siteUrl) -or [string]::IsNullOrEmpty($userName) ) {return}
  
  $admCtx = $tenant.Context
  try {
    write-host
    write-host "Adding '$userName' to site admins..." -NoNewLine
    $user = $tenant.SetSiteAdmin($siteUrl, $userName, $true)
    $admCtx.Load($user)
    $admCtx.ExecuteQuery()
    write-host "done"
  } catch {
    write-host
    write-host -ForegroundColor Red "Error adding $userName to site admins: $($_.Exception.Message)"
  }
}

function AddSiteVisitor(){
param (
  [Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[string]$siteUrl
  ,[string]$userName
)
  if( $tenant -eq $null -or [string]::IsNullOrEmpty($siteUrl) -or [string]::IsNullOrEmpty($userName) ) {return}
  
  try {
    write-host
    write-host "Adding '$userName' to the group of visitors..." -NoNewLine
    $admCtx = $tenant.Context
    $site = $tenant.GetSiteByUrl($siteUrl)
    $user = $site.RootWeb.EnsureUser($userName)
    $addedUser = $site.RootWeb.AssociatedVisitorGroup.Users.AddUser($user)
    $admCtx.Load($addedUser)
    $admCtx.ExecuteQuery()
    write-host "done"
  } catch {
    write-host
    write-host -ForegroundColor Red "Error adding $userName to the group of visitors: $($_.Exception.Message)"
  }
}

function ApplyCustomWebTemplate(){
param(
  [Parameter(Mandatory=$true)][string]$siteUrl
  ,[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$template
  ,[Parameter(Mandatory=$true)][ref]$isApplied
  ,[Parameter(Mandatory=$true)][ref]$appliedErrorMessage
  ,[bool]$deployLegacySolutions = $false
  ,[int]$language = 1033
  ,[string]$dateTimeStampFormatShort = "HH:mm:ss"
)
  
  if( $deployLegacySolutions ) {
    & $PSScriptRoot\4_DeployLegacySolutions.ps1 `
      -siteCollectionUrl $siteUrl -disableCustomizations $false -force $false -regexDeploymentUrls "."
  }
  
  $saved = $siteCollectionUrl
  $ctx = $null
  try {
    $siteCollectionUrl = $siteUrl
    $ctx = GetClientContext
    if( $ctx -eq $null ) {return}  # Some error; already displayed.

    write-host "Looking for the custom web template $template..." -NoNewLine

    $availableWebTemplates = $ctx.Site.RootWeb.GetAvailableWebTemplates($language, $false)
    $ctx.Load($availableWebTemplates)
    $ctx.ExecuteQuery()
    
    $template = $template.TrimStart("#")
    $customWebTemplates = ($availableWebTemplates | ? {$_.Name -ieq $template -or $_.Name -imatch ("#" + $template)})
    if( $customWebTemplates.Length -eq 0 ) {
      # Try to get cross language ones.
      $availableWebTemplates = $ctx.Site.RootWeb.GetAvailableWebTemplates($language, $true)
      $ctx.Load($availableWebTemplates)
      $ctx.ExecuteQuery()
      $customWebTemplates = ($availableWebTemplates | ? {$_.Name -ieq $template -or $_.Name -imatch ("#" + $template)})
    }
    if( $customWebTemplates.Length -gt 0 ) {
      $wt = $customWebTemplates[0].Name
      write-host "found $wt"
      write-host "Applying the web template $wt to $siteUrl started at $((get-date).ToString($dateTimeStampFormatShort))"
      write-host "Note if the operation times out the template will still be applied after some time." `
        -ForegroundColor Yellow
      $ctx.Site.RootWeb.ApplyWebTemplate($wt)
      try {
        $ctx.ExecuteQuery()
        $isApplied.Value = $true
      } catch {
        $message = $_.Exception.Message
        if( $message ) {
          $message = $message.ToString().Substring(
            $message.IndexOf(":") + 1).Trim().TrimStart('"').TrimEnd('"').TrimEnd('.')
        }
        if( $message -imatch "timed out"  ) {
          $isApplied.Value = $true
          write-host "$((get-date).ToString($dateTimeStampFormatShort)): $($message). Waiting for completion..." -NoNewLine
          try {
            $scriptBlock = {param($ctx)
              $ctx.Load($ctx.Site.RootWeb.RootFolder)
              $ctx.ExecuteQuery()
              # An assumption: a welcome page is not yet set while this custom template is being applied.
              if( [string]::IsNullOrEmpty($ctx.Site.RootWeb.RootFolder.WelcomePage) ) {
                throw "The opeartion is still in progress."
              }
            }
            # Try 20 times with interval 15 seconds (i.e. wait for completion max. for 5 min.)
            ExecuteAndRetryIfFailed -scriptBlock $scriptBlock -scriptBlockParams @($ctx) `
              -amountOfRetryAttempts 20 -waitIntervalSeconds 15
            write-host "done at $((get-date).ToString($dateTimeStampFormatShort))"
          } catch {
            $appliedErrorMessage.Value = $message
            write-host $message -ForegroundColor Red
          }
        } elseif( $message ) {
          $isApplied.Value = $false
          write-host $message -ForegroundColor Red
          $appliedErrorMessage.Value = $message
        }
      }
    } else {
      $isApplied.Value = $false
      write-host "not found"
      write-host "Custom web template $template not found on $siteUrl." -ForegroundColor Red
    }    
  } finally {
    $siteCollectionUrl = $saved
  }
}

function CreateSite(){
param(
  [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][string]$targetUrl
  ,[Parameter(Mandatory=$true)][string]$owner
  ,[Parameter(Mandatory=$true)][string]$title
  ,[string]$template = [string]::Empty
  ,[ref]$needsCustomTemplate
  ,[int]$language = 1033
  ,[int]$compatibilityLevel = 15
  ,[int]$storageMaximumLevel = 100
  ,[int]$userCodeMaximumLevel = 100
  ,[int]$maxRetryAttemptsToRefreshStatus = 5
  ,[string]$dateTimeStampFormatShort = "HH:mm:ss"
)
  try {
    if( [string]::IsNullOrEmpty($template) -or $template -ieq 	"__SELECTLATER" ) {
      $template = [string]::Empty
    } else {
      write-host
      write-host "Validating web template..." -NoNewLine
      $availableWebTemplates = $tenant.GetSPOTenantWebTemplates($language, $compatibilityLevel)
      $tenant.Context.Load($availableWebTemplates)
      $tenant.Context.ExecuteQuery()
      if( ($availableWebTemplates | ? {$_.Name -ieq $template}).Length -eq 0 ) {
        $message = "Web template '" + $template + "' not found. The site is being created with an empty template."
        write-host
        write-host $message -ForegroundColor Yellow
        $template = [string]::Empty
        $needsCustomTemplate.Value = $true
      } else {
        write-host "OK"
      }
    }
    
    $owner = $currentUser.LoginName
    $pipeIndex = $owner.LastIndexOf('|')
    if( $pipeIndex -gt -1 ) {$owner = $owner.Substring($pipeIndex + 1)}
    
    $siteCreationProperties = new-object "Microsoft.Online.SharePoint.TenantAdministration.SiteCreationProperties"
    $siteCreationProperties.Url = $targetUrl
    $siteCreationProperties.Owner = $owner
    $siteCreationProperties.Title = $title
    $siteCreationProperties.Template = $template
    $siteCreationProperties.Lcid = $language
    $siteCreationProperties.CompatibilityLevel = $compatibilityLevel
    $siteCreationProperties.StorageMaximumLevel = $storageMaximumLevel
    $siteCreationProperties.UserCodeMaximumLevel = $userCodeMaximumLevel

    write-host
    write-host "Creating the site $targetUrl started at $((get-date).ToString($dateTimeStampFormatShort))" -NoNewLine

    $op = $tenant.CreateSite($siteCreationProperties)
    $tenant.Context.Load($tenant)
    $tenant.Context.Load($op)
    $tenant.Context.ExecuteQuery()

    $attempts = 0
    while( !$op.IsComplete ) {
      write-host '.' -NoNewLine
      # Wait for 15 seconds and try again.
      Start-Sleep -s 15
      $op.RefreshLoad()
      try {
        $tenant.Context.ExecuteQuery()
      } catch {
        $message = $_.Exception.Message + $_.Exception.StackTrace
        $attempts++
        if( $attempts -ge $maxRetryAttemptsToRefreshStatus ) {
          throw $message
        } else {
          write-host " Status failure $($attempts) of $($maxRetryAttemptsToRefreshStatus) " -ForegroundColor Yellow -NoNewLine
        }
      }
    }
    write-host "Created at $((get-date).ToString($dateTimeStampFormatShort))"
  } catch {
    throw $_.Exception
  }
}

function DeleteSite() {
param(
  [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][string]$targetUrl
  ,[string]$dateTimeStampFormatShort = "HH:mm:ss"
)

  $isFound = DoesSiteExist -tenant $tenant -targetUrl $targetUrl
  if( !$isFound ) {return}
  
  try {
    write-host
    write-host "Deleting the site $targetUrl started at $((get-date).ToString($dateTimeStampFormatShort))" -NoNewLine
    $op = $tenant.RemoveSite($targetUrl)
    $tenant.Context.Load($tenant)
    $tenant.Context.Load($op)
    $tenant.Context.ExecuteQuery()

    while( !$op.IsComplete ) {
      write-host '.' -NoNewLine
      # Wait for 5 seconds and try again.
      Start-Sleep -s 5
      $op.RefreshLoad()
      $tenant.Context.ExecuteQuery()
    }
    write-host "deleted at $((get-date).ToString($dateTimeStampFormatShort))"
  } catch {
    throw $_.Exception
  }
}

function DoesSiteExist(){
param(
  [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][string]$targetUrl
  ,[bool]$suppressMessage = $false
) 
  $isFound = $false
  if( !$suppressMessage ) {
    write-host
    write-host "Checking for an existing site $targetUrl..." -NoNewLine
  }
  $site = $tenant.GetSiteByUrl($targetUrl)
  $tenant.Context.Load($site)
  try {
    $tenant.Context.ExecuteQuery()
    $isFound = ($site -ne $null)
    if( !$suppressMessage ) {
      if( $isFound ) {
        write-host "found"
      } else {
        write-host "Not found" # Or access denied
      }
    }
  } catch {
    $isFound = $false
    if( !$suppressMessage ) {
      write-host "Not found" # Or access denied
    }
  }
  
  return $isFound
}

function RemoveDeletedSiteFromRecycleBin() {
param(
  [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][string]$targetUrl
  ,[string]$dateTimeStampFormatShort = "HH:mm:ss"
)
  write-host
  write-host "Checking the recycle bin for a deleted site $targetUrl..." -NoNewLine
  $deletedSite = $tenant.GetDeletedSitePropertiesByUrl($targetUrl)
  $tenant.Context.Load($deletedSite)
  try {
    $tenant.Context.ExecuteQuery()
    try {
      write-host "found"
      write-host
      write-host `
        "Removing the deleted site $targetUrl started at $((get-date).ToString($dateTimeStampFormatShort))" -NoNewLine
      $op = $tenant.RemoveDeletedSite($targetUrl)
      $tenant.Context.Load($tenant)
      $tenant.Context.Load($op)
      $tenant.Context.ExecuteQuery()
      while( !$op.IsComplete ) {
        write-host '.' -NoNewLine
        # Wait for 5 seconds and try again.
        Start-Sleep -s 5
        $op.RefreshLoad()
        $tenant.Context.ExecuteQuery()
      }
      write-host "removed at $((get-date).ToString($dateTimeStampFormatShort))"
    } catch {
      throw $_.Exception
    }
  } catch {
    write-host "Not found" # Or access denied
  }
}
####################################################### //FUNCTIONS ########################################################

####################################################### EXECUTION ##########################################################
# Step 1. Connect to SharePoint environment amd get a client context.
write-host
write-host "Connecting to SharePoint environment..."
. $PSScriptRoot\__LoadContext.ps1 -initContextOnLoad:$false
$contextUri = new-object System.Uri($siteCollectionUrl) # $siteCollectionUrl is loaded by __LoadContext.ps1
$context = $null
if( $staticUrlWithCustomizations ) {
  $siteUrl = $contextUri.Scheme + "://" + $contextUri.Authority + "/" + $staticUrlWithCustomizations.TrimStart('/')
  $siteCollectionUrl = $siteUrl
  $context = GetClientContext
} else {
  $context = GetClientContext
}
if( $context -eq $null ) {
  write-host "Connection failed. Further execution is impossible."
  return
}
write-host
write-host ($context.Url)

try {
  # Connect to tenant's admin center.
  write-host
  write-host "Connecting to Tenant admin..."
  $adminContext = GetAdminClientContext
  if( $adminContext -ne $null ) {
    $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($adminContext)
    $adminContext.Load($tenant)
    $adminContext.ExecuteQuery()
  }
} catch {
  $message = $_.Exception.Message
  $tenant = $null
  $adminContext = $null
  write-host $message -ForegroundColor Red
}

if( $tenant -eq $null ) {
  write-host "Connection failed. Further execution is impossible."
  return
}

# Step 2. Get deployment requests from the correspondent list.
$listDeploymentRequests = GetListByFolderUrl -web $context.Web -webRelativeFolderUrl $listUrlDeploymentRequests
if( $listDeploymentRequests -eq $null ) {
  $message = "Required list of deployment requests not found on this web: " + `
    $context.Web.Url + "/" + $listUrlDeploymentRequests.TrimStart('/')
  throw $message
}

# Step 3. Check for and clear any hanging requests. Hanging requests are the ones that have 
# endings "...ing" in the "Status" field for continuous time, which exceeded the specified timeout.
write-host
write-host "Checks for hanging requests..." -NoNewLine
try {
  $camlQuery = new-object "Microsoft.SharePoint.Client.CamlQuery"
  $camlQuery.ViewXml = "<View><Query><Where><Contains><FieldRef Name='Status'/>" + `
    "<Value Type='Text'>ing</Value></Contains></Where></Query><RowLimit>100000</RowLimit></View>";
  $hangingRequests = $listDeploymentRequests.GetItems($camlQuery)
  $context.Load($hangingRequests)
  $context.ExecuteQuery();

  if( $hangingRequests.Count -gt 0 ) {
    write-host "found $($hangingRequests.Count)"
    $hangingRequests | % {
      $hangingItemId = $_.Id
      $hangingItem = $listDeploymentRequests.GetItemById($hangingItemId)
      $context.Load($hangingItem)
      $context.ExecuteQuery()
      $title = $hangingItem["Title"]
      $modified = $hangingItem["Modified"]
      $status = $hangingItem["Status"]
      $diff = ((get-date) - $modified)
      $diffMinutes = $diff.Days*1440 + $diff.Hours*60 + $diff.Minutes
      if( $diffMinutes -gt $timeoutToClearHangingRequestsMinutes ) {
        $message = "Pending request for $($title): reset Status from $status back to Requested."
        write-host $message
        $scriptBlock = {param($context, $hangingItem, $dateTimeStampFormatLong)
          $hangingItem["Status"] = "Requested"
          $hangingItem["UpdatedStatusMessage"] = `
            "Status reset to Requested, $((get-date).ToString($dateTimeStampFormatLong))"
          $hangingItem.Update()
          $context.ExecuteQuery()
        }
        ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
          -scriptBlockParams @($context, $hangingItem, $dateTimeStampFormatLong)
      }
    }
  } else {
    write-host "not found (OK)."
  }
} catch {
  write-host
  write-host "Failed to reset status of hanging requests: $($_.Exception.Message)" -ForegroundColor Red
}

# Step 4. Get relevant list items.
$camlQuery = new-object "Microsoft.SharePoint.Client.CamlQuery"
$camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Status'/>" + `
  "<Value Type='Text'>Requested</Value></Eq></Where></Query><RowLimit>" + `
  $maxSitesToProcessInSingleRun + "</RowLimit></View>"
$listItems = $listDeploymentRequests.GetItems($camlQuery)
$currentUser = $context.Web.CurrentUser
$context.Load($listItems)
$context.Load($currentUser)
$context.ExecuteQuery();
if( $? -eq $false ) {return}

if( $listItems.Count -eq 0 ) {
  write-host
  write-host "Active deployment requests to process not found."
  return
}

$listItems | % {
  $itemId = $_.Id
  $item = $listDeploymentRequests.GetItemById($itemId)
  $context.Load($item)
  $context.ExecuteQuery()
  if( $? -eq $false ) {return}

  $title = $item["Title"]; if($title -ne $null){$title = $title.ToString().Trim()}
  $managedPath = $item["ManagedPath"]; if($managedPath -ne $null){$managedPath = $managedPath.ToString().Trim()}
  $siteUrl = $item["SiteUrl"]; if($siteUrl -ne $null){$siteUrl = $siteUrl.ToString().Trim()}
  $template = $item["WebTemplate"]; if($template -ne $null){$template = $template.ToString().Trim()}
  $language = $item["Language"]; if($language -ne $null){$language = $language.ToString().Trim()}
  $status = $item["Status"]; if($status -ne $null){$status = $status.ToString().Trim()}
  $statusMessage = $item["StatusMessage"]; if($statusMessage -ne $null){$statusMessage = $statusMessage.ToString().Trim()}

  $contextUri = new-object System.Uri($context.Url)
  $targetUrl = $contextUri.Scheme + "://" + $contextUri.Authority + "/" + `
    $managedPath.TrimStart('/').TrimEnd('/') + "/" + $siteUrl.TrimStart('/').TrimEnd('/')
  
  try {
    if( $recreateSiteIfExists -or $removeSiteOnly ) {
      try {
        $scriptBlock = {param($context, $item, $dateTimeStampFormatLong)
          $item["Status"] = "Deleting"
          $item["UpdatedStatusMessage"] = "Delete started, $((get-date).ToString($dateTimeStampFormatLong))"
          $item.Update()
          $context.ExecuteQuery()
        }
        ExecuteAndRetryIfFailed -scriptBlock $scriptBlock -scriptBlockParams @($context, $item, $dateTimeStampFormatLong)
        DeleteSite -tenant $tenant -targetUrl $targetUrl
        RemoveDeletedSiteFromRecycleBin -tenant $tenant -targetUrl $targetUrl
        $scriptBlock = {param($context, $item, $dateTimeStampFormatLong)
          $item["Status"] = "Deleted"
          $item["UpdatedStatusMessage"] = "Deleted, $((get-date).ToString($dateTimeStampFormatLong))"
          $item.Update()
          $context.ExecuteQuery()
        }
        ExecuteAndRetryIfFailed -scriptBlock $scriptBlock -scriptBlockParams @($context, $item, $dateTimeStampFormatLong)
      } catch {
        $scriptBlock = {param($context, $item, $status, $statusMessage, $dateTimeStampFormatLong)
          $item["Status"] = $status
          $item["StatusMessage"] = $statusMessage
          $item["UpdatedStatusMessage"] = "Delete failed, $((get-date).ToString($dateTimeStampFormatLong))"
          $item.Update()
          $context.ExecuteQuery()
        }
        ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
          -scriptBlockParams @($context, $item, $status, $statusMessage, $dateTimeStampFormatLong)
        throw
      }
      if( $removeSiteOnly ) {
        write-host
        write-host "-removeSiteOnly is true. Creation is skipped."
        return
      }
    }

    $isFound = DoesSiteExist -tenant $tenant -targetUrl $targetUrl
    if( !$isFound ) {
      $scriptBlock = {param($context, $item, $dateTimeStampFormatLong)
        $item["Status"] = "Creating"
        $item["StatusMessage"] = "Creation started, $((get-date).ToString($dateTimeStampFormatLong))"
        $item.Update()
        $context.ExecuteQuery()
      }
      ExecuteAndRetryIfFailed -scriptBlock $scriptBlock -scriptBlockParams @($context, $item, $dateTimeStampFormatLong)
      try {
        $needsCustomTemplate=$false
        CreateSite -tenant $tenant -targetUrl $targetUrl -owner $currentUser.LoginName -title $title `
          -template $template -needsCustomTemplate ([ref]$needsCustomTemplate) -language $language `
          -compatibilityLevel $compatibilityLevel -storageMaximumLevel $storageMaximumLevel `
          -userCodeMaximumLevel $userCodeMaximumLevel

        # Adding default site administrators, if required.
        if( $defaultSiteAdministrators -and $defaultSiteAdministrators.Length -gt 0 ) {
          $defaultSiteAdministrators | % {
            try {
              AddSiteAdmin -tenant $tenant -siteUrl $targetUrl -userName $_
            } catch {
              write-host -ForegroundColor Red "Error adding $_ to site admins: $($_.Exception.Message)"
            }
          }
        }
        
        # Adding default visitors, if required.
        if( $defaultVisitors -and $defaultVisitors.Length -gt 0 ) {
          $defaultVisitors | % {
            try {
              AddSiteVisitor -tenant $tenant -siteUrl $targetUrl -userName $_
            } catch {
              write-host -ForegroundColor Red "Error adding $_ to the group of visitors: $($_.Exception.Message)"
            }
          }
        }
        
        $scriptBlock = {param($context, $item, $needsCustomTemplate, $dateTimeStampFormatLong)
          $statusMessage = "Created, $((get-date).ToString($dateTimeStampFormatLong))"
          if( !$needsCustomTemplate ) {
            $item["Status"] = "Created"
          } else {
            $item["Status"] = "CreatedNeedsCustomTemplate"
          }
          $item["StatusMessage"] = $statusMessage
          $item["UpdatedStatusMessage"] = ""
          $item.Update()
          $context.ExecuteQuery()
        }
        ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
          -scriptBlockParams @($context, $item, $needsCustomTemplate, $dateTimeStampFormatLong)

        # Try to deploy legacy WSP-solutions, search for and apply this custom template
        if( $needsCustomTemplate -and $deployLegacySolutionsForCustomWebTemplates ) {
          $appliedErrorMessage = $null
          $isApplied = $false
          try {
            ApplyCustomWebTemplate -siteUrl $targetUrl -template $template -isApplied ([ref]$isApplied) `
              -appliedErrorMessage ([ref]$appliedErrorMessage) -deployLegacySolutions $true `
              -language $language -dateTimeStampFormatShort $dateTimeStampFormatShort
            $errorMessage = $appliedErrorMessage
            if( $errorMessage ) {$errorMessage = ", " + $errorMessage}
            $scriptBlock = {param($context, $item, $isApplied, $errorMessage, $dateTimeStampFormatLong)
              if( $isApplied ) {
                $item["Status"] = "CreatedCustomTemplateApplied"
              } else {
                $item["Status"] = "CreatedCustomTemplateFailed"
              }
              $item["UpdatedStatusMessage"] = `
                "Custom web-template applied $((get-date).ToString($dateTimeStampFormatLong))$errorMessage"
              $item.Update()
              $context.ExecuteQuery()
            }
            ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
              -scriptBlockParams @($context, $item, $isApplied, $errorMessage, $dateTimeStampFormatLong)
          } catch {
            $scriptBlock = {param($context, $item, $dateTimeStampFormatLong)
              $item["Status"] = "CreatedCustomTemplateFailed"
              $item["UpdatedStatusMessage"] = `
                "Applying custom web-template failed, $((get-date).ToString($dateTimeStampFormatLong))"
              $item.Update()
              $context.ExecuteQuery()
            }
            ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
              -scriptBlockParams @($context, $item, $dateTimeStampFormatLong)
          }
        }
          
      } catch {
        $scriptBlock = {param($context, $item, $status, $statusMessage, $dateTimeStampFormatLong)
          $item["Status"] = $status
          $item["StatusMessage"] = $statusMessage
          $item["UpdatedStatusMessage"] = "Creation failed, $((get-date).ToString($dateTimeStampFormatLong))"
          $item.Update()
          $context.ExecuteQuery()
        }
        ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
          -scriptBlockParams @($context, $item, $status, $statusMessage, $dateTimeStampFormatLong)
        throw
      }
    } else {
      write-host
      write-host "The site with URL $targetUrl already exists. Creation failed." -ForegroundColor Yellow
      $scriptBlock = {param($context, $item, $targetUrl, $dateTimeStampFormatLong)
        $statusMessage = "The site with URL $targetUrl already exists, $((get-date).ToString($dateTimeStampFormatLong))"
        $item["Status"] = "Failed"
        $item["UpdatedStatusMessage"] = $statusMessage
        $item.Update()
        $context.ExecuteQuery()
      }
      ExecuteAndRetryIfFailed -scriptBlock $scriptBlock `
        -scriptBlockParams @($context, $item, $targetUrl, $dateTimeStampFormatLong)
    }
    write-host
  } catch {
    # This exception is only thrown in case when deletion, removal of an existing, 
    # or creation of a new site has actually failed. If the site does not exist this exception should not happen.
    $message = $_.Exception.Message + $_.Exception.StackTrace
    write-host
    write-host $message -ForegroundColor Red
    return # Continue to the next site
  }
}

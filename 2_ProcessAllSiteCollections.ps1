param(
  [bool]$unattended = $true
  ,[bool]$disableCustomizations = $false
  ,[bool]$forceOnFailedOnly = $true
  ,[int]$maxFailedAttempts = 0
  ,[bool]$forceOnSucceededOnly = $false
  
  # The next flag instructs the logic to prefer search query to get list of sites in case when the Tenant malfunctions.
  # - Remember that a newly created site is not instantly available in search results due to crawl processing delays.
  # - Security trimming is always applied to search results thus some content can be unavailable in compare with Tenant.
  # - The default logic uses retrieval of site collection properties using a Tenant object.
  # - Using tenant is more reliable in compare to search queries however its service can be randomly unavailable.
  # - If Tenant is unavailable (i.e. the site <tenant>-admin.sharepoint.com is down), change $preferSearchQuery to $true
  # - If $preferSearchQuery = $true the filters of $excludeBySearchProperties are used instead of $excludeBySiteProperties
  ,[bool]$preferSearchQuery = $false
  
  ,[Hashtable]$excludeBySiteProperties = @{
    # Explicitly exclude site collections restricted from changes
    DenyAddAndCustomizePages = @{Value = "Enabled"; Match = $true};
    # Include only unlocked site collections and exclude the locked ones
    #LockState = @{Value = "Unlock"; Match = $false};
    # Include only active site collections and exclude inactive ones
    Status = @{Value = "((Active)|())"; Match = $false};
    # Explicitly exclude site collections having these URLs
    Url = @{Value = "(?i)((/portals/hub)|(/portals/community)|(-public\.))"; Match = $true};
    # Include site collections with this URL pattern only
    #Url2 = @{Value = "(?i)((://[^/]+/$)|(/search$)|(/sites/t)|(/sites/p))"; Match = $false};
    # Include site collections of these templates and exclude others
    #Template = @{Value = "(?i)((STS)|(CMSPUBLISHING))"; Match = $false};
    # Explicitly exclude site collections of these templates
    Template2 = @{Value = "(?i)((SPSMSITEHOST)|(SPSPERS))"; Match = $true};
  }
  ,[Hashtable]$excludeBySearchProperties = @{
    Path = $excludeBySiteProperties.Url;
    #Path2 = $excludeBySiteProperties.Url2;
    #WebTemplate = $excludeBySiteProperties.Template1;
    WebTemplate2 = $excludeBySiteProperties.Template2
  }
   # regex $legacyUrls identifies legacy sites by their URLs.
   # Set $legacyUrls = "." to deploy WSPs to all site collections or $legacyUrls = $null to prevent deployment to all.
  ,[string]$legacyUrls = "(?i)/sites/pp2"
  ,[string]$logFile = "$PSScriptRoot\logs\log" + (get-date).ToString("-yyyy-MM-dd-HH-mm") + ".txt"
  ,[string]$logFileAdminsPending = "$PSScriptRoot\logs\__admins-pending.csv"
  ,[int]$maxLogFiles = 1440

  # The next flag allows to suppress error messages on possible lack of permissions
  # in site collections for the executing account. It instructs to add a site collection admin 
  # while processing a site collection and remove it after the processing is done.
  # If a site admin was added but could not be removed after processing was finished 
  # the incident is reported to the file $logFileAdminsPending and may require manual removal.
  ,[bool]$addSiteAdminWhileProcessing = $true
  ,[bool]$keepSiteAdminAfterProcessing = $false
  ,[bool]$suppressTranscript = $false           # Service flag used during initialization of the environment.
  ,[bool]$suppressSiteCollectionUpdate = $false # Service flag used during initialization of the environment.
  ,[Hashtable]$processedSites = @{}
)

######################################################## FUNCTIONS ########################################################
function AddSiteAdmin(){
param (
  [Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant,
  [string]$siteUrl
)
  if( $tenant -eq $null -or [string]::IsNullOrEmpty($siteUrl) ) {return $false}
  
  $admCtx = $tenant.Context
  $shouldAdd = $false
  try {
    $site = $tenant.GetSiteByUrl($siteUrl)
    $admCtx.Load($site)
    $admCtx.ExecuteQuery()

    $admCtx.Load($site.RootWeb)
    $admCtx.Load($site.RootWeb.CurrentUser)
    $admCtx.ExecuteQuery()
    if( !$site.RootWeb.CurrentUser.IsSiteAdmin ) {
      $shouldAdd = $true
    }
  } catch {
    # If the site was loaded but its root web was not most probably the user has no permissions on this site
    $shouldAdd = $true
  }
  
  if( $shouldAdd ) {
    # Add a site collection admin, temporarily.
    try {
      write-host "Adding a site collection administrator..."
      $user = $tenant.SetSiteAdmin($siteUrl, $admCtx.Credentials.UserName, $true)
      $admCtx.Load($user)
      $admCtx.ExecuteQuery()
      return $true
    } catch {}
  }
  
  return $false
}

function RemoveSiteAdmin(){
param (
  [Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant,
  [string]$siteUrl
)
  if( $tenant -eq $null -or [string]::IsNullOrEmpty($siteUrl) ) {return $false}

  $admCtx = $tenant.Context
  try {
    $site = $tenant.GetSiteByUrl($siteUrl)
    $admCtx.Load($site)
    $admCtx.ExecuteQuery()

    $admCtx.Load($site.RootWeb)
    $admCtx.Load($site.RootWeb.CurrentUser)
    $admCtx.ExecuteQuery()
    if( $site.RootWeb.CurrentUser.IsSiteAdmin ) {
      write-host
      write-host "Removing a site collection administrator..."
      $user = $tenant.SetSiteAdmin($siteUrl, $admCtx.Credentials.UserName, $false)
      $admCtx.Load($user)
      $admCtx.ExecuteQuery()
      return $true
    }
  } catch {}

  return $false
}

function GetSitePropertiesViaSearch(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context
  ,[Parameter(Mandatory=$true)][Hashtable]$excludeBySearchProperties
  ,[Parameter(Mandatory=$true)][Hashtable]$processedSites
)
  $props = ExecuteSearchQuery $context "ContentClass:STS_Site" @("Created Desc")
  # Exclude certain site collections from further processing if they match specific rules. 
  $validSiteCollections = GetValidSiteCollections -props $props -excludeByProperties $excludeBySearchProperties `
    -processedSites $processedSites
  return $validSiteCollections
}

function GetSitePropertiesViaTenant(){
param(
  [Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][Hashtable]$excludeBySiteProperties
  ,[Parameter(Mandatory=$true)][Hashtable]$processedSites
  ,[bool]$suppressSiteCollectionUpdate = $false
)
  if( $tenant -eq $null ) {
    throw "Tenant admin could not be contacted."
  }

  $admCtx = $tenant.Context
  $props = $tenant.GetSiteProperties(0, $true)
  $admCtx.Load($props)
  try {
    $admCtx.ExecuteQuery()
  } catch {
    throw
  }

  # Special case: disable the setting DenyAddAndCustomizePages enabled by default on the top root site.
  $updatable = $false
  $props | % {
    $url = $_.Url
    $uri = new-object System.Uri($url)
    if( $suppressSiteCollectionUpdate -and $uri.AbsolutePath -eq '/' `
        -and $_.DenyAddAndCustomizePages -ne "Disabled" -and !($_.Template -imatch "SPSMSITEHOST") ) {
      write-host
      write-host "Disabling the property DenyAddAndCustomizePages on the top root site collection $url"
      # Setting of the next property may fail with an error like "Current user is not a Tenant admin",
      # but it should still make the required changes.
      $_.DenyAddAndCustomizePages = "Disabled"
      $updatable = $true
    }
  }
  if( $updatable ) {
    $props.Update() | out-null
    try {
      $admCtx.ExecuteQuery()
    } catch {
      # The logic at this point should not cause a fallback to search
      write-host -ForegroundColor Red $_.Exception.Message
    }
  }

  # Exclude certain site collections from further processing if they match specific rules. 
  $validSiteCollections = GetValidSiteCollections -props $props -excludeByProperties $excludeBySiteProperties `
    -processedSites $processedSites
  return ($validSiteCollections | sort -Property LastContentModifiedDate -Descending)
}

function GetValidSiteCollections() {
param (
  [Parameter(Mandatory=$true)][object]$props # Dictionary or Hashtable
  ,[Parameter(Mandatory=$true)][Hashtable]$excludeByProperties
  ,[Parameter(Mandatory=$true)][Hashtable]$processedSites
)
  # Exclude certain site collections from further processing if they match specific rules. 
  write-host
  $validSiteCollections = @()
  $props | % {
    $sc = $_
    $excluded = $false
    $excludeByProperties.GetEnumerator() | % {
      $settingName = $_.Name
      $name = [regex]::Replace($_.Name, "\d+$", [string]::Empty)
      $value = $_.Value.Value
      $match = $_.Value.Match
      try {
        $propValue = $null
        $propValue = $sc.$name
      } catch {}
      if( $propValue -eq $null ) {return}
      
      if( !$excluded -and `
          (($match -and [regex]::IsMatch($propValue, $value)) -or `
          (!$match -and ![regex]::IsMatch($propValue, $value))) ) {
        $url = $sc.Url
        if( $url -eq $null ) {$url = $sc.Path} # In case of using search results
        write-host $url": " -NoNewLine
        write-host "excluded due to the setting $settingName=$propValue" -ForegroundColor Yellow
        $excluded = $true
      }
    } 
    if( !$excluded ) {
      $validSiteCollections += $sc
    } else {
      $lcUrl = $url.ToLower().TrimEnd('/')
      $processedSites.Remove($lcUrl)
    }
  }
  return $validSiteCollections
}

function ProcessSiteCollections(){
param(
  [Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][object[]]$validSiteCollections
  ,[Parameter(Mandatory=$true)][string]$propNameUrl
  ,[bool]$addSiteAdmin
  ,[bool]$keepSiteAdmin
  ,[bool]$disableCustomizations
  ,[bool]$unattendedExecution = $unattended
  ,[Parameter(Mandatory=$true)][Hashtable]$processedSites
  ,[bool]$suppressSiteCollectionUpdate = $false
)
  
  # Step 4. Process all site collections. 
  if( !$validSiteCollections -or $validSiteCollections.Count -eq $null `
      -or $validSiteCollections.Count -lt 1 -or !$propNameUrl ) {
    return
  }
  
  $message = "Processing " + $validSiteCollections.Length + " site collection"
  if( $validSiteCollections.Length -gt 1 ) {
    $message += "s"
  }
  $message += ":"

  write-host
  write-host $message
  write-host
  
  $i = 0
  $validSiteCollections | % {
    $i++
    $url = ($_.$propNameUrl)
    if( !$url ) {return}
    $url = $url.TrimEnd('/')
    $lcUrl= $url.ToLower()
    $prefix = $i.ToString() + " of " + $validSiteCollections.Length.ToString() + ": "
    
    $lastProcessed = $processedSites[$lcUrl].LastProcessed
    $succeeded = $processedSites[$lcUrl].Succeeded
    $customized = $processedSites[$lcUrl].Customized
    $failedAttempts = $processedSites[$lcUrl].FailedAttempts
    $reachedMaxCount = $false
    if( $failedAttempts -eq $null ) {
      $failedAttempts = 0
    } elseif( $maxFailedAttempts -gt 0 -and $failedAttempts -ge $maxFailedAttempts ) {
      $reachedMaxCount = $true
    }

    $force = `
      ($forceOnFailedOnly -and $succeeded -eq $false `
       -and ($maxFailedAttempts -eq 0 -or ($failedAttempts -lt $maxFailedAttempts))) `
       -or ($forceOnSucceededOnly -and $succeeded -eq $true)

    if( $lastProcessed -and ($reachedMaxCount -or $customized -eq !$disableCustomizations) -and !$force ) {
      $message = $prefix + "$url, " + $lastProcessed.ToString("d.M HH:mm") + ","
      write-host $message -NoNewLine
      if( $succeeded ) {
        write-host " OK" -ForegroundColor Green -NoNewLine
      } else {
        $message = " FAILED ($failedAttempts"
        if( $maxFailedAttempts -gt 0 ) {
          $message += " of $maxFailedAttempts"
        }
        $message += ")"
        write-host $message   -ForegroundColor Yellow -NoNewLine
      }
      if( $customized ) {
        write-host ", changed"
      } else {
        write-host ", unchanged"
      }
      return
    }
    $continue = 'n'
    if( $unattendedExecution ) {
      $continue = 'y'
    } else {
      write-host
      write-host "Process site collection $url? [y/n] " -ForegroundColor Yellow -NoNewLine
      $continue = read-host
    }
    if($continue -ne 'y') {return}
    
    #write-host $url
    $wasAdminAdded = $false
    try {     
      if( $addSiteAdmin ) {
        $wasAdminAdded = AddSiteAdmin -tenant $tenant -siteUrl $url
        write-host
      }
    } catch {}
    
    $noCache = $null
    try {
      $message = $prefix + $url
      write-host $message
      if( !$suppressSiteCollectionUpdate ) { # Service flag used during initialization of the environment.
        if( !$legacyUrls -or ![regex]::IsMatch($url, $legacyUrls) ) {
          write-host " Applying modern branding to $url"
          & $PSScriptRoot\3_UpdateSiteCollection.ps1 `
            -siteCollectionUrl $url -disableCustomizations $disableCustomizations -force $force
        } else {
          write-host " Applying legacy WSP-based branding to $url" -ForegroundColor Yellow
          & $PSScriptRoot\4_DeployLegacySolutions.ps1 `
            -siteCollectionUrl $url -disableCustomizations $disableCustomizations -force $force `
            -regexDeploymentUrls $legacyUrls
        }
      }
      $failedAttempts = 0
      $processedSites[$lcUrl] = @{
        LastProcessed = (get-date);
        Succeeded = $true;
        Customized = !$disableCustomizations;
        FailedAttempts = $failedAttempts;
      }
    } catch {
      write-host $_.Exception -ForegroundColor Red
      $failedAttempts++
      if( $customized -eq $null ) {$customized = $false}
      $processedSites[$lcUrl] = @{
        LastProcessed = (get-date); 
        Succeeded = $false; 
        Customized = $customized; 
        FailedAttempts = $failedAttempts;
      }
    } finally {
      if( $noCache ) {
        $processedSites.Remove($lcUrl)  
      }
      if( $wasAdminAdded ) {
        $wasAdminRemoved = $false
        if( !$keepSiteAdmin ) {
          try {
            $wasAdminRemoved = RemoveSiteAdmin -tenant $tenant -siteUrl $url
          } catch {}
        }
        
        if( $wasAdminAdded -and !$wasAdminRemoved) {
          $message = $url + "`t" +  $userName + "`t" + (get-date).ToString("yyyy-MM-dd HH:mm")
          if( test-path $logFileAdminsPending ) {
            $content = [System.IO.File]::ReadAllText($logFileAdminsPending)
            if( $content.IndexOf($message) -eq -1 ) {
              $message >> $logFileAdminsPending
            }
          } else {
            "Url`tSiteAdmin`tAdded" > $logFileAdminsPending
            $message >> $logFileAdminsPending
          }
          if( !$keepSiteAdmin ) {
            write-host "Site admin was added for processing but could not be removed afterwards." -ForegroundColor Yellow
            write-host "The incident was reported to the file $logFileAdminsPending." -ForegroundColor Yellow
          } else {
            write-host "Site admin was kept due to the setting value keepSiteAdminAfterProcessing=true." -ForegroundColor Yellow
            write-host "A report on this was added to the file $logFileAdminsPending." -ForegroundColor Yellow
          }
        }
      }
    }
  }
}
####################################################### //FUNCTIONS #######################################################

####################################################### EXECUTION #########################################################
# Step 1. Start transcripting actions into the log file that can be reviewed later in case of unclear errors.
$tmp = $logFile.Substring(0, $logFile.LastIndexOf('\') + 1)
if( $tmp ) {new-item -ItemType Directory -Force -Path $tmp | out-null}

if( !$suppressTranscript ) {
  try{stop-transcript -erroraction SilentlyContinue | out-null}catch{}
  start-transcript -path $logfile -append | out-null
  write-host
  write-host "Script started $((Get-Date).ToLocalTime())" -ForegroundColor DarkYellow
  write-host "----------------------------------" -ForegroundColor DarkYellow
}
$message = $null

# Step 2. Connect to SharePoint environment amd get a client context.
write-host
write-host "Connecting to SharePoint environment..."
try {
  . $PSScriptRoot\__LoadContext.ps1
} catch {
  # This can mean "Access denied" or invalid URL specified in the __LoadContext.ps1. Try using tenant instead.
  write-host "$siteCollectionUrl is unavailable (no access or invalid URL). Using Tenant's context instead." `
    -ForegroundColor Yellow
}

# Step 3. Get the list of existing site collections. 
$adminContext = $null
$tenant = $null
$validSiteCollectionsBySearch = $null
$validSiteCollectionsByTenant = $null

try {
  # Connect to tenant's admin and get the list of site collections.
  # Note this list does not contain personal sites although may inclide my sites' host.
  write-host
  write-host "Connecting to Tenant admin..."
  $adminContext = GetAdminClientContext
  if( $adminContext -ne $null ) {
    $tenant = New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($adminContext)
    # Try to set the request timeout to 10 minutes; unfortunately, this value is usually ignored by the SPO.
    $adminContext.RequestTimeOut = (60*10*1000)
    $adminContext.Load($tenant)
    $adminContext.ExecuteQuery()
  }
} catch {
  $tenant = $null
  $adminContext = $null
  write-host $_.Exception -ForegroundColor Red
}

if( $context -eq $null -and $tenant -eq $null ) {
  write-host "Connection failed. Further execution is impossible."
  return $processedSites
}

if( $preferSearchQuery ) {  # See the description of this flag in the script's section "param" above.
  write-host
  write-host "Requesting search engine to get available site collections..."
  if( $context -ne $null ) {
    $validSiteCollectionsBySearch = GetSitePropertiesViaSearch -context $context `
      -excludeBySearchProperties $excludeBySearchProperties -processedSites $processedSites
  } else {
    $validSiteCollectionsBySearch = GetSitePropertiesViaSearch -context $tenant.Context `
      -excludeBySearchProperties $excludeBySearchProperties -processedSites $processedSites
  }
} else {
  try {
    write-host
    write-host "Requesting Tenant admin to get available site collections..."
    $validSiteCollectionsByTenant = GetSitePropertiesViaTenant -tenant $tenant `
      -excludeBySiteProperties $excludeBySiteProperties -processedSites $processedSites `
      -suppressSiteCollectionUpdate $suppressSiteCollectionUpdate
  } catch {
    # A timeout exception can happen when the site <tenant>-admin.sharepoint.com is down; fallback to search in this case.
    $validSiteCollectionsByTenant = $null
    # Tenant's web services can be randomly unavailable and throw critical exceptions. 
    # In case of exception try getting site collections via regular search although this way provides less reliable results.
    # Note this list contains personal sites although does not contain my sites' host and search sites.
    write-host
    write-host "Attempt to connect to Tenant admin has failed. Reason: $(GetErrorMessage)" -ForegroundColor Red
    write-host
    write-host "Falling back to search engine to get available site collections..."
    if( $context -ne $null ) {
      $validSiteCollectionsBySearch = GetSitePropertiesViaSearch -context $context `
        -excludeBySearchProperties $excludeBySearchProperties -processedSites $processedSites
    } else {
      $validSiteCollectionsBySearch = GetSitePropertiesViaSearch -context $tenant.Context `
        -excludeBySearchProperties $excludeBySearchProperties -processedSites $processedSites
    }
  }
}

# Step 4. Process all site collections. 
if( $validSiteCollectionsByTenant ) {
  ProcessSiteCollections `
    -tenant $tenant -validSiteCollections $validSiteCollectionsByTenant -propNameUrl "Url" `
    -addSiteAdmin $addSiteAdminWhileProcessing -keepSiteAdmin $keepSiteAdminAfterProcessing `
    -disableCustomizations $disableCustomizations -unattendedExecution $unattended -processedSites $processedSites `
    -suppressSiteCollectionUpdate $suppressSiteCollectionUpdate
} elseif( $validSiteCollectionsBySearch ) {
  ProcessSiteCollections `
    -tenant $tenant -validSiteCollections $validSiteCollectionsBySearch -propNameUrl "Path" `
    -addSiteAdmin $addSiteAdminWhileProcessing -keepSiteAdmin $keepSiteAdminAfterProcessing `
    -disableCustomizations $disableCustomizations -unattendedExecution $unattended -processedSites $processedSites `
    -suppressSiteCollectionUpdate $suppressSiteCollectionUpdate
}

if( !$suppressTranscript ) {
  write-host
  write-host "Script ended $((Get-Date).ToLocalTime())" -ForegroundColor DarkYellow
  write-host "--------------------------------" -ForegroundColor DarkYellow

  # Step 5. Stop transcripting actions and format the output file.
  stop-transcript | out-null
  # Format the log file for more convenient reading in the simplest notepad for Windows.
  $logcontent = [string]::Join("`r`n", (get-content $logfile))
  $logcontent | out-file $logfile

  # Step 6. Cleanup old log files if needed. 
  $allLogFiles = get-childitem $PSScriptRoot `
   -Filter ("*" + $logfile.Substring($logfile.LastIndexOf('-') + 1)) | sort -Property LastWriteTime
  if( $allLogFiles.Length -gt $maxLogFiles ) {
    for( $i = 0; $i -lt ($allLogFiles.Length - $maxLogFiles); $i++ ) {
      try {
        remove-item $allLogFiles[$i] -force -confirm:$false
      } catch {}
    }
  }
}

return $processedSites

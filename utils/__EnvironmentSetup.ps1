param(
  [string]$staticUrlWithCustomizations = "/"
  ,[bool]$disableCustomizations = $false
  ,[string]$webRelativeUrlTargetFolder = "_catalogs/masterpage/customizations/scripts"
  ,[string]$filesForStandardWebTemplates = "customizations\scripts\wt-standard.js"
  ,[int]$defaultLocale = 1033
  ,[int[]]$supportedLocales = @(
    1025, 1068, 1069, 5146, 1026, 1027, 2052, 1028, 1050, 1029, 1030, 1033, 1164, 1043, 1061, 1035, 1036, 1110, 1031,
    1032, 1037, 1081, 1038, 1057, 2108, 1040, 1041, 1087, 1042, 1062, 1063, 1071, 1086, 1044, 1045, 1046, 2070, 1048, 
    1049, 10266, 9242, 1051, 1060, 3082, 1053, 1054, 1055, 1058, 1066, 1106)
   # If $allowOverwritingSupportedLocales = $true the list above is dynamically overwritten in case if the request 
   # to get all available languages succeeds. In case of error the default list of 50 languages shown above is used.
   # Note all 50 dynamically retrievable locales may produce ~400kb of JSON-data. You can improve the JS-load performance
   # by reducing the list of supported locates, for example, $supportedLocales = @(1033, 1035) and setting
   # $allowOverwritingSupportedLocales = $false
  ,[bool]$allowOverwritingSupportedLocales = $true
  ,[int]$compatibilityLevel = 15
)

######################################################## FUNCTIONS ########################################################
function GetAvailableLanguages(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context

)

  $request = CreateWebRequest -requestUrl ($context.Url.TrimEnd('/') + "/_layouts/15/muisetng.aspx?isDlg=1")
  $errorMessage = $null
  $content = ExecuteWebRequest $request -errorMessage ([ref]$errorMessage)
  if( $? -and !$errorMessage ) {
    $availableLanguages = @()
    [regex]::Matches($content, '(?i)<input[^>]+CblAlternateLanguages[^>]+value="([^"]+)"', "SingleLine") | % {
      $availableLanguages += $_.Groups[1].Value
    }
    return $availableLanguages
  }
  
  return $null
}

function GetTenantWebTemplatesAsJson() {
param(
  [Parameter(Mandatory=$true)][Microsoft.Online.SharePoint.TenantAdministration.Tenant]$tenant
  ,[Parameter(Mandatory=$true)][int[]]$supportedLocales
  ,[Parameter(Mandatory=$true)][int]$defaultLocale
  ,[Parameter(Mandatory=$true)][int]$compatibilityLevel
)
  $webTemplates = $null
  try {
    $webTemplates = "(function(){window.webTemplates=window.webTemplates||{};window.webTemplates.Standard=function(){return {Default:$defaultLocale"
    $supportedLocales | % {
      $locale = $_
      $allStandardWebTemplates = $tenant.GetSPOTenantWebTemplates($locale, $compatibilityLevel)
      $tenant.Context.Load($allStandardWebTemplates)
      $tenant.Context.ExecuteQuery()
      if( $? -eq $false ) {return} # Some error

      $webTemplates += "," + $locale.ToString() + ":" + "{"
      
      $categories = $allStandardWebTemplates | select DisplayCategory | sort -Property DisplayCategory -unique
      $categories | % {
        $category = $_.DisplayCategory
        $templates = $allStandardWebTemplates | ? {$_.DisplayCategory -eq $category} | select Name,Title,Description
        $webTemplates += '"' + $category + '":{' 
        $templates | % {
          $webTemplates += '"' + $_.Name + '"' + `
            ":{Title:" + '"' + $_.Title + '"' + ",Description:" + '"' + $_.Description + '"' + "},"
        }
        $webTemplates = $webTemplates.TrimEnd(',')
        $webTemplates += "},"
      }
      
      $webTemplates = $webTemplates.TrimEnd(',')
      $webTemplates += "}"
    }
    $webTemplates += "}};})();"
    #$webTemplates = $webTemplates.Replace("{", "{" + [System.Environment]::NewLine).Replace(
    #  ",", "," + [System.Environment]::NewLine)
  } catch {
    throw $_.Exception
  }
  
  return $webTemplates
}
####################################################### //FUNCTIONS #######################################################

####################################################### EXECUTION #########################################################
# Step 1. Adjust settings on the main site that contains files of customizations.
# A filter to process only the site with customizations.
$excludeBySiteProperties = @{
  Url = @{Value = "(?i)" + $staticUrlWithCustomizations + "$"; Match = $false}; # Include in processing
  Template = @{Value = "(?i)SPSMSITEHOST"; Match = $true}                       # Exclude from processing
}
$excludeBySearchProperties = @{Path = $excludeBySiteProperties.Url; WebTemplate = $excludeBySiteProperties.Template}
$legacyUrls = $null

$processedSites = . $PSScriptRoot\..\2_ProcessAllSiteCollections.ps1 -disableCustomizations $disableCustomizations `
  -preferSearchQuery $false -excludeBySiteProperties $excludeBySiteProperties `
  -excludeBySearchProperties $excludeBySearchProperties -legacyUrls $legacyUrls -addSiteAdminWhileProcessing $true `
  -keepSiteAdminAfterProcessing $true -suppressTranscript $true -suppressSiteCollectionUpdate $true

if( $tenant -eq $null ) {
  write-host "Connection to Tenant Administration failed. Further execution is impossible."
  return
}

# Step 2. Connect to SharePoint environment amd get a client context.
write-host
write-host "Connecting to SharePoint environment..."
if( $staticUrlWithCustomizations ) {
  $contextUri = new-object System.Uri($siteCollectionUrl) # $siteCollectionUrl is loaded by __LoadContext.ps1
  $context = $null
  $siteUrl = $contextUri.Scheme + "://" + $contextUri.Authority + "/" + $staticUrlWithCustomizations.TrimStart('/')
  $siteCollectionUrl = $siteUrl
  $context = GetClientContext
} elseif( $context -eq $null ) {
  $context = GetClientContext
}
if( $context -eq $null ) {
  write-host "Connection failed. Further execution is impossible."
  return
}
write-host
write-host ($context.Url)

# Step 3. Generate a list of available site templates and deploy it to the site with customizations.
if( !$disableCustomizations ) {
  write-host
  write-host "Generating a list of available web templates..." -NoNewLine
  $allSupportedLocales = $null
  if( $allowOverwritingSupportedLocales ) {
    $allSupportedLocales = GetAvailableLanguages -context $context
    if( $? -eq $false -or !$locales.Length ) {
      $allSupportedLocales = $supportedLocales
    }
  } else {
    $allSupportedLocales = $supportedLocales
  }

  $webTemplates = GetTenantWebTemplatesAsJson -tenant $tenant -supportedLocales $allSupportedLocales -defaultLocale $defaultLocale `
    -compatibilityLevel $compatibilityLevel
  $webTemplates > $PSScriptRoot\..\$($filesForStandardWebTemplates.TrimStart('\'))
  write-host "done"

  write-host
  write-host "Deploying customizations to $url..."
  & $PSScriptRoot\..\3_UpdateSiteCollection.ps1 -siteCollectionUrl $context.Url `
    -staticUrlWithCustomizations $staticUrlWithCustomizations -disableCustomizations $false -force $true
} else {
  & $PSScriptRoot\..\3_UpdateSiteCollection.ps1 -siteCollectionUrl $context.Url `
    -staticUrlWithCustomizations $staticUrlWithCustomizations -disableCustomizations $true -force $true
}

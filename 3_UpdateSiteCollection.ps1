param(
  [string]$siteCollectionUrl = $null # The default value is $null. In this case site collection specified in __LoadContext.ps1 will be processed.
  ,[bool]$disableCustomizations = $false
  ,[bool]$force = $false
   # The setting $evaluateOnly allows emulating the complete processing without making actual changes to the processed site collection.
  ,[bool]$evaluateOnly = $false
  ,[string]$staticUrlWithCustomizations = "/"
  ,[string]$webRelativeUrlTargetFolder = "_catalogs/masterpage/customizations"
  ,[string]$rootFolderWithCustomizations = "$PSScriptRoot\customizations"
  ,[string[]]$customActionFiles = @(
     "scripts/jquery-3.1.0.min.js", "scripts/custom-ui.js", "css/custom-ui.css", 
     "scripts/SPO-Responsive.js", "css/SPO-Responsive.css"
   )
   # The next reference "customActionLegacyFiles" is only used to identify if a site has legacy customizations.
  ,[string[]]$customActionLegacyFiles = @("scripts/custombranding.js")
  ,[Hashtable]$customActionMenuItems = @{
     CreateSite = @{
       Url = "scripts/custom-sa-create-site.js";
       LocalizedTexts = @{
         Default = 1033;
         1033 = @{Title = "Add new site collection"; Description = ""};
         1035 = @{Title = "Lisää uusi sivustokokoelma"; Description = ""}
       };
       Order = 1;
       Rights = "FullMask";
       SiteAdminsOnly = $true;
       PermittedOnUrls = @($staticUrlWithCustomizations)
     }
   }
  ,[string[]]$disableFeatures = @(
    # Minimal Download Strategy, web feature (https://<sc>/_layouts/15/start.aspx#/SitePages/Home.aspx --> https://<sc>/SitePages/Home.aspx)
    [guid]("87294c72-f260-42f3-a41b-981a2ffce37a")
    # Sideloading (site collection feature)
    #,[guid]("ae3a1339-61f5-4f8f-81a7-abd2da956a7d")
  )
  ,[string[]]$enableFeatures = @(
    # Minimal Download Strategy, web feature (https://<sc>/_layouts/15/start.aspx#/SitePages/Home.aspx --> https://<sc>/SitePages/Home.aspx)
    #[guid]("87294c72-f260-42f3-a41b-981a2ffce37a")
    # Sideloading (site collection feature)
    #,[guid]("ae3a1339-61f5-4f8f-81a7-abd2da956a7d")
  )
  ,[hashtable]$NavigationNodes = @{
    Default = 1033;
    1033 = @{
      "Enterprise" = @{Url = "https://products.office.com/en/business/office-365-enterprise-e3-business-software"; IsExternal = $true; Order = 1};
      "Cloud" = @{Url = "https://azure.microsoft.com/"; IsExternal = $true; Order = 2};
      "Services" = @{Url = "https://www.microsoft.com/en-us/cloud-platform/sql-server"; IsExternal = $true; Order = 3};
    };
    1035 = @{
      "Yrityksen" = @{Url = "https://products.office.com/fi-FI/business/office-365-enterprise-e3-business-software"; IsExternal = $true; Order = 1};
      "Pilvi" = @{Url = "https://azure.microsoft.com/"; IsExternal = $true; Order = 2};
      "Palvelut" = @{Url = "https://www.microsoft.com/en-us/cloud-platform/sql-server"; IsExternal = $true; Order = 3};
    }
  }
  ,[string]$webTemplatesWithNoWelcomePage = "(?i)((POLICYCTR)|(OFFILE)|(BDR))"
  ,[string]$webTemplatesWithNoTopNavigation = "(?i)POLICYCTR"
)
######################################################## FUNCTIONS #########################################################
function AddAlternateCssUrl(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$fileUrlCss,
  [string]$staticCustomizationUrl,
  [bool]$evaluateOnly
) 
  if( $staticCustomizationUrl ) {
    $url = $staticCustomizationUrl.TrimEnd('/') + '/' + $fileUrlCss.TrimStart('/')
  } else {
    $url = $context.Site.RootWeb.ServerRelativeUrl.TrimEnd('/') + '/' + $fileUrlCss.TrimStart('/')
  }
  if( !$evaluateOnly) {
    $context.Site.RootWeb.AlternateCssUrl = $url
    $context.ExecuteQuery()
  }
}

function AddCustomSiteAction(){
param (
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$fileUrlJs,
  [Parameter(Mandatory=$true)][string]$title,
  [Parameter(Mandatory=$true)][int]$sequence,
  [string]$description,
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.PermissionKind]$rights,
  [bool]$siteAdminsOnly,
  [string[]]$permittedOnUrls,
  [string]$staticCustomizationUrl,
  [bool]$evaluateOnly
)
  $isPermitted = $false
  if( $permittedOnUrls -ne $null -and $permittedOnUrls.Length -gt 0 ) {
    $permittedOnUrls | % {
      $permittedUrl = $_
      if( $permittedUrl.TrimEnd('/') -ieq $context.Site.RootWeb.ServerRelativeUrl.TrimEnd('/') ) {
        $isPermitted = $true
      }
    }
  } else {
    $isPermitted = $true
  }
  if( !$isPermitted ) {
    $message = ' "' + $title + '"' + " is not required on this site; skipped."
    write-host $message
    return
  }
  
  $siteCustomActions = $context.Site.UserCustomActions
  $context.Load($siteCustomActions)
  $context.ExecuteQuery()

  $url = $null
  if( $staticCustomizationUrl ) {
    $url = $staticCustomizationUrl.TrimEnd('/') + '/' + $fileUrlJs.TrimStart('/')
  } else {
    $url = $context.Site.RootWeb.ServerRelativeUrl.TrimEnd('/') + '/' + $fileUrlJs.TrimStart('/')
  }
  $existingCustomActions = ($siteCustomActions | ? {$_.Url -imatch $fileUrlJs -and $_.Title -ieq $title})
  if( $existingCustomActions.Length -eq 0 ) {
    write-host " adding custom action for $($url.Substring($url.LastIndexOf('/') + 1))"
    if( !$evaluateOnly) {
      $basePermissions = new-object "Microsoft.SharePoint.Client.BasePermissions"
      $basePermissions.Set($rights)
      
      $customAction = $siteCustomActions.Add()
      $uniqueId = $null
      if( $url.IndexOf('?') -gt -1 ) {
        $uniqueId = ' + "&rnd=" + Math.random().toString(16).slice(2)'
      } else {
        $uniqueId = ' +"?rnd=" + Math.random().toString(16).slice(2)'
      }
      $innerScript = `
        'var el = document.createElement("script");' + `
        'el.type = "text/javascript";' + `
        'el.src = "' + $url + '"' + $uniqueId + ';' + `
        'document.getElementsByTagName("head")[0].appendChild(el);'
      $script = $null
      if( $siteAdminsOnly ) {
        $script = `
        'javascript:(function(){' + `
          'SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {' + `
            'var ctx = new SP.ClientContext.get_current();' + `
            'var currentUser = ctx.get_web().get_currentUser();' + `
            'ctx.load(currentUser);' + `
            'ctx.executeQueryAsync(success, error);' + `
            'function success() {' + `
            '  if( !currentUser.get_isSiteAdmin() ) {' + `
            '    alert("This operation is permitted only to site administrators.");' + `
            '  } else {' + `
            '    ' + $innerScript + `
            '  }' + `
            '}' + `
            'function error(sender, args) {' + `
            '  alert("Error occured: " + args.get_message());' + `
            '}' + `
          '});' + `
        '})();'
      } else {
        $script = 'javascript:(function(){' + $innerScript + '})();'
      }
      #$customAction.Url = `
      #  'javascript:(function(){var el=document.createElement("script");el.type="text/javascript";el.src="' `
      #  + $url + '";document.getElementsByTagName("head")[0].appendChild(el);})();'
      $customAction.Url = $script
      $customAction.Location = "Microsoft.SharePoint.StandardMenu"
      $customAction.Group = "SiteActions"
      $customAction.Sequence = $sequence
      $customAction.Title = $title
      $customAction.Description = $description
      $customAction.Rights = $basePermissions
      $customAction.Update()
      try {
        $context.ExecuteQuery()
      } catch {
        #$_.Exception.StackTrace
        throw  # Cancels execution of this script.
      }
    }
  } elseif( $existingCustomActions.Length -gt 1 ) {
    # Delete possible duplicates of this custom action
    for( $i = $existingCustomActions.Length - 1; $i -gt 0; $i-- ) {
      write-host " deleting a duplicate of $($url.Substring($url.LastIndexOf('/') + 1))"
      if( !$evaluateOnly) {
        $existingCustomActions[$i].DeleteObject()
        $context.ExecuteQuery()
      }
    }
  }
}

function AddSiteScriptCustomAction(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$fileUrlCssOrJs,
  [Parameter(Mandatory=$true)][int]$sequence,
  [string]$staticCustomizationUrl,
  [bool]$evaluateOnly
)
  $siteCustomActions = $context.Site.UserCustomActions
  $context.Load($siteCustomActions)
  $context.ExecuteQuery()
   
  $isCss = $false
  $existingCustomActions = $null
  $url = $null
  if( [regex]::IsMatch($fileUrlCssOrJs, "(?i)\.css\??") ) {
    $isCss = $true
    if( $staticCustomizationUrl ) {
      $url = $staticCustomizationUrl.TrimEnd('/') + '/' + $fileUrlCssOrJs.TrimStart('/')
    } else {
      $url = $context.Site.RootWeb.ServerRelativeUrl.TrimEnd('/') + '/' + $fileUrlCssOrJs.TrimStart('/')
    }
    $existingCustomActions = ($siteCustomActions | ? {$_.ScriptBlock -imatch $url})
  } else {
    if( $staticCustomizationUrl ) {
      $uri = new-object System.Uri($context.Url)
      $url = $uri.Scheme + "://" + $uri.Authority
      $url += $staticCustomizationUrl.TrimEnd('/') + '/' + $fileUrlCssOrJs.TrimStart('/')
    } else {
      if( !$useLocalEnvironment ) { # $useLocalEnvironment is defined in __LoadContext.ps1
        $url = "~SiteCollection/" + $fileUrlCssOrJs.TrimStart('/')
      } else {
        # To have common compatibility between SPO and "on-premises"
        $url = $context.Site.Url.TrimEnd('/') + '/' + $fileUrlCssOrJs.TrimStart('/')
      }
    }
    $existingCustomActions = ($siteCustomActions | ? {$_.ScriptSrc -ieq $url})
  }
  if( $existingCustomActions.Length -eq 0 ) {
    write-host " adding custom action for $($url.Substring($url.LastIndexOf('/') + 1))"
    if( !$evaluateOnly) {
      $customAction = $siteCustomActions.Add()
      if( $isCSS ) {
        $customAction.ScriptBlock = `
          '(function(){var el=document.createElement("link");el.rel="stylesheet";el.type="text/css";el.href="' `
          + $url + '";document.getElementsByTagName("head")[0].appendChild(el);})();'
      } else {
        if( !$useLocalEnvironment ) { # $useLocalEnvironment is defined in __LoadContext.ps1
          $customAction.ScriptSrc = $url
        } else {
          # This way causes the problem with Microsoft.SharePoint.Utilities.SPUtility.MakeBrowserCacheSafeLayoutsUrl in SharePoint "on-premises"
          # The next way works seamlessly in SharePoint "on-premises".
          $customAction.ScriptBlock = "document.write('" + '<script src="' + $url + '"><\/script>' + "');"
        }
      }
      $customAction.Location = "ScriptLink"
      $customAction.Sequence = $sequence  

      $customAction.Update()
      try {
        $context.ExecuteQuery()
      } catch {
        throw  # Cancels execution of this script.
      }
    }
  } elseif( $existingCustomActions.Length -gt 1 ) {
    # Delete possible duplicates of this custom action
    for( $i = $existingCustomActions.Length - 1; $i -gt 0; $i-- ) {
      write-host " deleting a duplicate of $($url.Substring($url.LastIndexOf('/') + 1))"
      if( !$evaluateOnly) {
        $existingCustomActions[$i].DeleteObject()
        $context.ExecuteQuery()
      }
    }
  }
}

function DeployCustomizations(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$webRelativeUrlListRootFolder,
  [Parameter(Mandatory=$true)][string]$fileSystemRootFolderWithCustomizations,
  [bool]$evaluateOnly
)
  write-host
  write-host "Uploading customizations..."
  get-childitem -path $fileSystemRootFolderWithCustomizations -Recurse | ? {!$_.PSIsContainer} | % {
    $subFolderUrl = $_.FullName.Substring( `
      $fileSystemRootFolderWithCustomizations.TrimEnd('\').Length + 1).Replace('\','/').TrimStart('/')
    $lastSlashIndex = $subFolderUrl.LastIndexOf('/')
    if( $lastSlashIndex -gt -1 ) {
      $subFolderUrl = $webRelativeUrlListRootFolder.TrimEnd('/') + '/' + $subFolderUrl.Substring(0, $lastSlashIndex)
    } else {
      $subFolderUrl = $webRelativeUrlListRootFolder.TrimEnd('/')
    }
    $fileUrl = $context.Web.ServerRelativeUrl.TrimEnd('/') + '/' + $subFolderUrl + '/' + $_.Name
    write-host " $fileUrl"
    if( !$evaluateOnly) {
      UploadFile -web $context.Web -filePath $_.FullName -webRelativeUrl $subFolderUrl -evaluateOnly $evaluateOnly
    }
  }
}

function DisableFeature(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][guid]$featureId,
  [bool]$force = $true,
  [bool]$evaluateOnly
)
  $context.Load($context.Site.Features)
  $context.Load($context.Web.Features)
  $context.ExecuteQuery()

  $context.Site.Features | ? {$_.DefinitionId -eq $featureId} | % {
    write-host " $featureId"
    if( !$evaluateOnly) {
      $context.Site.Features.Remove($featureId, $force)
      $context.ExecuteQuery()
    }
  }
  $context.Web.Features | ? {$_.DefinitionId -eq $featureId} | % {
    write-host " $featureId"
    if( !$evaluateOnly) {
      $context.Web.Features.Remove($featureId, $force)
      $context.ExecuteQuery()
    }
  } 
}

function EnableFeature(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][guid]$featureId,
  [bool]$force = $true,
  [bool]$evaluateOnly
)
  $context.Load($context.Site.Features)
  $context.Load($context.Web.Features)
  $context.ExecuteQuery()

  if( ($context.Site.Features | ? {$_.DefinitionId -eq $featureId}).Length -eq 0 ) {
    if( !$evaluateOnly) {
      $context.Site.Features.Add( `
        $featureId, $force, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None) | out-null
    }
    try {
      if( !$evaluateOnly) {
        $context.ExecuteQuery()
      }
      write-host " $featureId"
    } catch {}
  }
  if( ($context.Web.Features | ? {$_.DefinitionId -eq $featureId}).Length -eq 0 ) {
    if( !$evaluateOnly) {
      $context.Web.Features.Add( `
        $featureId, $force, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None) | out-null
    }
    try {
      if( !$evaluateOnly) {
        $context.ExecuteQuery()
      }
      write-host " $featureId"
    } catch {}
  }
}

function EnsureCustomFolders(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$webRelativeUrlListRootFolder,
  [Parameter(Mandatory=$true)][string]$fileSystemRootFolderWithCustomizations,
  [bool]$evaluateOnly
)
  write-host
  write-host "Creating necessary folders..."
  $allSubFolders = get-childitem $fileSystemRootFolderWithCustomizations -Recurse | ? {$_.PSIsContainer} 
  if( $allSubFolders.Length -gt 0 ) {
    $allSubFolders| % {
      $subFolderUrl = $webRelativeUrlListRootFolder.TrimEnd('/') + '/' `
        + $_.FullName.Substring( `
          $fileSystemRootFolderWithCustomizations.TrimEnd('\').Length + 1).replace('\','/').TrimStart('/')
      $folderUrl = $context.Web.ServerRelativeUrl.TrimEnd('/') + '/' + $subFolderUrl
      write-host " $folderUrl"
      if( !$evaluateOnly) {
        EnsureFolders -context $context -parentFolder $context.Web.RootFolder `
          -webRelativeFolderUrl $subFolderUrl -evaluateOnly $evaluateOnly
      }
    }
  } else {
    $folderUrl = $context.Web.ServerRelativeUrl.TrimEnd('/') + '/' + $webRelativeUrlListRootFolder.TrimStart('/')
    write-host " $folderUrl"
    if( !$evaluateOnly) {
      EnsureFolders -context $context -parentFolder $context.Web.RootFolder `
        -webRelativeFolderUrl $webRelativeUrlListRootFolder.TrimStart('/') -evaluateOnly $evaluateOnly
    }
  }
}

function EnsureTopNavigationNode(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$title,
  [Parameter(Mandatory=$true)][string]$url,
  [bool]$isExternal = $true,
  [bool]$ignoreDuplicates = $false,
  [bool]$evaluateOnly
)
  $topNavigationBar = $context.Web.Navigation.TopNavigationBar
  $context.Load($topNavigationBar)
  $context.ExecuteQuery()

  $count = 0
  $duplicateNodes = @()
  if( !$ignoreDuplicates ) {
    $topNavigationBar | % {
      if( $url -imatch $_.Url -and $_.Url -ne "/" ) {
        $count++
        if( $count -gt 1 ) {
          $duplicateNodes += $_
        }
      }
    }
  }

  if( $count -eq 1 ) {
    return # This node already exists
  } elseif( $duplicateNodes.Length -gt 0 ) {
    $duplicateNodes | % {
      write-host " deleting a duplicate navigation node $($_.Title): $($_.Url)"
      if( !$evaluateOnly) {
        $_.DeleteObject()
        $context.ExecuteQuery()
      }
    }
    return
  }

  write-host " $($title): $url"
  
  if( !$isExternal ) {
    if( $url -match "://" ) {  # If this is an absolute url.
      $uri = new-object System.Uri($url)
      if( !$uri.AbsolutePath.StartsWith(
            $context.Site.RootWeb.ServerRelativeUrl, [System.StringComparison]::InvariantCultureIgnoreCase) ) {
        $isExternal = $true
      }
    } elseif ( !$url.StartsWith(
        $context.Site.RootWeb.ServerRelativeUrl, [System.StringComparison]::InvariantCultureIgnoreCase) ) {
      $isExternal = $true
    }
  }
  
  $newNode = new-object "Microsoft.SharePoint.Client.NavigationNodeCreationInformation"
  $newNode.IsExternal = $isExternal
  $newNode.Title = $title
  $newNode.Url = $url
  $newNode.AsLastNode = $true

  if( !$evaluateOnly) {
    # http://sharepoint.stackexchange.com/questions/96595/how-to-set-navigation-in-office-365-site-using-powershell
    $context.Load($topNavigationBar.Add($newNode))
    $context.ExecuteQuery()
  }
}

function HasCustomizations() {
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context
  ,[Parameter(Mandatory=$true)][string[]]$customActionFiles
)

  $hasCustomizations = $false
  
  $siteCustomActions = $context.Site.UserCustomActions
  $context.Load($siteCustomActions)
  $context.ExecuteQuery()

  $customActionFiles | % {
    if( !$hasCustomizations ) {
      $url = $_
      $siteCustomActions | % {
        if( $_.ScriptBlock -imatch $url -or $_.ScriptSrc -imatch $url ) {
          $hasCustomizations = $true
        }
      }
    }
  }
  
  return $hasCustomizations
}

function HasLegacyCustomizations(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context
  ,[string]$defaultMaster = "seattle.master"
  ,[string[]]$customActionLegacyFiles
)

  $hasLegacyCustomizations = $false
  if( !($context.Site.RootWeb.CustomMasterUrl -imatch ("/" + $defaultMaster.TrimStart('/').TrimEnd('$') + "$")) `
      -or ![string]::IsNullOrEmpty($context.Site.RootWeb.AlternateCssUrl) ) {
    $hasLegacyCustomizations = $true
  } elseif( $customActionLegacyFiles -ne $null ) {
    $hasLegacyCustomizations = HasCustomizations -context $context -customActionFiles $customActionLegacyFiles
  }
  
  return $hasLegacyCustomizations
}

function RemoveAlternateCssUrl(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [bool]$evaluateOnly
)
  if( !$evaluateOnly) {
    $context.Site.RootWeb.AlternateCssUrl = $null
    $context.Site.RootWeb.Update()
    $context.ExecuteQuery()
  }
}

function RemoveSiteCustomActions(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string[]]$customActionFiles,
  [bool]$evaluateOnly
)
  $siteCustomActions = $context.Site.UserCustomActions
  $context.Load($siteCustomActions)
  $context.ExecuteQuery()

  $existingCustomActions = @()
  $customActionFiles | % {
    $url = $_
    $siteCustomActions | % {
      $customAction = $_
      if( $_.ScriptBlock -imatch $url -or $_.ScriptSrc -imatch $url -or $_.Url -imatch $url ) {
        if( ($existingCustomActions | ? {$_ -eq $customAction}).length -eq 0 ) {
          $existingCustomActions += $customAction
        }
      }
    }
  }

  #write-host " before: $($siteCustomActions.Count)"
  $existingCustomActions | % {
    #write-host " Deleting custom action"
    if( !$evaluateOnly) {
      $_.DeleteObject()
      $context.ExecuteQuery()
    }
  }
  #write-host " after: $($siteCustomActions.Count)"
}

function RemoveTopNavigationNodes(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)]$nodes,
  [bool]$evaluateOnly
)
  $topNavigationBar = $context.Web.Navigation.TopNavigationBar
  $context.Load($topNavigationBar)
  $context.ExecuteQuery()

  $existingNodes = @()
  $nodes | % {
    $url = $_.Value.Url
    $topNavigationBar | % {
      $uri = $null
      if( [System.Uri]::TryCreate($_.Url, "Absolute", [ref]$uri) ) {
        if( $url -ieq $uri.AbsoluteUri ) {
          $existingNodes += $_
        }
      } elseif( $url -ieq $_.Url ) {
        $existingNodes += $_
      }
    }
  }
  
  #write-host " before: $($topNavigationBar.Count)"
  $existingNodes | % {
    write-host " deleting a navigation node $($_.Title): $($_.Url)"
    if( !$evaluateOnly) {
      $_.DeleteObject()
      $context.ExecuteQuery()
    }
  }
  #write-host " after: $($topNavigationBar.Count)"
}

function UpdateCurrentNavigation(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource]$source,
  [bool]$evaluateOnly
)
  if( !(IsPublishingWeb -web $context.Site.RootWeb) ) {
    return
  } elseif( !$source ) {
    $source = [Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource]::PortalProvider
  }
  $sourceId = 
    [int][enum]::Parse([Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource],$source)
  if( $? -eq $false ) {return} # Invalid value of $source

  $navigation = new-object Microsoft.SharePoint.Client.Publishing.Navigation.WebNavigationSettings(
    $context, $context.Site.RootWeb)
  $currentNavigation = $navigation.CurrentNavigation
  $context.Load($navigation)
  $context.Load($currentNavigation)
  $context.ExecuteQuery()

  if( $currentNavigation.Source -ne $source ) {
    if( !$evaluateOnly ) {
      $currentNavigation.Source = $sourceId
      $navigation.Update($null)
      $context.ExecuteQuery()
    }     
  }

  if( !$evaluateOnly `
      -and $source -eq [Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource]::PortalProvider) {
    $web = $context.Site.RootWeb
    #$web.AllProperties["__NavigationOrderingMethod"] = 0         # 0 - sort automatically, 1 - sort manually
    #$web.AllProperties["__NavigationAutomaticSortingMethod"] = 1 # 0 - by Title, 1 - by Created Date, 2 - by Last Modified Date
    #$web.AllProperties["__NavigationSortAscending"] = $true.ToString()
    $web.AllProperties["__CurrentNavigationIncludeTypes"] = 3     # 0 - none, 1 - Sub-sites only, 2 - Pages only, 3 - Sub-sites and Pages
    $web.Update()
    $context.ExecuteQuery()
  }
}

function UpdateGlobalNavigation(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource]$source,
  [bool]$evaluateOnly
)
  if( !(IsPublishingWeb -web $context.Site.RootWeb) ) {
    return
  } elseif( !$source ) {
    $source = [Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource]::PortalProvider
  }
  $sourceId = 
    [int][enum]::Parse([Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource],$source)
  if( $? -eq $false ) {return} # Invalid value of $source
  
  $navigation = new-object Microsoft.SharePoint.Client.Publishing.Navigation.WebNavigationSettings(
    $context, $context.Site.RootWeb)
  $globalNavigation = $navigation.GlobalNavigation
  $context.Load($navigation)
  $context.Load($globalNavigation)
  $context.ExecuteQuery()

  if( $globalNavigation.Source -ne $source ) {
    if( !$evaluateOnly ) {
      $globalNavigation.Source = $sourceId
      $navigation.Update($null)
      $context.ExecuteQuery()
    }
  }

  if( !$evaluateOnly `
      -and $source -eq [Microsoft.SharePoint.Client.Publishing.Navigation.StandardNavigationSource]::PortalProvider) {
    $web = $context.Site.RootWeb
    $web.AllProperties["__GlobalNavigationIncludeTypes"] = 0    # 0 - none, 1 - Sub-sites only, 2 - Pages only, 3 - Sub-sites and Pages
    $web.Update()
    $context.ExecuteQuery()
  }
}
####################################################### //FUNCTIONS ########################################################
   
####################################################### EXECUTION ##########################################################
if( $evaluateOnly ) {
  write-host 
  write-host "EVALUATION MODE: no real changes applied." -ForegroundCOlor Yellow
}

# Step 1. Connect to SharePoint environment amd get a client context.
write-host
write-host "Connecting to SharePoint environment..."
if( $siteCollectionUrl ) {
  . $PSScriptRoot\__LoadContext.ps1 -siteCollectionUrl $siteCollectionUrl
} else {
  . $PSScriptRoot\__LoadContext.ps1
}
if( $context -eq $null ) {
  write-host "Connection failed. Further execution is impossible."
  return
}
write-host
write-host ($context.Url)

# Step 2. Check if this site has been created without any web template (to be set later).
$context.Load($context.Site.RootWeb.RootFolder)
$context.ExecuteQuery()
# Sites created from special templates do not have a welcome page.
$hasWebTemplate = ![string]::IsNullOrEmpty($context.Site.RootWeb.RootFolder.WelcomePage) `
  -or $context.Site.RootWeb.WebTemplate -imatch $webTemplatesWithNoWelcomePage
if( !$hasWebTemplate `
    -and ($context.Site.RootWeb.ServerRelativeUrl.TrimEnd('/') -ine $staticUrlWithCustomizations.TrimEnd('/')) ) {
  # If this site has no web template and it is not a special site with customizations.
  write-host
  write-host "The site $($context.Site.RootWeb.Url) does not have a web template. Customization is omitted."
  write-host
  # Set the value of upper scope variable. Reference type cannot be used here because this script 
  # should be also able to run as standalone.
  Set-Variable -Name noCache -Value $true -Scope 1
  return
}

# Step 3. Check if customizations have already been applied and if they have to be removed.
$hasLegacyCustomizations = HasLegacyCustomizations -context $context -customActionLegacyFiles $customActionLegacyFiles
if( $hasLegacyCustomizations ) {  # If this site has any legacy WSP-based customizations.
  write-host
  write-host "Legacy customizations have already been applied to this site."
  return
}

$hasCustomizations = HasCustomizations -context $context -customActionFiles $customActionFiles
if( $disableCustomizations ) {
  write-host
  write-host "Disabling customizations..."
  if( $hasCustomizations -or $force ) {
    write-host
    write-host "Deleting custom actions..."
    $fileUrls = @()
    $customActionFiles | % {$fileUrls += $_}
    $customActionMenuItems.GetEnumerator() | % {$fileUrls += $_.Value.Url}
    RemoveSiteCustomActions -context $context -customActionFiles $fileUrls -evaluateOnly $evaluateOnly

    # Compliance Policy Centar cannot have top navigation nodes (exemption?)
    if( $hasWebTemplate -and !($context.Site.RootWeb.WebTemplate -imatch "POLICYCTR") ) {
      write-host
      write-host "Deleting custom navigation nodes..."
      $lcid = [int]$context.Site.RootWeb.Language
      $tmp = $NavigationNodes.($lcid)
      $nodes = $null
      if( $tmp ) {
        $nodes = $tmp.GetEnumerator()
      } else {
        $nodes = $NavigationNodes.($NavigationNodes.Default).GetEnumerator()
      }
      RemoveTopNavigationNodes -context $context -nodes $nodes -evaluateOnly $evaluateOnly
      # Update navigation sources (note: it has effect on publishing sites only).
      UpdateGlobalNavigation $context "TaxonomyProvider"
      UpdateCurrentNavigation $context "TaxonomyProvider"
    }
    #RemoveAlternateCssUrl -context $context -evaluateOnly $evaluateOnly
    write-host
  }
  return
}

if( $hasCustomizations -and !$force ) {
  write-host
  write-host "Customizations have already been applied earlier to this site."
  write-host "Use the parameters -disableCustomizations or -force to process them."
  return
}

$useStaticCustomizationUrl = ![string]::IsNullOrEmpty($staticUrlWithCustomizations) -and `
  !$context.Site.RootWeb.ServerRelativeUrl.TrimEnd('/').Equals($staticUrlWithCustomizations.TrimEnd('/'))

# Step 4. Upload customizations to the target document library.
if( !$useStaticCustomizationUrl ) {
  write-host
  write-host "Obtaining the target list for $webRelativeUrlTargetFolder and deploying files..."
  $scriptBlock = {param($context, $webRelativeUrlListRootFolder, $fileSystemRootFolderWithCustomizations)
    EnsureCustomFolders -context $context -webRelativeUrlListRootFolder $webRelativeUrlListRootFolder `
      -fileSystemRootFolderWithCustomizations $fileSystemRootFolderWithCustomizations -evaluateOnly $evaluateOnly
    DeployCustomizations -context $context -webRelativeUrlListRootFolder $webRelativeUrlListRootFolder `
      -fileSystemRootFolderWithCustomizations $fileSystemRootFolderWithCustomizations -evaluateOnly $evaluateOnly
  }
  ProcessFilesInDocumentLibrary `
    -web $context.Site.RootWeb -webRelativeUrlListRootFolder $webRelativeUrlTargetFolder -scriptBlock $scriptBlock `
    -scriptBlockParams @($context, $webRelativeUrlTargetFolder, $rootFolderWithCustomizations) -evaluateOnly $evaluateOnly
}

# Step 5. Add custom actions.
write-host
write-host "Adding self-executing custom scripts to the site collection..."
[int]$sequence = 1100
$customActionFiles | % {
  $fileUrl = $webRelativeUrlTargetFolder.TrimStart('/').TrimEnd('/') + '/' + $_.TrimStart('/')
  if( $useStaticCustomizationUrl ) {
    AddSiteScriptCustomAction -context $context -fileUrlCssOrJs $fileUrl -sequence $sequence `
      -staticCustomizationUrl $staticUrlWithCustomizations -evaluateOnly $evaluateOnly
  } else {
    AddSiteScriptCustomAction -context $context -fileUrlCssOrJs $fileUrl -sequence $sequence `
      -staticCustomizationUrl $null -evaluateOnly $evaluateOnly
  }
  $sequence += 10
}
write-host
write-host "Adding custom menu items to site settings..."
$customActionMenuItems.GetEnumerator() | sort {$_.Value.Order} | % {
  $custoSitemAction = $_.Value
  $fileUrljs = $webRelativeUrlTargetFolder.TrimStart('/').TrimEnd('/') + '/' + $custoSitemAction.Url.TrimStart('/')
  $lcid = [int]$context.Site.RootWeb.Language
  $texts = $custoSitemAction.LocalizedTexts.($lcid)
  if( $texts -eq $null ) {
    $texts = $custoSitemAction.LocalizedTexts.($custoSitemAction.LocalizedTexts.Default)
  }
  $title = $texts.Title
  $description = $texts.Description
  $rights = $custoSitemAction.Rights
  $siteAdminsOnly = $custoSitemAction.SiteAdminsOnly
  $permittedOnUrls = $custoSitemAction.PermittedOnUrls
  if( !$rights ) {
    $rights = "EmptyMask"
  }
  if( $useStaticCustomizationUrl ) {
    AddCustomSiteAction -context $context -fileUrlJs $fileUrlJs -title $title -sequence $sequence `
      -description $description -rights $rights -siteAdminsOnly $siteAdminsOnly -permittedOnUrls $permittedOnUrls `
      -staticCustomizationUrl $staticUrlWithCustomizations -evaluateOnly $evaluateOnly
  } else {
    AddCustomSiteAction -context $context -fileUrlJs $fileUrlJs -title $title -sequence $sequence `
      -description $description -rights $rights -siteAdminsOnly $siteAdminsOnly -permittedOnUrls $permittedOnUrls `
      -staticCustomizationUrl $null -evaluateOnly $evaluateOnly
  }
  $sequence += 10
}

# Step 6. Enable and disable features if necessary.
write-host
write-host "Disabling undesired features..."
$disableFeatures | % {
  $featureId = $_
  DisableFeature -context $context -featureId $featureId -evaluateOnly $evaluateOnly
}

write-host
write-host "Enabling desired features..."
$enableFeatures | % {
  $featureId = $_
  EnableFeature -context $context -featureId $featureId -evaluateOnly $evaluateOnly
}

# Step 7. Adjust navigation settings for a publishing web.
if( (IsPublishingWeb $context.Web) ) {
  write-host
  write-host "Adjusting settings of the global navigation..."
  UpdateGlobalNavigation -context $context -evaluateOnly $evaluateOnly
  write-host
  write-host "Adjusting settings of the current navigation..."
  UpdateCurrentNavigation -context $context -evaluateOnly $evaluateOnly
}

# Step 8. Add custom nodes to the top navigation.
# The next command may help to remove all provisioned custom navigation nodes, 
# for example, if some were mistekenly added: RemoveAllTopNavigationNodes $context
# Compliance Policy Centar cannot have top navigation nodes (exemption?)
if( $hasWebTemplate -and !($context.Site.RootWeb.WebTemplate -imatch $webTemplatesWithNoTopNavigation) ) {
  write-host
  write-host "Adding custom nodes to the top navigation..."

  $lcid = [int]$context.Site.RootWeb.Language
  $tmp = $NavigationNodes.($lcid)
  $nodes = $null
  if( $tmp ) {
    $nodes = $tmp.GetEnumerator()
  } else {
    $nodes = $NavigationNodes.($NavigationNodes.Default).GetEnumerator()
  }
  $nodes | sort -Property {$_.Value.Order} | % {
    EnsureTopNavigationNode -context $context -title $_.Name -url $_.Value.Url `
      -isExternal $_.Value.IsExternal -evaluateOnly $evaluateOnly
  }
}
write-host
write-host "Done"
write-host
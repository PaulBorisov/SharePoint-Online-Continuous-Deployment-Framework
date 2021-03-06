param(
  [string]$siteCollectionUrl = "https://<your-tenant>.sharepoint.com"                 # URL of either SharePoint Online or "on-premises" site collection
  ,[string]$username = "<spo-admin-or-global-admin>@<your-tenant>.onmicrosoft.com"    # Login name of either SharePoint Online account or "on-premises" Windows user in format DOMAIN\account
  ,[string]$password = "<password>"                                                   # Password of either SharePoint Online account or "on-premises" Windows user
  ,[bool]$useLocalEnvironment = $false    # Set to $true if you intend to use local SharePoint environment. This can be usefull in case of "on-premises" site collections (not in the SharePoint Online).
  ,[bool]$useDefaultCredentials = $false  # Set to $true if you intend to use default network credentials instead of username and password
  ,[string[]]$pathsToCsomDlls = @(
     "$PSScriptRoot\Microsoft.SharePoint.Client*.dll"
     ,"$PSScriptRoot\Microsoft.Online.SharePoint.Client*.dll"
     #"C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client*.dll"
     #,"C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.Client*.dll"
  )
  ,[bool]$initContextOnLoad = $true
)
######################################################## FUNCTIONS #########################################################  
function CreateWebRequest(){
param (
  [Parameter(Mandatory=$true)][string]$requestUrl
)
  $cookieContainer = $null
  if( !$useLocalEnvironment -and !$useDefaultNetworkCredentials ) {
    $cookieContainer = GetAuthenticationCookie
    if( $cookieContainer -eq $null ) {
      return
    }
  }

  $request = $context.WebRequestExecutorFactory.CreateWebRequestExecutor($context, $requestUrl).WebRequest
  $request.Method = "GET"
  $request.Accept = "text/html, application/xhtml+xml, */*"
  if( $cookieContainer -ne $null )  {
    $request.CookieContainer = $cookieContainer
  } elseif ( $useDefaultNetworkCredentials ) {
    $request.UseDefaultCredentials = $true
  } elseif( $useLocalEnvironment ) {
    $request.Credentials = $context.Credentials
  }
  $request.UserAgent = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
  $request.Headers["Cache-Control"] = "no-cache"
  $request.Headers["Accept-Encoding"] = "gzip, deflate"
  return $request
}

function EnsureFolders(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Folder]$parentFolder,
  [Parameter(Mandatory=$true)][string]$webRelativeFolderUrl,
  [bool]$evaluateOnly
)
  $allFolders = $webRelativeFolderUrl.Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)
  if( $allFolders.Length -eq 0 ) {
    return
  }
  $folderName = $allFolders[0]
  $folder = $null
  if( !$evaluateOnly ) {
    $folder = $parentFolder.Folders.Add($folderName)
    $context.Load($folder)
    $context.ExecuteQuery()
    if( $context -eq $null ) {return}
  }
  
  if( $allFolders.Length -gt 1 ) {
    $allFolders = [string]::Join('/', $allFolders, 1, ($allFolders.Length - 1))
    EnsureFolders -context $context -parentFolder $folder -webRelativeFolderUrl $allFolders -evaluateOnly $evaluateOnly
  }
}

function ExecuteAndRetryIfFailed(){
param(
  [Parameter(Mandatory=$true)][ScriptBlock]$scriptBlock
  ,[object[]]$scriptBlockParams
  ,[int]$amountOfRetryAttempts = 5
  ,[int]$waitIntervalSeconds = 3
)
  for( $i = 0; $i -lt $amountOfRetryAttempts; $i++ ) {
    try {
      Invoke-Command $scriptBlock -ArgumentList $scriptBlockParams
    } catch {
      # Wait and retry once again if failed
      Start-Sleep -s $waitIntervalSeconds
      if( ($i + 1) -ge $amountOfRetryAttempts ) {
        write-host $scriptBlock -ForegroundColor Red
        throw
      }
    }
  }
}

function ExecuteSearchQuery(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$queryText,
  [string[]]$sorting = @("LastModifiedTime Desc"),
  [Hashtable]$queryParameters = @{
    StartRow = 0;
    RowLimit = 500;
    ProcessBestBets = $false;
    BypassResultTypes = $true;
    EnableInterleaving = $false;
    EnableQueryRules = $false;
    EnableStemming = $true;
    TrimDuplicates = $false
  }
)
  
  $keywordQuery = New-Object Microsoft.SharePoint.Client.Search.Query.KeywordQuery($context)
  $keywordQuery.QueryText = $queryText
  if( $queryParameters ) {
    $queryParameters.GetEnumerator() | % {
      $keywordQuery.($_.Name) = $($_.Value)
    }
  }
  
  if( $sorting.Count -gt 0 ) {
    $sorting | % {
      $criteria = $_
      $descensing = $false
      $lastSpaceIndex = $_.LastIndexOf(" ")
      if( $lastSpaceIndex -gt -1 ) {
        $descensing = [regex]::IsMatch($criteria.Substring($lastSpaceIndex + 1), "(?i)^desc")
        $criteria = $criteria.Substring(0, $lastSpaceIndex)
      }
      if( $descensing ) {
        $keywordQuery.SortList.Add($criteria, "Descending")
      } else {
        $keywordQuery.SortList.Add($criteria, "Ascending")
      }
    }
  }
  
  $searchExecutor = new-object Microsoft.SharePoint.Client.Search.Query.SearchExecutor($context)
  $results = $searchExecutor.ExecuteQuery($keywordQuery)
  $context.ExecuteQuery()

  $foundItems = @()
  $results.Value | ? {$_.TableType -eq [Microsoft.SharePoint.Client.Search.Query.KnownTableTypes]::RelevantResults} | % {
    $_.ResultRows | % {
      $row = $_
      $singleItem = @{}
      $_.Keys | % {
        $singleItem[$_] = $row[$_]
      }
      if( $? ) {
        $foundItems += $singleItem
      }
    }
  }
  return $foundItems
}

function ExecuteWebRequest(){
param (
  [Parameter(Mandatory=$true)][System.Net.HttpWebRequest]$request,
  [Hashtable]$requestData,
  [ref]$errorMessage,
  [ref]$successMessage
)
  if( $requestData -ne $null -and $requestData.Keys -ne $null ) {
    # Format inputs as postback data string
    $strData = $null
    foreach( $inputKey in $requestData.Keys )  {
      if( -not([String]::IsNullOrEmpty($inputKey)) ) {
        $strData += [System.Web.HttpUtility]::UrlEncode($inputKey) `
          + "=" + [System.Web.HttpUtility]::UrlEncode($requestData[$inputKey]) + "&"
      }
    }
    $strData = $strData.TrimEnd('&')
    $requestDataBytes = [System.Text.Encoding]::UTF8.GetBytes($strData)
    $request.ContentLength = $requestDataBytes.Length
    
    # Add postback data to the request stream
    $stream = $request.GetRequestStream()
    try {
      $stream.Write($requestDataBytes, 0, $requestDataBytes.Length)
    } finally {
      $stream.Close()
      $stream.Dispose()
    }
  } else {
    $request.ContentLength = 0
  }
     
  $response = $request.GetResponse()
  $stream = $response.GetResponseStream()
  try {
    if( -not([string]::IsNullOrEmpty($response.Headers["Content-Encoding"])) ) {
      if( $response.Headers["Content-Encoding"].ToLower().Contains("gzip") ) {
        $stream = New-Object System.IO.Compression.GZipStream($stream, 
          [System.IO.Compression.CompressionMode]::Decompress)
      } elseif ( $response.Headers["Content-Encoding"].ToLower().Contains("deflate") ) {
        $stream = New-Object System.IO.Compression.DeflateStream($stream, 
          [System.IO.Compression.CompressionMode]::Decompress)
      }
    }
         
    # Retrieve response content as string
    $reader = New-Object System.IO.StreamReader($stream)
    $content = $reader.ReadToEnd()
    #$content > "$PSScriptRoot\request-content.html"
    if( $successMessage.Value -ne $null ) {
      $match = [regex]::Match($content, $successMessage.Value, "Singleline")
      $successMessage.Value = $null
      if( $match.Success ) {
        for( $i = 1; $i -lt $match.Groups.Count; $i++ ) {
          $tmp = $match.Groups[$i].Value
          if( $tmp.Length -gt 1 ) {
            $tmp = $tmp.Substring(0,1).ToUpper() + $tmp.Substring(1)
          }
          $successMessage.Value += $tmp
        }
      }     
    }
    $match = $null
    if( $errorMessage.Value -ne $null ) {
      $match = [regex]::Match($content, $errorMessage.Value, "Singleline")
    } else {
      $match = [regex]::Match($content, 
        '(?i)<div id="ms-error">.+<span id="ctl\d+_PlaceHolderMain_LabelMessage">(.[^<]+)<\/span>', "Singleline")
    }
    $errorMessage.Value = $null
    if( $match.Success ) {
      for( $i = 1; $i -lt $match.Groups.Count; $i++ ) {
        $tmp = $match.Groups[$i].Value
        if( $tmp.Length -gt 1 ) {
          $tmp = $tmp.Substring(0,1).ToUpper() + $tmp.Substring(1)
        }
        $errorMessage.Value += $tmp
      }
    }
    $reader.Close()
    $reader.Dispose()
    return $content
  } finally {        
    $stream.Close()
    $stream.Dispose()
  }
  return $null
}

function GetAdminClientContext() {
  $uri = new-object System.Uri($siteCollectionUrl)
  $adminUrl = $null
  if( $useLocalEnvironment ) {
    # Just the top root site collection
    $adminUrl = [string]::Format("{0}://{1}", $uri.Scheme, $uri.Authority)
  } else {
    $fistDashIndex = $uri.Authority.IndexOf('-')
    $fistDotIndex = $uri.Authority.IndexOf('.')
    if( $fistDotIndex -gt -1 ) {
      # SPO specific tenant admin site collection
      if( $fistDashIndex -gt -1 -and $fistDashIndex -lt $fistDotIndex ) {
        $adminUrl = [string]::Format("{0}://{1}", $uri.Scheme, $uri.Authority.Insert(
          $fistDotIndex, "-admin").Remove($fistDashIndex, $fistDotIndex - $fistDashIndex))
      } else {
        $adminUrl = [string]::Format("{0}://{1}", $uri.Scheme, $uri.Authority.Insert($fistDotIndex, "-admin"))
      }
    } else {
      throw "Invalid URL. Site collection $siteCollectionUrl does not belong to SharePoint Online"
    }
  }

  $adminContext = New-Object Microsoft.SharePoint.Client.ClientContext($adminUrl)
  if( $useDefaultNetworkCredentials ) {
    $adminContext.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  } elseif( $useLocalEnvironment ) {
    #$domain = $null
    #$login = $null
    #if( $username.Contains("@") ) {
    #  $domain = $username.Substring($username.IndexOf('@') + 1)
    #  $login = $username.Substring(0, $username.IndexOf('@'))
    #  $adminContext.Credentials = (New-Object System.Net.NetworkCredential($login, $password, $domain))
    #} elseif( $username.Contains("\") ) {
    #  $domain = $username.Substring(0, $username.IndexOf('\'))
    #  $login = $username.Substring($username.IndexOf('\') + 1)
    #  $adminContext.Credentials = (New-Object System.Net.NetworkCredential($login, $password, $domain))
    #} else {
    $adminContext.Credentials = (New-Object System.Net.NetworkCredential($username, $password))
    #}
  } else {
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword)
    $adminContext.Credentials = $credentials
  }
  $adminContext.Load($adminContext.Site)
  $adminContext.Load($adminContext.Site.RootWeb)
  $adminContext.Load($adminContext.Web)
  $adminContext.ExecuteQuery()
  if( $? -eq $false ) {
    return $null
  } else {
    return $adminContext
  }
}

function GetAuthenticationCookie() {
  $sharePointUri = New-Object System.Uri($context.Url)
  $authCookie = $context.Credentials.GetAuthenticationCookie($sharePointUri)
  if( $? -eq $false ) {
    return $null
  } else {
    $fedAuthString = $authCookie.TrimStart("SPOIDCRL=".ToCharArray())
    $cookieContainer = New-Object System.Net.CookieContainer
    $cookieContainer.Add($sharePointUri, (New-Object System.Net.Cookie("SPOIDCRL", $fedAuthString)))
    return $cookieContainer
  }
}

function GetClientContext() {
  $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteCollectionUrl)
  if( $useDefaultNetworkCredentials ) {
    $context.Credentials = [System.Net.CredentialCache]::DefaultCredentials
  } elseif( $useLocalEnvironment ) {
    #$domain = $null
    #$login = $null
    #if( $username.Contains("@") ) {
    #  $domain = $username.Substring($username.IndexOf('@') + 1)
    #  $login = $username.Substring(0, $username.IndexOf('@'))
    #  $context.Credentials = (New-Object System.Net.NetworkCredential($login, $password, $domain))
    #} elseif( $username.Contains("\") ) {
    #  $domain = $username.Substring(0, $username.IndexOf('\'))
    #  $login = $username.Substring($username.IndexOf('\') + 1)
    #  $context.Credentials = (New-Object System.Net.NetworkCredential($login, $password, $domain))
    #} else {
    $context.Credentials = (New-Object System.Net.NetworkCredential($username, $password))
    #}
  } else {
    $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $securePassword)
    $context.Credentials = $credentials
  }
  $context.Load($context.Site)
  $context.Load($context.Site.RootWeb)
  $context.Load($context.Web)
  $context.ExecuteQuery()
  if( $? -eq $false ) {
    return $null
  } else {
    return $context
  }
}

function GetErrorMessage() {
  if( $_.Exception ) {
    if( $_.Exception.InnerException ) {
      return $_.Exception.InnerException.ToString()
    } elseif ( $_.Exception.Message ) {
      return $_.Exception.Message.Substring($_.Exception.Message.IndexOf(":") + 1).Trim().TrimStart('"').TrimEnd('"')
    } else {
      return $_.Exception.ToString()
    }
  } else {
    return [string]::Empty
  }
}

function GetListByFolderUrl(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Web]$web,
  [Parameter(Mandatory=$true)][string]$webRelativeFolderUrl,
  [int]$skipFoldersCount = 1
)
  $urlParts = $webRelativeFolderUrl.TrimStart('/').TrimEnd('/').Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)
  $url = $web.ServerRelativeUrl.TrimEnd('/')
  $web.Context.Load($web.Lists)
  $web.Context.ExecuteQuery()
  if( $context -eq $null ) {return $null}
  
  $web.Lists | % {
    $web.Context.Load($_.RootFolder)
    $web.Context.ExecuteQuery()
  }
  
  $i = 0
  do {
    $url += '/' + $urlParts[$i]
    $i++
    if( $i -le $skipFoldersCount ) { # For example, skip "_catalog"
      continue
    }
    $lists = $web.Lists | ? {$_.RootFolder.ServerRelativeUrl -ieq $url}
    if( $lists.Length -gt 0 ) {
      return $lists[0]
    } 
  } while( $i -lt $urlParts.Length )
  
  return $null
}

function IsAbsoluteUrl(){
param (
  [string]$url
)
  $result = $null
  return [System.Uri]::TryCreate($url, "Absolute", [ref]$result)
}

## http://sharepoint.stackexchange.com/questions/126221/spo-retrieve-hasuniqueroleassignements-property-using-powershell
function InvokeLoadMethod(){
param(
   [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientObject]$clientObject,
   [Parameter(Mandatory=$true)][string]$propertyName
) 
  $ctx = $clientObject.Context
  $load = [Microsoft.SharePoint.Client.ClientContext].GetMethod("Load")
  $type = $clientObject.GetType()
  $clientLoad = $load.MakeGenericMethod($type)

  $parameter = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
  $expression = [System.Linq.Expressions.Expression]::Lambda(
  [System.Linq.Expressions.Expression]::Convert(
    [System.Linq.Expressions.Expression]::PropertyOrField($parameter,$propertyName),
    [System.Object]
  ), $($parameter)
  )
  $expressionArray = [System.Array]::CreateInstance($expression.GetType(), 1)
  $expressionArray.SetValue($expression, 0)
  $clientLoad.Invoke($ctx,@($clientObject,$expressionArray))
}

function IsPublishingWeb(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Web]$web
)
  $web.Context.Load($web.Features)
  $web.Context.ExecuteQuery()
  return ($web.Features | ? {$_.DefinitionId -eq [guid]("94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb")}).Length -gt 0
}

function ProcessFilesInDocumentLibrary(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Web]$web,
  [Parameter(Mandatory=$true)][string]$webRelativeUrlListRootFolder,
  [Parameter(Mandatory=$true)][ScriptBlock]$scriptBlock,
  [object[]]$scriptBlockParams,
  [bool]$evaluateOnly
)
  # Obtain a parent list.
  $list = GetListByFolderUrl $web $webRelativeUrlListRootFolder
  if( $list -eq $null ) {
    throw "List not found. Further processing is impossible."
  }

  # Temporarily change list's settings to suppress unnecessary multiple check-out, check-in, publish and approval operations.
  $EnableModeration = $list.EnableModeration
  $EnableVersioning = $list.EnableVersioning
  $EnableMinorVersions = $list.EnableMinorVersions
  $DraftVersionVisibility = $list.DraftVersionVisibility
  $ForceCheckout = $list.ForceCheckout

  if( !$evaluateOnly ) {
    $list.EnableModeration = $false
    $list.ForceCheckout = $false
    $list.EnableVersioning = $true
    $list.EnableMinorVersions = $false
    $list.Update()
    $web.Context.ExecuteQuery()
    if( $? -eq $false ) {
      if( $web.ServerRelativeUrl -eq "/" ) {
        throw "The top root site of a tenant may require specific additional permissions." `
          + " Use the following command to grant them: Set-SPOsite $($web.Url) -DenyAddAndCustomizePages 0"
      } else {
        throw "Further processing stopped due to the above exception."
      }
      return
    }
  }
  
  try {
    Invoke-Command $scriptBlock -ArgumentList $scriptBlockParams
  } finally {
    if( !$evaluateOnly ) {
      # Restore the initial list's settings.
      $list.EnableModeration = $EnableModeration
      $list.ForceCheckout = $ForceCheckout
      $list.EnableVersioning = $EnableVersioning
      $list.EnableMinorVersions = $EnableMinorVersions
      $list.DraftVersionVisibility = $DraftVersionVisibility
      $list.Update()
      $web.Context.ExecuteQuery()
    }
  }
}

function UploadFile(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.Web]$web,
  [Parameter(Mandatory=$true)][string]$filePath,
  [Parameter(Mandatory=$true)][string]$webRelativeUrl,
  [bool]$evaluateOnly
)
  $folderUrl = $web.ServerRelativeUrl.TrimEnd('/') + '/' + $webRelativeUrl.TrimStart('/')
  $folder = $web.GetFolderByServerRelativeUrl($folderUrl)
  $web.Context.Load($folder)
  $web.Context.ExecuteQuery()
  if( $context -eq $null ) {return}
  
  if( !$evaluateOnly ) {
    $fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $fileInfo.Content = [System.IO.File]::ReadAllBytes($filePath)
    $fileInfo.Url = $filePath.Substring($filePath.LastIndexOf('\') + 1)
    $fileInfo.Overwrite = $true
     
    $uploadedFile = $folder.Files.Add($fileInfo)
    $web.Context.Load($uploadedFile)
    $web.Context.ExecuteQuery()
  }
}
####################################################### //FUNCTIONS ########################################################
   
####################################################### EXECUTION ##########################################################
# Step 1. Check major version of Powershell. It must be 3 or higher to work with SharePoint Client Context.
if( $PSVersionTable.PSVersion.Major -lt 3 ) {
  throw "This script requires Powershell version 3 or higher to run." + `
    "You can download Powershell 3 from https://www.microsoft.com/en-us/download/details.aspx?id=34595"
}

# Step 2. Ensure the client components loaded
$hasCsomLoadErrors = $false
$pathsToCsomDlls | % {
  $lastBackSlashIndex = $_.LastIndexOf('\')
  Get-ChildItem -Path $($_.Substring(0, $lastBackSlashIndex + 1)) `
    -Include $($_.Substring($lastBackSlashIndex + 1)) -Exclude "" -Recurse | % {
    # Usually, this is sufficient to load the following CSOM-dlls:
    # - Microsoft.SharePoint.Client.dll
    # - Microsoft.SharePoint.Client.Runtime.dll
    # - Microsoft.SharePoint.Client.Taxonomy.dll
    # - Microsoft.SharePoint.Client.Publishing.dll
    # - Microsoft.SharePoint.Client.Search.dll
    # - Microsoft.Online.SharePoint.Client.Tenant.dll # This is used for the portal management
    [Reflection.Assembly]::LoadFile($_) | out-null
    if( $? -eq $false ) {
      $hasCsomLoadErrors = $true
    }
  }
}

if( $initContextOnLoad -and !$hasCsomLoadErrors ) {
  $context = GetClientContext
} else {
  $context = $null
}

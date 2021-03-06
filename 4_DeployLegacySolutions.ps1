param(
  # The default value is $null. In this case site collection specified in __LoadContext.ps1 will be processed.
  [string]$siteCollectionUrl = $null
  
  # Full path to the folder, which contains sandbox wsp-solutions to be (re)deployed
  # If you need to activate solutions in a correct order use correspondent prefix in file names, 
  # for example, 1_firstSolution.wsp, 2_anotherSolution.wsp, 3_extraSolution.wsp
  ,[string]$pathToFolderWIthWspSolutions = "$PSScriptRoot\legacy\sandbox-wsps"
  
  ,[string]$regexDeploymentUrls = "(?i)/sites/" # Set to "." if WSPs should be deployed on all site collections
  ,[bool]$disableCustomizations = $false
  ,[bool]$force = $false
  ,[string[]]$dismissAdditionalSolutions = @(
     # List of optional extra solutions that need to be deactivated before
     # activating the ones that present in the folder $pathToFolderWIthWspSolutions.
     # This is useful when old and new solutions may conflict with each other.
   )
)
######################################################## FUNCTIONS #########################################################
# This function allows establishing more refined rules for deployment of legacy WSP sandbox solutions.
function AllowDeployment(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [string]$regexDeploymentUrls
)
  $isPermitted = $false
  $url = $context.Site.RootWeb.ServerRelativeUrl
  # Step 1. Check if URL is matching the permitted ones.
  if( $regexDeploymentUrls ) {
    if( ![regex]::IsMatch($url, $regexDeploymentUrls) ) {
      $isPermitted = $false
    } else {
      $isPermitted = $true
    }
  }
  # Step 2. Check if the site is matching very specific requirements.
  # In this particular example site's URL must end with an even number to allow deployment of legacy WSPs.
  $match = [regex]::Match($url, "ts(\d+)$")
  $number = $null
  if( $match.Success -and [int]::TryParse($match.Groups[1].Value, [ref]$number) ) {
    if( $number % 2 -eq 0 ) {
      $isPermitted = $true
    } else {
      $isPermitted = $false
    }
  }
  
  if( !$isPermitted ) {
    write-host " WSPs are not required on this site collection; skipped."
  }
  return $isPermitted
}

function ActivateOrDeactivateSolution(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][int]$solutionId,
  [Parameter(Mandatory=$true)][bool]$activate,
  [ref]$errorMessage
)
  $op = $null
  if( $activate ) { 
    $op = "ACT"
  } else { 
    $op = "DEA"
  }
     
  $activationUrl = $context.Site.Url + "/_catalogs/solutions/forms/activate.aspx?Op=$op&ID=$solutionId"  
  $request = CreateWebRequest -requestUrl $activationUrl
  $content = ExecuteWebRequest -request $request -errorMessage ([ref]$errorMessage)
   
  $inputMatches = $content | Select-String -AllMatches -Pattern "<input.+?\/??>" | select -Expand Matches
  $inputs = @{}
       
  # Iterate through inputs and add specific values to the dictionary for postback
  foreach( $match in $inputMatches ) {
    if( -not($match[0] -imatch "name=\""(.+?)\""") ) {
      continue
    }
    $name = $matches[1]
    if( $name.EndsWith("iidIOGoBack") ) {
      continue
    }
    
    if( -not($match[0] -imatch "value=\""(.+?)\""") ) {
      continue
    }
    $value = $matches[1]
   
    $inputs.Add($name, $value)
  }
   
  # Search for the id of the button "activate"
  $searchString = $null
  if ($activate) {
    $searchString = "ActivateSolutionItem"
  } else {
    $searchString = "DeactivateSolutionItem"
  }
           
  $match = $content -imatch "__doPostBack\(\&\#39\;(.*?$searchString)\&\#39\;"
  $inputs.Add("__EVENTTARGET", $matches[1])

  # Build postback request
  $activateRequest = CreateWebRequest -requestUrl $activationUrl
  $activateRequest.CookieContainer = GetAuthenticationCookie
  $activateRequest.Method = "POST"
  $activateRequest.ContentType = "application/x-www-form-urlencoded"
  $content = ExecuteWebRequest -request $activateRequest -requestData $inputs -errorMessage ([ref]$errorMessage)
}
   
function GetSolutionFileWithListItem(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$solutionName
)
  
  $solutionFileUrl = $context.Site.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/solutions/" + $solutionName
  $solutionFile = $context.Site.RootWeb.GetFileByServerRelativeUrl($solutionFileUrl)
  $context.Load($solutionFile.ListItemAllFields)
  $context.ExecuteQuery()
  return $solutionFile
}
  
function GetSolutionId(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$solutionName
)
  $solutionFile = GetSolutionFileWithListItem $context $solutionName
  return $solutionFile.ListItemAllFields.Id
}
  
function GetSolutionStatus(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$solutionName
)
  $solutionFile = GetSolutionFileWithListItem $context $solutionName
  return $solutionFile.ListItemAllFields["Status"]
}
<#   
function UploadSolution(){
param(
  [Parameter(Mandatory=$true)][Microsoft.SharePoint.Client.ClientContext]$context,
  [Parameter(Mandatory=$true)][string]$filePath
)
  $fileInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
  $fileInfo.Content = [System.IO.File]::ReadAllBytes($filePath)
  $fileInfo.Url = $filePath.Substring($filePath.LastIndexOf('\') + 1)
  $fileInfo.Overwrite = $true
   
  $folderUrl = $context.Site.Url + "/_catalogs/solutions"
  $folderUri = New-Object System.Uri($folderUrl)
   
  $solutionsFolder = $context.Site.RootWeb.GetFolderByServerRelativeUrl($folderUri.AbsolutePath)
  $uploadedFile = $solutionsFolder.Files.Add($fileInfo)
  $context.Load($uploadedFile)
  $context.ExecuteQuery()
}
#>
####################################################### //FUNCTIONS ########################################################
   
####################################################### EXECUTION ##########################################################
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

if( !(AllowDeployment -context $context -regexDeploymentUrls $regexDeploymentUrls) ) {
  return
}

# Step 2. Check optional (possibly existing) solutions.
$fileSystemSolutions = Get-ChildItem -Path $pathToFolderWIthWspSolutions -Filter "*.wsp" | ? {!$_.PSIsContainer}

if( $dismissAdditionalSolutions -ne $null -and $dismissAdditionalSolutions.Length -gt 0 ) {
  # Exclude matching solutions that may exist in the file system.
  $tmp1 = @(); $fileSystemSolutions | % {$tmp1 += $_.Name}
  $tmp2 = @(); $dismissAdditionalSolutions | % {$tmp2 += $_}
  $dismissSolutions = Compare-Object -ReferenceObject $tmp1 -DifferenceObject $tmp2 -PassThru | ? {$_.SideIndicator -eq "=>"}

  $dismissSolutions | % {
    $extraSolutionName = $_
    $extraSolutionId = GetSolutionId $context $extraSolutionName
    if( $extraSolutionId -is [int] ) {
      $message = "Deactivating an old"
      $message2 = " optional "
      $message3 = "solution $extraSolutionName..."
      write-host
      write-host $message -NoNewLine
      write-host $message2 -NoNewLine -ForegroundColor Yellow
      write-host $message3
      try {
        $errorMessage = $null
        ActivateOrDeactivateSolution -context $context -solutionId $extraSolutionId `
          -activate $false ([ref]$errorMessage)
      } catch {
        #The action can fail if the solution has already been deactivated earlier.
      }
    }
  }
}

# Step 3. Deactivate existing solutions, (re)deploy, and (re)activate new versions of solutions
Get-ChildItem -Path $pathToFolderWIthWspSolutions -Filter "*.wsp" | ? {!$_.PSIsContainer} | % {
  $filePath = $_.FullName
  $solutionName = $_.Name
  $solutionId = GetSolutionId $context $solutionName
  $alreadyCustomized = $false
  if( $solutionId -is [int] ) {
    if( $disableCustomizations -or $force ) {
      $message = "Deactivating an old solution $solutionName..."
      write-host
      write-host $message -NoNewLine
      try {
        $errorMessage = $null
        ActivateOrDeactivateSolution -context $context -solutionId $solutionId `
          -activate $false ([ref]$errorMessage)
      } catch {
        #The action can fail if the solution has already been deactivated earlier.
      }
    } else {
      write-host
      write-host "The solution $solutionName already exists on this site."
      write-host "Use the parameters -disableCustomizations or -force to process it."
      $alreadyCustomized = $true
    }
  }
   
  if( !$disableCustomizations -and !$alreadyCustomized ) {
    $message = "Uploading a new solution $solutionName..."
    write-host
    write-host $message -NoNewLine
    #UploadSolution -context $context -filePath $filePath
    UploadFile -web $context.Site.RootWeb -filePath $filePath -webRelativeUrl "_catalogs/solutions"
    # In case of an uploade error; no need for any further actions on this solution. Skip and continue to the next one.
    if( $? -eq $false ) {return}
     
    # Refresh client context
    $context = GetClientContext
     
    $solutionId = GetSolutionId -context $context -solutionName $solutionName
    if( $solutionId -is [int] ) {
      $message = "Activating a new solution $solutionName..."
      write-host
      write-host $message -NoNewLine
   
      $errorMessage = $null
      ActivateOrDeactivateSolution -context $context -solutionId $solutionId -activate $true -errorMessage ([ref]$errorMessage)
   
      $status = GetSolutionStatus -context $context -solutionName $solutionName
      if( $status -eq $null ) {
        write-host "Not activated" -ForegroundColor Red
      } elseif( $status.LookupValue -eq 1) {
        write-host "Activated" -ForegroundColor Green
      } else {
        write-host "Status is undefined. Check in the library /_catalogs/solutions." -ForegroundColor Yellow
      }
      if( $errorMessage -ne $null ) {
        write-host $errorMessage -ForegroundColor Yellow
        $errorMessage = $null
      }
    }
  }
}
write-host

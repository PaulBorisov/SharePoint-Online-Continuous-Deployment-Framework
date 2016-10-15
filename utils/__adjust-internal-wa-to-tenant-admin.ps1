<#
Inspired by https://github.com/OfficeDev/PnP/blob/master/Samples/Provisioning.SiteCol.OnPrem/EnableOnPremSiteCol.ps1
Enable the remote site collection creation for on-prem in web application level.
If this is not done, unknown object exception is raised by the CSOM code on attempts to use Tenant's object.
#>
. "$PSScriptRoot\__LoadContext.ps1" -initContextOnLoad:$false
$uri = new-object System.Uri($siteCollectionUrl)
$webApplicationUrl = [string]::Format("{0}://{1}", $uri.Scheme, $uri.Authority)
write-host
write-host "Web application: $webApplicationUrl"

$message = "Adjusting settings on the web-application " + $webApplicationUrl
write-host $message

# Load SharePoint Snapin
$snapin = get-pssnapin | ? {$_.Name -ieq 'Microsoft.SharePoint.Powershell'}
if ($snapin -eq $null) {
	write-host "Loading SharePoint Powershell Snapin..." -NoNewLine
	add-pssnapin "Microsoft.SharePoint.Powershell" | out-null
  write-host "done"
}	

$wa = Get-SPWebApplication $webApplicationUrl
$rootSite = $null
$wa.Sites | % {
  if( $_.Url -ieq $wa.Url.TrimEnd('/') ) {
    $rootSite = $_
  } else {
    $_.Dispose()
  }
}
if( $rootSite -eq $null ) {
  write-host "Creating a root site to support tenant administration..." -NoNewLine
  $rootSite = New-SPSite -URL $wa.Url -OwnerAlias $userName -Language 1033 -CompatibilityLevel 15
  if($? -eq $false) {return} # Some error

  write-host "done, $($rootSite.Url)"
  write-host
}

$newProxyLibrary = New-Object "Microsoft.SharePoint.Administration.SPClientCallableProxyLibrary"
$newProxyLibrary.AssemblyName = "Microsoft.Online.SharePoint.Dedicated.TenantAdmin.ServerStub, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"
$existingAssemblies = ($wa.ClientCallableSettings.ProxyLibraries | ? {$_.AssemblyName -ieq $newProxyLibrary.AssemblyName})
if( $existingAssemblies.Length -eq 0 ) {
  $newProxyLibrary.SupportAppAuthentication = $true
  $wa.ClientCallableSettings.ProxyLibraries.Add($newProxyLibrary)
  $wa.ClientCallableSettings.ExecutionTimeout = [System.TimeSpan]::FromMinutes(3.0)
  $wa.SelfServiceSiteCreationEnabled = $true
  $wa.Update()
  
  write-host
  write-host "Successfully added TenantAdmin ServerStub to ClientCallableProxyLibrary."
  # Reset the memory of the web application
  write-host "IISReset..."    
  restart-service W3SVC,WAS -force
  write-host "IISReset complete on this server, remember other servers in farm as well."    
  
} elseif( $existingAssemblies.Length -gt 1 ) {
  # Remove duplicates (if any).
  while( $existingAssemblies.Count -gt 1 ) {
    $wa.ClientCallableSettings.ProxyLibraries.RemoveAt($existingAssemblies.Count-1)
    $wa.Update()
    $existingAssemblies = ($wa.ClientCallableSettings.ProxyLibraries | ? {$_.AssemblyName -ieq $newProxyLibrary.AssemblyName})
  }
}

# Ensure admin site type property for the root site.
# Note this property needs to be set for the site collection which is used as the "Connection point"
# (Tenant admin) for the CSOM to be able to create site collections in this "on-premises" web-application.
write-host
write-host "Adjusting the root site to become a connection point (Tenant Admin)..." -NoNewLine
$rootSite.AdministrationSiteType = [Microsoft.SharePoint.SPAdministrationSiteType]::TenantAdministration
write-host "done"

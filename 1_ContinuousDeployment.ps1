param(
  [int]$secondsToRepeat = 30                                             # Interval between iterations, in seconds
  ,[int]$maxIterations = 0                                               # 0 or negative means infinite execution.
  ,[int]$maxFailedAttempts = 5                                           # 0 or negative means infinite execution.
  ,[string]$filePathSiteStates = "$PSScriptRoot\logs\__site-states.csv"  # Path to the file that contains cached statuses.
  ,[bool]$readCacheFileOnEveryIteration = $true                          # Forces to refresh statuses from the cached file.
)

######################################################## FUNCTIONS #########################################################
function StartProcessUnlessActive() {
param(
  [string]$name = "powershell.exe"
  ,[string]$argumentList
  ,[string]$identifier
  ,[Hashtable]$activeProcesses
)

  $id = $activeProcesses[$identifier]
  if( $id ) {
    $isActive = (get-process | ? {$_.Id -eq $id}).Length -gt 0
    if( !$isActive ) {
      $activeProcesses[$identifier] = $null
      $id = $null
    }
  }

  if( !$id ) {
    $started = start-process $name -ArgumentList $argumentList -PassThru
    $activeProcesses[$identifier] = $started.Id
  }
}

function ReadCacheFile(){
param(
  [Parameter(Mandatory=$true)][string]$filePathSiteStates
)
  $processedSites = @{}
  if( test-path $filePathSiteStates ) {
    $lineNumber = 0
    [System.IO.File]::ReadAllLines($filePathSiteStates) | % {
      $lineNumber++
      #if( $lineNumber -eq 1 ) {
      #  return
      #}
      $line = $_.Trim()
      if( $line.Length -eq 0 -or !$line.StartsWith("http", "InvariantCultureIgnoreCase") ) {return}

      $parts = $line.Split("`t")
      if( $parts.Length -lt 4 ) {return}
      
      try {
        $processedSites[$parts[0]] = @{
          LastProcessed = [datetime]::Parse($parts[1]);
          Succeeded = [bool]::Parse($parts[2]);
          Customized = [bool]::Parse($parts[3]);
          FailedAttempts = [int]::Parse($parts[4])
        }
      } catch {
        write-host
        write-host "$($filePathSiteStates): line $lineNumber is invalid and ignored." -ForegroundColor Red
      }
    }
  }
  return $processedSites
}
####################################################### //FUNCTIONS ########################################################

####################################################### EXECUTION ##########################################################
$activeProcesses = @{}
$processedSites = ReadCacheFile -filePathSiteStates $filePathSiteStates
$iteration = 1
$shouldRun = $true

while( $shouldRun ) {
  write-host
  write-host "Start of iteration $iteration"
  
  $processedSites = & "$PSScriptRoot\2_ProcessAllSiteCollections.ps1" -processedSites $processedSites `
    -maxFailedAttempts $maxFailedAttempts
  $processedSitesArray = @()
  if( $processedSites -ne $null -and $processedSites.Length -gt 0 ) {
    $processedSites.GetEnumerator() | % {
      $processedSitesArray += (
        $_.Name + "`t" + `
        $_.Value.LastProcessed.ToString("yyyy-MM-dd HH:mm") + "`t" + `
        $_.Value.Succeeded.ToString().ToLower() + "`t" + `
        $_.Value.Customized.ToString().ToLower() + "`t" + `
        $_.Value.FailedAttempts.ToString().ToLower()
      )
    }
    if( $processedSitesArray.Length -gt 0 ) {
       # The first line contains headers for convenience of reading
      "Url`tLastProcessed`tSucceeded`tCustomized`tFailedAttempts" > $filePathSiteStates
       # The next lines contains data
      [string]::Join([environment]::NewLine, ($processedSitesArray | sort)) >> $filePathSiteStates
    }
  } 

  $iteration++
  if( $maxIterations -gt 0 -and $iteration -gt $maxIterations ) {
    $shouldRun = $false
    write-host
    return
  }
  
  StartProcessUnlessActive -argumentList "$PSScriptRoot\5_CreateRequestedSites.ps1" -identifier CreateSites -activeProcesses $activeProcesses

  write-host
  write-host "Waiting for release of resources..." -NoNewLine
  [console]::CursorVisible=$false
  $j = $secondsToRepeat
  while( $j -gt 0 ) {
    Start-Sleep -s 1
    $j--
    $fgc = $null
    $value = $null
    if( $j -gt 19 ) {
      $fgc = "Green"
      $value = "$j"
    } elseif ( $j -gt 9 ) {
      $fgc = "Yellow"    
      $value = "$j"
    } else {
      $fgc = "White"
      $value = "0$j"
    }
    write-host -NoNewLine $value -ForegroundColor $fgc
    $origpos = $host.UI.RawUI.CursorPosition
    $origpos.X -= 2
    $host.UI.RawUI.CursorPosition = $origpos
  }
  write-host "  "
  [console]::CursorVisible=$true
  
  if( $readCacheFileOnEveryIteration ) {
    $processedSites = ReadCacheFile -filePathSiteStates $filePathSiteStates
  }
}

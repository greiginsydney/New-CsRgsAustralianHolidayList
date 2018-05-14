<#  
.SYNOPSIS  
    This script creates RGS holidaysets for Australian states based on live data from the Australian Government website


.DESCRIPTION  
    Created by James Arber. www.skype4badmin.com
    Although every effort has been made to ensure this list is correct, dates change and sometimes I goof. 
    Please use at your own risk.
    Holiday Data taken from http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays
	    
	
.NOTES  
    Version      	      : 2.2
    Date			          : 11/05/2018
    Lync Version		    : Tested against Skype4B Server 2015 and Lync Server 2013
    Author    			    : James Arber
    Header stolen from  : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

    Revision History	: v2.2: Cleaned Up Code
                      : Fixed a bug with logging system culture
                      : Removed some old redundant code
                      : Passed script through ISESteriods PSSharper and applied corrections
                      : Fixed a few typos
                      : Updated Pat Richard's website
                      : Removed PowerShell 5.1 cmdlet (Get-Timezone), using a WMI query instead
                       

                    : v2.1: Added Script logging
                      : Updated to use my new autoupdate code
                      : Added ability to switch between devel/master branches
                      : Added timezone offset detection / warning
                      : Added SSL support for the new Govt website requirements



                    : v2.01: Migrated to GitHub
                      : Minor Typo corrections
                      : Check for and prompt user for updates
                      : Fixed a bug with multiple pool selction
                      : Fixed issues with double spaced event names
                      : Added better timeout handling to XML downloads
                      : Added better user feedback when downloading XML file
                      : Fixed bug with proxy detection failing to execute
                      : Removed redundant code for XML lookup
                      : Fixed an unattened run bug
                      : Fixed commandline switch descriptions


                    : v2.0: Update for XML Support
                      : Added Autodetecton of single RGS pool
                      : Complete Rewrite of existing rule rewrite code , Should make less red text now.
                      : Added Region detection, Will prompt to change regions or try to use US date format
                      : More user friendly and better instructions
                      : Fixed a few typo's causing dates to be incorrect.
                      : Fixed alot of gramatical errors
                      : Added XML download and implementation with proxy support
                      : Auto removes any dates not listed by the Australian Government (such as old dates) if the $RemoveExistingRules is set
                      : Script no longer deletes existing timeframes, No need to re-assign to workflows!

	
                    : v1.1: Fix for Typo in Victora Holiday set
                      : Fix ForEach loop not correctly removing old time frames
                      : Fix Documentation not including the SID for ServiceID parameter
                        
                    : v1.0: Initial Release
						
    Disclaimer: Whilst I take considerable effort to ensure this script is error free and wont harm your enviroment.
                I have no way to test every possible senario it may be used in. I provide these scripts free
                to the Lync and Skype4B community AS IS without any warranty on its appropriateness for use in
                your enviroment. I disclaim all implied warranties including,
                without limitation, any implied warranties of merchantability or of fitness for a particular
                purpose. The entire risk arising out of the use or performance of the sample scripts and
                documentation remains with you. In no event shall I be liable for any damages whatsoever
                (including, without limitation, damages for loss of business profits, business interruption,
                loss of business information, or other pecuniary loss) arising out of the use of or inability
                to use the script or documentation.

    Acknowledgements 	
         : Testing and Advice
             Greig Sheriden https://greiginsydney.com/about/ @greiginsydney

         : Auto Update Code
             Pat Richard https://ucunleashed.com @patrichard

         : Proxy Detection
             Michel de Rooij	http://eightwone.com

  								
.INPUTS 
    None. New-CsRgsAustralianHolidayList.ps1 does not accept pipelined input.

.OUTPUTS
    New-CsRgsAustralianHolidayList.ps1 creates multiple new instances of the Microsoft.Rtc.Rgs.Management.WritableSettings.HolidaySet object and cannot be piped.

.PARAMETER -ServiceID <RgsIdentity> 
    Service where the new holiday set will be hosted. For example: -ServiceID "service:ApplicationServer:SFBFE01.Skype4badmin.com/1987d3c2-4544-489d-bbe3-59f79f530a83".
    To obtain your service ID, run Get-CsRgsConfiguration -Identity FEPool01.skype4badmin.com
    If you dont specify a ServiceID or FrontEndPool, the script will try and guess the frontend to put the holidays on.

.PARAMETER -FrontEndPool <FrontEnd FQDN> 
    Frontend Pool where the new holiday set will be hosted. 
    If you dont specify a ServiceID or FrontEndPool, the script will try and guess the frontend to put the holidays on.
    Specifiying this instead of ServiceID will cause the script to confirm the pool unless -Unattended is specified

.PARAMETER -RGSPrepend <String>
    String to Prepend to Listnames to suit your enviroment

.PARAMETER -DisableScriptUpdate
    Stops the script from checking online for an update and prompting the user to download. Ideal for scheduled tasks

.PARAMETER -RemoveExistingRules
    Deprecated. Script now updates existing rulesets rather than removing them. Kept for backwards compatability

.PARAMETER -Unattended
    Assumes yes for pool selection critera when multiple pools are present and Poolfqdn is specified.
    Also stops the script from checking for updates
    Check the script works before using this!

.LINK  
    http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/


.EXAMPLE

    PS C:\> New-CsRgsAustralianHolidayList.ps1 -ServiceID "service:ApplicationServer:SFBFE01.skype4badmin.com/1987d3c2-4544-489d-bbe3-59f79f530a83" -RGSPrepend "RGS-AU-"

    PS C:\> New-CsRgsAustralianHolidayList.ps1 

    PS C:\> New-CsRgsAustralianHolidayList.ps1 -DisableScriptUpdate -FrontEndPool AUMELSFBFE.Skype4badmin.local -Unattended

#>
# Script Config
#Requires -Version 3
#Requires -Modules SkypeforBusiness
[CmdletBinding(DefaultParametersetName='Common')]
param(
  [Parameter(Position=1)] [string]$ServiceID,
  [Parameter(Position=2)] [string]$RGSPrepend,
  [Parameter(Position=3)] [string]$FrontEndPool,
  [Parameter(Position=4)] [switch]$DisableScriptUpdate,
  [Parameter(Position=4)] [switch]$Unattended,
  [Parameter(Position=5)] [switch]$RemoveExistingRules,
  [Parameter(Position=6)] [string]$LogFileLocation
  )
#region config
  [Net.ServicePointManager]::SecurityProtocol = 'tls12, tls11, tls'
  $MaxCacheAge = 7 # Max age for XML cache, older than this # days will force info refresh
  $SessionCache = Join-Path -Path $PSScriptRoot -ChildPath 'AustralianHolidays.xml' #Filename for the XML data
  If (!$LogFileLocation) {$LogFileLocation = $PSCommandPath -replace '.ps1','.log'}
  [float]$ScriptVersion = '2.2'
  [string]$GithubRepo = 'New-CsRgsAustralianHolidayList'
  [string]$GithubBranch = 'master' #todo
  [string]$BlogPost = 'http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/'
#endregion config


#region Functions
Function Write-Log {
  <#
      .SYNOPSIS
      Fucntion to output messages to the console based on their severity and create log files

      .DESCRIPTION
      Add a more complete description of what the function does.

      .PARAMETER Message
      The message to write

      .PARAMETER Path
      The location of the logfile.

      .PARAMETER severity
      Sets the sevirity of the log message, Higher severitys will call Write-Warning or Write-Error

      .PARAMETER component
      Used to track the module or function that called "Write-Log" 

      .PARAMETER logonly
      Forces Write-Log to not display anything to the user

      .EXAMPLE
      Write-Log -Message 'This is a log message' -Severity 3 -component 'Example Component'
      Writes a log file message and displays a warning to the user

      .NOTES
      N/A

      .LINK
      http://www.skype4badmin.com

      .INPUTS
      This function does not accept pipelined input

      .OUTPUTS
      This function does not create pipelined output
  #>
    [CmdletBinding()]
    PARAM(
        [String]$Message,
        [String]$Path = $LogFileLocation,
        [int]$severity = 1,
        [string]$component = 'Default',
        [switch]$logonly
		  )
    $Date= Get-Date -Format 'HH:mm:ss'
    $Date2= Get-Date -Format 'MM-dd-yyyy'
    $MaxLogFileSizeMB = 10
         If(Test-Path -Path $Path)
         {
            if(((Get-ChildItem -Path $Path).length/1MB) -gt $MaxLogFileSizeMB) # Check the size of the log file and archive if over the limit.
            {
                $ArchLogfile = $Path.replace('.log', "_$(Get-Date -Format dd-MM-yyy_hh-mm-ss).lo_")
        Rename-Item -Path ren -NewName $Path -Path $ArchLogfile
            }
         }
         
     "$env:ComputerName date=$([char]34)$date2$([char]34) time=$([char]34)$date$([char]34) component=$([char]34)$component$([char]34) type=$([char]34)$severity$([char]34) Message=$([char]34)$Message$([char]34)"| Out-File -FilePath $Path -Append -NoClobber -Encoding default
    If (!$logonly) { #If LogOnly is set, we dont want to write anything to the screen as we are capturing data that might look bad onscreen
      #If the log entry is just informational (less than 2), output it to write verbose
      if ($severity -le 2) {"Info: $Message"| Write-Host -ForegroundColor Green}
      #If the log entry has a severity of 3 assume its a warning and write it to write-warning
      if ($severity -eq 3) {"$date $Message"| Write-Warning}
      #If the log entry has a severity of 4 or higher, assume its an error and display an error message (Note, critical errors are caught by throw statements so may not appear here)
      if ($severity -ge 4) {"$date $Message"| Write-Error}
    }
  } #end WriteLog

Function Get-IEProxy {

  Write-Log -Message 'Checking for proxy settings' -severity 1
        If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) {
            $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
            if ($proxies) {
                if ($proxies -ilike '*=*') {
                    return $proxies -replace '=', '://' -split (';') | Select-Object -First 1
                }
                Else {
                    return ('http://{0}' -f $proxies)
                }
            }
            Else {
                return $null
            }
        }
        Else {
            return $null
        }
    }

Function Get-ScriptUpdate {
 

  if ($DisableScriptUpdate -eq $false) {
    Write-Log -component 'Self Update' -Message 'Checking for Script Update' -severity 1
    Write-Log -component 'Self Update' -Message 'Checking for Proxy' -severity 1
      $ProxyURL = Get-IEProxy
    If ( $ProxyURL) {
      Write-Log -component 'Self Update' -Message "Using proxy address $ProxyURL" -severity 1
       }
    Else {
      Write-Log -component 'Self Update' -Message 'No proxy setting detected, using direct connection' -severity 1
        }
    }
    $GitHubScriptVersion = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/atreidae/$GitHubRepo/$GitHubBranch/version" -TimeoutSec 10 -Proxy $ProxyURL
        If ($GitHubScriptVersion.Content.length -eq 0) {
      Write-Log -component 'Self Update' -Message 'Error checking for new version. You can check manualy here' -severity 3
      Write-Log -component 'Self Update' -Message $BlogPost -severity 1
      Write-Log -component 'Self Update' -Message 'Pausing for 5 seconds' -severity 1
            start-sleep -Seconds 5
            }
        else { 
                if ([float]$GitHubScriptVersion.Content -gt [float]$ScriptVersion) {
        Write-Log -component 'Self Update' -Message 'New Version Available' -severity 3
                   #New Version available

                    #Prompt user to download
        $title = 'Update Available'
        $message = 'an update to this script is available, did you want to download it?'

        $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
          'Launches a browser window with the update'

        $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
          'No thanks.'

        $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

        switch ($result)
          {
            0 {
              Write-Log -component 'Self Update' -Message 'User opted to download update' -severity 1
              Start-Process -FilePath $BlogPost #todo F
              Write-Log -component 'Self Update' -Message 'Exiting Script' -severity 3
              Exit
            }
            1 {Write-Log -component 'Self Update' -Message 'User opted to skip update' -severity 1
									
              }
							
          }
                 }   
                 Else{
                 Write-Log -component 'Self Update' -Message "Script is up to date on $GithubBranch branch" -severity 1
                 }
        
         }

  }

#endregion Functions


Write-Log -Message "New-CsRgsAustralianHolidayList.ps1 Version $scriptversion" -severity 1
$culture = (Get-Culture)
$GMTOffset = (Get-WmiObject -Query 'Select Bias from Win32_TimeZone')
Write-Log -Message 'Current system culture'
Write-Log -Message $Culture
Write-Log -Message 'Current Timezone'
Write-Log -Message $GMTOffset.bias
Write-Log -Message 'Checking UTC Offset'
If ($GMTOffset.bias -lt 480) {
  Write-Log -Message 'UTC Base offset less than +8 hours'
  Write-log -Message 'Your timezone appears to be misconfigured. This script may not function as expected' -severity 3
  pause}

$National = $RGSPrepend+'National'

if ($Unattended) {$DisableScriptUpdate = $true}
if ($RemoveExistingRules -eq $true) {
  Write-log -Message 'RemoveExistingRules parameter set to True. Script will automatically delete existing entries from rules' -severity 3
    Write-Log -Message 'Pausing for 5 seconds' -severity 1
    start-sleep -Seconds 5
  }
#Get Proxy Details
      $ProxyURL = Get-IEProxy
    If ( $ProxyURL) {
        Write-Log -Message "Using proxy address $ProxyURL" -severity 1
    }
    Else {
        Write-Log -Message 'No proxy setting detected, using direct connection' -severity 1
    }

if ($DisableScriptUpdate -eq $false) {
  Write-Log -Message 'Checking for Script Update' -severity 1 #todo
   Get-ScriptUpdate

  }

Write-Log -Message 'Importing modules' -severity 1
#$VerbosePreference="SilentlyContinue" #Stops powershell showing Every cmdlet it imports
Import-Module -Name Lync
Import-module -Name SkypeForBusiness
#$VerbosePreference="Continue" #Comment out if you dont want to see whats going on



Write-Log -Message 'Checking for XML file' -severity 1


#Check for XML file and download it 
 $SessionCacheValid = $false
    If ( Test-Path -Path $SessionCache) {
        Try {
            If ( (Get-childItem -Path $SessionCache).LastWriteTime -ge (Get-Date).AddDays( - $MaxCacheAge)) {
                Write-Log -Message 'XML file found. Reading data' -severity 1
                [xml]$XMLdata = Get-Content -Path $SessionCache 
                $EventCount = ($XMLdata.OuterXml | select-string -Pattern '<event' -AllMatches)
                $XMLCount = ($EventCount.Matches.Count)
                Write-Log -Message "Imported file with $XMLCount event tags"  -severity 1
                if ($XMLCount -le 10) {
                         Write-Log -Message 'Imported file doesnt appear to contain correct data'  -severity 1
                         throw 'Imported file doesnt appear to contain correct data'
                         }
                
                $SessionCacheValid = $true
            }
            Else {
                Write-log -Message 'XML file expired. Will re-download XML from website' -severity 3
            }
        }
        Catch {
            Write-log -Message 'Error reading XML file or XML file invalid - Will re-download' -severity 3
        }
    }
   If ( -not( $SessionCacheValid)) {

        Write-Log -Message 'Downloading Date list from Australian Government Website' -severity 1
        Try {

            Invoke-WebRequest -Uri 'https://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays/xml' -TimeoutSec 20 -OutFile $SessionCache -Proxy $ProxyURL #-PassThru
             Write-Log -Message 'XML file downloaded. Reading data' -severity 1

                [xml]$XMLdata = Get-Content -Path $SessionCache 
                $EventCount = ($XMLdata.OuterXml | select-string -Pattern '<event' -AllMatches)
                $XMLCount = ($EventCount.Matches.Count)
                Write-Log -Message "Imported file with $XMLCount event tags"  -severity 1
                if ($XMLCount -le 10) {
                         Write-Log -Message 'Downloaded file doesnt appear to contain correct data'  -severity 1
                         throw 'Imported file doesnt appear to contain correct data'
                         }
                
                $SessionCacheValid = $true
            }
        Catch {
      Write-log -Message 'An error occurred attempting to download XML file automatically' -severity 3
      Write-log -Message 'Download the file from the URI below, name it "AustralianHolidays.xml" and place it in the same folder as this script' -severity 3
      Write-Log -Message 'https://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays/xml' -ForegroundColor Blue
      Throw ('Problem retrieving XML file {0}' -f $error[0])
            Exit 1
      }
     }



Write-Log -Message 'Gathering Front End Pool Data' -severity 1
$Pools = (Get-CsService -Registrar)

Write-Log -Message 'Checking Region Info' -severity 1
$ConvertTime = $false
$region = (Get-Culture)
if ($region.Name -ne 'en-AU') {
  #We're not running en-AU region setting, Warn the user and prompt them to change
  Write-log -Message 'This script is only supported on systems running the en-AU region culture' -severity 3
  Write-log -Message 'This is due to the way the New-CsRgsHoliday cmdlet processes date strings' -severity 3
  Write-log -Message 'More information is available at the url below' -severity 3
  Write-log -Message 'https://docs.microsoft.com/en-us/powershell/module/skype/new-csrgsholiday?view=skype-ps' -severity 3
  Write-log -Message 'The script will now prompt you to change regions. If you continue without changing regions I will output everything in US date format and hope for the best.' -severity 3

	
  #Prompt user to switch culture
  Write-Log -Message 'prompting user to change region'
        $title = 'Switch Windows Region?'
        $message = 'Update the Windows Region (Culture) to en-AU?'

        $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
          'Changes the Region Settings to en-AU and exits'

        $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
          'No, I like my date format, please convert the values.'

        $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

        switch ($result)
          {
            0 {Write-Log -Message 'Updating System Culture' -severity 1
              Set-Culture -CultureInfo en-AU
              Write-log -Message 'System Culture Updated, Script will exit.' -severity 3
              Write-log -Message 'Close any PowerShell windows and run the script again' -severity 3
              Exit
            }
            1 {Write-log -Message 'Unsupported Region. Setting compatability mode' -severity 3
              $ConvertTime = $true
              }
							
          }
  }




Write-Log -Message 'Parsing command line parameters' -severity 1

# Detect and deal with null service ID
If ($ServiceID -eq $null) {
    Write-log -Message 'No ServiceID entered, Searching for valid ServiceID' -severity 3
    Write-Log -Message 'Looking for Front End Pools' -severity 1
    $PoolNumber = ($Pools).count
    if ($PoolNumber -eq 1) { 
      Write-Log -Message "Only found 1 Front End Pool, $Pools.poolfqdn, Selecting it" -severity 1
      $RGSIDs = (Get-CsRgsConfiguration -Identity $pools.PoolFqdn)
      $Poolfqdn = $Pools.poolfqdn
      #Prompt user to confirm
        Write-Log -Message "Found RGS Service ID $RGSIDs" -severity 1
        $title = 'Use this Front End Pool?'
        $message = "Use the Response Group Server on $poolfqdn ?"

        $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
          'Continues using the selected Front End Pool.'

        $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
          'Aborts the script.'

        $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

        $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

        switch ($result)
          {
            0 {Write-Log -Message 'Updating ServiceID parameter' -severity 1
              $ServiceID = $RGSIDs.Identity.tostring()}
            1 {Write-log -Message "Couldn't Autolocate RGS pool. Aborting script" -severity 3
              Throw "Couldn't Autolocate RGS pool. Abort script"}
							
          }

        }
	

  Else {
    #More than 1 Pool Detected and the user didnt specify anything
    Write-Log -Message "Found $PoolNumber Front End Pools" -severity 1
	
    If ($FrontEndPool -eq $null) {
      Write-Log -Message 'Prompting user to select Front End Pool' -severity 1
      Write-log -Message "Couldn't Locate ServiceID or PoolFQDN on the command line and more than one Front End Pool was detected" -severity 3
			
      #Menu code thanks to Grieg.
      #First figure out the maximum width of the pools name (for the tabular menu):
      $width=0
      foreach ($Pool in ($Pools)) {
        if ($Pool.Poolfqdn.Length -gt $width) {
          $width = $Pool.Poolfqdn.Length
        }
      }

      #Provide an on-screen menu of Front End Pools for the user to choose from:
      $index = 0
      write-host ('Index  '), ('Pool FQDN'.Padright($width + 1),' '), 'Site ID'
      foreach ($Pool in ($Pools)) {
        write-host ($index.ToString()).PadRight(7,' '), ($Pool.Poolfqdn.Padright($width + 1),' '), $pool.siteid.ToString()
        $index++
        }
      $index--	#Undo that last increment
      Write-Host
      Write-Host 'Choose the Front End Pool you wish to use'
      $chosen = read-host -Prompt 'Or any other value to quit'
      Write-log -Message "User input $chosen" -severity 1
      if ($chosen -notmatch '^\d$') {Exit}
      if ([int]$chosen -lt 0) {Exit}
      if ([int]$chosen -gt $index) {Exit}
      $FrontEndPool = $pools[$chosen].PoolFqdn
      $Poolfqdn = $FrontEndPool
      $RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)
    }


    #User specified the pool at the commandline or we collected it earlier
		
    Write-Log -Message "Using Front End Pool $FrontendPool" -severity 1
    $RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)
    $Poolfqdn = $FrontEndPool



    if (!$Unattended) {
      #Prompt user to confirm
      $title = 'Use this Pool?'
      $message = "Use the Response Group Server on $poolfqdn ?"

      $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', `
      'Continues using the selected Front End Pool.'

      $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&No', `
      'Aborts the script.'

      $options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)

      $result = $host.ui.PromptForChoice($title, $message, $options, 0) 

      switch ($result)
      {
        0 {Write-Log -Message 'Updating ServiceID'  -severity 1
          $ServiceID = $RGSIDs.Identity.tostring()}
        1 {Write-log -Message 'Couldnt Autolocate RGS pool. Abort script' -severity 3
          Throw 'Couldnt Autolocate RGS pool. Abort script'}
      }
        }

  } 

}
#We should have a valid service ID by now

Write-Log -Message 'Parsing XML data' -severity 1
foreach ($State in $XMLData.ausgovEvents.jurisdiction) {
    switch ($state.jurisdictionName) 
      { 
        'ACT' {$StateName = ($RGSPrepend+'Australian Capital Territory')
              $StateID = 0 }

        'NSW' {$StateName = ($RGSPrepend+'New South Wales')
              $StateID = 1} 

        'NT' {$StateName = ($RGSPrepend+'Northern Territory')
             $StateID = 2}  

        'QLD' {$StateName = ($RGSPrepend+'Queensland')
              $StateID = 3} 

        'SA' {$StateName = ($RGSPrepend+'South Australia')
             $StateID = 4} 

        'TAS' {$StateName = ($RGSPrepend+'Tasmania') 
              $StateID = 5} 

        'VIC' {$StateName = ($RGSPrepend+'Victoria')
              $StateID = 6} 

        'WA' {$StateName = ($RGSPrepend+'Western Australia')
             $StateID = 7} 
       }
  

    Write-Log -Message "Processing events in $statename" -severity 1
    #Find and clear the existing RGS Object
    try {
      Write-Log -Message "Checking for existing $StateName Holiday Set" -severity 1
      $holidayset = (Get-CsRgsHolidaySet -Name "$StateName")
      Write-Log -Message "Removing old entries from $StateName" -severity 1
      $holidayset.HolidayList.clear()
    Write-Log -Message "Existing entries from Holiday Set $StateName removed" -severity 1
    }
  catch {Write-Log -Message "Didnt find $StateName Holiday Set. Creating" -severity 1
        $PlaceholderDate = (New-CsRgsHoliday -StartDate '11/11/1970 12:00 AM' -EndDate '12/11/1970 12:00 AM' -Name 'Placeholder. Shouldnt Exist')
        $holidayset = (New-CsRgsHolidaySet -Parent $ServiceID -Name "$Statename" -HolidayList $PlaceholderDate -ErrorAction silentlycontinue)
        Write-Log -Message 'Removing Placeholder Date' -severity 1
        $holidayset.HolidayList.clear()            
        }
 
        #Process Events in that State
        foreach ($event in $XMLData.ausgovEvents.jurisdiction[$stateID].events.event){
         
    #Deal with Unix date format
         $udate = get-date -Date '1/1/1970'
         if ($ConvertTime) {
                #American Date format
                $StartDate = ($Udate.AddSeconds($event.rawDate).ToLocalTime() | get-date -Format MM/dd/yyyy)
                $EndDate = ($Udate.AddSeconds(([int]$event.rawDate+86400)).ToLocalTime() | get-date -Format MM/dd/yyyy)     
                         }
            else {
                #Aussie Date format
                $Startdate = ($Udate.AddSeconds($event.rawDate).ToLocalTime() | get-date -Format dd/MM/yyyy)
                $EndDate = ($Udate.AddSeconds(([int]$event.rawDate+86400)).ToLocalTime() | get-date -Format dd/MM/yyyy)
                 }

        #Create the event in Skype format
         $EventName = ($event.holidayTitle)      
     $EventName = ($EventName -replace '  ' , ' ') #Remove Double Spaces in eventname
         $CurrentEvent = (New-CsRgsHoliday -StartDate "$StartDate 12:00 AM" -EndDate "$EndDate 12:00 AM" -Name "$StateName $EventName")
         #$CurrentEvent
        #add it to the variable.
        Write-Log -Message "Adding $EventName to $StateName" -severity 1
        $HolidaySet.HolidayList.Add($CurrentEvent)
        }
    Write-Log -Message 'Finished adding events' -severity 1
        Write-Log -Message "Writing $StateName to Database" -severity 1
        Try {Set-CsRgsHolidaySet -Instance $holidayset}

        Catch {Write-log -Message 'Something went wrong attempting to commit holidayset to database' -severity 3
          $ErrorMessage = $_.Exception.Message
          $FailedItem = $_.Exception.ItemName
          Write-Log -Message "$FailedItem failed. The error message was $ErrorMessage" -severity 4
          Throw $errormessage}
               
}


#Okay, now deal with National Holidays

 try {
      Write-Log -Message "Checking for existing $National Holiday Set" -severity 1
      $holidayset = (Get-CsRgsHolidaySet -Name "$National")
      Write-Log -Message "Removing old entries from $National" -severity 1
      $holidayset.HolidayList.clear()
    Write-Log -Message "Existing entries from Holiday Set $National removed" -severity 1
    }
  catch {Write-Log -Message "Didnt find $National Holiday Set. Creating" -severity 1
        $PlaceholderDate = (New-CsRgsHoliday -StartDate '11/11/1970 12:00 AM' -EndDate '12/11/1970 12:00 AM' -Name 'Placeholder. Shouldnt Exist')
        $holidayset = (New-CsRgsHolidaySet -Parent $ServiceID -Name "$National" -HolidayList $PlaceholderDate -ErrorAction silentlycontinue)
        Write-Log -Message 'Removing Placeholder Date' -severity 1
        $holidayset.HolidayList.clear()            
        }

#Find dates that are in every state

 Write-Log -Message 'Finding National Holidays (This can take a while)' -severity 1
$i =0
$RawNatHolidayset = $null
$NatHolidayset = $null

$RawNatHolidayset = @()

foreach ($State in $XMLData.ausgovEvents.jurisdiction) {

    #Process Events in that State
        foreach ($event in $XMLData.ausgovEvents.jurisdiction[$i].events.event) {
        $RawNatHolidayset += ($event)
        }
        $i ++

  }

  $NatHolidayset = ($RawNatHolidayset | Sort-Object -Property rawDate -Unique)
  ForEach($Uniquedate in $NatHolidaySet){

             $SEARCH_RESULT=$RawNatHolidaySet|Where-Object{$_.rawDate -eq $Uniquedate.rawdate}

             if ( $SEARCH_RESULT.Count -eq 8)
             {      
                    $event = ($SEARCH_RESULT | select-object -first 1)
                  

                    #Deal with Unix date format
                         $udate = get-date -Date '1/1/1970'
                         if ($ConvertTime) {
                                #American Date format
                                $StartDate = ($Udate.AddSeconds($event.rawDate).ToLocalTime() | get-date -Format MM/dd/yyyy)
                                $EndDate = ($Udate.AddSeconds(([int]$event.rawDate+86400)).ToLocalTime() | get-date -Format MM/dd/yyyy)     
                                         }
                            else {
                                #Aussie Date format
                                $Startdate = ($Udate.AddSeconds($event.rawDate).ToLocalTime() | get-date -Format dd/MM/yyyy)
                                $EndDate = ($Udate.AddSeconds(([int]$event.rawDate+86400)).ToLocalTime() | get-date -Format dd/MM/yyyy)
                                 }
                                 
                            #Create the event in Skype format
                            Write-Log -Message "Found $EventName" -severity 1
                             $EventName = ($event.holidayTitle)      
                             $CurrentEvent = (New-CsRgsHoliday -StartDate "$StartDate 12:00 AM" -EndDate "$EndDate 12:00 AM" -Name "$StateName $EventName")
                             $HolidaySet.HolidayList.Add($CurrentEvent)
             }

}
 Write-Log -Message 'Finished adding events' -severity 1
 Write-Log -Message "Writing $National to Database" -severity 1
 Try {Set-CsRgsHolidaySet -Instance $holidayset}
            Catch {Write-log -Message 'Something went wrong attempting to commit holidayset to database' -severity 3
      $ErrorMessage = $_.Exception.Message
      $FailedItem = $_.Exception.ItemName
      Write-Log -Message "$FailedItem failed. The error message was $ErrorMessage" -ForegroundColor Red
      Throw $errormessage}



Write-Log -Message ''
Write-Log -Message ''
Write-Log -Message 'Looks like everything went okay. Here are your current RGS Holiday Sets' -severity 1
Get-CsRgsHolidaySet | Select-Object -Property name
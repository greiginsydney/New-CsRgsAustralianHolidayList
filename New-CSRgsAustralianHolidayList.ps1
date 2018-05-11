<#  
.SYNOPSIS  
	This script creates RGS holidaysets for Australian states based on live data from the Australian Government website


.DESCRIPTION  
	Created by James Arber. www.skype4badmin.com
	Although every effort has been made to ensure this list is correct, dates change and sometimes I goof. 
	Please use at your own risk.
	Data taken from http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays
	    
	
.NOTES  
    Version      	   	: 2.01
	Date			    : 9/12/2017
	Lync Version		: Tested against Skype4B Server 2015 and Lync Server 2013
    Author    			: James Arber
	Header stolen from  : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

	Revision History	: v2.01: Migrated to GitHub
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
						
	Disclaimer   		: Whilst I take considerable effort to ensure this script is error free and wont harm your enviroment.
								I have no way to test every possible senario it may be used in. I provide these scripts free
								to the Lync and Skype4B community AS IS without any warranty on its appropriateness for use in
								your enviroment. I disclaim all implied warranties including,
  								without limitation, any implied warranties of merchantability or of fitness for a particular
  								purpose. The entire risk arising out of the use or performance of the sample scripts and
  								documentation remains with you. In no event shall I be liable for any damages whatsoever
  								(including, without limitation, damages for loss of business profits, business interruption,
  								loss of business information, or other pecuniary loss) arising out of the use of or inability
  								to use the script or documentation.

	Acknowledgements 	: Testing and Advice
  								Greig Sheriden https://greiginsydney.com/about/ @greiginsydney

						: Auto Update Code
								Pat Richard http://www.ehloworld.com @patrichard

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
[CmdletBinding(DefaultParametersetName="Common")]
param(
	[Parameter(Mandatory=$false, Position=1)] $ServiceID,
	[Parameter(Mandatory=$false, Position=2)] $RGSPrepend,
	[Parameter(Mandatory=$false, Position=3)] $FrontEndPool,
	[Parameter(Mandatory=$false, Position=4)] [switch]$DisableScriptUpdate,
    [Parameter(Mandatory=$false, Position=4)] [switch]$Unattended,
	[Parameter(Mandatory=$false, Position=5)] [switch]$RemoveExistingRules
	)
#region config
 
    $MaxCacheAge = 7 # Max age for XML cache, older than this # days will force info refresh
	$SessionCache = Join-Path $PSScriptRoot 'AustralianHolidays.xml' #Filename for the XML data
	[single]$Version = "2.01"
#endregion config


#region Fucntions
Function Get-IEProxy {
	Write-Host "Info: Checking for proxy settings" -ForegroundColor Green
        If ( (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').ProxyEnable -ne 0) {
            $proxies = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings').proxyServer
            if ($proxies) {
                if ($proxies -ilike "*=*") {
                    return $proxies -replace "=", "://" -split (';') | Select-Object -First 1
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


#endregion Functions




#Define Listnames
Write-Host "Info: New-CsRgsAustralianHolidayList.ps1 Version $version" -ForegroundColor Green
$National = $RGSPrepend+"National"
$Vic = $RGSPrepend+"Victoria"
$NSW = $RGSPrepend+"New South Wales"
$QLD = $RGSPrepend+"Queensland"
$ACT = $RGSPrepend+"Australian Capital Territory"
$NT = $RGSPrepend+"Northern Territory"
$SA = $RGSPrepend+"South Australia"
$WA = $RGSPrepend+"Western Australia"
$Tas = $RGSPrepend+"Tasmania"

$AllStates = @()
$allstates += $National
$allstates += $Vic
$allstates += $NSW
$allstates += $QLD
$allstates += $ACT
$allstates += $NT
$allstates += $SA
$allstates += $WA
$allstates += $TAS
if ($Unattended) {$DisableScriptUpdate = $true}
if ($RemoveExistingRules -eq $true) {
	Write-Warning "RemoveExistingRules parameter set to True. Script will automatically delete existing entries from rules"
    Write-Host "Info: Pausing for 5 seconds" -ForegroundColor Green
    start-sleep 5
	}
#Get Proxy Details
	    $ProxyURL = Get-IEProxy
    If ( $ProxyURL) {
        Write-Host "Info: Using proxy address $ProxyURL" -ForegroundColor Green
    }
    Else {
        Write-Host "Info: No proxy setting detected, using direct connection" -ForegroundColor Green
    }

if ($DisableScriptUpdate -eq $false) {
	Write-Host "Info: Checking for Script Update" -ForegroundColor Green #todo
    $GitHubScriptVersion = Invoke-WebRequest https://raw.githubusercontent.com/atreidae/New-CsRgsAustralianHolidayList/master/version -TimeoutSec 10 -Proxy $ProxyURL
        If ($GitHubScriptVersion.Content.length -eq 0) {

            Write-Warning "Error checking for new version. You can check manualy here"
            Write-Warning "http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/"
            Write-Host "Info: Pausing for 5 seconds" -ForegroundColor Green
            start-sleep 5
            }
        else { 
                if ([single]$GitHubScriptVersion.Content -gt [single]$version) {
                 Write-Host "Info: New Version Available" -ForegroundColor Green
                    #New Version available

                    #Prompt user to download
				$title = "Update Available"
				$message = "an update to this script is available, did you want to download it?"

				$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
					"Launches a browser window with the update"

				$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
					"No thanks."

				$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

				$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

				switch ($result)
					{
						0 {Write-Host "Info: User opted to download update" -ForegroundColor Green
							start "http://www.skype4badmin.com/australian-holiday-rulesets-for-response-group-service/"
							Write-Warning "Exiting script"
							Exit
						}
						1 {Write-Host "Info: User opted to skip update" -ForegroundColor Green
							
							}
							
					}
                 }   
                 Else{
                 Write-Host "Info: Script is up to date" -ForegroundColor Green
                 }
        
	       }

	}

Write-Host "Info: Importing modules" -ForegroundColor Green
#$VerbosePreference="SilentlyContinue" #Stops powershell showing Every cmdlet it imports
Import-Module Lync
Import-module SkypeForBusiness
#$VerbosePreference="Continue" #Comment out if you dont want to see whats going on



Write-Host "Info: Checking for XML file" -ForegroundColor Green


#Check for XML file and download it 
 $SessionCacheValid = $false
    If ( Test-Path $SessionCache) {
        Try {
            If ( (Get-childItem -Path $SessionCache).LastWriteTime -ge (Get-Date).AddDays( - $MaxCacheAge)) {
                Write-Host 'Info: XML file found. Reading data' -ForegroundColor Green
                [xml]$XMLdata = Get-Content -Path $SessionCache 
                $EventCount = ($XMLdata.OuterXml | select-string "<event" -AllMatches)
                $XMLCount = ($EventCount.Matches.Count)
                Write-Host "Info: Imported file with $XMLCount event tags"  -ForegroundColor Green
                if ($XMLCount -le 10) {
                         Write-host "Info: Imported file doesnt appear to contain correct data"  -ForegroundColor Green
                         throw "Imported file doesnt appear to contain correct data"
                         }
                
                $SessionCacheValid = $true
            }
            Else {
                Write-Warning 'Info: XML file expired. Will re-download XML from website' -ForegroundColor Green
            }
        }
        Catch {
            Write-Warning 'Error reading XML file or XML file invalid - Will re-download'
        }
    }
	 If ( -not( $SessionCacheValid)) {

        Write-Host 'Info: Downloading Date list from Australian Government Website' -ForegroundColor Green
        Try {
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri 'http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays/xml' -TimeoutSec 20 -OutFile $SessionCache -Proxy $ProxyURL #-PassThru
             Write-Host 'Info: XML file downloaded. Reading data' -ForegroundColor Green
                [xml]$XMLdata = Get-Content -Path $SessionCache 
                $EventCount = ($XMLdata.OuterXml | select-string "<event" -AllMatches)
                $XMLCount = ($EventCount.Matches.Count)
                Write-Host "Info: Imported file with $XMLCount event tags"  -ForegroundColor Green
                if ($XMLCount -le 10) {
                         Write-host "Info: Downloaded file doesnt appear to contain correct data"  -ForegroundColor Green
                         throw "Imported file doesnt appear to contain correct data"
                         }
                
                $SessionCacheValid = $true
            }
        Catch {
			Write-Warning "An error occurred attempting to download XML file automatically"
			Write-Warning 'Download the file from the URI below, name it "AustralianHolidays.xml" and place it in the same folder as this script'
			Write-host "http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays/xml" -ForegroundColor Blue
			Throw ('Problem retrieving XML file {0}' -f $error[0])
            Exit 1
			}
		 }



Write-Host "Info: Gathering Front End Pool Data" -ForegroundColor Green
$Pools = (Get-CsService -Registrar)

Write-Host "Info: Checking Region Info" -ForegroundColor Green
$ConvertTime = $false
$region = (Get-Culture)
if ($region.Name -ne "en-AU") {
	#We're not running en-AU region setting, Warn the user and prompt them to change
	Write-Warning "This script is only supported on systems running the en-AU region culture"
	Write-Warning "This is due to the way the New-CsRgsHoliday cmdlet processes date strings"
	Write-Warning "More information is available at the url below"
	Write-Warning "https://docs.microsoft.com/en-us/powershell/module/skype/new-csrgsholiday?view=skype-ps"
	Write-Warning "The script will now prompt you to change regions. If you continue without changing regions I will output everything in US date format and hope for the best."

	
	#Prompt user to switch culture
				$title = "Switch Windows Region?"
				$message = "Update the Windows Region (Culture) to en-AU?"

				$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
					"Changes the Region Settings to en-AU and exits"

				$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
					"No, I like my date format, please convert the values."

				$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

				$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

				switch ($result)
					{
						0 {Write-Host "Info: Updating System Culture" -ForegroundColor Green
							Set-Culture en-AU
							Write-Warning "System Culture Updated, Script will exit."
							Write-Warning "Close any PowerShell windows and run the script again"
							Exit
						}
						1 {Write-Warning "Unsupported Region. Setting compatability mode"
							$ConvertTime = $true
							}
							
					}
	}




Write-Host "Info: Parsing command line parameters" -ForegroundColor Green

# Detect and deal with null service ID
If ($ServiceID -eq $null) {
		Write-Warning "No ServiceID entered, Searching for valid ServiceID"
		Write-Host "Info: Looking for Front End Pools" -ForegroundColor Green
		$PoolNumber = ($Pools).count
		if ($PoolNumber -eq 1) { 
			Write-Host "Info: Only found 1 Front End Pool, $Pools.poolfqdn, Selecting it" -ForegroundColor Green
			$RGSIDs = (Get-CsRgsConfiguration -Identity $pools.PoolFqdn)
			$Poolfqdn = $Pools.poolfqdn
			#Prompt user to confirm
				Write-Host "Info: Found RGS Service ID $RGSIDs" -ForegroundColor Green
				$title = "Use this Front End Pool?"
				$message = "Use the Response Group Server on $poolfqdn ?"

				$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
					"Continues using the selected Front End Pool."

				$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
					"Aborts the script."

				$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

				$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

				switch ($result)
					{
						0 {Write-Host "Info: Updating ServiceID parameter" -ForegroundColor Green
							$ServiceID = $RGSIDs.Identity.tostring()}
						1 {Write-Warning "Couldn't Autolocate RGS pool. Aborting script"
							Throw "Couldn't Autolocate RGS pool. Abort script"}
							
					}

				}
	

	Else {
	#More than 1 Pool Detected and the user didnt specify anything
	Write-Host "Info: Found $PoolNumber Front End Pools" -ForegroundColor Green
	
		If ($FrontEndPool -eq $null) {
			Write-Host "Info: Prompting user to select Front End Pool" -ForegroundColor Green
			Write-Warning "Couldn't Locate ServiceID or PoolFQDN on the command line and more than one Front End Pool was detected"
			
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
			write-host ("Index  "), ("Pool FQDN".Padright($width + 1)," "), "Site ID"
			foreach ($Pool in ($Pools)) {
				write-host ($index.ToString()).PadRight(7," "), ($Pool.Poolfqdn.Padright($width + 1)," "), $pool.siteid.ToString()
				$index++
				}
			$index--	#Undo that last increment
			Write-Host
			Write-Host "Choose the Front End Pool you wish to use"
			$chosen = read-host "Or any other value to quit"

			if ($chosen -notmatch '^\d$') {Exit}
			if ([int]$chosen -lt 0) {Exit}
			if ([int]$chosen -gt $index) {Exit}
			$FrontEndPool = $pools[$chosen].PoolFqdn
			$Poolfqdn = $FrontEndPool
			$RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)
		}


	#User specified the pool at the commandline or we collected it earlier
		
	Write-Host "Info: Using Front End Pool $FrontendPool" -ForegroundColor Green
	$RGSIDs = (Get-CsRgsConfiguration -Identity $FrontEndPool)
	$Poolfqdn = $FrontEndPool



if (!$Unattended) {
	#Prompt user to confirm
		$title = "Use this Pool?"
		$message = "Use the Response Group Server on $poolfqdn ?"

		$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
			"Continues using the selected Front End Pool."

		$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
			"Aborts the script."

		$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

		$result = $host.ui.PromptForChoice($title, $message, $options, 0) 

		switch ($result)
			{
				0 {Write-Host "Info: Updating ServiceID"  -ForegroundColor Green
					$ServiceID = $RGSIDs.Identity.tostring()}
				1 {Write-Warning "Couldnt Autolocate RGS pool. Abort script"
					Throw "Couldnt Autolocate RGS pool. Abort script"}
			}
        }

	} 

}
#We should have a valid service ID by now


 $removedsomething = $false
 $alreadyexists = $false

Write-Host "Info: Parsing XML data" -ForegroundColor Green
foreach ($State in $XMLData.ausgovEvents.jurisdiction) {
    switch ($state.jurisdictionName) 
      { 
        "ACT" {$StateName = ($RGSPrepend+"Australian Capital Territory")
              $StateID = 0 }

        "NSW" {$StateName = ($RGSPrepend+"New South Wales")
              $StateID = 1} 

        "NT" {$StateName = ($RGSPrepend+"Northern Territory")
             $StateID = 2}  

        "QLD" {$StateName = ($RGSPrepend+"Queensland")
              $StateID = 3} 

        "SA" {$StateName = ($RGSPrepend+"South Australia")
             $StateID = 4} 

        "TAS" {$StateName = ($RGSPrepend+"Tasmania") 
              $StateID = 5} 

        "VIC" {$StateName = ($RGSPrepend+"Victoria")
              $StateID = 6} 

        "WA" {$StateName = ($RGSPrepend+"Western Australia")
             $StateID = 7} 
       }
  

    Write-Host "Info: Processing events in $statename" -ForegroundColor Green
    #Find and clear the existing RGS Object
    try {
	    Write-Host "Info: Checking for existing $StateName Holiday Set" -ForegroundColor Green
	    $holidayset = (Get-CsRgsHolidaySet -Name "$StateName")
	    Write-host "Info: Removing old entries from $StateName" -ForegroundColor Green
	    $holidayset.HolidayList.clear()
		Write-Host "Info: Existing entries from Holiday Set $StateName removed" -ForegroundColor Green
		}
	catch {Write-Host "Info: Didnt find $StateName Holiday Set. Creating" -ForegroundColor Green
        $PlaceholderDate = (New-CsRgsHoliday -StartDate "11/11/1970 12:00 AM" -EndDate "12/11/1970 12:00 AM" -Name "Placeholder. Shouldnt Exist")
        $holidayset = (New-CsRgsHolidaySet -Parent $ServiceID -Name "$Statename" -HolidayList $PlaceholderDate -ErrorAction silentlycontinue)
        Write-Host "Info: Removing Placeholder Date" -ForegroundColor Green
        $holidayset.HolidayList.clear()            
        }
 
        #Process Events in that State
        foreach ($event in $XMLData.ausgovEvents.jurisdiction[$stateID].events.event){
         
		#Deal with Unix date format
         $udate = get-date "1/1/1970"
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
        Write-host "Info: Adding $EventName to $StateName" -ForegroundColor Green
        $HolidaySet.HolidayList.Add($CurrentEvent)
        }
		Write-Host "Info: Finished adding events" -ForegroundColor Green
        Write-host "Info: Writing $StateName to Database" -ForegroundColor Green
        Try {Set-CsRgsHolidaySet -Instance $holidayset}
            Catch {Write-Warning "Something went wrong attempting to commit holidayset to database"
			$ErrorMessage = $_.Exception.Message
			$FailedItem = $_.Exception.ItemName
			Write-host "$FailedItem failed. The error message was $ErrorMessage" -ForegroundColor Red
			Throw $errormessage}
               
}


#Okay, now deal with National Holidays

 try {
	    Write-Host "Info: Checking for existing $National Holiday Set" -ForegroundColor Green
	    $holidayset = (Get-CsRgsHolidaySet -Name "$National")
	    Write-host "Info: Removing old entries from $National" -ForegroundColor Green
	    $holidayset.HolidayList.clear()
		Write-Host "Info: Existing entries from Holiday Set $National removed" -ForegroundColor Green
		}
	catch {Write-Host "Info: Didnt find $National Holiday Set. Creating" -ForegroundColor Green
        $PlaceholderDate = (New-CsRgsHoliday -StartDate "11/11/1970 12:00 AM" -EndDate "12/11/1970 12:00 AM" -Name "Placeholder. Shouldnt Exist")
        $holidayset = (New-CsRgsHolidaySet -Parent $ServiceID -Name "$National" -HolidayList $PlaceholderDate -ErrorAction silentlycontinue)
        Write-Host "Info: Removing Placeholder Date" -ForegroundColor Green
        $holidayset.HolidayList.clear()            
        }

#Find dates that are in every state

 Write-Host "Info: Finding National Holidays (This can take a while)" -ForegroundColor Green
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

             $SEARCH_RESULT=$RawNatHolidaySet|?{$_.rawDate -eq $Uniquedate.rawdate}

             if ( $SEARCH_RESULT.Count -eq 8)
             {      
                    $event = ($SEARCH_RESULT | select-object -first 1)
                  

                    #Deal with Unix date format
                         $udate = get-date "1/1/1970"
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
                            Write-Host "Info: Found $EventName" -ForegroundColor Green
                             $EventName = ($event.holidayTitle)      
                             $CurrentEvent = (New-CsRgsHoliday -StartDate "$StartDate 12:00 AM" -EndDate "$EndDate 12:00 AM" -Name "$StateName $EventName")
                             $HolidaySet.HolidayList.Add($CurrentEvent)
             }

}
 Write-Host "Info: Finished adding events" -ForegroundColor Green
 Write-host "Info: Writing $National to Database" -ForegroundColor Green
 Try {Set-CsRgsHolidaySet -Instance $holidayset}
            Catch {Write-Warning "Something went wrong attempting to commit holidayset to database"
			$ErrorMessage = $_.Exception.Message
			$FailedItem = $_.Exception.ItemName
			Write-host "$FailedItem failed. The error message was $ErrorMessage" -ForegroundColor Red
			Throw $errormessage}

Write-Host ""
Write-Host ""
Write-Host "Info: Looks like everything went okay. Here are your current RGS Holiday Sets" -ForegroundColor Green
Get-CsRgsHolidaySet | select name

.SYNOPSIS  
	This script creates RGS holidaysets for Australian states based on live data from the Australian Government website


.DESCRIPTION  
	Created by James Arber. www.skype4badmin.com
	Although every effort has been made to ensure this list is correct, dates change and sometimes I goof. 
	Please use at your own risk.
	Data taken from http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays
	    
	
.NOTES  
    Version      	   	: 2.10
	Date			    : 31/03/2018
	Lync Version		: Tested against Skype4B Server 2015 and Lync Server 2013
    Author    			: James Arber
	Header stolen from  : Greig Sheridan who stole it from Pat Richard's amazing "Get-CsConnections.ps1"

	Revision History	: v2.1: Added Script logging
                              : Updated to use my new autoupdate code
                              : Added ability to switch between devel/master branches
                              : Added timezone offset detection / warning

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

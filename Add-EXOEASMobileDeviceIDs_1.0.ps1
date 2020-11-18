#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Exchange Online Device partnership inventory
# Add-EXOEASMobileDeviceIDs.ps1
#  
# Created by: Austin McCollum 2/11/2018 austinmc@microsoft.com
# Updated by: Kevin Bloom and Garrin Thompson 11/3/2020 Kevin.Bloom@Microsoft.com garrint@microsoft.com *** "Borrowed" a few quality-of-life functions from Start-RobustCloudCommand.ps1 and added EXOv2 connection
#
#########################################################################################
# This script reads the output file from EXO_MobileDevice_Inventory_3.6.ps1 and adds mobile device IDs for non-Outlook Mobile devices that are Allowed and Externally Managed
#
#########################################################################################

# Writes output to a log file with a time date stamp
Function Write-Log {
	Param ([string]$string)
	$NonInteractive = 1
	# Get the current date
	[string]$date = Get-Date -Format G
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Host
	}
}

# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
	Param([int]$sleeptime)
	# Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		# Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		# Sleep 1 second
		start-sleep 1
	}
	Write-Progress -Completed -Activity "Sleeping"
}

# Setup a new O365 Powershell Session using RobustCloudCommand concepts
Function New-CleanO365Session {
	#Prompt for UPN used to login to EXO 
   Write-log ("Removing all PS Sessions")

   # Destroy any outstanding PS Session
   Get-PSSession | Remove-PSSession -Confirm:$false
   
   # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
   [System.GC]::Collect()
   
   # Sleep 10s to allow the sessions to tear down fully
   Write-Log ("Sleeping 10 seconds to clear existing PS sessions")
   Start-Sleep -Seconds 10

   # Clear out all errors
   $Error.Clear()
   
   # Create the session
   Write-Log ("Creating new PS Session")
	#OLD BasicAuth method create session
	#$Exchangesession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
   # Check for an error while creating the session
	If ($Error.Count -gt 0){
		Write-log ("[ERROR] - Error while setting up session")
		Write-log ($Error)
		# Increment our error count so we abort after so many attempts to set up the session
		$ErrorCount++
		# If we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		If ($ErrorCount -gt 3){
			Write-log ("[ERROR] - Failed to setup session after multiple tries")
			Write-log ("[ERROR] - Aborting Script")
			exit		
		}	
		# If we are not aborting then sleep 60s in the hope that the issue is transient
		Write-log ("Sleeping 60s then trying again...standby")
		Start-SleepWithProgress -sleeptime 60
		
		# Attempt to set up the sesion again
		New-CleanO365Session
	}
   
   # If the session setup worked then we need to set $errorcount to 0
   else {
	   $ErrorCount = 0
   }
   # Import the PS session/connect to EXO
	$null = Connect-ExchangeOnline -UserPrincipalName $EXOLogonUPN -DelegatedOrganization $EXOtenant -ShowProgress:$false -ShowBanner:$false
   # Set the Start time for the current session
	Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy; Goes ahead and resets it every "$ResetSeconds" number of seconds (14.5 mins) either way 
Function Test-O365Session {
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession
	# Make sure we found a session
	if ($SessionInfo -eq $null) { 
		Write-log ("[ERROR] - No Session Found")
		Write-log ("Recreating Session")
		New-CleanO365Session
	}	
	# Make sure it is in an opened state if not log and recreate
	elseif ($SessionInfo.State -ne "Opened"){
		Write-log ("[ERROR] - Session not in Open State")
		Write-log ($SessionInfo | fl | Out-String )
		Write-log ("Recreating Session")
		New-CleanO365Session
	}
	# If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		Write-log ("Rebuilding Connection")
		
		# Estimate the throttle delay needed since the last session rebuild
		# Amount of time the session was allowed to run * our activethrottle value
		# Divide by 2 to account for network time, script delays, and a fudge factor
		# Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		[int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		# If the delay is >15s then sleep that amount for throttle to recover
		if ($DelayinSeconds -gt 0){
			Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			Start-SleepWithProgress -SleepTime $DelayinSeconds
		}
		# If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		else {
			Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		}
		# new O365 session and reset our object processed count
		New-CleanO365Session
	}
	else {
		# If session is active and it hasn't been open too long then do nothing and keep going
	}
	# If we have a manual throttle value then sleep for that many milliseconds
	if ($ManualThrottle -gt 0){
		Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

#------------------v
#ScriptSetupSection
#------------------v

#Set Variables
$logfilename = '\Add-EXOEASMobileDeviceIDs_logfile_'
$outputfilename = '\Add-EXOEASMobileDeviceIDs_Output_'
$execpol = get-executionpolicy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force  #this is just for the session running this script
Write-Host;$EXOLogonUPN=Read-host "Type in UPN for account that will execute this script"
$EXOtenant=Read-host "Type in your tenant domain name (eg <domain>.onmicrosoft.com)";write-host "...pleasewait...connecting to EXO..."
$SmtpCreds = (get-credential -Message "Provide EXO account Pasword" -UserName "$EXOLogonUPN")
# Set $OutputFolder to Current PowerShell Directory
[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
$outputFolder = [IO.Directory]::GetCurrentDirectory()
$DateTicks = (Get-Date).Ticks
$logFile = $outputFolder + $logfilename + $DateTicks + ".txt"
$OutputFile= $outputfolder + $outputfilename + $DateTicks + ".csv"
[int]$ManualThrottle=0
[double]$ActiveThrottle=.25
[int]$ResetSeconds=870

#Imports the previous file and filters items into a new array that are not Outlook for iOS and Android and are Allowed and are Externally Managed
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{InitialDirectory = $outputFolder}
$Null = $FileBrowser.ShowDialog()

Write-Log ("Importing CSV data file")
$ExoMobileDevices = Import-Csv $FileBrowser.FileName
Write-Log ("Filtering Mobile Devices")
$ExoMobileDevicesFiltered = @()
foreach ($Item in $ExoMobileDevices)
{
    if ($Item.DeviceModel -ne 'Outlook for iOS and Android' -and $Item.AccessState -eq 'Allowed' -and $Item.AccessReason -eq 'ExternallyManaged') {
        $ExoMobileDevicesFiltered += $Item
    }
}
Remove-Variable ExoMobileDevices 
Write-Log ("Sorting Mobile Devices")
$ExoMobileDevicesFilteredAndSorted = $ExoMobileDevicesFiltered | Sort-Object -Property PrimarySmtpAddress
Remove-Variable ExoMobileDevicesFiltered

#Creates and populates a hash-table with the user information and their missing Device IDs
Write-Log ("Organizing Mobile Devices per their associated user")
$TempUser = $null
[array]$TempDeviceIDs = @()
[array]$ExoMobileDevicesHashed = @()
Foreach ($Item in $ExoMobileDevicesFilteredAndSorted)
#Foreach ($Item in $TempExoMobile)
{
	If ($Item.PrimarySmtpAddress -eq $TempUser)
	{
		$TempDeviceIDs += $Item.DeviceID
		#Write-Host "hits If" -ForegroundColor Cyan
		#Write-Host $TempDeviceIDs -ForegroundColor cyan
	}
	else 
	{
		#Adds previous record to the collection if the user is populated
		if ($TempUser -ne $null)
		{
			$Record = "" | select PrimarySMTPAddress,DeviceIDs
			$Record.PrimarySMTPAddress = $TempUser
			$Record.DeviceIDs = $TempDeviceIDs
			#Write-Host $Record -ForegroundColor DarkGreen
			$ExoMobileDevicesHashed += $Record
		}
		#Prepps the variables for the next loop run
		$TempUser = $Item.PrimarySmtpAddress
		$TempDeviceIDs = $Item.DeviceID
		#Write-Host $TempUser -ForegroundColor DarkYellow
		#Write-Host $TempDeviceIDs -ForegroundColor DarkYellow
	}
}
#Adds the last record to the collection
$Record = "" | select PrimarySMTPAddress,DeviceIDs
$Record.PrimarySMTPAddress = $TempUser
$Record.DeviceIDs = $TempDeviceIDs
$ExoMobileDevicesHashed += $Record
#Write-Host $Record -ForegroundColor DarkGreen
#$ExoMobileDevicesHashed

#Exports the $ExoMobileDevicesHashed and $ExoMobileDevicesFilteredAndSorted to file for reference if needed
Write-Log ("Exporting Filtered and Organized Users and their Mobile Devices")
$ExoMobileDevicesHashedFileName = $outputfolder + "\ExoMobileDevicesHashed" + $DateTicks + ".csv"
$ExoMobileDevicesFilteredAndSortedFileName = $outputfolder + "\ExoMobileDevicesFilteredAndSorted" + $DateTicks + ".csv"
$ExoMobileDevicesHashed | select PrimarySMTPAddress, @{L = "DeviceIDs"; E = { $_.DeviceIDs -join ";"}} | export-csv -path $ExoMobileDevicesHashedFileName -notypeinformation
$ExoMobileDevicesFilteredAndSorted | export-csv -path $ExoMobileDevicesFilteredAndSortedFileName -notypeinformation

# Setup our first session to O365
$ErrorCount = 0
New-CleanO365Session
Write-Log ("Connected to Exchange Online")
write-host;write-host -ForegroundColor Green "...Connected to Exchange Online as $EXOLogonUPN";write-host

# Get when we started the script for estimating time to completion
$ScriptStartTime = Get-Date
$startDate = Get-Date
write-progress -id 1 -activity "Beginning..." -PercentComplete (1) -Status "initializing variables"

# Clear the error log so that sending errors to file relate only to this run of the script
$error.clear()

#-------------------------v
#Start CUSTOM CODE Section
#-------------------------v

# Set a counter and some variables to use for periodic write/flush and reporting for loop to create Hashtable
$currentProgress = 1
[TimeSpan]$caseCheckTotalTime=0
# report counter
$c = 0
# running counter
$i = 0
# Set the number of objects to cycle before writing to disk and sending stats, i'd consider 5000 max
$statLimit = 250
# Get the total number of objects, which we use in some stat calculations
$t = $ExoMobileDevicesHashed.Count
# Set some timedate variables for the stats report
$loopStartTime = Get-Date
$loopCurrentTime = Get-Date

$progressActions = $ExoMobileDevicesHashed.Count
[array]$ExoMobileDevicesHashedRunningList = @()
$Line = "" | select PrimarySMTPAddress,DeviceIDs

Write-Log ("Starting to add Device IDs")

# Update Screen
Write-host;Write-host -foregroundcolor Cyan "Starting to add Device IDs...(counters will display for every 250 users updated)";;sleep 2;write-host "-------------------------------------------------"
Write-Host -NoNewline "Total EXO UserMailboxes being updated: ";Write-Host -ForegroundColor Green $progressActions

#Loop through the users and add the device IDs
Foreach ($Item in $ExoMobileDevicesHashed)
{
    #Check that we still have a valid EXO PS session
	Test-O365Session
    # Total up the running count 
	$i++	
	# Dump the $ResultsList to CSV at every $statLimit number of objects (defined above); also send status e-mail with some metrics at each dump.
	If (++$c -eq $statLimit)
	{
		$ExoMobileDevicesHashedRunningList | select PrimarySMTPAddress, @{L = "DeviceIDs"; E = { $_.DeviceIDs -join ";"}} | export-csv -path $OutputFile -notypeinformation -Append
        $loopLastTime = $loopCurrentTime
        $loopCurrentTime = Get-Date
        $currentRate = $statLimit/($loopCurrentTime-$loopLastTime).TotalHours
		$avgRate = $i/($loopCurrentTime-$loopStartTime).TotalHours
		# Update Log
		Write-Log ("Counter: $i out of $t objects at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))")
		# Update Screen
		Write-host "Counter: $i out of $t objects at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))" 
		# Clear StatLimit and $ResultsLost for next run
		$c = 0
		$ExoMobileDevicesHashedRunningList.Clear()
	}
	#Adds the Device IDs to the user's CasMailbox ActiveSyncAllowedDeviceIDs attribute
	$DeviceIDs = $item | select @{L = "DeviceIDs"; E = { $_.DeviceIDs -join "; "}}
	Write-Log ("Adding $DeviceIDs to $($Item.PrimarySMTPAddress)")
	Set-CASMailbox -Identity $item.PrimarySMTPAddress -ActiveSyncAllowedDeviceIDs @{add=$($item.DeviceIDs)}
	$Line = "" | select PrimarySMTPAddress,DeviceIDs
	$line.PrimarySMTPAddress = $item.PrimarySMTPAddress
	$line.DeviceIDs = $DeviceIDs
	$ExoMobileDevicesHashedRunningList += $Line

	# Update Progress
	$currentProgress++
}

# Update Elapsed Time
$invokeEndDate = Get-Date
$invokeElapsedTime = $invokeEndDate - $startDate
Write-Host -NoNewline "Elapsed time to add device IDs to users";write-host -ForegroundColor Yellow "$($invokeElapsedTime)"

# Disconnect from EXO and cleanup the PS session
Get-PSSession | Remove-PSSession -Confirm:$false -ErrorAction silentlycontinue

# Create the Output File (report) using the attributes created in the Hashtable by exporting to CSV
# Update Progress
write-progress -id 1 -activity "Creating Output Report" -PercentComplete (95) -Status "$outputFolder"
# Create Report
$ExoMobileDevicesHashedRunningList | select PrimarySMTPAddress, @{L = "DeviceIDs"; E = { $_.DeviceIDs -join ";"}} | export-csv -path $OutputFile -notypeinformation -Append

# Separately capture any PowerShell errors and output to an errorfile
$errfilename = $outputfolder + '\' + $logfilename + "_ERRORs_" + (Get-Date).Ticks + ".txt" 
write-progress -id 1 -activity "Error logging" -PercentComplete (99) -Status "$errfilename"
ForEach ($err in $error) {  
    $logdata = $null 
    $logdata = $err 
    If ($logdata) 
        { 
            out-file -filepath $errfilename -Inputobject $logData -Append 
        } 
}

#Clean Up and Show Completion in session and logs
# Update Progress	
write-progress -id 1 -activity "Complete" -PercentComplete (100) -Status "Success!"
# Update Log
    $endDate = Get-Date
    $elapsedTime = $endDate - $startDate
    Write-Log ("Report started at: " + $($startDate));Write-Log ("Report ended at: " + $($endDate));Write-Log ("Total Elapsed Time: " + $($elapsedTime)); Write-Log ("Adding Device IDs Completed!")
# Update Screen
    write-host;Write-Host -NoNewLine "Report started at    ";write-host -ForegroundColor Yellow "$($startDate)"
    Write-Host -NoNewLine "Report ended at      ";write-host -ForegroundColor Yellow "$($endDate)"
    Write-Host -NoNewLine "Total Elapsed Time:   ";write-host -ForegroundColor Yellow "$($elapsedTime)"
    Write-host "-------------------------------------------------";write-host -foregroundcolor Cyan "Adding Device IDs Completed!";write-host;write-host -foregroundcolor Green "...The $OutputFile CSV and execution logs were created in $outputFolder";write-host;write-host;sleep 1

#------------------------^
#End CUSTOM CODE Section
#------------------------^
#End of Script
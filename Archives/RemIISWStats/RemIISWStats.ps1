<#
Script Name:	RemIISWStats.ps1
Author:			Fabian Caignard
Date:			19/12/2018
Version:		3.0
-----------------------------------------------------------
Updates:	v1.0:	Initial coding: Reuse and adjustment of IISWebStats.ps1 script for a remote
					use.
			v2.0	Add Comments Based Help to enable users to use Get-Help cmdlet on the script.
					Bug fix for computer name in output file header: local computer name is now replaced
					by remote server name.
			v3.0	Bug: When requested fields are missing, script used 1st data column then duplicates
					intermediate and final deletions were not efficients, and then script run for long time
					and sometimes display powershell errors. Fix: Check if there is a log Date Header line
					and check if each requested fields are available in the log files. Log Files which does
					not have all requested fields are then skipped.
#>

<#
.SYNOPSIS

Extract and Concatenate IIS Web or FTP logs on a server over the network.

.DESCRIPTION


Remotely get IIS Web or FTP logs from a server over the network (It can
be used when server does not have Windows Powershell installed).
Concatenation of Ms IIS Web or FTP logs for analysis of IIS usage.
Simplifying URL to only applications root, only by Date, URL, Username
and Client IP address, and removing of duplicates.

This script need some user interactions and can not be run in a simple
command line.

.PARAMETER Server

Hostname of the target Microsoft Internet Information Services server.

.PARAMETER Username

Username to autenticate connection to remote server.
Username can be in [Domain]\[User] or [User]@[Domain] format.

.PARAMETER Password

Password of the user provided in Username parameter to authenticate
connection to remote server.
In some cases, the password may be interpreted by the script, so
ignore Username and Password settings and let the script ask for
them during processing.

.PARAMETER LogType

This parameter is only used when no logs path (LogPath) is provided:
the script will then look for default IIS logs locations:
   C:\inetput\logs\LogFiles
     or
   C:\Windows\system32\LogFiles
Type of log the script should search for (when LogPath parameter is
not provided): W3 for Web IIS log directories (W3SVC*) or FTP for FTP
IIS log directories (MSFTPSVC* or FTPSVC*)

.PARAMETER LogPath

Path of a directory containing Microsoft IIS log files (only *.log files)
or
Path of a log file (no file extension limitations), but content should be
a IIS log file format (see notes).
Path can be the name of the remote share or the local drive letter on the
remote server. The script will automatically translate drive letter into
system hidden share name (ie. C: -> C$).

.PARAMETER SinceLast

Number of logs history to process. By default, the script will process
on all found log files. This parameter let you limit the script to only
process logs from 1 to 12 last months. It always get logs from 1st day of
the last define months (by running script on 15th December and with this
parameter defined to 3, the script will process all log files created 
between 1st September and 15th December).

.PARAMETER OutFile

Path of a directory where script will save its results (he script will
automatically generate output file name)
or
Path of an output file name.
This path should be asbolute path.
If this parameter is ignored then script will save its results in its
location and automatically generate the file name.
The Default file name is:
<W3|FTP|Custom>_logs_export-<ALL|L01-L12>-<yyyyMMdd-HHmm_ss>.txt
  Examples:
    W3_logs_export-All-20180115-1055_38.txt
    FTP_logs_export-L06-20180116-0912_19.txt

.INPUTS

None. You cannot pipe objects to RemIISWStats.ps1.

.OUTPUTS

None. RemIISWStats.ps1 does not generate any output.

.NOTES

To be valid, the log file format should be in text, and data should
be separated by spaces. No other format can be accepted.

First, when the script is looking for log files, it read first line
of each log file and, in 50 first characters, looks for:
"#Software: Microsoft Internet Information Services"

If the script does not find this string, it does not consider the
log file as an IIS valid log file.

Secondly, to determine how to find the appropriate content, it read
each line, then it need to find an header line containing list of
fields or columns starting by "#Fields:" and following by columns
titles. This header line should be placed before data line to let
the script analyze logs.

Used titles are "date", "cs-uri-stem", "cs-username", "c-ip" and
"s-port". If all these fields are not in the log file the script
will no be able to analyze.
Theses fields or columns in the file can be in any order: the data
should simply be in the same order than the headers.

.EXAMPLE
.\RemIISWStats.ps1 MYSERVER

Process on ALL default IIS Web log files on MYSERVER remote server with
current Windows PowerShell console user context.

.EXAMPLE
.\RemIISWStats.ps1 MYSERVER -SinceLast 2

Process on last 2 months default IIS Web log files on MYSERVER remote
server with current Windows PowerShell console user context.

.EXAMPLE
.\RemIISWStats.ps1 -Server MYSERVER Contoso\JohnDoe MyPassword FTP

Process on ALL default IIS FTP log files on MYSERVER remote server with
'Contoso\JohnDoe' user credentials.

.EXAMPLE
.\RemIISWStats.ps1 MYSERVER C:\Temp\LogFiles

Process on ALL *.log files in C:\Temp\Logfiles directory on MYSERVER
remote server with current Windows PowerShell console user context.


.EXAMPLE
.\RemIISWStats.ps1 MYSERVER C$\Temp\LogFiles\20181225.txt

Process on C:\Temp\LogFiles\20181225.txt log file only on MYSERVER
remote server with current Windows PowerShell console user context.

#>


#region Script Input Parameters
[CmdletBinding(DefaultParametersetName='SDefaultLogs')]
param (
	[Parameter(Mandatory=$true,Position=0)]
	[Alias('Computer','Srv')]
	[ValidateScript({if($_ -notmatch "\\\\\w+") {$true} else {Throw "Bad server or computer name format! Simply provide Hostname or FQDN without any backslash."}})]
	[string]$Server,
	[Parameter(ParameterSetName="ADefaultLogs",Mandatory=$true,Position=1)]
	[Parameter(ParameterSetName="ACustomLogs",Mandatory=$true,Position=1)]
	[Alias('User','Usr')]
	[string]$Username,
	[Parameter(ParameterSetName="ADefaultLogs",Mandatory=$true,Position=2)]
	[Parameter(ParameterSetName="ACustomLogs",Mandatory=$true,Position=2)]
	[Alias('Pwd','Passwd','Pass')]
	[string]$Password,
	[Parameter(ParameterSetName="SDefaultLogs",Mandatory=$false,Position=1)]
	[Parameter(ParameterSetName="ADefaultLogs",Mandatory=$false,Position=3)]
	[Alias('Type')]
	[ValidateSet("W3","www","Web","FTP")]
	[string]$LogType="W3",
	[Parameter(ParameterSetName="SCustomLogs",Mandatory=$true,Position=1)]
	[Parameter(ParameterSetName="ACustomLogs",Mandatory=$true,Position=3)]
	[Alias('Log')]
	[ValidateScript({if(($_ -match "[^\\]([a-zA-Z$]){1,}\\\w+") -or ($_ -match "(.){1}:\\\w+")) {$true} else {Throw "Bad Log path! This parameter can only start by local drive letter from the remote server or sharename (accept hidden shares with ""$"")."}})]
	[string]$LogPath,
	[Parameter(Mandatory=$false)]
	[Alias('Since','Last')]
	[ValidateRange(1,12)]
	[int]$SinceLast=0,
	[Parameter(Mandatory=$false)]
	[Alias('Output','Out','File')]
	[ValidateScript({If(Test-Path (Split-Path -Path $_)){$True} else {Throw "The OutFile target does not exists! Please provide a path of an existing directory or of a file in an existing directory."}})]
	[string]$OutFile
	)
#endregion
#--------------------------------------------------------------------
#region Script Variables
#--------------------------------------------------------------------
$script:OutputFile = $OutFile	#Final Output file path (string)
$script:LFList = @()			#Logs directories list (array)
$script:LogFiles = @()			#Logs files list (array)
$script:SvcType = "W3"			#Service Type (string): this string define what kind of default IIS logs directories to search.
$script:TmpOpened = $false		#Flag to indicate to some script's functions that temporary output file is already opened.
$script:OutOpened = $false		#Flag to indicate to some script's functions that final output file is already opened.
$script:TmpData = @()			#Buffuring array of data to write to temporary output file.
#Default Choices Yes and No options for script's prompts
$script:yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Yes', 'Allows script to continue'
$script:no = New-Object System.Management.Automation.Host.ChoiceDescription '&No', 'Does not allow to continue'
#endregion
#--------------------------------------------------------------------
#region Principal script function
<#
The "ScriptRun" function check script's parameters and call other script's functions
#>
function ScriptRun {
	#Test connection to the remote server provided in parameters, then if not reachable dispplay a warning message and exit script.
	if (!(Test-Connection $Server -Count 1 -ErrorAction SilentlyContinue)){
		Write-Warning "The ""$Server"" remote server was not found!"
		Exit
	} else { write-host "remote server OK!"}
	#If username is provided in script parameters, then build Credentials object.
	if(($Username -ne "") -and ($Username -ne $null)) {
		#Convert clear text password to a secure string
		$SecPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
		#Build Credentials object with username provided to the script and password secure string
		$RemCred = New-Object System.Management.Automation.PSCredential($Username,$SecPassword)
		#Check credentials, if failed, then exit script.
		if(!(Test-Cred $RemCred)) {Exit}
	#else (no user and password provided), create empty credential object
	} else {[System.Management.Automation.PSCredential]$RemCred = [System.Management.Automation.PSCredential]::Empty}
	#By default type of log are fixed to W3 (IIS Web), if LogType parameter is passed to "FTP", then redefine it to FTP.
	If ($LogType -eq "FTP") {$script:SvcType = "*$LogType"}
	#Run ChkOutFile function to determine the script output file (or exit script)
	ChkOutFile
	#If no log path was provided in script parameter, then look for default IIS log directories
	if(($LogPath -eq "") -or ($LogPath -eq $null)){
		#Connection to remote server
		OpenRemoteSrv "C:" $RemCred
		Write-Warning "No logs path provided : Looking for default IIS logs..."
		Write-Host "Looking for ""C:\inetpub\logs\LogFiles"" on $Server..."
		ChkLogFolders "RDrive:\inetpub\logs" -DefaultChk
		Write-Host "Looking for ""C:\Windows\system32\LogFiles"" on $Server..."
		ChkLogFolders "RDrive:\Windows\system32" -DefaultChk
#----------------- Only for testing the script -----------------------
#		Write-Host "Looking for ""C:\Temp\LogFiles"" on $Server..."
#		ChkLogFolders "RDrive:\Temp" -DefaultChk
#---------------------------------------------------------------------
	#Else (A log path was provided to the script), then...
	} else {
		#Split Log Path to connect to remote server share root
		#Connection to remote server
		OpenRemoteSrv $LogPath $RemCred
		#Define remote Path from provided path.
		if ($LogPath.IndexOf("\") -ne -1) {[string]$TargetPath = "RDrive:{0}" -f $LogPath.Substring($LogPath.IndexOf("\"))} else {[string]$TargetPath = "RDrive:"}
		#Check if path exists, then get type of target (file or directory).
		if (Test-Path $TargetPath) {
			#If target is a file, then run checking file function "ChkLogFile", else run checking folder function "ChkLogFolders"
			if (Test-Path $TargetPath -Type Leaf) {ChkLogFile $TargetPath} else {ChkLogFolders $TargetPath}
		} else {
			#Else (Target does not exists), then display warning message, unmount remote drive then exit script
			Write-Warning """$LogPath"" (on $Server server) not found or is not a valid path!  Please provide valid a folder path or a log file path."
			Remove-PSDrive RDrive
			Exit
		}
	}
	#If Log directories list is not empty, then for each directory in the list...
	if ($script:LFList.length -gt 0) {
		foreach ($logdir in $script:LFList) {
			#If $SinceLast script parameter (filter logs on only x last months), then...
			If ($SinceLast -gt 0) {
				#Define variable of 1st day of the month x months earlier (x = $SinceLast parameter value)
				[datetime]$filterdate = Get-Date (Get-Date).AddMonths(-$SinceLast) -Day 1 -Hour 0 -Minute 0 -Second 0
				#Run checking file function for each file *.log since x last month in the directory.
				Write-Host "------------------------------------------------------------------`nCheck all log files created between $filterdate and Now..."
				Get-ChildItem $logdir\*.log | ? {$_.CreationTime -gt $filterdate} | foreach {ChkLogFile $_.FullName}
			} else {
				#Run checking file function for each file *.log in the directory.
				Write-Host "------------------------------------------------------------------`nCheck all log files (no creation date limit)..."
				get-ChildItem $logdir\*.log | foreach {ChkLogFile $_.Fullname}
			}
		}
	}
	#If list of found log files is not empty, then...
	If ($script:LogFiles.Length -gt 0) {
		Write-Host "------------------------------------------------------------------"
		#Create Prompting object to display amount of log files found, then ask to continue.
		$details = New-Object System.Management.Automation.Host.ChoiceDescription 'See &details', 'Display list of IIS log files found'
		$ContOpt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $details)
		if ($script:LogFiles.Length -gt 1) {$amountintro = "There are ";$amountfil = "s"} else {$amountintro = "There is ";$amountfil=""}
		if ($script:LFList.Length -gt 1) {$amountdir = "ies"} else {$amountdir = "y"}
		$resultMsg = "{0}{1} IIS log file{2} in the process list." -f $amountintro, $script:LogFiles.Length, $amountfil
		$answer = $host.ui.PromptForChoice($resultMsg, "Do you want to proceed?", $ContOpt, 0)
		Switch ($answer) {
			#If No (do not proceed), then exit script
			1 {
				Write-Host "Script cancel"
				Remove-PSDrive RDrive
				Exit
			}
			#If See details, then display list of IIS log files found.
			2 {
				$resultMsg = "List of found log files:`n"
				$script:LogFiles | foreach {$resultMsg += "$_`n"}
				$ContOpt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
				#Ask if user want to continue and then proceed.
				$answerbis = $host.ui.PromptForChoice($resultMsg, "Do you want to proceed?", $ContOpt, 0)
				#If No, then exit script, else proceed.
				if ($answerbis -eq 1) {
					Write-Host "Script canceled"
					Remove-PSDrive RDrive
					Exit
				}
			}
		}
	} else {
		#Else (no log file in the list) exit script.
		Write-Warning "No valid log file was found!"
		Remove-PSDrive RDrive
		Exit
	}
	Write-Host "------------------------------------------------------------------`nProceed..."
	#For each log file in the list, run ReadLogFile function (to proceed with)
#	$LogFiles | foreach {ReadLogFile $_}
	For ($NbLF = 0; $NbLF -lt $LogFiles.Length; $NbLF++) {
		ReadLogFile $LogFiles[$NbLF] $NbLF
	}
	#Unmount remote drive
	Remove-PSDrive RDrive
	#If Output file is open, then close it.
	If ($TmpOpened) {$script:MyOutstream.Close()}
	#Delete double entries
	WriteToOutput
}
#endregion
#--------------------------------------------------------------------
#region ChkOutFile
<#
The "ChkOutFile" function check if passed output file name exists or not, if no file is passed to the script
the function will define a default output file name in the script directory.
#>
function ChkOutFile {
	#Set function flag to $false
	$OutIsDir = $false
	#Get script location and genrerate a timestamp
	$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
	$fildate = [DateTime]::Now.ToString("yyyyMMdd-HHmm_ss")
	#Create Prompting object to ask if overwrite export file then to continue.
	$ContOpt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
#	$msg = "OutputFile variable is not:`n`tEmpty: {0}`n`tNull: {1}" -f ($OutputFile -ne ""),($OutputFile -ne $null)
#	write-host $msg
	#If output path was provided in script's parameters, then test if already exists.
	if (($OutputFile -ne "") -and ($OutputFile -ne $null)) {
		if(Test-Path $OutputFile) {
			#Test if provided output path is a file
			if (Test-Path $OutputFile -Type Leaf) {
				#Alert user and ask to continue (then overwrite file) or not.
				write-warning """$OutputFile"" already exists !"
				$COFAE = $host.ui.PromptForChoice("If you continue, the script will overwrite the content of this file.", "Do you want to overwrite file?", $ContOpt, 1)
				#If user don't want to overwrite the file, then prupose him to use default outputfile.
				if ($COFAE -eq 1) {
					$CUDOF = $host.ui.PromptForChoice("This script can save its results in a default output file.", "Do you want to proceed?", $ContOpt, 0)
					#If user don't want to use default output file, then exit script.
					if ($CUDOF -eq 0) {
						$script:OutputFile = ""
					} else {Exit}
				}
			} else {$OutIsDir = $true} #Else (provided path exists but is a directory)...
		}
	}
	#Build Date scope for the output file name (how many last months of logs history).
	If ($SinceLast -gt 0) {$datescope = "L{0:00}" -f $SinceLast} else {$datescope = "All"}
	#If no output file was provided in script's parameters (or if user previously accept to use default file) or If provided output path is a directory, then...
	if (($OutputFile -eq "") -or ($OutputFile -eq $null) -or ($OutIsDir)) {
		#If flag is to $true (means the provided output path is a directory), then...
		if ($OutIsDir) {
			#Define Type of logs to "Custom"
			$script:MyST = "Custom"
			#Set ouput file directory to path provided in script's parameters
			$MyOutDir = $OutputFile.Trim("\")
		} else {
			#Define Type of logs for Output file name
			if ($script:SvcType -eq "W3") {$script:MyST = $SvcType} else {$script:MyST = "FTP"}
			#Set output file directory to the script directory
			$MyOutDir = $ScriptDir
		}
		#Build default output filename and path ([ScriptPath]\[Type of Log]_logs_export-[Timestamp].txt)
		$script:OutputFile = "{0}\{1}_logs_export-{3}-{2}.txt" -f $MyOutDir, $MyST, $fildate, $datescope
		#Inform user about outputfile name and location
		Write-Warning "No export file path provided! Results of the script will be saved in`n   --> $OutputFile <--"
	}
	$script:TmpFile = "{0}\~{1}_logs_export-{3}-{2}.tmp" -f $ScriptDir, $MyST, $fildate, $datescope
}
#endregion
#--------------------------------------------------------------------
#region ChkLogFolders function
<#
The "ChkLogFolders" function check the logs folder passed in its parameters.
$MyLogPath refer to the Path to check
$DefaultChk is a switch parameter to let the function looking for a subfolder named "LogFiles"
and if it contains subfolders whare name starts by "W3SVC" or "FTPSVC".
#>
function ChkLogFolders([string]$MyLogPath,[switch]$DefaultChk) {
#Create objects for Host prompting choices (Yes or No)
	$ContOpt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
#If the Path of the folder received in parameter exists, then get the target type (file or directory).
	if (Test-Path $MyLogPath) {
#If $DefaultChk switch parameter is enabled, then...
		if ($DefaultChk) {
			$searchdirstr = "{0}\LogFiles\{1}SVC*" -f $MyLogPath, $script:SvcType
			Write-Host "`tLooking for $searchdirstr subdirectories..."
#Look for \LogFiles\W3SVC* or \LogFiles\*FTPSVC* subfolder.
			$fnddirs = @(Get-ChildItem $searchdirstr)
			$foundFolders = @($fnddirs | ?{$_.PSIsContainer})
			switch ($foundFolders.Count)	{
#If no subfolder found, simply display a warning message
				0 {
					Write-Warning "No IIS logs directory found in ""$MyLogPath\LogFiles""!"
				}
#If only 1 subfolder is found ask for confirmation to work on this subfolder
				1 {
					$FFolderName = $foundFolders[0].Name
					$answer = $host.ui.PromptForChoice("""$FFolderName"" IIS logs directory found in ""$MyLogPath\LogFiles""!", "Do you want to continue with this folder?", $ContOpt, 0)
#If Yes, Add the subfolder in the list
					if ($answer -eq 0) {$script:LFList += "$MyLogPath\LogFiles\$FFolderName"}
				}
#If there are several correspondinf subfolders, then ask for selection of 1 or all directories in the list
				{$_ -gt 1} {
					Write-Host "------------------------------------------------------------------`nSeveral IIS logs directories found in ""$MyLogPath\LogFiles""!"
					$OptList = @()
					$tmpFList = @()
					$i = 0
#Create a choice object and message string to prompt a folder selection.
					$foundFolders.Name | foreach {
						$i++;$tmpID = ('{0:X}'-f $i).toString()
						$OptList += New-Object System.Management.Automation.Host.ChoiceDescription $tmpID, $_
						$FldList = "$FldList$tmpID`t-`t$_`n"
						$tmpFList += "$MyLogPath\LogFiles\$_"
					}
					$OptList += New-Object System.Management.Automation.Host.ChoiceDescription 'All', "All listed folders"
					$defOpt = $OptList.Length -1
#Prompt for selection of 1 of the found logs directories (or all of them)
					$FolderSel = $host.ui.PromptForChoice("Please select with which logs directory the script should proceed:", $FldList , $OptList, $defOpt)
					if ($FolderSel -gt ($tmpFList.Length -1)) {
						$tmpFList | foreach {ChkLogFolders $_}
					} else {
						ChkLogFolders $tmpFList[$FolderSel]
					}
				}
			}
		} else {
#Else (if $DefaultChk is not enabled), then...
			#Check if logs folder contains log, then add folder in Folder list
			$chkIfLogs = (Get-ChildItem $MyLogPath\*.log).Count
			If ($chkIfLogs -gt 0) {
				$script:LFList += $MyLogPath
			} else { Write-Warning """$MyLogPath"" does not contains any proper log file!" }
		}
#Else (if path was not found), then display a warning message.
	} else { Write-Warning """$MyLogPath"" not found or is not a valid path!" }
}
#endregion
#--------------------------------------------------------------------
#region ChkLogFile function
<#
The "ChkLogFile" function check if the log file passed in its parameter starts by "#Software: Microsoft Internet Information Services", then add the file path in the list of log files.
#>
function ChkLogFile([string]$MyLogFile) {
	#$file = new-object System.IO.StreamReader($MyLogFile)
	#$tmplogfirstline = $file.ReadLine()
	#$file.close()
	$tmplogfirstline = Get-Content $MyLogFile -TotalCount 1
	if (($tmplogfirstline.Length -gt 50) -and ($tmplogfirstline.Substring(0,50) -eq "#Software: Microsoft Internet Information Services")) {
		$script:LogFiles += $MyLogFile
	} else {
		Write-Warning """$MyLogFile"": KO - This is not an IIS log file!"
	}
}
#endregion
#--------------------------------------------------------------------
#region ReadLogFile function
function ReadLogFile ([string]$MyLogFile,[int]$MyLogID) {
	$progression = "[{0}/{1}]" -f ($MyLogID + 1),$LogFiles.Length
	Write-Host "$progression Reading and analyzing ""$MyLogFile""..."
	#Set columns ID mapping variables
	$FIDdate = -1
	$FIDuri = -1
	$FIDuser = -1
	$FIDip = -1
	$FIDport = -1
	#Set/Reset Date data Field to empty string
	$FieldDate = ""
	#Read each line of the log file passed in parameter of the function
	$alllines = Get-Content $MyLogFile
	#foreach ($line in [System.IO.File]::ReadLines($MyLogFile)) {
	foreach ($line in $alllines) {
		#Set default value for User and client IP fields in case of not available.
		$FieldIP = "-"
		$FieldUser = "-"
		#Split line content in a temporary array
		$tmparr = $line.Split(" ")
		#If line starts by "#" (it is an Header line), then ...
		if ($line.Substring(0,1) -eq "#") {
			#If line starts by "#Fields:" (it is the header line which describes columns content of following lines), then...
			if ($line.Substring(0,8) -eq "#Fields:") {
				#Set check fileds flag to $true
				$chkFields = $true
				#Get size of the array
				$tmpnbcol = $tmparr.Length
				#For each item in the array, check if it match with "date", "cs-uri-stem", "cs-username", or "c-ip"
				for ($i = 0 ; $i -lt $tmpnbcol ; $i++) {
					#Depending of matching, store ID of the column is the corresponding mapping variable.
					Switch ($tmparr[$i]) {
						"date" {$FIDdate = $i - 1}
						"cs-uri-stem" {$FIDuri = $i - 1}
						"cs-username" {$FIDuser = $i - 1}
						"c-ip" {$FIDip = $i - 1}
						"s-port" {$FIDport = $i -1}
					}
				}
				#If 1 of the requested fields is not found, then...
				if(($FIDip -eq -1) -or ($FIDuri -eq -1) -or ($FIDuser -eq -1)) {
					#Set missing fields list variable
					$missFields = ""
					#If URL column ID equals to -1, then add URL field name in the list of missing fields and set check flag to $false (because field is madatory)
					if ($FIDuri -eq -1) {$missFields += "`n - cs-uri-stem";$chkFields = $false}
					#If User column ID equals to -1, then add User field name in the list of missing fields
					if ($FIDuser -eq -1) {$missFields += "`n - c-username"}
					#If client IP column ID equals to -1, then add client IP field name in the list of missing fields
					if ($FIDip -eq -1) {$missFields += "`n - c-ip"}
					#If Client IP AND User are missing then set chek flag to $false because at least 1 of these both field should be available in the log file
					if (($FIDuser + $FIDip) -eq -2) {$chkFields = $false}
					#Display message with missing fields
					Write-Host "Requested fields are missing in the log file! Missing fields are:$missFields"
					#If one of the mandatory requested field is missing then display a warning message the skip the current log file by exiting the function
					If (!$chkFields) {
						Write-Warning "Because one of the missing field is mandatory this log file is skipped!"
						Return
					}
				}
			#Else if line starts by "#Date:" (it is the header line which provide logs date), then...
			} elseif ($line.Substring(0,6) -eq "#Date:") {
				#Set the data field Date with header date value
				$FieldDate = $tmparr[1]
			} #Because no "Else" statement, other header lines than columns headers are ignored.
		} else {
		#Else (it is a line of data), then...
			#If Date field column ID is not equal to -1 (means columns exists), then set Date data field value.
			if ($FIDdate -ne -1) {$FieldDate = $tmparr[$FIDdate]} else {
				#else If Date field is empty, then display a warning message the skip the current log file by exiting the function
				if ($FieldDate -eq "") {
					Write-Warning "Because the date field is mandatory and is missing this log file is skipped!"
					Return
				}
			}
			if ($MyST -eq "FTP"){$tmpUrlStr = "-FTP-"} else {$tmpUrlStr = $tmparr[$FIDuri]}
			#If Url in log file is not equal to "-" and to "***", then... (to exclude non efficient entries)
			if (($tmpUrlStr -ne "-") -and ($tmpUrlStr -ne "***")) {
				#If the URL contains more than 1 occurence of "/", then extract only content before second "/" (to only keep root web site browsed by a user)
				if ($tmpUrlStr.IndexOfAny("/",1) -gt 0) {$tmpUrlStr = $tmpUrlStr.Substring(0,$tmpUrlStr.IndexOfAny("/",1))}
				#If Server port field exists, then...
				If ($FIDport -ne -1) {
					#If port equals to 21, then replace URL by "-FTP-" (to simplify reading on data: FTP details are not efficient here)
					If ($tmparr[$FIDport] -eq "21") {$tmpUrlStr = "-FTP-"}
				}
				if ($FIDuser -ne -1) {$FieldUser = $tmparr[$FIDuser]}
				if ($FIDip -ne -1) {$FieldIP = $tmparr[$FIDip]}
				#Build data string to be write in file.
				$MyArrData = $FieldDate,$tmpUrlStr,$FieldUser,$FieldIP
				#Run WriteToTmp function to send the data string in a temporary output file.
				WriteToTmp $MyArrData
			}
		}
	}
}
#endregion
#--------------------------------------------------------------------
#region WriteToTmp function
function WriteToTmp($MyArrData) {
	#If Output file is not yet opened, then...
	if (!($TmpOpened)) {
		#Open StreamWriter on the temporary file
		$script:MyOutstream = New-Object System.IO.StreamWriter($TmpFile)
		#Write a file Header
		$myTmpHeader = "#IIS {0} logs exports on {1}" -f $MyST, $Server.ToUpper()
		$MyOutstream.WriteLine($myTmpHeader)
		#Write columns headers line in the file
		$MyOutstream.WriteLine("Date`tURL`tUsername`tClient_IP")
		#Set Output file opening flag to yes
		$script:TmpOpened = $true
	}
	#Build data line for output file
	$MyDataLine = "{0}`t{1}`t{2}`t{3}" -f $MyArrData[0],$MyArrData[1],$MyArrData[2],$MyArrData[3]
	#Check size of Temporary data Buffer
	Switch ($TmpData.Count) {
		#If less than 30 data lines in Buffer, then Add new data line
		{$_ -lt 30}	{
			$script:TmpData += $MyDataLine
		}
		#If there are 30 data lines in buffer, then...
		default {
			#Add current data line in buffer
			$script:TmpData += $MyDataLine
			$ctdb = $script:TmpData.Count
			#Delete duplicates in buffer array
			$script:TmpData = $script:TmpData | Select-Object -Unique
			$script:ctda += $ctdb - $script:TmpData.Count
			#Write buffer data lines in the output file
			$script:TmpData | foreach {$MyOutstream.WriteLine("$_")}
			#wipe buffer data lines
			$script:TmpData = @()
		}
	}
	#$MyOutstream.WriteLine($MyDataLine)
	#Write-Host $MyDataLine
}
#endregion
#--------------------------------------------------------------------
#region WriteToOutput function
function WriteToOutput {
	Write-Host "Read temporary data file..."
	#Read content of temporary file
	$tmpresult = [System.IO.File]::ReadAllLines($TmpFile)
	$ctritem = $tmpresult.Count + $script:ctda
	Write-Host "There were $ctritem data lines in temp file!"
	#Delete duplicates from temporary file
	Write-Host "Delete duplicates..."
	$outresult = $tmpresult | Select-Object -Unique
	$coritem = $outresult.Count
	Write-Host "There are $coritem data lines in total!"
	#Open StreamWriter on the output file
	Write-Host "Write data to Output file..."
	$script:MyOutstream = New-Object System.IO.StreamWriter($OutputFile)
	#Save dataset to output file
	$outresult | foreach{$MyOutstream.WriteLine("$_")}
	$MyOutstream.Close()
	#Delete temporary file
	Remove-Item -Path $TmpFile -Force
	Write-Host "------------------------------------------------------------------`n`nOperation done!`n`nPlease check content of following output file:`n`n$OutputFile`n"
}
#endregion
#--------------------------------------------------------------------
#region Test-Cred function
function Test-Cred ([System.Management.Automation.PSCredential]$Credentials){
	$Domain = $null
	$Root = $null
	$tmpUsr = $null
	$tmpPwd = $null
	#Try to authenticate with credentials provided to the function
	Try {
		# Split username and password
		$tmpUsr = $Credentials.username
		$tmpPwd = $Credentials.GetNetworkCredential().password
		# Get Domain
		$Root = "LDAP://" + ([ADSI]'').distinguishedName
		$Domain = New-Object System.DirectoryServices.DirectoryEntry($Root,$tmpUsr,$tmpPwd)
	}
	#Catch any error, then continue
	Catch {
		$_.Exception.Message
		Continue
	}
	#If Domain variable is $null or get an error, then display a warning message and return false
	If(!$domain) {
		Write-Warning "An error occured when trying to authenticate you!"
		return $false
	#Else, Domain is not $null or didn't get error, then...
	} Else {
		#If Domain name is not $null (means authentication success), then retrun $true, else display a warning message then return $false
		If ($domain.name -ne $null) {
			return $true
		} Else {
			Write-Warning "Authentication failed!"
			return $false
		}
	}
}
#endregion
#--------------------------------------------------------------------
#region OpenRemoteSrv funtion
function OpenRemoteSrv ([string]$Path,[System.Management.Automation.PSCredential]$Cred=[System.Management.Automation.PSCredential]::Empty) {
	#Set succes flag to $false
	$ConSuccess = $false
	#Define Drive/share root from provided path : extract first characters before first occurence of "\".
	if ($Path.IndexOf("\") -ne -1) {[string]$Drive = $Path.Substring(0,$Path.IndexOf("\"))} else {$Drive = $Path}
	#Get position of ":" in $Drive variable, then depend of position...
	Switch($Drive.IndexOf(":")) {
		#If 1 (then second character in the string), the its is a drive letter like "C:", so get drive letter only and suffix it by "$".
		1 {
			[string]$Drive = "{0}$" -f $Drive.Substring(0,1)
		}
		-1 {
			#If -1 (then no ":" found in the string), do nothing.
			#Need to trap this to prevent next (default) statement
		}
		#If any other position of ":" in the string, then it is not a valid drive letter, so dispaly a warning message then return $false.
		default {
			Write-Warning "Bad Log Path! Drive root of log path should be either in format ""[Drive]:"", either in format ""[ShareRoot][$]""."
			Write-Host "Unable to connect on remote server, then exit script !"
			Exit
		}
	}
	#Set New Powershell drive path
	$PSDRoot = "\\{0}\{1}" -f $Server, $Drive
	#If mount of PSDrive fails (Means that credentials are not allowed to connect on the target server or the target share), then...
	if (!(New-PSDrive -Name RDrive -PSProvider FileSystem -Root $PSDRoot -Credential $Cred -Scope Script -ErrorAction SilentlyContinue)) {
		#Display warning message then ask to user if he want to provide other credentials.
		Write-Warning "Script can't connect to remote computer with provided credentials!"
		$ContOpt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
		$answer = $host.ui.PromptForChoice("The current credentials does not have permissions to connect on $Server server.", "Do you want to provide other credentials?", $ContOpt, 1)
		#If User answer Yes (then he want to provide other credentials), then...
		if ($answer -eq 0) {
			#Ask for username and password by using Read-Host instead of Get-Credential cmdlet
			#(to prevent any Ms Windows style interaction and stay in text PS Console)
			$tmpUsr = Read-Host "Username"
			$tmpPwd = Read-Host "Password" -AsSecureString
			#Build Credentials object with username provided to the script and password secure string
			$tmpCred = New-Object System.Management.Automation.PSCredential($tmpUsr,$tmpPwd)
			#If retry of PSDrive mouting with new credentials also fails, then set flag to $true
			if (New-PSDrive -Name RDrive -PSProvider FileSystem -Root $PSDRoot -Credential $tmpCred -Scope Script -ErrorAction SilentlyContinue) {$ConSuccess = $true}
		}
	#Else (the mount of PsDrive is successfull), then set flag to $true.
	} else {$ConSuccess = $true}
	#If Success flag equals to $false, then display message then exit script.
	if (!$ConSuccess) {
		Write-Host "Unable to connect on remote server, then exit script !"
		Exit
	}
}
#endregion
#--------------------------------------------------------------------
#region Script Running command
#--------------------------------------------------------------------
# Here is the command to start the script !
#--------------------------------------------------------------------
ScriptRun #Calling the main script's function
#--------------------------------------------------------------------
#endregion
#--------------------------------------------------------------------
#--------------------------------------------------------------------
#--------------------------------------------------------------------
#region SandBox code

#Software: Microsoft Internet Information Services

#Fields: date time s-ip cs-method cs-uri-stem cs-uri-query s-port cs-username c-ip cs(User-Agent) sc-status sc-substatus sc-win32-status time-taken
#Fields: date time s-sitename s-ip cs-method cs-uri-stem cs-uri-query s-port cs-username c-ip cs(User-Agent) sc-status sc-substatus sc-win32-status 
#Fields: date time c-ip cs-username s-ip s-port cs-method cs-uri-stem sc-status sc-win32-status sc-substatus x-session x-fullpath

#date cs-uri-stem c-ip cs-username 

#endregion
#--------------------------------------------------------------------
#--------------------------------------------------------------------
#--------------------------------------------------------------------


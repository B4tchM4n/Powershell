<#
Library Name:	_02-SharedFunctions.ps1
Authors:		Fabian Caignard
Date:			March 7th 2019
Version:		1.0
Description:	This Powershell library contains all functions and script blocks which
                can be used several times by IISLogsAnalyser.ps1 script (and which can
                be reuse in other scripts).
#>

#-------------------------------------------------------------------------------
#region: Get-RemoteDirectories function (get log directories for computers corresponding to provided PSDrives, then load files lists in LogDirectory objects)
function Get-RemoteDirectories ([string[]]$drives) {
    # Show Application Loading form
    Show-AppLoading
    # Getting default log directories on local system drive (then for local computer), then for each found...
    $drives | ForEach-Object {
        # Update loading message
        $MsgDisplay_Lbl.Text = "Loading"
        $MsgDisplay_Lbl.Refresh()
        # set counter to 0
        $localldc = 0
        # Get Computer ID from PSDrive name
        [int]$ID = $_.Substring(6)
        [string]$drvpath = $_ + ':\'
        # Get Log directories corresponding to PSDrive, then for each found log directory...
        Get-LogDirectories -Path $drvpath | ForEach-Object {
            # Add the found log directory to LogDirectory objects collection with computer ID
            $script:LogDirectoryList.Add([LogDirectory]::New($_, $ID))
            # Increase counter
            $localldc++
        }
        # Update loading message
        $MsgDisplay_Lbl.Text += "."
        $MsgDisplay_Lbl.Refresh()
        # Get Computer name corresponding to ID
        $ComputerName = ($script:ComputerList | Where-Object -Property ID -EQ $ID).Name
        # If no log directory found (counter equals 0), Display a warning message box, else...
        if ($localldc -eq 0) { Show-Warning "The $ComputerName computer seems to not be an IIS server!" "No IIS logs found" } else {
            # For each LogDirectory Object corresponding to current PSDrive ID (in fact computer ID) in the collection...
            $script:LogDirectoryList | Where-Object -Property ServerID -EQ $ID | ForEach-Object {
                # Update loading message
                $MsgDisplay_Lbl.Text = "Please wait"
                $MsgDisplay_Lbl.Refresh()
                # Get log files list by calling GetFiles method of LogDirectory Object.
                $_.GetFiles()
                $_
            }
        }
    }
    # Closing Application loading form.
    $AppLoading_Frm.Close()
}
#endregion
#-------------------------------------------------------------------------------
#region: Get-LogDirectories function (search for default log directories in provided path)
function Get-LogDirectories {
    param (
        # Specifies a path to one or more locations.
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Path to one or more locations.")]
        [Alias("PSPath")]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Path
    )
    # Allow processing each input object from Pipeline
    process {
        # For each path provided to Path parameter...
        foreach ($item in $Path) {
            # Test if path exists and is a directory. If True then...
            if (Test-Path -Path $item -PathType Container) {
                # Remove ending "\" (if present) from path and store in temporary variable
                $tmppath = $item.TrimEnd('\')
                $tmpout = @()
                # Look for W3SVC* and *FTPSVC* directories in path or in its LogFiles sub-directory and add them in temporary variable
                $tmpout += Get-ChildItem -Path "$tmppath\LogFiles\W3SVC*" -Directory -ErrorAction SilentlyContinue
                $tmpout += Get-ChildItem -Path "$tmppath\W3SVC*" -Directory -ErrorAction SilentlyContinue
                $tmpout += Get-ChildItem -Path "$tmppath\LogFiles\*FTPSVC*" -Directory -ErrorAction SilentlyContinue
                $tmpout += Get-ChildItem -Path "$tmppath\*FTPSVC*" -Directory -ErrorAction SilentlyContinue
                # Set an array of strings with default IIS locations in case of provided path is the system drive (C:\).
                [string[]]$defaultPathes = "$tmppath\inetpub\logs", "$tmppath\Windows\System32"
                # For each default IIS location...
                foreach ($defpath in $defaultPathes) {
                    # Look for W3SVC* and *FTPSVC* directories in default location or in its LogFiles sub-directory and add them in temporary variable
                    $tmpout += Get-ChildItem -Path "$defpath\LogFiles\W3SVC*" -Directory -ErrorAction SilentlyContinue
                    $tmpout += Get-ChildItem -Path "$defpath\W3SVC*" -Directory -ErrorAction SilentlyContinue
                    $tmpout += Get-ChildItem -Path "$defpath\LogFiles\*FTPSVC*" -Directory -ErrorAction SilentlyContinue
                    $tmpout += Get-ChildItem -Path "$defpath\*FTPSVC*" -Directory -ErrorAction SilentlyContinue
                }
                # If W3SVC* and/or *FTPSVC* directories where found...
                if ($tmpout.Count) {
                    # Return found directories
                    $tmpout
                }
                # else, simply return directory item corresponding to provided path.
                elseif ((Get-ChildItem "$item\*.log" -File).Count -gt 0) { Get-Item "$item\" }
            }
        } # End of Foreach loop on Path parameter
    } # End of Process Block
}
#endregion
#-------------------------------------------------------------------------------
#region: Test-OutDirSettings function (checks WorkDir and OutputPath variables)
function Test-OutDirSettings {
    if ($script:WorkDir -eq $script:OutputPath) { Remove-Variable -Name WorkDir -Scope Script -Force -ErrorAction SilentlyContinue }
}
#endregion
#-------------------------------------------------------------------------------
#region: Test-Credentials function (test if passed credentials are correct - Thanks to Pawel Janowicz at https://bit.ly/2FmRdnG)
function Test-Credentials ([System.Management.Automation.PSCredential]$Credentials) {
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
        $Domain = New-Object System.DirectoryServices.DirectoryEntry($Root, $tmpUsr, $tmpPwd)
    }
    #Catch any error, then continue
    Catch {
        #$_.Exception.Message
        Continue
    }
    #If Domain variable is $null or get an error, then display a warning message and return false
    If (!$domain) {
        Show-Warning "An error occured when trying to check provided credentials!" "Credentials issue"
        #Write-Warning "An error occured when trying to authenticate you!"
        return $false
        #Else, Domain is not $null or didn't get error, then...
    }
    Else {
        #If Domain name is not $null (means authentication success), then retrun $true, else display a warning message then return $false
        If ($null -ne $domain.name) {
            return $true
        }
        Else {
            Show-Warning "Authentication fialed!`nPlease check the provided user and password and retry." "Bad credentials"
            #Write-Warning "Authentication failed!"
            return $false
        }
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Get-TestedCreds function (use Get-Credential cmdlet, test credentials, then return a credential or null)
function Get-TestedCreds () {
    $tmpCred = $null
    # Display a credentials dialog
    Try { $tmpCred = Get-Credential -Credential "" -ErrorAction Stop } Catch { $tmpCred = $null }
    # If user provide crendentials (do not click on cancel button), then...
    if ($tmpCred) {
        # If test of credentials returns an error, then set credentials to empty one.
        if (!(Test-Credentials $tmpCred)) { $tmpCred = $null }
    }
    return $tmpCred
}
#endregion
#-------------------------------------------------------------------------------
#region: Connect-Computer function (connect to a specified computer, mount a PSDrive and store in Computer object & returns the PSDrive name)
function Connect-Computer {
    param (
        # Specifies one or more computer IDs.
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "One or more computer ID.")]
        [ValidateNotNullOrEmpty()]
        [int[]]
        $ID,
        [Parameter(Mandatory = $false,
            Position = 1,
            HelpMessage = "Credential to use to connect on computer")]
        [System.Management.Automation.PSCredential]
        $Credential
    )
    # Allow processing each input object from Pipeline
    process {
        # For each provided ID...
        foreach ($item in $ID) {
            # Set temporary credentials variable to null
            $tmpCred = $null
            # Get Computer Object from collection corresponding to provided ID
            $tmpCompObj = $script:ComputerList | Where-Object -Property ID -EQ $item
            # If no credential is provided to this function...
            if (!$Credendial) {
                # If authentication status for custom credentials is true, set temporary credentials to computer's custom credentials
                if ($tmpCompObj.Authentications.withCredentials) { $tmpCred = $tmpCompObj.scred }
                # else, if authentication status for current user is true, set temporary credentials to empty credentials
                elseif ($tmpCompObj.Authentications.LocalUser) { $tmpCred = [System.Management.Automation.PSCredential]::Empty }
                # else (credentials were provided to the function), set temporary credentials to provided credentials
            }
            else { $tmpCred = $Credential }
            # If tmpCred variable is set (contain credentials in fact)...
            if ($tmpCred) {
                # Set PSDrive root path ("\\<Computer_Name>\C$")
                $psdroot = '\\' + $tmpCompObj.Name + '\C$'
                # Set PsDrive name ("ILADrv<Computer_ID_on_2_digits>" - I L A for IIS Log Analyser)
                $psdname = 'ILADrv' + "{0:00}" -f $item
                # If new PSDrive is successfully mounted...
                if (New-PSDrive -Name $psdname -PSProvider FileSystem -Root $psdroot -Credential $tmpCred -Scope Script -ErrorAction SilentlyContinue) {
                    # If access test of PSDrive root path failed, Force remove (unmount) the PSDrive
                    if (!(Test-Path $psdroot -ErrorAction SilentlyContinue)) {
                        Get-PSDrive $psdname | Remove-PSDrive -Force
                        # Else return the new PSDrive name
                    }
                    else {
                        $tmpCompObj.PSDrive = $psdname
                        $psdname
                    }
                }
            }
        }
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Get-SelectedLogFiles function (get IIS log files object from LogDirectory objects collection and according to applied dates filters)
function Get-SelectedLogFiles {
    # Allow processing each input object from Pipeline
    process {
        # For each selected LogDirectory in the collection...
        $script:LogDirectoryList.where( { $_.Selected }) | ForEach-Object {
            # Get computer name corresponding to current LogDirectory Object
            $server = $($script:ComputerList | Where-Object -Property ID -EQ $_.ServerID).Name
            # Get log directory name
            $instance = $_.Name
            # If LogDirectory object's filter is not set...
            if ($_.Filter.Oldest -eq "Not set") {
                # For each IIS LogFiles...
                $_.LogFiles.Where( { $_.isIISLog }) | ForEach-Object {
                    # Create new LogFileData object (servername, log directory as instance, and log file object)
                    $tmpLDF = [LogFileData]::new($server, $instance, $_.fileobject)
                    # Return newly created object
                    $tmpLDF
                }
                # Else (a filter is set)...
            }
            else {
                # Get Filtering start and end dates from LogDirectory object
                [datetime]$startFilter = (Get-Date -Date $_.Filter.Oldest)
                [datetime]$endFilter = (Get-Date -Date $_.Filter.Latest)
                # For each IIS log files which are between filtering dates...
                $_.LogFiles.Where( { $_.isIISLog -and $_.CreationDate -ge $startFilter -and $_.CreationDate -le $endFilter }) | ForEach-Object {
                    # Create new LogFileData object (servername, log directory as instance, and log file object)
                    $tmpLDF = [LogFileData]::new($server, $instance, $_.fileobject)
                    # Return newly created object
                    $tmpLDF
                }
            }
        }
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Invoke-LogDataFromFile function (Read data lines form provided LogFileData objects. In case of multiple objects sent,
#        it can update a progressbar and a related label objects if they are provided in parameter and if the total count of
#        sent objects is also provided)
function Invoke-LogDataFromFile {
    param (
        # Specifies one or more IIS Log file data.
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "One or more IIS Log file data.")]
        [ValidateNotNullOrEmpty()]
        [LogFileData[]]
        $LogFileData,
        [Parameter(Mandatory = $false,
            Position = 1)]
        [int]
        $TotalCount,
        [Parameter(Mandatory = $false,
            Position = 2)]
        $ProgressBarObj = $null,
        [Parameter(Mandatory = $false,
            Position = 3)]
        $ProgressTextObj = $null
    )
    begin {
        # If a ProgressBar object is provided in parameters...
        if ($ProgressBarObj) {
            # Depending of amount of log files in the scope of script's analysis...
            switch ($TotalCount * 10) {
                # If there is no log file, exit switch loop
                { $_ -eq 0 } { break }
                # If there are less or equals to 100 logfiles in the scope...
                { $_ -le 100 } {
                    # Set progressbar maximum value to 100
                    $ProgressBarObj.Maximum = 100
                    # Set progressbar step value to 100/Amount of log files in scope
                    $ProgressBarObj.Step = [math]::Truncate(100 / $TotalCount)
                }
                # If there are more than 100 log files in the scope...
                { $_ -gt 100 } {
                    # Set progressbar maximum value to the amount of files
                    $ProgressBarObj.Maximum = $TotalCount
                    # Set progressbar step value to 1
                    $ProgressBarObj.Step = 1
                }
            }
        }
    }
    # Allow processing each input object from Pipeline
    process {
        # For each provided LogFileData...
        :logfilesLbl foreach ($item in $LogFileData) {
            # Check if corresponding log file is not locked by another application, then...
            try {
                $FileStream = [System.IO.File]::Open($item.Logfile, 'Open', 'Write', 'ReadWrite')
                $FileStream.Close()
                $FileStream.Dispose()
                $isLocked = $false
            }
            catch {
                # If check failed, then skip current log file
                $isLocked = $true
                continue logfilesLbl
            }
            if (!$isLocked) {
                # Create StreamReader and open the current log file
                $StreamReader = [System.IO.StreamReader]::new($item.Logfile.FullName)
                # If StreamReader is created...
                if ($StreamReader) {
                    # If a Label object is provided in parameters, set its text
                    if ($ProgressTextObj) { $ProgressTextObj.Text = "Analysis of " + $item.Logfile.FullName; Start-Sleep -Milliseconds 1 }
                    # Read each line (while line is not not null)
                    :loglineLbl while ($null -ne ($line = $StreamReader.ReadLine())) {
                        # Split line string in an array of strings
                        [string[]]$line = $line.Split(' ')
                        # Depending of first value in the line array...
                        switch ($line[0]) {
                            { $_ -eq '#Software:' -or $_ -eq '#Version:' } { continue loglineLbl }
                            # If it contains  "#Date:", set a "CommonDate variable with found date string"
                            '#Date:' { $CommonDate = $line[1] }
                            # If it contains "#Fileds:"...
                            '#Fields:' {
                                # Remove fisrt value of the array (to only keep column headers form log file)
                                $line = $line[1..($line.Count - 1)]
                                # Set counter to 0
                                $i = 0
                                # For each string in the line array...
                                $line | ForEach-Object {
                                    # For Log field mapping item corresponding to current header string, set the ColumnId to counter value
                                    $script:LogFieldsMapping | Where-Object -Property Name -EQ $_ | ForEach-Object { $_.ColumnID = $i }
                                    # Increase counter value
                                    $i++
                                }
                                # If there is any LogFieldsMapping item with no defined column ID, skip current log file
                                if ($script:LogFieldsMapping.Where( { $_.ColumnID -eq -1 -and $_.Mandatory })) { continue logfilesLbl }
                            }
                            # in all other cases
                            Default {
                                # Create empty hashtable for output data line
                                $outline = @{ }
                                # Add server name in the output data line
                                $outline.Add('Server', $item.Server)
                                # Add "instance" form current LogFileData object (in fact log directory name) in the output data line
                                $outline.add('Instance', $item.Instance)
                                $script:LogFieldsMapping | ForEach-Object {
                                    # If the current log field mapping item name equals 'date' and CommonDate string is not empty or null, set temporary data to CommonDate string
                                    if ($_.Name -eq 'date' -and ![string]::IsNullOrEmpty($CommonDate)) {
                                        $tmpData = $CommonDate
                                        # Else if current log field maping item name equals 'cs-uri-stem' (this is the browsed URL in the log file)...
                                    }
                                    elseif ($_.Name -eq 'cs-uri-stem') {
                                        # Set the temporary data to string from corresponding column in the log file data line
                                        $tmpData = $line[$_.ColumnID]
                                        # If "instance" of current LogFileData contains 'FTP' and temporary data does not contain any '/', temporary data equals then to an empty string.
                                        if ($item.Instance -like '*ftp*' -and $tmpData.IndexOf('/') -eq -1) { [string]$tmpData = "" }
                                        # Remove first "/" (if there is) of the string then split by "/" as an array
                                        [string[]]$tmpArrData = $tmpData.TrimStart('/').Split('/')
                                        # If 1st value in the array contains a ".", temporary data equals an empty string, else set temporary data to 1st value of the array as a string
                                        if ($tmpArrData[0] -like '*.*' -or $tmpArrData[0] -eq '-') { [string]$tmpData = "" } else { [string]$tmpData = $tmpArrData[0] }
                                        # Else if columnID is defined (it is not date column AND there is no CommonDate defined: it can then be date)
                                    }
                                    elseif ($_.ColumnID -ne -1) {
                                        # Set the temporary data to string from corresponding column in the log file data line
                                        $tmpData = $line[$_.ColumnID]
                                        # Else (the column ID equals then to -1: so not defined)
                                    }
                                    else {
                                        # Set Temporary data to empty string
                                        $tmpData = ""
                                    }
                                    # Building the output line :
                                    # Add current temporary data value to output
                                    $outline.Add($_.Name, $tmpData)
                                }
                                # Output the data line as a custom object
                                [PSCustomObject]$outline
                            }
                        }
                    }
                    # Free up the log file by closing the Streamreader on it
                    $StreamReader.Close()
                    $StreamReader.Dispose()
                }
            }
            # If a ProgressBar object is provided in parameters, perform step
            if ($ProgressBarObj) { $ProgressBarObj.PerformStep() }
        }
    }
    end {
        # If a Label object is provided in parameters, set its text to empty string
        if ($ProgressTextObj) { $ProgressTextObj.Text = "" }
        # If a ProgressBar object is provided in parameters, set its value to its maximum
        if ($ProgressBarObj) { $ProgressBarObj.Value = $ProgressBarObj.Maximum }
        # Try to free up memory
        [GC]::Collect()
    }
}
#endregion
#-------------------------------------------------------------------------------
#region: Invoke-DedupData function (delete duplicates based on provided filtering criteria in provided datasets and output them by set based on grouping criteria)
function Invoke-DedupData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Collection of datalines.")]
        [Alias("DataLine")]
        [ValidateNotNullOrEmpty()]
        [psobject[]]
        $Data,
        [Parameter(Mandatory = $false,
            Position = 1)]
        [Alias("BufferSize")]
        [ValidateRange(5, 500)]
        [int]
        $Buffer = 100,
        [Parameter(Mandatory = $false,
            Position = 2)]
        [Alias("Property")]
        [string[]]
        $Filter = "*",
        [Parameter(Mandatory = $false,
            Position = 3)]
        [Alias("OutputGroup")]
        [string[]]
        $GroupBy
    )
    begin {
        # If GroupBy parameter is defined...
        if ($GroupBy) {
            # Create an empty hashtable
            $hashgroup = @{ }
            # For each string in the GroupBy parameter, add it as a key without any value in the hashtable
            $GroupBy | ForEach-Object { $hashgroup.Add($_, "") }
            # Convert the hashtable into a custom object named Groups (with each GroupBy item as an object's property)
            $Groups = [PSCustomObject]$hashgroup
            # Delete hashtable variable
            Remove-Variable -Name hashgroup -Force -ErrorAction SilentlyContinue
            # Try to free up memory
            [GC]::Collect()
        }
    }
    # Allow processing each input object from Pipeline
    process {
        # Initialise a counter
        if ($null -eq $i) { $i = 0 }
        # For each data item passed in "Data" collection...
        foreach ($item in $Data) {
            # If GroupBy parameter is defined...
            if ($GroupBy) {
                # For each item in GroupBy...
                $GroupBy | ForEach-Object {
                    # If Groups object's property value (corresponding to current GoupBy parameter item) equals an empty string...
                    if ($Groups.$_ -eq "") {
                        # Set the Groups property to current item corresponding property value (in fact we define the currently working with grouping property value)
                        $Groups.$_ = $item.$_
                        # Else, if Groups object's property value equals to value of current item corresponding property...
                    }
                    elseif ($Groups.$_ -ne $item.$_) {
                        # Set the Groups property to current item corresponding property value (in fact we define the currently working with grouping property value)
                        $Groups.$_ = $item.$_
                        # If output data array exists (to prevent a double - or more - action when more than 1 grouping criteria)...
                        if ($outArrData) {
                            # If temporary data array is not empty (for same reasons than before)...
                            if ($tmpArrData.Count -gt 0) {
                                # Remove duplicates from temporary data array, then store its items in output data array
                                $outArrData += $tmpArrData | Sort-Object -Property $Filter -Unique
                                # Delete tmpArrData variable
                                Remove-Variable -Name tmpArrData -Force -ErrorAction SilentlyContinue
                                # Try to free up memory
                                [GC]::Collect()
                                # Reset $i counter to 0
                                $i = 0
                            }
                            # Remove remaining duplicates in the output data array, then output it
                            , $outArrData | Sort-Object -Property $Filter -Unique
                            # Delete outArrData variable
                            Remove-Variable -Name outArrData -Force -ErrorAction SilentlyContinue
                            # Try to free up memory
                            [GC]::Collect()
                        }
                    }
                }
            }
            # If output data array does not exists, then create it as an empty array
            if (!$outArrData) { $outArrData = @() }
            # If $i counter equals to Buffer value...
            if ($i -eq $Buffer) {
                # Remove duplicates from temporary data array, then store its items in output data array
                $outArrData += $tmpArrData | Sort-Object -Property $Filter -Unique
                # Delete tmpArrData variable
                Remove-Variable -Name tmpArrData -Force -ErrorAction SilentlyContinue
                # Try to free up memory
                [GC]::Collect()
                # Reset $i counter to 0
                $i = 0
            }
            # If temporary data array does not exists, then create it as an empty array
            if (!$tmpArrData) { $tmpArrData = @() }
            # Add current item to temporary data array
            $tmpArrData += $item
            # Increase $i counter
            $i++
        }
    }
    end {
        # If output data array exists...
        if ($outArrData) {
            # If temporary data array is not empty...
            if ($tmpArrData.Count -gt 0) {
                # Remove duplicates from temporary data array, then store its items in output data array
                $outArrData += $tmpArrData | Sort-Object -Property $Filter -Unique
                # Delete tmpArrData variable
                Remove-Variable -Name tmpArrData -Force -ErrorAction SilentlyContinue
            }
            # Remove remaining duplicates in the output data array, then output it
            , $outArrData | Sort-Object -Property $Filter -Unique
            # Delete outArrData variable
            Remove-Variable -Name outArrData -Force -ErrorAction SilentlyContinue
            # Try to free up memory
            [GC]::Collect()
        }
    }
}
#endregion
#-------------------------------------------------------------------------------

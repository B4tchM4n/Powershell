<#
Script Name:    IISLogAnalyser.ps1
Author:         Fabian Caignard
Date:           February 18th 2019
Version:        1.0
-----------------------------------------------------------
Updates:    v1.0:   Initial coding
#>


<#
.SYNOPSIS
    Graphical Interface to Extract and Concatenate IIS logs on a local computer or on remote servers over the network.
.DESCRIPTION
    Get IIS logs (in W3C format) from the local computer or on remote servers over the network (It can be used when servers does not have Windows Powershell installed).
    Concatenation of Microsoft IIS logs for analysis of IIS usage.
    Simplifying URL to only applications root.
    Removing of duplicates from logs entries based on selected fields.

.INPUTS
    None: All parameters and settings have to be done in the Graphical interface.
.OUTPUTS
    Text files (by default in the same directory than the script itself).
.NOTES
    Script developped by Fabian Caignard in 2019.

    This script requires Windows Powershell v5 minimum and use Windows Forms.

    Sources:
    - http://poshgui.com GUI Editor: Most of GUI Forms
    - Thomas Levesque (http://bit.ly/1KmLgyN) for IconExtractor code
    - Marco Blessing (https://bit.ly/2ExphgA) helping me on read-only class properties
    - Michael Willis (https://bit.ly/2T1N94D) helping me on Classes coding
    - Microsoft Technet forums: some other parts of code
    - Microsoft Docs & MSDN: help and code examples
#>


#-------------------------------------------------------------------------------
#region: SCRIPT REQUIREMENTS
#Requires -Version 5
#endregion
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
#region: Script details
$MyScript = [PSCustomObject]@{
    Name        = "IIS Logs Analyzer"
    Version     = "1.0"
    Directory   = Split-Path $MyInvocation.MyCommand.Path
    Description = @"
    This script is the Graphical User Interface version of IISWebStats.ps1 and RemIISWStats.ps1 scripts.

It allow you to read, analyze and concatenate all IIS Web and/or FTP log files on local computer or on one or several remote servers. It automatically remove duplicates (based on same Date, User, Client IP and URLs in log entries).

Results of the script analysis are saved in a text output file. By default, the output file is saved in the same directory than the script.
It is possible to define an other directory to save output file(s) but it is not possible to define the output file(s) name.

Click "Next" button and Enjoy!
"@
}
#endregion
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
#region: Import and load Script's Libraries
# For each PSL file found in provided path...
foreach ($file in (Get-ChildItem $("{0}\*.psl" -f $($MyScript.Directory).TrimEnd('\')) -File)) {
    # Display a loading output and load in dot source decoded scriptblock from Base64 content.
    Write-Host $("Loading {0} in current script." -f $file.Name)
    . ([scriptblock]::Create([System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($(Get-Content $file.FullName)))))
}
#endregion
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
#region: Default GUI Font face
$MyFonts = [GuiDefaultFonts]::new("Segoe UI")
# Help: To use these default GUI Font faces you can use:
#    -----------------------------     ----------------------------
#    Code                              Value
#    -----------------------------     ----------------------------
#    [string]$MyFonts.MainTitle        "Segoe UI,12,style=Bold"
#    [string]$MyFonts.Title            "Segoe UI,10,style=Bold"
#    [string]$MyFonts.Normal           "Segoe UI,9,style=Regular"
#    [string]$MyFonts.Strong           "Segoe UI,9,style=Bold"
#    -----------------------------     ----------------------------
#endregion
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
#region: Loading Checks
#-------------------------------------------------------------------------------
# Checks Screen Resolution: if screen width is less than 1024 or screen height is less than 720, displays an error message box then exit script.
if (([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width -lt 1024) -or ([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height -lt 720)) {
    $msg = "The current screen resolution is not appropriate to run this script!`nPlease make sure you have a minimum resolution of 1024x768 or 1280x720."
    [System.Windows.MessageBox]::Show($msg, "Unsufficient screen resolution", 0, 16)
    exit
}
#-------------------------------------------------------------------------------
# Checks the Windows DPI rate: if different than 100% (Windows Text Size setting), displays a warning message box.
if ($curDPIrate -ne 100) {
    # Formatting the warning message string.
    $msg = "The current Windows Text size setting is set to {0}%. To get the best experience with {1} the recommended setting is 100%.`n`nClick on OK button to continue, or click Cancel button to exit." -f $curDPIrate, $MyScript.Name
    # If user click on Cancel button, then Exit script.
    if ([System.Windows.MessageBox]::Show($msg, "Not appropriate Windows setting detected", 1, 48) -eq [System.Windows.MessageBoxResult]::Cancel) { exit }
}
#-------------------------------------------------------------------------------
#endregion
#-------------------------------------------------------------------------------

#-------------------------------------------------------------------------------
#region: Script running
# If user clicks on Next button or press Enter key in splash screen, then continue the Script's process.
$ret = Show-WelcomeForm
If ($ret -eq "OK") {
    # Show Application Loading form
    Show-AppLoading
    # Update loading message
    $MsgDisplay_Lbl.Text = "Loading"
    $MsgDisplay_Lbl.Refresh()
    #region: Common variables setting
    [string]$script:OutputPath = $MyScript.Directory
    [System.Management.Automation.PSCredential]$script:SharedCredentials = [System.Management.Automation.PSCredential]::Empty
    [bool]$script:UseSharedCreds = $false
    $script:ComputerList = [System.Collections.ObjectModel.Collection[Computer]]::new()
    $script:LogDirectoryList = [System.Collections.ObjectModel.Collection[LogDirectory]]::new()
    #endregion
    # Update loading message
    $MsgDisplay_Lbl.Text += "."
    $MsgDisplay_Lbl.Refresh()
    #region: Adding local computer and related log directories to script lists
    # Adding of Local computer in computers list.
    $script:ComputerList.Add([Computer]::new(".", 0))
    # Update loading message
    $MsgDisplay_Lbl.Text += "."
    $MsgDisplay_Lbl.Refresh()
    # Initiate a counter of Log Directories to 0
    [int]$localldc = 0
    # Getting default log directories on local system drive (then for local computer), then for each found...
    Get-LogDirectories -Path $env:SystemDrive | ForEach-Object {
        # Add the found log directory to LogDirectory objects collection with local computer ID (here 0)
        $script:LogDirectoryList.Add([LogDirectory]::New($_, 0))
        # Increase counter
        $localldc++
    }
    # Update loading message
    $MsgDisplay_Lbl.Text += "."
    $MsgDisplay_Lbl.Refresh()
    # If no log directory found (counter equals 0), Display a warning message box, else...
    if ($localldc -eq 0) { Show-Warning "The local computer seems to not be an IIS server!" "No IIS logs found" } else {
        # For each LogDirectory Object in the collection...
        $script:LogDirectoryList | ForEach-Object {
            # Update loading message
            $MsgDisplay_Lbl.Text += "."
            $MsgDisplay_Lbl.Refresh()
            # Get log files list by calling GetFiles method of LogDirectory Object.
            $_.GetFiles()
        }
    }
    # Closing Application loading form.
    $AppLoading_Frm.Close()
    #endregion
    #region: Displaying script main window.
    Show-MainForm
    #endregion
}
#endregion
#-------------------------------------------------------------------------------

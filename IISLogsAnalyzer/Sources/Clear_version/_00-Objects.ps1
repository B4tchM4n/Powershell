<#
Library Name:   _00-Objects.ps1
Authors:        Fabian Caignard
Date:           March 7th 2019
Version:        1.0
Description:    This Powershell library file contains Enumerations and Classes of custom
                Objects used by IISLogsAnalyser.ps1 script.
                Custom Objects contained in this library:
                -   AuthStatus (2 booleans flags Object for "LocalUser" & "withCredentials"
                    authentication status)
                -   DateRange (Dates range Object of 2 dates: "Oldest" & "Latest")
                -   Computer (representing a Computer and storing Credentials, AuthStatus,
                    etc.)
                -   LogFile (representing a log file, storing corresponding File Object, and
                    checking if it is an IIS W3C log file)
                -   LogFileData (representing a log file data object, storing corresponding server,
                    log directory as "Instance" and logfile object)
                -   LogField (representing a W3C Log file field)
                -   LogDirectory (representing a log folder, storing a collection of LogFile
                    objects, DateRange objects, etc.)
                -   GuiFont (representing a Font family and it's style and size)
                -   GuiDefaultFonts (representing a set of 4 GuiFont objects: MainTitle, Title,
                    Normal, adn Strong)
#>

#-------------------------------------------------------------------------------
Enum ComputerType {
    Local = 0
    Remote = 1
}
#-------------------------------------------------------------------------------
#region: AuthStatus Object Class
Class AuthStatus {
    #region: Properties
    # LU private property for Local user authentication status
    hidden [bool]$LU = $false
    # SC private property for specific credentials authentication status
    hidden [bool]$SC = $false
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        # Adding of LocalUser public property: Set property read-only. Returns the value of LU private property.
        $this | Add-Member -MemberType ScriptProperty -Name LocalUser -Value { $this.LU } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of withCredentials public property: Set property read-only. Returns the value of SC private property.
        $this | Add-Member -MemberType ScriptProperty -Name withCredentials -Value { $this.SC } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
    }
    #endregion
    #region: Constructor (call AddPublicMembers private method, then set appropriate properties)
    AuthStatus() { $this.AddPublicMembers() }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: DateRange Object Class
Class DateRange {
    #region: Properties
    # OD private property for oldest/starting date in the range.
    hidden [datetime]$OD = [Datetime]::MinValue
    # LD private property for latest/ending date in the range.
    hidden [datetime]$LD = [datetime]::MinValue
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        # Adding of Oldest public property: Set property read-only. Returns "Not set" if no date defined or date from OD private property.
        $this | Add-Member -MemberType ScriptProperty -Name Oldest -Value {
            if ($this.OD -eq [datetime]::MinValue) { "Not set" } else { Get-Date -Date $this.OD -Format d }
        } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of Latest public property: Set property read-only. Returns "Not set" if no date defined or date from LD private property.
        $this | Add-Member -MemberType ScriptProperty -Name Latest -Value {
            if ($this.LD -eq [datetime]::MinValue) { "Not set" } else { Get-Date -Date $this.LD -Format d }
        } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
    }
    # SetRange public method to set DateRange dates (2nd method allow applying a dates limit range)
    [void]SetRange ([datetime]$oldest, [datetime]$latest) {
        if ($oldest -gt $latest) { Write-Warning "Invalid dates! The oldest date cannot be later than the latest date." } else {
            $this.OD = Get-Date -Date $oldest -Hour 0 -Minute 0 -Second 0 -Millisecond 0
            $this.LD = Get-Date -Date $latest -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        }
    }
    [void]SetRange ([datetime]$oldest, [datetime]$latest, [DateRange]$limitrange) {
        if ($oldest -gt $latest) { Write-Warning "Invalid dates! The oldest date cannot be later than the latest date." } else {
            if (($oldest -lt $limitrange.OD) -or ($latest -gt $limitrange.LD)) { Write-Warning "Invalid dates! One of the provided dates is out of limit range." } else {
                $this.OD = Get-Date -Date $oldest -Hour 0 -Minute 0 -Second 0 -Millisecond 0
                $this.LD = Get-Date -Date $latest -Hour 0 -Minute 0 -Second 0 -Millisecond 0
            }
        }
    }
    # ClearRange public method to reset DateRange dates to "undefined" datetime values
    [void]ClearRange () {
        $this.OD = [datetime]::MinValue
        $this.LD = [datetime]::MinValue
    }
    #endregion
    #region: Constructor (call AddPublicMembers private method, then set appropriate properties)
    DateRange () { $this.AddPublicMembers() }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: Computer Object Class
Class Computer {
    #region: Properties
    # hostname private property for Hostname of computer Object (should match Hostname RegEx pattern).
    [ValidatePattern('^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$')]
    hidden [string]$hostname
    # hosttype private property for type of computer (Local or Remote).
    hidden [ComputerType]$hosttype
    # reachable private property as boolean (True if computer is reachable over the network or if is Local computer)
    hidden [bool]$reachable = $false
    # scred private property for specific credentials to connect on computer (empty credential by default - current user)
    hidden [System.Management.Automation.PSCredential]$scred = [System.Management.Automation.PSCredential]::Empty
    # auths private property containing an AuthStatus (see AuthStatus ObjectClass) to determine if connection on computer is permitted.
    hidden [AuthStatus]$auths = $([AuthStatus]::new())
    # ID public property to store an identifier as an interger for this computer object.
    [int]$ID = -1
    # PSDrive public property to store the PSDrive name connected for this computer object
    [string]$PSDrive
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        # Adding of Name public property: Setter checks passed hostname by hostname_chk private method.
        $this | Add-Member -MemberType ScriptProperty -Name Name -Value { $this.hostname } -SecondValue { $this.hostname_chk($([String]$args[0])) }
        # Adding of Type public property: Set property read-only. Returns the value of hosttype private property.
        $this | Add-Member -MemberType ScriptProperty -Name Type -Value { $this.hosttype } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of Authentications public property: Set property read-only. Returns the value of auths private property.
        $this | Add-Member -MemberType ScriptProperty -Name Authentications -Value { $this.auths } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of IsReachable public property: Set property read-only. Returns the value of reachable private property.
        $this | Add-Member -MemberType ScriptProperty -Name IsReachable -Value { $this.reachable } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of Credential public property: Returns "(None)" if no specific credentials defined, or Credentials Username from scred private property if defined.
        # If hosttype private property equals to "Remote", checks and set passed credentials in scred private property, then checks authentication to computer.
        $this | Add-Member -MemberType ScriptProperty -Name Credential -Value {
            if ($this.scred -eq [System.Management.Automation.PSCredential]::Empty) { "(None)" } else { $this.scred.UserName }
        } -SecondValue {
            if ($this.hosttype -eq [ComputerType]::Remote) {
                if (!($args[0])) {
                    $this.scred = [System.Management.Automation.PSCredential]::Empty
                }
                else {
                    $this.scred = [System.Management.Automation.PSCredential]$($args[0])
                    $this.auths = $this.auth_chk()
                }
            }
            else { Write-Warning "Credentials applying error! Local computer object cannot accept any specific credential." }
        }
    }
    # CheckAvailability public method to check if computer (by Hostname from hostname private property) is reachable over the network. Returns True or False.
    [bool]CheckAvailability() {
        if (Test-Connection -ComputerName $this.hostname -Count 1 -ErrorAction SilentlyContinue) { return $true } else { return $false }
    }
    # test_auth private method (called by TestAuthentication public method) to check if passed credentials are permitted to connect on computer.
    hidden [bool]test_auth([System.Management.Automation.PSCredential]$credentials, [string]$path) {
        # If computer is not reachable, test if reachable again, then, in case of success, perform authentication tests, else returns False.
        if (!($this.reachable)) { if ($this.CheckAvailability()) { $this.reachable = $true } else { return $false } }
        $psdroot = '\\' + $this.hostname + '\' + $path
        if (!(New-PSDrive -Name objTmpDrv -PSProvider FileSystem -Root $psdroot -Credential $credentials -ErrorAction SilentlyContinue)) { return $false } else {
            if (!(Test-Path $psdroot -ErrorAction SilentlyContinue)) {
                Get-PSDrive objTmpDrv | Remove-PSDrive -Force
                return $false
            }
            Get-PSDrive objTmpDrv | Remove-PSDrive -Force
        }
        return $true
    }
    # TestAuthentication public method to check if passed credentials are permitted to connect on computer.
    [bool]TestAuthentication() {
        [System.Management.Automation.PSCredential]$credentials = [System.Management.Automation.PSCredential]::Empty
        [string]$path = 'C$'
        return $this.test_auth($credentials, $path)
    }
    [bool]TestAuthentication([System.Management.Automation.PSCredential]$credentials) {
        [string]$path = 'C$'
        return $this.test_auth($credentials, $path)
    }
    [bool]TestAuthentication([System.Management.Automation.PSCredential]$credentials, [string]$path) {
        return $this.test_auth($credentials, $path)
    }
    # auth_chk private method to check Current User and defined specific credentials authentications (by calling TestAuthentication public method) for Remote computer only.
    hidden [AuthStatus]auth_chk() {
        if ($this.hosttype -eq [ComputerType]::Local) { $ret = $this.auths } else {
            $ret = [AuthStatus]::New()
            $ret.set_LU($($this.TestAuthentication()))
            $ret.set_SC($($this.TestAuthentication($this.scred)))
        }
        return $ret
    }
    # hostname_chk private method to check passed hostname before store it in hostname private property. If "." is passed, then use local computername.
    hidden [void]hostname_chk($name) {
        if ($name -eq ".") { $name = $env:COMPUTERNAME }
        $name = $name.ToUpper()
        $this.hostname = $name
        if ($($env:COMPUTERNAME).ToUpper() -ne $this.hostname) {
            $this.hosttype = [ComputerType]::Remote
            $this.reachable = $this.CheckAvailability()
        }
        else {
            $this.hosttype = [ComputerType]::Local
            $this.reachable = $true
            $this.auths.set_LU($true)
        }
    }
    #endregion
    #region: Constructors (All of them call AddPublicMembers private method, then set appropriate properties)
    Computer () { $this.AddPublicMembers() }
    Computer ([string]$name) {
        $this.AddPublicMembers()
        $this.hostname_chk($name)
    }
    Computer ([string]$name, [int]$identifier) {
        $this.AddPublicMembers()
        $this.hostname_chk($name)
        $this.ID = $identifier
    }
    Computer ([string]$name, [System.Management.Automation.PSCredential]$credentials) {
        $this.AddPublicMembers()
        $this.hostname_chk($name)
        $this.Credential = $credentials
    }
    Computer ([string]$name, [System.Management.Automation.PSCredential]$credentials, [int]$identifier) {
        $this.AddPublicMembers()
        $this.hostname_chk($name)
        $this.Credential = $credentials
        $this.ID = $identifier
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: LogFile Object Class
Class LogFile {
    #region: Properties
    # fileobject private property to store FileInfo Object of LogFile
    hidden [System.IO.FileInfo]$fileobject
    # isiislogfile private proterty as boolean flag to determine if log file is an IIS log file
    hidden [bool]$isiislogfile = $false
    # alreadychecked private property as boolean flag to determine if IIS log file type was already checked or not.
    hidden [bool]$alreadychecked = $false
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        # Adding of Name public property: Set property read-only. Returns the value of Name property of object stored in fileobject private property.
        $this | Add-Member -MemberType ScriptProperty -Name Name -Value { $this.fileobject.Name } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of CreationDate public property: Set property read-only. Returns the value of CreationTime property of object stored in fileobject private property.
        $this | Add-Member -MemberType ScriptProperty -Name CreationDate -Value { [datetime](Get-Date -Date $this.fileobject.CreationTime -Hour 0 -Minute 0 -Second 0 -Millisecond 0) } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of isIISLog public property: Set property read-only. Returns the value of isiislogfile private property (if not already checked, runs the Test_iisW3Clog method).
        $this | Add-Member -MemberType ScriptProperty -Name isIISLog -Value {
            if (!($this.alreadychecked)) {
                $this.isiislogfile = $this.Test_iisW3Clog()
            }
            $this.isiislogfile
        } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
    }
    # Test_iisW3Clog private method to check if the file object in fileobject private property is an IIS log file in W3C format. Returns a boolean value.
    hidden [bool]Test_iisW3Clog () {
        [string]$firstfileline = Get-Content -Path $this.fileobject.FullName -TotalCount 1 -Force -ErrorAction SilentlyContinue
        $this.alreadychecked = $true
        if ($firstfileline.Length -gt 50) { return ($firstfileline.Substring(0, 50) -eq '#Software: Microsoft Internet Information Services') } else { return $false }
    }
    #endregion
    #region: Constructor (call AddPublicMembers private method, then set appropriate properties)
    LogFile ([string]$fileName) {
        if (!(Test-Path $fileName -PathType Leaf)) { Throw "Invalid filename!" } else {
            $this.fileobject = [System.IO.FileInfo]::New($fileName)
        }
        $this.AddPublicMembers()
    }
    LogFile ([System.IO.FileInfo]$fileinfoOject) {
        $this.fileobject = $fileinfoOject
        $this.AddPublicMembers()
    }
    LogFile ([string]$fileName, [bool]$checkiislog) {
        if (!(Test-Path $fileName -PathType Leaf)) { Throw "Invalid filename!" } else {
            $this.fileobject = [System.IO.FileInfo]::New($fileName)
        }
        $this.AddPublicMembers()
        if ($checkiislog) { $this.isiislogfile = $this.Test_iisW3Clog() }
    }
    LogFile ([System.IO.FileInfo]$fileinfoOject, [bool]$checkiislog) {
        $this.fileobject = $fileinfoOject
        $this.AddPublicMembers()
        if ($checkiislog) { $this.isiislogfile = $this.Test_iisW3Clog() }
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: LogFileData Object Class
Class LogFileData {
    #region: Properties
    [string]$ServerName
    [string]$InstanceName
    [System.IO.FileInfo]$LogFileObject
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        $this | Add-Member -MemberType AliasProperty -Name Server -Value ServerName
        $this | Add-Member -MemberType AliasProperty -Name Instance -Value InstanceName
        $this | Add-Member -MemberType AliasProperty -Name Logfile -Value LogFileObject
    }
    #endregion
    #region: Constructor
    LogFileData () { $this.AddPublicMembers() }
    LogFileData ([string]$server, [string]$instance, [System.IO.FileInfo]$file) {
        $this.AddPublicMembers()
        $this.ServerName = $server
        $this.InstanceName = $instance
        $this.LogFileObject = $file
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: LogField Object Class
Class LogField {
    #region: Properties
    # DisplayName public property of User friendly name of log field
    [string]$DisplayName
    # WN private proterty for the technical W3C name of the log field
    hidden [string]$WN
    # Description public property for official log field description
    [string]$Description
    # Selected public property to determine if log filed is selected
    [bool]$Selected
    # hFMD private property for the W3C name of another field that the current field mandatory status depends on
    hidden [string]$hFMD
    # Mandatory public property to determine if log field is mandatory
    [bool]$Mandatory
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        # Adding of W3Cname public property: Set property if string does not contain any whitespace. Returns the value of WN private property.
        $this | Add-Member -MemberType ScriptProperty -Name W3Cname -Value { $this.WN } -SecondValue {
            if ($args[0] -match '\ +') { Write-Warning "The W3C log field name cannot contain any whitespace!" } else { $this.WN = $args[0] }
        }
        # Adding of FMD public property: Set property if string does not contain any whitespace. Returns the value of hFMD private property.
        $this | Add-Member -MemberType ScriptProperty -Name FMD -Value { $this.hFMD } -SecondValue {
            if ($args[0] -match '\ +') { Write-Warning "The W3C log field name for this property cannot contain any whitespace!" } else { $this.hFMD = $args[0] }
        }
        # Adding of FieldMandatoryDependency public alias property of FMD public property.
        $this | Add-Member -MemberType AliasProperty -Name FieldMandatoryDependency -Value FMD
    }
    #endregion
    #region: Constructor (call AddPublicMembers private method, then set appropriate properties)
    LogField () { $this.AddPublicMembers() }
    LogField ([string]$name) {
        $this.AddPublicMembers()
        if ($name -match '\ +') { throw "The W3C log field name cannot contain any whitespace!" } else { $this.WN = $name }
    }
    LogField ([string]$name, [string]$DisplayName, [bool]$ismandatory, [string]$dependency) {
        $this.AddPublicMembers()
        if ($name -match '\ +') { throw "The W3C log field name cannot contain any whitespace!" } else { $this.WN = $name }
        if ($dependency -match '\ +') { throw "The W3C log field name for dependency parameter cannot contain any whitespace!" } else { $this.FMD = $dependency }
        $this.Mandatory = $ismandatory
        if ($this.Mandatory) { $this.Selected = $true }
        $this.DisplayName = $DisplayName
    }
    LogField ([string]$name, [string]$DisplayName, [string]$description, [bool]$ismandatory, [string]$dependency) {
        $this.AddPublicMembers()
        if ($name -match '\ +') { throw "The W3C log field name cannot contain any whitespace!" } else { $this.WN = $name }
        if ($dependency -match '\ +') { throw "The W3C log field name for dependency parameter cannot contain any whitespace!" } else { $this.FMD = $dependency }
        $this.Mandatory = $ismandatory
        if ($this.Mandatory) { $this.Selected = $true }
        $this.DisplayName = $DisplayName
        $this.Description = $description
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: LogDirectory Object Class
Class LogDirectory {
    #region: Properties
    # dirpath private property for directory path string.
    hidden [string]$dirpath
    # logfilesdates private property for range of creation dates of log files found in the directory
    hidden [DateRange]$logfilesdates = [DateRange]::new()
    # filterdates private property for range of dates to filter log files by their creation dates
    hidden [DateRange]$filterdates = [DateRange]::new()
    # Selected public property to define if Log Directory is selected or not (by default, it is set to true).
    [bool]$Selected = $true
    # ServerID public property for identifier of corresponding Computer
    [int]$ServerID = -99
    # ID public property to store an identifier as an integer for this Log Directory object
    [int]$ID = -1
    # LogFiles public property as a collection of LogFile objects corresponding to log files found in this directory
    $LogFiles = $(New-Object System.Collections.ObjectModel.Collection[LogFile])
    #endregion
    #region: Methods
    # AddPublicMembers private method to add public properties with specific getters and/or setters.
    hidden [void]AddPublicMembers () {
        # Adding of Path public property: Set property read-only. Returns the value of dirpath private property.
        $this | Add-Member -MemberType ScriptProperty -Name Path -Value { $this.dirpath } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of Name public property: Set property read-only. Returns the directory name from dirpath private property.
        $this | Add-Member -MemberType ScriptProperty -Name Name -Value { $(Split-Path -Path $this.dirpath -Leaf).ToUpper() } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of RootDrive public property: Set property read-only. Returns the directory root drive or unc path root from dirpath private property.
        $this | Add-Member -MemberType ScriptProperty -Name RootDrive -Value {
            if ($this.dirpath.Substring(2) -eq '\\') {
                $uncrootpos = $this.dirpath.IndexOf('\', $this.dirpath.IndexOf('\', 2) + 1)
                if ($uncrootpos -eq -1) { $this.dirpath } else { $this.dirpath.Substring(0, $uncrootpos) }
            }
            else {
                Split-Path -Path $this.dirpath -Qualifier
            }
        } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of LogsDates public property: Set property read-only. Returns the log files dates range from logfilesdates private property.
        $this | Add-Member -MemberType ScriptProperty -Name LogsDates -Value {
            if ($this.LogFiles.Count -gt 0) {
                $oldest = $($this.LogFiles | Where-Object -Property isIISLog -EQ $true | Measure-Object -Property CreationDate -Minimum).Minimum
                $latest = $($this.LogFiles | Where-Object -Property isIISLog -EQ $true | Measure-Object -Property CreationDate -Maximum).Maximum
                $this.logfilesdates.SetRange($oldest, $latest)
            }
            $this.logfilesdates
        } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
        # Adding of Filter public property: Set property read-only. Returns the filtering dates range from filterdates private property.
        $this | Add-Member -MemberType ScriptProperty -Name Filter -Value { $this.filterdates } -SecondValue { Write-Warning "This property cannot be modified! This is a read-only property." }
    }
    # GetFiles public method to fill LogFiles property with Log files in the directory
    [void]GetFiles () {
        Get-ChildItem -Path $("{0}*.log" -f $this.Path) -File | ForEach-Object { $this.LogFiles.Add([LogFile]::New($_, $true)) }
    }
    #endregion
    #region: Constructors (All of them call AddPublicMembers private method, then set appropriate properties - does not accept empty string for $path parameter)
    LogDirectory ([string]$path) {
        $this.AddPublicMembers()
        if ($path -eq "") { Throw "Invalid parameter: path cannot be an empty string!" }
        if (!(Test-Path $path -PathType Container)) { Throw "The provided path is not a valid directory path!" } else {
            if ($($path)[-1] -ne '\') { $this.dirpath = $path + '\' } else { $this.dirpath = $path }
        }
    }
    LogDirectory ([System.IO.DirectoryInfo]$directory) {
        $this.AddPublicMembers()
        $path = $directory.FullName
        if ($($path)[-1] -ne '\') { $this.dirpath = $path + '\' } else { $this.dirpath = $path }
    }
    LogDirectory ([string]$path, [int]$serveridentifier) {
        $this.AddPublicMembers()
        if ($path -eq "") { Throw "Invalid parameter: path cannot be an empty string!" }
        if (!(Test-Path $path -PathType Container)) { Throw "The provided path is not a valid directory path!" } else {
            if ($($path)[-1] -ne '\') { $this.dirpath = $path + '\' } else { $this.dirpath = $path }
        }
        $this.ServerID = $serveridentifier
    }
    LogDirectory ([System.IO.DirectoryInfo]$directory, [int]$serveridentifier) {
        $this.AddPublicMembers()
        $path = $directory.FullName
        if ($($path)[-1] -ne '\') { $this.dirpath = $path + '\' } else { $this.dirpath = $path }
        $this.ServerID = $serveridentifier
    }
    LogDirectory ([int]$identifier, [string]$path) {
        $this.AddPublicMembers()
        if ($path -eq "") { Throw "Invalid parameter: path cannot be an empty string!" }
        if (!(Test-Path $path -PathType Container)) { Throw "The provided path is not a valid directory path!" } else {
            if ($($path)[-1] -ne '\') { $this.dirpath = $path + '\' } else { $this.dirpath = $path }
        }
        $this.ID = $identifier
    }
    LogDirectory ([int]$identifier, [string]$path, [int]$serveridentifier) {
        $this.AddPublicMembers()
        if ($path -eq "") { Throw "Invalid parameter: path cannot be an empty string!" }
        if (!(Test-Path $path -PathType Container)) { Throw "The provided path is not a valid directory path!" } else {
            if ($($path)[-1] -ne '\') { $this.dirpath = $path + '\' } else { $this.dirpath = $path }
        }
        $this.ID = $identifier
        $this.ServerID = $serveridentifier
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
enum GuiFontStyle {
    Bold
    Italic
    Strikeout
    Underline
}
#-------------------------------------------------------------------------------
#region: GuiFont Object Class
class GuiFont {
    #region: Properties
    [string]$FamilyName
    [GuiFontStyle[]]$Style
    [ValidateRange(1, 48)]
    [int]$Size = 9
    #endregion
    #region: Methods
    # Override of ToString() method: Define and force output format when caller expect a string
    [string]ToString() {
        $tmpStyle = ""
        $this.Style | ForEach-Object { if ($tmpStyle -ne "") { $tmpStyle += "," } ; $tmpStyle += [string]$($_) }
        if ($tmpStyle -eq "") { $tmpStyle = "Regular" }
        if ([string]::IsNullOrWhiteSpace($this.FamilyName)) { $tmpFamily = "Microsoft Sans Serif" } else { $tmpFamily = $this.FamilyName }
        # String output is: "<Font_Family_Name>,<Font_Size>,style=<Font_Styles_Comma_Separated>"
        return $("{0},{1},style={2}" -f $tmpFamily, $this.Size, $tmpStyle)
    }
    # Adding private and static AddStyle Method to add a style to a GuiFont object.
    hidden static [GuiFont]AddStyle([GuiFont]$font, [GuiFontStyle]$style) {
        $tmpFont = [GuiFont]::new()
        $tmpFont.FamilyName = $font.FamilyName
        $tmpFont.Style = $font.Style
        if ($tmpfont.Style -notcontains $style) { $tmpFont.Style += $style }
        $tmpFont.Size = $font.Size
        return $tmpFont
    }
    # Adding private and static FromString Method to return a GuiFont object from a string.
    hidden static [GuiFont]FromString([string]$font) {
        $tmparr = $font.Split(',')
        $arrCount = $tmparr.Count
        if ($arrCount -lt 3) { Throw "Invalid Font String format! Impossible to convert ""string"" in to a ""GuiFont"" object." }
        $tmpFont = [GuiFont]::new()
        $tmpFont.FamilyName = $tmparr[0]
        $tmpFont.Size = $tmparr[1]
        $firstStyle = $($tmparr[2].Split('='))[1]
        if ($arrCount -eq 3) {
            if ($firstStyle -ne "Regular") { $tmpFont.Style += $firstStyle }
        }
        else {
            $tmpFont.Style += $firstStyle
            for ($i = 3; $i -le $($arrCount - 1); $i++) {
                $tmpFont.Style += $tmparr[$i]
            }
        }
        return $tmpFont
    }
    # Adding public static methods to add Italic style to a GuiFont object.
    static [GuiFont]ToItalic([GuiFont]$font) {
        return [GuiFont]::AddStyle($font, [GuiFontStyle]::Italic)
    }
    static [GuiFont]ToItalic([string]$font) {
        return [GuiFont]::AddStyle($([GuiFont]::FromString($font)), [GuiFontStyle]::Italic)
    }
    # Adding public static methods to add Bold style to a GuiFont object.
    static [GuiFont]ToBold([GuiFont]$font) {
        return [GuiFont]::AddStyle($font, [GuiFontStyle]::Bold)
    }
    static [GuiFont]ToBold([string]$font) {
        return [GuiFont]::AddStyle($([GuiFont]::FromString($font)), [GuiFontStyle]::Bold)
    }
    # Adding public static methods to add Underline style to a GuiFont object.
    static [GuiFont]ToUnderline([GuiFont]$font) {
        return [GuiFont]::AddStyle($font, [GuiFontStyle]::Underline)
    }
    static [GuiFont]ToUnderline([string]$font) {
        return [GuiFont]::AddStyle($([GuiFont]::FromString($font)), [GuiFontStyle]::Underline)
    }
    # Adding public static methods to change Size of a GuiFont object.
    static [GuiFont]Size([GuiFont]$font, [int]$size) {
        $tmpFont = [GuiFont]::new()
        $tmpFont.FamilyName = $font.FamilyName
        $tmpFont.Style = $font.Style
        $tmpFont.Size = $size
        return $tmpFont
    }
    static [GuiFont]Size([string]$font, [int]$size) {
        [GUIFont]$font = $([GuiFont]::FromString($font))
        $tmpFont = [GuiFont]::new()
        $tmpFont.FamilyName = $font.FamilyName
        $tmpFont.Style = $font.Style
        $tmpFont.Size = $size
        return $tmpFont
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------
#region: GuiDefaultFonts Object Class
class GuiDefaultFonts {
    #region: Properties
    [GuiFont]$MainTitle = [GuiFont]::new()
    [GuiFont]$Title = [GuiFont]::new()
    [GuiFont]$Normal = [GuiFont]::new()
    [GuiFont]$Strong = [GuiFont]::new()
    hidden[string]$defftfam
    #endregion
    #region: Methods
    hidden [void]AddPublicMembers() {
        # Adding DefaultFamily public property: Displays default Font Family ; Set a common Font Family for all unset properties
        $this | Add-Member -MemberType ScriptProperty -Name DefaultFamily -Value {
            if (!([string]::IsNullOrWhiteSpace($this.defftfam))) { $this.defftfam } else { "Not set" }
        } -SecondValue {
            $this.defftfam = [string]$args[0]
            if ([string]::IsNullOrWhiteSpace($this.MainTitle.FamilyName)) { $this.MainTitle.FamilyName = [string]$args[0] }
            if ([string]::IsNullOrWhiteSpace($this.Title.FamilyName)) { $this.Title.FamilyName = [string]$args[0] }
            if ([string]::IsNullOrWhiteSpace($this.Normal.FamilyName)) { $this.Normal.FamilyName = [string]$args[0] }
            if ([string]::IsNullOrWhiteSpace($this.Strong.FamilyName)) { $this.Strong.FamilyName = [string]$args[0] }
        }
    }
    # SetDefaultSize public method: Set provided <size> for Normal & Strong, <size>+1 for Title and <size+3> for MainTitle.
    [void]SetDefaultSize ([int]$fontsize) {
        If (!($fontsize -in 1..48)) { Write-Warning "Out of range size! Font size must be between 1 and 48." } else {
            $this.Normal.Size = $fontsize
            $this.Strong.Size = $fontsize
            $this.Title.Size = $fontsize + 1
            $this.MainTitle.Size = $fontsize + 3
        }
    }
    # Override of ToString() method: Output string output of Normal property
    [string]ToString() {
        return $([string]$this.Normal)
    }
    #endregion
    #region: Constructor
    GuiDefaultFonts() {
        $this.AddPublicMembers()
        $this.SetDefaultSize(9)
        $this.MainTitle.Style += "Bold"
        $this.Title.Style += "Bold"
        $this.Strong.Style += "Bold"
    }
    GuiDefaultFonts([string]$familyName) {
        if ([string]::IsNullOrWhiteSpace($familyName)) { Throw "Bad parameter! Fonf family name cannot be null or empty." }
        $this.AddPublicMembers()
        $this.DefaultFamily = $familyName
        $this.SetDefaultSize(9)
        $this.MainTitle.Style += "Bold"
        $this.Title.Style += "Bold"
        $this.Strong.Style += "Bold"
    }
    #endregion
}
#endregion
#-------------------------------------------------------------------------------

<#
    .SYNOPSIS
        Get or set Dell BIOS settings, or set the boot order, using the DellBIOSProvider module.

    .DESCRIPTION
        Automatically configure Dell BIOS settings

    .PARAMETER GetSettings
        Instruct the script to get a list of current BIOS settings

    .PARAMETER SetSettings
        Instruct the script to set BIOS settings

    .PARAMETER SetBootOrder
        The desired boot order to be set on the system. Specify a comma separated list of device numbers or short form names (for example: 6,1,0 or hdd.1,embnicipv4,embnicipv6)

    .PARAMETER BootMode
        Used with the SetBootOrder parameter. Specifies the boot mode the boot order should be set for. Acceptable values are (UEFI or Legacy)

    .PARAMETER CsvPath
        The path to the CSV file to be imported or exported

    .PARAMETER AdminPassword
        The current BIOS password

    .PARAMETER AdminPasswordCmsFile
        Specify the path to a CMS-encrypted file containing the current BIOS (admin) password. The file is decrypted in memory at runtime using the device's document-encryption certificate. Use this instead of AdminPassword to keep the password off the command line. Cannot be combined with AdminPassword.

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        PS C:\> Manage-DellBiosSettings-DellBIOSProvider.ps1 -SetSettings -AdminPassword ExamplePassword

        Set BIOS settings supplied in the script.

    .EXAMPLE
        PS C:\> Manage-DellBiosSettings-DellBIOSProvider.ps1 -SetSettings -CsvPath C:\Temp\Settings.csv -AdminPassword ExamplePassword

        Set BIOS settings supplied in a CSV file.

    .EXAMPLE
        PS C:\> Manage-DellBiosSettings-DellBIOSProvider.ps1 -SetSettings -AdminPasswordCmsFile C:\Temp\admin.cms

        Set BIOS settings using an admin password sourced from a CMS-encrypted file.

    .EXAMPLE
        PS C:\> Manage-DellBiosSettings-DellBIOSProvider.ps1 -SetBootOrder hdd.1,embnicipv4,embnicipv6 -BootMode UEFI -AdminPassword ExamplePassword

        Set the UEFI boot order.

    .EXAMPLE
        PS C:\> Manage-DellBiosSettings-DellBIOSProvider.ps1 -GetSettings

        Output a list of current BIOS settings to the screen.

    .EXAMPLE
        PS C:\> Manage-DellBiosSettings-DellBIOSProvider.ps1 -GetSettings -CsvPath C:\Temp\Settings.csv

        Output a list of current BIOS settings to a CSV file.

    .LINK
        https://www.configjon.com/dell-bios-settings-management/

    .NOTES
        Created by: Jon Anderson
        Version: 2.3.2
        Modified: 2026-10-02

        Changelog (newest first). See the related link for additional detail on each release.

        2.3.2 (2026-10-02)
            - Get-Help now displays the script's help. A non-standard section name stopped PowerShell from reading the help block, so the changelog is
              now part of the NOTES section (shown by Get-Help -Full). Added a synopsis, gave each example its own entry, and listed the blog post as
              a related link, so Get-Help -Online opens it.
            - Fixed a spelling error in the LogFile parameter's validation message.

        2.3.1 (2026-09-05)
            - A missing, empty, or mis-headed CSV file passed to CsvPath now stops the script with a logged error instead of applying nothing. The
              path is read literally, so file names containing brackets work.
            - Fixed the timezone bias written to log entries. Every entry after the first carried the bias with the wrong sign, which shifts those
              entries' times in viewers that honor it, such as CMTrace. The bias is now computed on every write.
            - The LogFile and CsvPath parameters now require their .log and .csv extensions at the end of the path instead of anywhere in it.
            - Maintenance: the task sequence detection helper initializes its variables before use so it holds up under strict mode. No behavior
              change.

        2.3.0 (2026-05-24)
            - Added secure password sourcing. New optional AdminPasswordCmsFile parameter sources the admin password from a CMS-encrypted file, decrypted in memory at runtime, so
              the password never appears on the command line. The plain-text AdminPassword parameter is unchanged; specifying both is rejected.
            - Renamed the script and its default log file from the -PSModule suffix to -DellBIOSProvider, to match the DellBIOSProvider installer script and the HP HPCMSL variant naming.

        2.1.0 (2026-05-23)
            - Added the SetBootOrder and BootMode parameters for boot order management via the DellBIOSProvider (DellSmbios:\BootSequence). SetDefaults is intentionally not added
              here because the DellBIOSProvider module has no load-defaults equivalent (it remains available in the WMI settings script).
            - Fixed a dead code path so the "The specified Admin password matches the currently set password" confirmation is now logged correctly.

        2.0.0 (2026-05-21)
            - Added a PossibleValue column to the GetSettings output to match the shape of the WMI settings script (Name, Value, PossibleValue), and rendered the BootSequence value
              as a readable ordered device list instead of the BootDevice array type name.
            - Added -ErrorAction SilentlyContinue to the per-category Get-ChildItem enumeration so benign DellBIOSProvider errors for unsupported UEFI variables (e.g. ForcedNetworkFlag)
              no longer print to the console.
            - Hardened the DellBIOSProvider module version detection to match Install-DellBiosProvider.ps1 (sorts multiple installed package versions and uses recursive
              DellBIOSProvider.psd1 discovery).
            - Added [CmdletBinding(PositionalBinding=$false)] so all arguments must be named.
            - Verified compatibility with DellBIOSProvider 2.10.1 and PowerShell 5.1 + 7.x.

        Pre-2.0.0 (legacy)
            2020-09-07 - Added a LogFile parameter. Changed the default log path in full Windows to $ENV:ProgramData\ConfigJonScripts\Dell. Changed the default log file name to Manage-DellBiosSettings-PSModule.log
                         Created a new function (Stop-Script) to consolidate some duplicate code and improve error reporting. Made a number of minor formatting and syntax changes
            2020-09-17 - Improved the log file path configuration

#>

#Parameters ===================================================================================================================

[CmdletBinding(PositionalBinding=$false)]
param(
    [Parameter(Mandatory=$false)][Switch]$GetSettings,
    [Parameter(Mandatory=$false)][Switch]$SetSettings,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$SetBootOrder,
    [Parameter(Mandatory=$false)][ValidateSet('UEFI','Legacy')]$BootMode,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$AdminPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$AdminPasswordCmsFile,
    [ValidateScript({
            if($_ -notmatch '\.csv$')
            {
                throw "The specified file must be a .csv file"
            }
            return $true
        })]
    [System.IO.FileInfo]$CsvPath,
    [Parameter(Mandatory=$false)][ValidateScript({
            if($_ -notmatch '\.log$')
            {
                throw "The file specified in the LogFile parameter must be a .log file"
            }
            return $true
        })]
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\Dell\Manage-DellBiosSettings-DellBIOSProvider.log"
)

#Script version
$Version = '2.3.2'

#Log component name
$Component = 'Manage-DellBiosSettings-DellBIOSProvider'

#List of settings to be configured ============================================================================================
#==============================================================================================================================
$Settings = (
    "FingerprintReader,Enabled",
    "FnLock,Enabled",
    "IntegratedAudio,Enabled",
    "NumLock,Enabled",
    "SecureBoot,Enabled",
    "TpmActivation,Enabled",
    "TpmClear,Disabled",
    "TpmPpiClearOverride,Disabled",
    "TpmPpiDpo,Disabled",
    "TpmPpiPo,Enabled",
    "TpmSecurity,Enabled",
    "UefiNwStack,Enabled",
    "Virtualization,Enabled",
    "VtForDirectIo,Enabled",
    "WakeOnLan,Disabled"
)
#==============================================================================================================================
#==============================================================================================================================

#Functions ====================================================================================================================

Function Get-TaskSequenceStatus
{
    #Determine if a task sequence is currently running. The check is two-stage on purpose: creating Microsoft.SMS.TSEnvironment succeeds on any device with the
    #ConfigMgr client installed, so that alone is a false positive; _SMSTSType is only readable while a task sequence is executing.
    #The variables are initialized before the try blocks so the function also holds up under strict mode.
    $TSEnv = $null
    try
    {
        $TSEnv = New-Object -ComObject 'Microsoft.SMS.TSEnvironment'
    }
    catch
    {
        return $False
    }
    if($NULL -eq $TSEnv)
    {
        return $False
    }
    $TSType = $null
    try
    {
        $TSType = $TSEnv.Value('_SMSTSType')
    }
    catch
    {
        return $False
    }
    return -not [string]::IsNullOrEmpty($TSType)
}

Function Stop-Script
{
    #Write an error to the log file and terminate the script

    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$ErrorMessage,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Exception
    )
    Write-LogEntry -Value $ErrorMessage -Severity 3
    if($Exception)
    {
        Write-LogEntry -Value "Exception Message: $Exception" -Severity 3
    }
    throw $ErrorMessage
}

Function Get-CmsPassword
{
    #Decrypt one or more CMS-encrypted password files and return the plain-text value(s).
    #Decryption uses the device's store-resident document-encryption certificate (the matching private key must be present in Cert:\LocalMachine\My or Cert:\CurrentUser\My).
    #One plain-text value is returned per file, preserving the input order so [String[]] password slots round-trip.
    #The decrypted value is never written to the log, only the file path and a generic failure reason.

    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$CmsFile
    )
    $Result = foreach($File in $CmsFile)
    {
        if(!(Test-Path -LiteralPath $File))
        {
            Stop-Script -ErrorMessage "CMS password file not found: $File"
        }
        try
        {
            Unprotect-CmsMessage -LiteralPath $File -ErrorAction Stop
        }
        catch
        {
            Stop-Script -ErrorMessage "Failed to decrypt the CMS password file: $File. Ensure the document-encryption certificate's private key is present in the certificate store" -Exception $_.Exception.Message
        }
    }
    return $Result
}

Function Import-ValidatedCsv
{
    #Import a CSV file and fail fast: an empty or header-only file, or a file missing any of the required columns, throws a clear error instead of leaving the caller with nothing to apply.
    #The path is read literally so file names containing wildcard characters such as [ ] work. Returns the rows as an array.

    param(
        [Parameter(Mandatory=$true)][String]$Path,
        [String[]]$RequiredColumns
    )
    $Rows = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    if($Rows.Count -eq 0)
    {
        throw "No data rows were found in '$Path'. The file may be empty or contain only a header row."
    }
    if($RequiredColumns)
    {
        $CsvColumns = @($Rows[0].PSObject.Properties.Name)
        $MissingColumns = @($RequiredColumns | Where-Object {$CsvColumns -notcontains $_})
        if($MissingColumns.Count -gt 0)
        {
            throw "The CSV file '$Path' is missing required column(s): $($MissingColumns -join ', ')"
        }
    }
    return , $Rows
}

Function Set-DellBiosSetting
{
    #Set a specific Dell BIOS setting or the boot order

    param(
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Name,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Value,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Password,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$NewBootOrder,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$BootMode
    )
    #Set the boot order
    if($NewBootOrder)
    {
        #Map the BootMode parameter to the value expected by the provider (PossibleValues are Legacy and Uefi)
        if($BootMode -eq 'UEFI')
        {
            $BootListValue = 'Uefi'
        }
        else
        {
            $BootListValue = 'Legacy'
        }
        try
        {
            if([String]::IsNullOrEmpty($Password))
            {
                Set-Item -Path DellSmbios:\BootSequence\BootList -Value $BootListValue -ErrorAction Stop
                Set-Item -Path DellSmbios:\BootSequence\BootSequence -Value ($NewBootOrder -join ',') -ErrorAction Stop
            }
            else
            {
                Set-Item -Path DellSmbios:\BootSequence\BootList -Value $BootListValue -Password $Password -ErrorAction Stop
                Set-Item -Path DellSmbios:\BootSequence\BootSequence -Value ($NewBootOrder -join ',') -Password $Password -ErrorAction Stop
            }
        }
        catch
        {
            $SettingSet = "Failed"
        }
        if($SettingSet -eq "Failed")
        {
            Write-LogEntry -Value "Failed to set the ""$BootMode"" boot order to ""$($NewBootOrder -join ', ')""" -Severity 3
            $Script:FailSet++
        }
        else
        {
            Write-LogEntry -Value "Successfully set the ""$BootMode"" boot order to ""$($NewBootOrder -join ', ')""" -Severity 1
            $Script:SuccessSet++
        }
    }
    #Set a specific BIOS setting
    else
    {
        #Ensure the specified setting exists and get the possible values
        $CurrentValue = $SettingList | Where-Object Attribute -eq $Name | Select-Object -ExpandProperty CurrentValue
        if($NULL -ne $CurrentValue)
        {
            #Setting is already set to specified value
            if($CurrentValue -eq $Value)
            {
                Write-LogEntry -Value "Setting ""$Name"" is already set to ""$Value""" -Severity 1
                $Script:AlreadySet++
            }
            #Setting is not set to specified value
            else
            {
                $SettingPath = $SettingList | Where-Object Attribute -eq $Name | Select-Object -ExpandProperty PSChildName
                if([String]::IsNullOrEmpty($Password))
                {
                    try
                    {
                        Set-Item -Path DellSmbios:\$SettingPath\$Name -Value $Value -ErrorAction Stop
                    }
                    catch
                    {
                        $SettingSet = "Failed"
                    }
                }
                else
                {
                    try
                    {
                        Set-Item -Path DellSmbios:\$SettingPath\$Name -Value $Value -Password $Password -ErrorAction Stop
                    }
                    catch
                    {
                        $SettingSet = "Failed"
                    }
                }
                if($SettingSet -eq "Failed")
                {
                    Write-LogEntry -Value "Failed to set ""$Name"" to ""$Value""." -Severity 3
                    $Script:FailSet++
                }
                else
                {
                    Write-LogEntry -Value "Successfully set ""$Name"" to ""$Value""" -Severity 1
                    $Script:SuccessSet++
                }
            }
        }
        #Setting not found
        else
        {
            Write-LogEntry -Value "Setting ""$Name"" not found" -Severity 2
            $Script:NotFound++
        }
    }
}

Function Write-LogEntry
{
    #Write data to a CMTrace compatible log file. (Credit to MSEndpointMgr - https://www.msendpointmgr.com/)

    param(
        [parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,
        [parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = ($script:LogFile | Split-Path -Leaf),
        [parameter(Mandatory = $false, HelpMessage = "Name of the component that the log entry will be associated with.")]
        [ValidateNotNullOrEmpty()]
        [string]$Component = $script:Component
    )
    #Determine log file location
    $LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName
    #Construct time stamp for log entry. CMTrace expresses the timezone bias as the minutes to ADD to local time to reach
    #UTC, so the sign is the inverse of the .NET offset. The bias is computed fresh on every call: the previous cached form
    #only flipped the sign on a local copy, so every log line after the first in a run carried the raw, wrong-signed value.
    [string]$Bias = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalMinutes
    if($Bias -match "^-")
    {
        $TimezoneBias = $Bias.Replace('-', '+')
    }
    else
    {
        $TimezoneBias = '-' + $Bias
    }
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), $TimezoneBias)
    #Construct date for log entry
    $Date = (Get-Date -Format "MM-dd-yyyy")
    #Construct context for log entry
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    #Construct final log entry
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($Component)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
    #Add value to log file
    try
    {
        Out-File -InputObject $LogText -Append -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    }
    catch [System.Exception]
    {
        Write-Warning -Message "Unable to append log entry to $FileName file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

#Main program =================================================================================================================

#Configure Logging and task sequence variables
if(Get-TaskSequenceStatus)
{
    $TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
    $LogsDirectory = $TSEnv.Value("_SMSTSLogPath")
}
else
{
    $LogsDirectory = ($LogFile | Split-Path)
    if([string]::IsNullOrEmpty($LogsDirectory))
    {
        $LogsDirectory = $PSScriptRoot
    }
    else
    {
        if(!(Test-Path -PathType Container $LogsDirectory))
        {
            try
            {
                New-Item -Path $LogsDirectory -ItemType "Directory" -Force -ErrorAction Stop | Out-Null
            }
            catch
            {
                throw "Failed to create the log file directory: $LogsDirectory. Exception Message: $($PSItem.Exception.Message)"
            }
        }
    }
}
Write-Output "Log path set to $LogFile"
Write-LogEntry -Value "START - Dell BIOS settings management script (version $Version)" -Severity 1

#Resolve secure password sources (decrypt any CMS-encrypted password file into the standard password variable)
if($AdminPassword -and $AdminPasswordCmsFile)
{
    Stop-Script -ErrorMessage "Specify either the AdminPassword or the AdminPasswordCmsFile parameter, not both"
}
if($AdminPasswordCmsFile)
{
    Write-LogEntry -Value "Decrypting the admin password from the supplied CMS file" -Severity 1
    $AdminPassword = Get-CmsPassword -CmsFile $AdminPasswordCmsFile
}

#Check if 32 or 64 bit architecture
if([System.Environment]::Is64BitOperatingSystem)
{
    $ModuleInstallPath = $env:ProgramFiles
}
else
{
    $ModuleInstallPath = ${env:ProgramFiles(x86)}
}

#Verify the DellBIOSProvider module is installed
Write-LogEntry -Value "Checking the version of the currently installed DellBIOSProvider module" -Severity 1
try
{
    $LocalVersion = Get-Package DellBIOSProvider -ErrorAction Stop |
        Select-Object -ExpandProperty Version -ErrorAction Stop |
        Sort-Object { [Version]$_ } -Descending |
        Select-Object -First 1
}
catch
{
    $Local = $true
    $LocalModuleRoot = "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider"
    $LocalPsd1 = if(Test-Path $LocalModuleRoot) { Get-ChildItem -Path $LocalModuleRoot -Filter 'DellBIOSProvider.psd1' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 } else { $null }
    if($LocalPsd1)
    {
        $LocalVersion = Get-Content $LocalPsd1.FullName | Select-String "ModuleVersion ="
        $LocalVersion = (([regex]".*'(.*)'").Matches($LocalVersion))[0].Groups[1].Value
        if($null -ne $LocalVersion)
        {
            Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
        }
        else
        {
            Stop-Script -ErrorMessage "DellBIOSProvider module not found on the local machine"
        }
    }
    else
    {
        Stop-Script -ErrorMessage "DellBIOSProvider module not found on the local machine"
    }
}
if(($null -ne $LocalVersion) -and (!($Local)))
{
    Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
}

#Verify the DellBIOSProvider module is imported
Write-LogEntry -Value "Verify the DellBIOSProvider module is imported" -Severity 1
$ModuleCheck = Get-Module DellBIOSProvider
if($ModuleCheck)
{
    Write-LogEntry -Value "The DellBIOSProvider module is already imported" -Severity 1
}
else
{
    Write-LogEntry -Value "Importing the DellBIOSProvider module" -Severity 1
    $Error.Clear()
    try
    {
        Import-Module DellBIOSProvider -Force -ErrorAction Stop
    }
    catch
    {
        Stop-Script -ErrorMessage "Failed to import the DellBIOSProvider module" -Exception $PSItem.Exception.Message
    }
    if(!($Error))
    {
        Write-LogEntry -Value "Successfully imported the DellBIOSProvider module" -Severity 1
    }
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1
if($GetSettings -and ($SetSettings -or $SetBootOrder))
{
    Stop-Script -ErrorMessage "Cannot specify the GetSettings and SetSettings or SetBootOrder parameters at the same time"
}
if(!($GetSettings -or $SetSettings -or $SetBootOrder))
{
    Stop-Script -ErrorMessage "One of the GetSettings or SetSettings or SetBootOrder parameters must be specified when running this script"
}
if($SetSettings -and !($Settings -or $CsvPath))
{
    Stop-Script -ErrorMessage "Settings must be specified using either the Settings variable in the script or the CsvPath parameter"
}
if($SetBootOrder -and !($BootMode))
{
    Stop-Script -ErrorMessage "When using the SetBootOrder parameter, the BootMode parameter must also be specified"
}
Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Set counters to 0
if($SetSettings -or $SetBootOrder)
{
    $AlreadySet = 0
    $SuccessSet = 0
    $FailSet = 0
    $NotFound = 0
}

#Get the current password status
if($SetSettings -or $SetBootOrder)
{
    Write-LogEntry -Value "Get the current password state" -Severity 1
    $AdminPasswordCheck = Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet | Select-Object -ExpandProperty CurrentValue
    if($AdminPasswordCheck -eq "True")
    {
        Write-LogEntry -Value "The Admin password is currently set" -Severity 1
        #Setup password set but parameter not specified
        if([String]::IsNullOrEmpty($AdminPassword))
        {
            Stop-Script -ErrorMessage "The Admin password is set, but no password was supplied. Use the AdminPassword parameter when a password is set"
        }
        #Setup password set correctly
        try
        {
            Set-Item -Path DellSmbios:\Security\AdminPassword $AdminPassword -Password $AdminPassword -ErrorAction Stop
        }
        catch
        {
            $AdminPasswordCheck = "Failed"
            Stop-Script -ErrorMessage "The specified Admin password does not match the currently set password"
        }
        Write-LogEntry -Value "The specified Admin password matches the currently set password" -Severity 1
    }
    else
    {
        Write-LogEntry -Value "The Admin password is not currently set" -Severity 1
    }
}

#Get a list of current BIOS settings
Write-LogEntry -Value "Getting a list of current BIOS settings" -Severity 1
$DellSmbios = Get-ChildItem -Path DellSmbios:\
$SettingCategory = $DellSmbios.Category
$SettingList = @()
$WarnPref = $WarningPreference #Get the current value of WarningPreference
$WarningPreference = 'SilentlyContinue' #Suppress warnings

if($SetSettings)
{
    foreach($Category in $SettingCategory){
        $SettingList += Get-ChildItem -Path "DellSmbios:\$($Category)" -ErrorAction SilentlyContinue | Select-Object Attribute,CurrentValue,PSChildName
    }
}

#Get the current settings
if($GetSettings)
{
    foreach($Category in $SettingCategory){
        $SettingList += Get-ChildItem -Path "DellSmbios:\$($Category)" -ErrorAction SilentlyContinue | Select-Object Attribute,CurrentValue,PossibleValues
    }
    $WarningPreference = $WarnPref #Revert WarningPreference back to the original value
    $SettingList = $SettingList | Sort-Object Attribute
    $SettingObject = ForEach($Setting in $SettingList){
        #BootSequence returns an array of BootDevice objects - render it as a readable, ordered device list
        if($Setting.Attribute -eq 'BootSequence')
        {
            $Value = ($Setting.CurrentValue.DeviceName) -join ' '
        }
        else
        {
            $Value = $Setting.CurrentValue
        }
        [PSCustomObject]@{
            Name = $Setting.Attribute
            Value = $Value
            PossibleValue = [String]$Setting.PossibleValues
        }
    }
    if($CsvPath)
    {
        $SettingObject | Export-Csv -Path $CsvPath -NoTypeInformation
        (Get-Content $CsvPath) | ForEach-Object {$_ -Replace '"',""} | Out-File $CsvPath -Force -Encoding ascii
    }
    else
    {
        Write-Output $SettingObject
    }
}
#Set settings and/or boot order
if($SetSettings -or $SetBootOrder)
{
    if($CsvPath)
    {
        try
        {
            $Settings = Import-ValidatedCsv -Path $CsvPath -RequiredColumns @('Name','Value')
        }
        catch
        {
            Stop-Script -ErrorMessage "Failed to import BIOS settings from the CSV file: $CsvPath" -Exception $_.Exception.Message
        }
    }
    #Set Dell BIOS settings - password is set
    if($AdminPasswordCheck -eq "True")
    {
        #Set the boot order
        if($SetBootOrder)
        {
            Set-DellBiosSetting -NewBootOrder $SetBootOrder -BootMode $BootMode -Password $AdminPassword
        }
        #Set all other settings
        if($SetSettings)
        {
            if($CsvPath)
            {
                ForEach($Setting in $Settings){
                    Set-DellBiosSetting -Name $Setting.Name -Value $Setting.Value -Password $AdminPassword
                }
            }
            else
            {
                ForEach($Setting in $Settings){
                    $Data = $Setting.Split(',')
                    Set-DellBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim() -Password $AdminPassword
                }
            }
        }
    }
    #Set Dell BIOS settings - password is not set
    else
    {
        #Set the boot order
        if($SetBootOrder)
        {
            Set-DellBiosSetting -NewBootOrder $SetBootOrder -BootMode $BootMode
        }
        #Set all other settings
        if($SetSettings)
        {
            if($CsvPath)
            {
                ForEach($Setting in $Settings){
                    Set-DellBiosSetting -Name $Setting.Name -Value $Setting.Value
                }
            }
            else
            {
                ForEach($Setting in $Settings){
                    $Data = $Setting.Split(',')
                    Set-DellBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim()
                }
            }
        }
    }
}

#Display results
if($SetSettings -or $SetBootOrder)
{
    Write-Output "$AlreadySet settings already set correctly"
    Write-LogEntry -Value "$AlreadySet settings already set correctly" -Severity 1
    Write-Output "$SuccessSet settings successfully set"
    Write-LogEntry -Value "$SuccessSet settings successfully set" -Severity 1
    Write-Output "$FailSet settings failed to set"
    Write-LogEntry -Value "$FailSet settings failed to set" -Severity 3
    Write-Output "$NotFound settings not found"
    Write-LogEntry -Value "$NotFound settings not found" -Severity 2
}
Write-Output "Dell BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - Dell BIOS settings management script" -Severity 1

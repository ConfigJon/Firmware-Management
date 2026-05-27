<#
    .DESCRIPTION
        Automatically configure HP BIOS settings
        This variant uses the HP Client Management Script Library (HPCMSL) instead of direct WMI calls.
        The HPCMSL must be installed before running this script (see Install-HPCMSL.ps1).

    .PARAMETER GetSettings
        Instruct the script to get a list of current BIOS settings

    .PARAMETER SetSettings
        Instruct the script to set BIOS settings

    .PARAMETER SetDefaults
        Instruct the script to reset all BIOS settings to their default values

    .PARAMETER CsvPath
        The path to the CSV file to be imported or exported

    .PARAMETER SetupPassword
        The current BIOS password

    .PARAMETER SetupPasswordCmsFile
        Specify the path to a CMS-encrypted file containing the current BIOS (setup) password. The file is decrypted in memory at runtime using the device's document-encryption certificate. Use this instead of SetupPassword to keep the password off the command line. Cannot be combined with SetupPassword.

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        #Set BIOS settings supplied in the script
        Manage-HPBiosSettings-HPCMSL.ps1 -SetSettings -SetupPassword ExamplePassword

        #Set BIOS settings supplied in a CSV file
        Manage-HPBiosSettings-HPCMSL.ps1 -SetSettings -CsvPath C:\Temp\Settings.csv -SetupPassword ExamplePassword

        #Set BIOS settings using a setup password sourced from a CMS-encrypted file
        Manage-HPBiosSettings-HPCMSL.ps1 -SetSettings -SetupPasswordCmsFile C:\Temp\setup.cms

        #Output a list of current BIOS settings to the screen
        Manage-HPBiosSettings-HPCMSL.ps1 -GetSettings

        #Output a list of current BIOS settings to a CSV file
        Manage-HPBiosSettings-HPCMSL.ps1 -GetSettings -CsvPath C:\Temp\Settings.csv

        #Reset all BIOS settings to their default values
        Manage-HPBiosSettings-HPCMSL.ps1 -SetDefaults -SetupPassword ExamplePassword

    .NOTES
        Created by: Jon Anderson
        Reference: https://www.configjon.com/hp-bios-settings-management-hpcmsl/
        Version: 2.3.1
        Modified: 2026-05-26

    .CHANGELOG
        See .NOTES Reference for additional detail on each release.

        2.3.1 (2026-05-26)
            - Fixed BIOS setting value parsing on HP models that return enumeration values with leading whitespace. The asterisk-prefix check that
              identifies the currently-set value now trims first, so enumeration values on affected models are correctly detected in GetSettings
              output.
            - Credit to @CharlesNRU for diagnosing the parsing issue and proposing the fix in PR #11.

        2.3.0 (2026-05-24)
            - Added secure password sourcing. New optional SetupPasswordCmsFile parameter sources the setup password from a CMS-encrypted file, decrypted in memory at runtime, so
              the password never appears on the command line. The plain-text SetupPassword parameter is unchanged; specifying both is rejected.

        2.2.0 (2026-05-23)
            - Added detection for HP Sure Admin (Enhanced BIOS Authentication Mode). When Sure Admin is enabled, SetSettings and SetDefaults now log a clear message and exit
              cleanly without attempting changes, since Sure Admin requires signed payloads or a local access key rather than a BIOS password. GetSettings is unaffected (reads do
              not require authorization).

        2.1.0 (2026-05-23)
            - Added the SetDefaults parameter to reset all BIOS settings to defaults via Set-HPBIOSSettingDefaults (no WMI equivalent, so this is only available in the HPCMSL
              variant).
            - The GetSettings output and CSV export now include a PossibleValue column to match the other manufacturer scripts.

        2.0.0 (2026-05-22)
            - Initial release. HPCMSL-based variant of Manage-HPBiosSettings, derived from the WMI script. Gets and sets BIOS settings using the HP Client Management Script Library
              cmdlets (Get-HPBIOSSettingsList, Get-HPBIOSSettingValue, Set-HPBIOSSettingValue) instead of direct WMI calls. Runs on both Windows PowerShell 5.1 and PowerShell 7.
              Requires the HPCMSL to be installed (see Install-HPCMSL.ps1).

#>

#Parameters ===================================================================================================================

[CmdletBinding(PositionalBinding=$false)]
param(
    [Parameter(Mandatory=$false)][Switch]$GetSettings,
    [Parameter(Mandatory=$false)][Switch]$SetSettings,
    [Parameter(Mandatory=$false)][Switch]$SetDefaults,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPasswordCmsFile,
    [ValidateScript({
            if($_ -notmatch "(\.csv)")
            {
                throw "The specified file must be a .csv file"
            }
            return $true
        })]
    [System.IO.FileInfo]$CsvPath,
    [Parameter(Mandatory=$false)][ValidateScript({
            if($_ -notmatch "(\.log)")
            {
                throw "The file specified in the LogFile paramter must be a .log file"
            }
            return $true
        })]
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\HP\Manage-HPBiosSettings-HPCMSL.log"
)

#Script version
$Version = '2.3.1'

#Log component name
$Component = 'Manage-HPBiosSettings-HPCMSL'

#List of settings to be configured ============================================================================================
#==============================================================================================================================
$Settings = (
    "Deep S3,Off",
    "Deep Sleep,Off",
    "S4/S5 Max Power Savings,Disable",
    "S5 Maximum Power Savings,Disable",
    "Fingerprint Device,Disable",
    "Num Lock State at Power-On,Off",
    "NumLock on at boot,Disable",
    "Numlock state at boot,Off",
    "Prompt for Admin password on F9 (Boot Menu),Enable",
    "Prompt for Admin password on F11 (System Recovery),Enable",
    "Prompt for Admin password on F12 (Network Boot),Enable",
    "PXE Internal IPV4 NIC boot,Enable",
    "PXE Internal IPV6 NIC boot,Enable",
    "PXE Internal NIC boot,Enable",
    "Wake On LAN,Boot to Hard Drive",
    "Swap Fn and Ctrl (Keys),Disable"
)
#==============================================================================================================================
#==============================================================================================================================

#Functions ====================================================================================================================

Function Get-TaskSequenceStatus
{
    #Determine if a task sequence is currently running
    try
    {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    }
    catch{}
    if($NULL -eq $TSEnv)
    {
        return $False
    }
    else
    {
        try
        {
            $SMSTSType = $TSEnv.Value("_SMSTSType")
        }
        catch{}
        if([string]::IsNullOrEmpty($SMSTSType))
        {
            return $False
        }
        else
        {
            return $True
        }
    }
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

Function Set-HPBiosSetting
{
    #Set a specific HP BIOS setting

    param(
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Name,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Value,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Password,
        [Parameter(Mandatory=$false)][Switch]$Defaults
    )
    #Load default BIOS settings
    if($Defaults)
    {
        try
        {
            if(!([String]::IsNullOrEmpty($Password)))
            {
                Set-HPBIOSSettingDefaults -Password $Password -ErrorAction Stop
            }
            else
            {
                Set-HPBIOSSettingDefaults -ErrorAction Stop
            }
            Write-LogEntry -Value "Successfully loaded default BIOS settings" -Severity 1
            $Script:DefaultSet = $True
        }
        catch
        {
            Write-LogEntry -Value "Failed to load default BIOS settings. $($PSItem.Exception.Message)" -Severity 3
            $Script:DefaultSet = $False
        }
        return
    }
    #Ensure the specified setting exists and get the current value
    try
    {
        $CurrentValue = Get-HPBIOSSettingValue -Name $Name -ErrorAction Stop
    }
    catch
    {
        Write-LogEntry -Value "Setting ""$Name"" not found" -Severity 2
        $Script:NotFound++
        return
    }
    #Setting is already set to specified value
    if($CurrentValue -eq $Value)
    {
        Write-LogEntry -Value "Setting ""$Name"" is already set to ""$Value""" -Severity 1
        $Script:AlreadySet++
    }
    #Setting is not set to specified value
    else
    {
        try
        {
            if(!([String]::IsNullOrEmpty($Password)))
            {
                Set-HPBIOSSettingValue -Name $Name -Value $Value -Password $Password -ErrorAction Stop
            }
            else
            {
                Set-HPBIOSSettingValue -Name $Name -Value $Value -ErrorAction Stop
            }
            Write-LogEntry -Value "Successfully set ""$Name"" to ""$Value""" -Severity 1
            $Script:SuccessSet++
        }
        catch
        {
            Write-LogEntry -Value "Failed to set ""$Name"" to ""$Value"". $($PSItem.Exception.Message)" -Severity 3
            $Script:FailSet++
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
    #Construct time stamp for log entry
    if(-not(Test-Path -Path 'variable:global:TimezoneBias'))
    {
        [string]$global:TimezoneBias = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalMinutes
        if($TimezoneBias -match "^-")
        {
            $TimezoneBias = $TimezoneBias.Replace('-', '+')
        }
        else
        {
            $TimezoneBias = '-' + $TimezoneBias
        }
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
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
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
Write-LogEntry -Value "START - HP BIOS settings management script (version $Version)" -Severity 1

#Resolve secure password sources (decrypt any CMS-encrypted password file into the standard password variable)
if($SetupPassword -and $SetupPasswordCmsFile)
{
    Stop-Script -ErrorMessage "Specify either the SetupPassword or the SetupPasswordCmsFile parameter, not both"
}
if($SetupPasswordCmsFile)
{
    Write-LogEntry -Value "Decrypting the setup password from the supplied CMS file" -Severity 1
    $SetupPassword = Get-CmsPassword -CmsFile $SetupPasswordCmsFile
}

#Verify the HP Client Management Script Library is installed
Write-LogEntry -Value "Verify the HP Client Management Script Library is installed" -Severity 1
$HPCMSL = Get-Module -Name HP.ClientManagement -ListAvailable | Sort-Object {[Version]$_.Version} -Descending | Select-Object -First 1
if($NULL -eq $HPCMSL)
{
    Stop-Script -ErrorMessage "The HP Client Management Script Library (HPCMSL) is not installed. Run Install-HPCMSL.ps1 before running this script"
}
try
{
    Import-Module HP.ClientManagement -Force -ErrorAction Stop
    Write-LogEntry -Value "Successfully imported the HP.ClientManagement module (version $($HPCMSL.Version))" -Severity 1
}
catch
{
    Stop-Script -ErrorMessage "Failed to import the HP.ClientManagement module" -Exception $PSItem.Exception.Message
}

#Check for HP Sure Admin (Enhanced BIOS Authentication Mode). When enabled, changes require signed payloads or a local access key rather than a password, which this script does not perform, so exit cleanly without attempting changes. GetSettings is allowed because reads do not require authorization.
$SureAdminMode = $null
try { $SureAdminMode = Get-HPBIOSSettingValue -Name "Enhanced BIOS Authentication Mode" -ErrorAction Stop } catch { $SureAdminMode = $null }
if(($SureAdminMode -eq "Enable") -and ($SetSettings -or $SetDefaults))
{
    Write-LogEntry -Value "HP Sure Admin (Enhanced BIOS Authentication Mode) is enabled. This script manages password-based BIOS authentication only. Use HP's Sure Admin tooling (signed payloads or the local access key) on Sure Admin enabled devices." -Severity 2
    Write-Output "HP Sure Admin is enabled. No settings actions were taken."
    Write-LogEntry -Value "END - HP BIOS settings management script" -Severity 1
    Exit 0
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1
if($GetSettings -and ($SetSettings -or $SetDefaults))
{
    Stop-Script -ErrorMessage "Cannot specify the GetSettings and SetSettings or SetDefaults parameters at the same time"
}
if(!($GetSettings -or $SetSettings -or $SetDefaults))
{
    Stop-Script -ErrorMessage "One of the GetSettings or SetSettings or SetDefaults parameters must be specified when running this script"
}
if($SetSettings -and !($Settings -or $CsvPath))
{
    Stop-Script -ErrorMessage "Settings must be specified using either the Settings variable in the script or the CsvPath parameter"
}
if($SetSettings -and $SetDefaults)
{
    Write-LogEntry -Value "Both the SetSettings and SetDefaults parameters have been used. The SetDefaults parameter will override any other settings" -Severity 2
}
if(($SetDefaults -and $CsvPath) -and !($SetSettings))
{
    Write-LogEntry -Value "The CsvPath parameter has been specified without the SetSettings parameter. The CSV file will be ignored" -Severity 2
}
Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Set counters to 0
if($SetSettings -or $SetDefaults)
{
    $AlreadySet = 0
    $SuccessSet = 0
    $FailSet = 0
    $NotFound = 0
}

#Get the current password status
if($SetSettings -or $SetDefaults)
{
    Write-LogEntry -Value "Check current BIOS setup password status" -Severity 1
    if(Get-HPBIOSSetupPasswordIsSet)
    {
        $PasswordCheck = 1
        #Setup password set but parameter not specified
        if([String]::IsNullOrEmpty($SetupPassword))
        {
            Stop-Script -ErrorMessage "The BIOS setup password is set, but no password was supplied. Use the SetupPassword parameter when a password is set"
        }
        #Verify the supplied setup password is correct (setting the password to itself succeeds only when it is the current password)
        try
        {
            Set-HPBIOSSetupPassword -NewPassword $SetupPassword -Password $SetupPassword -ErrorAction Stop
            Write-LogEntry -Value "The specified setup password matches the currently set password" -Severity 1
        }
        catch
        {
            Stop-Script -ErrorMessage "The specified setup password does not match the currently set password"
        }
    }
    else
    {
        $PasswordCheck = 0
        Write-LogEntry -Value "The BIOS setup password is not currently set" -Severity 1
    }
}

#Get the current settings
if($GetSettings)
{
    #Exclude password entries, which are returned by Get-HPBIOSSettingsList but are not configurable settings
    $SettingList = Get-HPBIOSSettingsList | Where-Object { $_.CimClass.CimClassName -ne 'HPBIOS_BIOSPassword' } | Sort-Object Name
    $SettingObject = ForEach($Setting in $SettingList){
        $SetValue = $Setting.Value
        $PossibleValue = ""
        #Enumeration values are a comma separated list with the currently set value marked by a leading asterisk (some HP models return values with leading whitespace; trim so the asterisk-prefix check works reliably)
        if($SetValue -match '\*')
        {
            $SettingSplit = $SetValue.Split(',') | ForEach-Object { $_.Trim() }
            $SplitCount = 0
            while($SplitCount -lt $SettingSplit.Count)
            {
                if($SettingSplit[$SplitCount].StartsWith('*'))
                {
                    $SetValue = ($SettingSplit[$SplitCount]).Substring(1)
                    break
                }
                else
                {
                    $SplitCount++
                }
            }
            #Build the list of possible values (strip the asterisk that marks the currently set value)
            $PossibleValue = ($SettingSplit | ForEach-Object { $_.TrimStart('*') }) -join ','
        }
        [PSCustomObject]@{
            Name = $Setting.Name
            Value = $SetValue
            PossibleValue = $PossibleValue
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
#Set settings
if($SetSettings -or $SetDefaults)
{
    if($CsvPath)
    {
        Clear-Variable Settings -ErrorAction SilentlyContinue
        $Settings = Import-Csv -Path $CsvPath
    }
    #Set HP BIOS settings - password is set
    if($PasswordCheck -eq 1)
    {
        if($SetSettings)
        {
            if($CsvPath)
            {
                ForEach($Setting in $Settings){
                    Set-HPBiosSetting -Name $Setting.Name -Value $Setting.Value -Password $SetupPassword
                }
            }
            else
            {
                ForEach($Setting in $Settings){
                    $Data = $Setting.Split(',')
                    Set-HPBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim() -Password $SetupPassword
                }
            }
        }
        if($SetDefaults)
        {
            Set-HPBiosSetting -Defaults -Password $SetupPassword
        }
    }
    #Set HP BIOS settings, password is not set
    else
    {
        if($SetSettings)
        {
            if($CsvPath)
            {
                ForEach($Setting in $Settings){
                    Set-HPBiosSetting -Name $Setting.Name -Value $Setting.Value
                }
            }
            else
            {
                ForEach($Setting in $Settings){
                    $Data = $Setting.Split(',')
                    Set-HPBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim()
                }
            }
        }
        if($SetDefaults)
        {
            Set-HPBiosSetting -Defaults
        }
    }
}

#Display results
if($SetSettings)
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
if($SetDefaults)
{
    if($DefaultSet -eq $True)
    {
        Write-Output "Successfully loaded default BIOS settings"
    }
    else
    {
        Write-Output "Failed to load default BIOS settings"
    }
}
Write-Output "HP BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - HP BIOS settings management script" -Severity 1

<#
    .DESCRIPTION
        Automatically configure Dell BIOS settings

    .PARAMETER GetSettings
        Instruct the script to get a list of current BIOS settings

    .PARAMETER SetSettings
        Instruct the script to set BIOS settings

    .PARAMETER CsvPath
        The path to the CSV file to be imported or exported

    .PARAMETER AdminPassword
        The current BIOS password

    .EXAMPLE
        #Set BIOS settings supplied in the script
        Manage-DellBiosSettings.ps1 -SetSettings -AdminPassword ExamplePassword

        #Set BIOS settings supplied in a CSV file
        Manage-DellBiosSettings.ps1 -SetSettings -CsvPath C:\Temp\Settings.csv -AdminPassword ExamplePassword

        #Output a list of current BIOS settings to the screen
        Manage-DellBiosSettings.ps1 -GetSettings

        #Output a list of current BIOS settings to a CSV file
        Manage-DellBiosSettings.ps1 -GetSettings -CsvPath C:\Temp\Settings.csv

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/dell-bios-settings-management/
        Modified: 2020-02-21
#>

#Parameters ===================================================================================================================

param(
    [Parameter(Mandatory=$false)][Switch]$GetSettings,
    [Parameter(Mandatory=$false)][Switch]$SetSettings,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$AdminPassword,
    [ValidateScript({
        if($_ -notmatch "(\.csv)")
        {
            throw "The specified file must be a .csv file"
        }
        return $true
    })]
    [System.IO.FileInfo]$CsvPath
)

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

#Determine if a task sequence is currently running
Function Get-TaskSequenceStatus
{
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

		if($NULL -eq $SMSTSType)
		{
			return $False
		}
		else
		{
			return $True
		}
	}
}

#Set a specific Dell BIOS setting
Function Set-DellBiosSetting
{
    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Name,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Value,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Password
    )

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

#Write data to a CMTrace compatible log file. (Credit to SCConfigMgr - https://www.scconfigmgr.com/)
Function Write-LogEntry
{
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
		[string]$FileName = "Manage-DellBiosSettings.log"
	)
    # Determine log file location
    $LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName

    # Construct time stamp for log entry
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

    # Construct date for log entry
    $Date = (Get-Date -Format "MM-dd-yyyy")

    # Construct context for log entry
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

    # Construct final log entry
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-DellBiosSettings"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"

    # Add value to log file
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
	$LogsDirectory = "$ENV:ProgramData\BiosScripts\Dell"
	if(!(Test-Path -PathType Container $LogsDirectory))
	{
		New-Item -Path $LogsDirectory -ItemType "Directory" -Force | Out-Null
	}
}
Write-Output "Log path set to $LogsDirectory\Manage-DellBiosSettings.log"
Write-LogEntry -Value "START - Dell BIOS settings management script" -Severity 1

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
    $LocalVersion = Get-Package DellBIOSProvider -ErrorAction Stop | Select-Object -ExpandProperty Version
}
catch
{
    $Local = $true
    if(Test-Path "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider")
    {
        $LocalVersion = Get-Content "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider\DellBIOSProvider.psd1" | Select-String "ModuleVersion ="
        $LocalVersion = (([regex]".*'(.*)'").Matches($LocalVersion))[0].Groups[1].Value
        if($NULL -ne $LocalVersion)
        {
            Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "DellBIOSProvider module not found on the local machine" -Severity 3
            throw "DellBIOSProvider module not found on the local machine"
        }
    }
    else
    {
        Write-LogEntry -Value "DellBIOSProvider module not found on the local machine" -Severity 3
        throw "DellBIOSProvider module not found on the local machine"
    }
}
if(($NULL -ne $LocalVersion) -and (!($Local)))
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
        Write-LogEntry -Value "Failed to import the DellBIOSProvider module" -Severity 3
        throw "Failed to import the DellBIOSProvider module"
    }
    if(!($Error))
    {
        Write-LogEntry -Value "Successfully imported the DellBIOSProvider module" -Severity 1
    }
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1

if($GetSettings -and $SetSettings)
{
	$ErrorMsg = "Cannot specify the GetSettings and SetSettings parameters at the same time"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(!($GetSettings -or $SetSettings))
{
	$ErrorMsg = "One of the GetSettings or SetSettings parameters must be specified when running this script"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if($SetSettings -and !($Settings -or $CsvPath))
{
	$ErrorMsg = "Settings must be specified using either the Settings variable in the script or the CsvPath parameter"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}

Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Set counters to 0
if($SetSettings)
{
    $AlreadySet = 0
    $SuccessSet = 0
    $FailSet = 0
    $NotFound = 0
}

#Get the current password status
if($SetSettings)
{
    Write-LogEntry -Value "Get the current password state" -Severity 1
    $AdminPasswordCheck = Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet | Select-Object -ExpandProperty CurrentValue

    if($AdminPasswordCheck -eq "True")
    {
        Write-LogEntry -Value "The Admin password is currently set" -Severity 1

        #Setup password set but parameter not specified
        if([String]::IsNullOrEmpty($AdminPassword))
        {
            Write-LogEntry -Value "The Admin password is set, but no password was supplied. Use the AdminPassword parameter when a password is set" -Severity 3
            throw "The Admin password is set, but no password was supplied. Use the AdminPassword parameter when a password is set"
        }
        #Setup password set correctly
        try
        {
            Set-Item -Path DellSmbios:\Security\AdminPassword $AdminPassword -Password $AdminPassword -ErrorAction Stop
        }
        catch
        {
            $AdminPasswordCheck = "Failed"
            Write-LogEntry -Value "The specified Admin password does not match the currently set password" -Severity 3
            throw "The specified Admin password does not match the currently set password"
        }
        if(!($AdminPasswordCheck))
        {
            Write-LogEntry -Value "The specified Admin password matches the currently set password" -Severity 1
        }
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
        $SettingList += Get-ChildItem -Path "DellSmbios:\$($Category)" | Select-Object Attribute,CurrentValue,PSChildName
    }
}

#Get the current settings
if($GetSettings)
{
    foreach($Category in $SettingCategory){
        $SettingList += Get-ChildItem -Path "DellSmbios:\$($Category)" | Select-Object Attribute,CurrentValue
    }
    $WarningPreference = $WarnPref #Revert WarningPreference back to the original value
    $SettingList = $SettingList | Sort-Object Attribute
    $SettingObject = ForEach($Setting in $SettingList){
        [PSCustomObject]@{
            Name = $Setting.Attribute
            Value = $Setting.CurrentValue
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

if($SetSettings)
{
    if($CsvPath)
    {
        Clear-Variable Settings -ErrorAction SilentlyContinue
        $Settings = Import-Csv -Path $CsvPath
    }

    #Set Dell BIOS settings - password is set
    if($AdminPasswordCheck -eq "True")
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
    #Set Dell BIOS settings - password is not set
    else
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
Write-Output "Dell BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - Dell BIOS settings management script" -Severity 1
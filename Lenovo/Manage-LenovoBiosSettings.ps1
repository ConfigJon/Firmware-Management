<#
    .DESCRIPTION
        Automatically configure Lenovo BIOS settings

    .PARAMETER GetSettings
        Instruct the script to get a list of current BIOS settings

    .PARAMETER SetSettings
        Instruct the script to set BIOS settings

    .PARAMETER CsvPath
        The path to the CSV file to be imported or exported

    .PARAMETER SupervisorPassword
        The current BIOS password

    .EXAMPLE
        #Set BIOS settings supplied in the script
        Manage-LenovoBiosSettings.ps1 -SetSettings -SupervisorPassword ExamplePassword

        #Set BIOS settings supplied in a CSV file
        Manage-LenovoBiosSettings.ps1 -SetSettings -CsvPath C:\Temp\Settings.csv -SupervisorPassword ExamplePassword

        #Output a list of current BIOS settings to the screen
        Manage-LenovoBiosSettings.ps1 -GetSettings

        #Output a list of current BIOS settings to a CSV file
        Manage-LenovoBiosSettings.ps1 -GetSettings -CsvPath C:\Temp\Settings.csv

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/lenovo-bios-settings-management/
        Modified: 2020-02-21

    .CHANGELOG
        2019-11-04 - Added additional logging. Changed the default log path to $ENV:ProgramData\BiosScripts\Lenovo.
        2020-02-10 - Fixed a bug where the script would ignore the supplied Supervisior Password when attempting to change settings.
        2020-02-21 - Added the ability to get a list of current BIOS settings on a system via the GetSettings parameter
                     Added the ability to read settings from or write settings to a csv file with the CsvPath parameter
                     Added the SetSettings parameter to indicate that the script should attempt to set settings
                     Changed the $Settings array in the script to be comma seperated instead of semi-colon seperated
                     Updated formatting
#>

#Parameters ===================================================================================================================

param(
    [Parameter(Mandatory=$false)][Switch]$GetSettings,
    [Parameter(Mandatory=$false)][Switch]$SetSettings,    
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SupervisorPassword,
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
    "PXE IPV4 Network Stack,Enabled",
    "IPv4NetworkStack,Enable",
    "PXE IPV6 Network Stack,Enabled",
    "IPv6NetworkStack,Enable",
    "Intel(R) Virtualization Technology,Enabled",
    "VirtualizationTechnology,Enable",
    "VT-d,Enabled",
    "VTdFeature,Enable",
    "Enhanced Power Saving Mode,Disabled",
    "Wake on LAN,Primary",
    "Require Admin. Pass. For F12 Boot,Yes",
    "Physical Presence for Provisioning,Disabled",
    "PhysicalPresenceForTpmProvision,Disable",
    "Physical Presnce for Clear,Disabled",
    "PhysicalPresenceForTpmClear,Disable",
    "Boot Up Num-Lock Status,Off"
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

#Set a specific Lenovo BIOS setting
Function Set-LenovoBiosSetting
{
    param (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Name,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Value,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Password
    )

    #Ensure the specified setting exists and get the possible values
    $CurrentSetting = $SettingList | Where-Object CurrentSetting -Like "$Name*" | Select-Object -ExpandProperty CurrentSetting
    if($NULL -ne $CurrentSetting)
    {
        #Check how the CurrentSetting data is formatted, then split the setting and current value
        if($CurrentSetting -match ';')
        {
            $FormattedSetting = $CurrentSetting.Substring(0, $CurrentSetting.IndexOf(';'))
            $CurrentSettingSplit = $FormattedSetting.Split(',')
        }
        else
        {
            $CurrentSettingSplit = $CurrentSetting.Split(',')
        }

        #Setting is already set to specified value
        if($CurrentSettingSplit[1] -eq $Value)
        {
            Write-LogEntry -Value "Setting ""$Name"" is already set to ""$Value""" -Severity 1
            $Script:AlreadySet++
        }
        #Setting is not set to specified value
        else
        {
            if(!([String]::IsNullOrEmpty($Password)))
            {
                $SettingResult = ($Interface.SetBIOSSetting("$Name,$Value,$Password,ascii,us")).Return
            }
            else
            {
                $SettingResult = ($Interface.SetBIOSSetting("$Name,$Value")).Return
            }
            

            if($SettingResult -eq "Success")
            {
                Write-LogEntry -Value "Successfully set ""$Name"" to ""$Value""" -Severity 1
                $Script:SuccessSet++
            }
            else
            {
                Write-LogEntry -Value "Failed to set ""$Name"" to ""$Value"". Return code: $SettingResult" -Severity 3
                $Script:FailSet++
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
		[string]$FileName = "Manage-LenovoBiosSettings.log"
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
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-LenovoBiosSettings"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
		
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
	$LogsDirectory = "$ENV:ProgramData\BiosScripts\Lenovo"
	if(!(Test-Path -PathType Container $LogsDirectory))
	{
		New-Item -Path $LogsDirectory -ItemType "Directory" -Force | Out-Null
	}
}
Write-Output "Log path set to $LogsDirectory\Manage-LenovoBiosSettings.log"
Write-LogEntry -Value "START - Lenovo BIOS settings management script" -Severity 1

#Connect to the Lenovo_BiosSetting WMI class
$Error.Clear()
try
{
    Write-LogEntry -Value "Connect to the Lenovo_BiosSetting WMI class" -Severity 1
    $SettingList = Get-WmiObject -Namespace root\wmi -Class Lenovo_BiosSetting
}
catch
{
    Write-LogEntry -Value "Unable to connect to the Lenovo_BiosSetting WMI class" -Severity 3
    throw "Unable to connect to the Lenovo_BiosSetting WMI class"
}
if(!($Error))
{
	Write-LogEntry -Value "Successfully connected to the Lenovo_BiosSetting WMI class" -Severity 1
}

#Connect to the Lenovo_SetBiosSetting WMI class
$Error.Clear()
try
{
    Write-LogEntry -Value "Connect to the Lenovo_SetBiosSetting WMI class" -Severity 1
    $Interface = Get-WmiObject -Namespace root\wmi -Class Lenovo_SetBiosSetting
}
catch
{
    Write-LogEntry -Value "Unable to connect to the Lenovo_SetBiosSetting WMI class" -Severity 3
    throw "Unable to connect to the Lenovo_SetBiosSetting WMI class"
}
if(!($Error))
{
	Write-LogEntry -Value "Successfully connected to the Lenovo_SetBiosSetting WMI class" -Severity 1
}

#Connect to the Lenovo_SaveBiosSettings WMI class
$Error.Clear()
try
{
    Write-LogEntry -Value "Connect to the Lenovo_SaveBiosSettings WMI class" -Severity 1
    $SaveSettings = Get-WmiObject -Namespace root\wmi -Class Lenovo_SaveBiosSettings
}
catch
{
    Write-LogEntry -Value "Unable to connect to the Lenovo_SaveBiosSettings WMI class" -Severity 3
    throw "Unable to connect to the Lenovo_SaveBiosSettings WMI class"
}
if(!($Error))
{
	Write-LogEntry -Value "Successfully connected to the Lenovo_SaveBiosSettings WMI class" -Severity 1
}

#Connect to the Lenovo_BiosPasswordSettings WMI class
$Error.Clear()
try
{
	Write-LogEntry -Value "Connect to the Lenovo_BiosPasswordSettings WMI class" -Severity 1
	$PasswordSettings = Get-WmiObject -Namespace root\wmi -Class Lenovo_BiosPasswordSettings
}
catch
{
	Write-LogEntry -Value "Unable to connect to the Lenovo_BiosPasswordSettings WMI class" -Severity 3
	throw "Unable to connect to the Lenovo_BiosPasswordSettings WMI class"
}
if(!($Error))
{
	Write-LogEntry -Value "Successfully connected to the Lenovo_BiosPasswordSettings WMI class" -Severity 1
}

#Connect to the Lenovo_SetBiosPassword WMI class
$Error.Clear()
try
{
	Write-LogEntry -Value "Connect to the Lenovo_SetBiosPassword WMI class" -Severity 1
	$PasswordSet = Get-WmiObject -Namespace root\wmi -Class Lenovo_SetBiosPassword
}
catch
{
	Write-LogEntry -Value "Unable to connect to the Lenovo_SetBiosPassword WMI class" -Severity 3
	throw "Unable to connect to the Lenovo_BiosPasswordSettings WMI class"
}
if(!($Error))
{
	Write-LogEntry -Value "Successfully connected to the Lenovo_SetBiosPassword WMI class" -Severity 1
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
    Write-LogEntry -Value "Check current BIOS supervisor password status" -Severity 1
    $PasswordCheck = $PasswordSettings.PasswordState
    if(($PasswordCheck -eq 2) -or ($PasswordCheck -eq 3) -or($PasswordCheck -eq 6) -or($PasswordCheck -eq 7))
    {
        $SupervisorPasswordSet = $true
    }
    else
    {
        $SupervisorPasswordSet = $false
    }
    
    if($SupervisorPasswordSet)
    {
        #Supervisor password set but parameter not specified
        if([String]::IsNullOrEmpty($SupervisorPassword))
        {
            Write-LogEntry -Value "The BIOS supervisor password is set, but no password was supplied. Use the SupervisorPassword parameter when a password is set" -Severity 3
            throw "The BIOS supervisor password is set, but no password was supplied. Use the SupervisorPassword parameter when a password is set"
        }
        #Supervisor password set correctly
        if($PasswordSet.SetBiosPassword("pap,$SupervisorPassword,$SupervisorPassword,ascii,us").Return -eq "Success")
	    {
		    Write-LogEntry -Value "The specified supervisor password matches the currently set password" -Severity 1
        }
        #Supervisor password not set correctly
        else
        {
            Write-LogEntry -Value "The specified supervisor password does not match the currently set password" -Severity 3
            $ReturnCode = $PasswordSet.SetBiosPassword("pap,$SupervisorPassword,$SupervisorPassword,ascii,us").Return
            Write-Output $ReturnCode
            throw "The specified supervisor password does not match the currently set password"
        }
    }
    else
    {
        Write-LogEntry -Value "The BIOS supervisor password is not currently set" -Severity 1
    }
}

#Get the current settings
if($GetSettings)
{
    $SettingList = $SettingList | Select-Object CurrentSetting | Sort-Object CurrentSetting

    $SettingObject = ForEach($Setting in $SettingList){
        #Split the current values
        $SettingSplit = ($Setting.CurrentSetting).Split(',')

        if($SettingSplit[0] -and $SettingSplit[1])
        {
            [PSCustomObject]@{
                Name = $SettingSplit[0]
                Value = $SettingSplit[1]
            }
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

    #Set Lenovo BIOS settings - supervisor password is set
    if($SupervisorPasswordSet)
    {
        if($CsvPath)
        {
            ForEach($Setting in $Settings){
                Set-LenovoBiosSetting -Name $Setting.Name -Value $Setting.Value -Password $SupervisorPassword
            }
        }
        else
        {
            ForEach($Setting in $Settings){
                $Data = $Setting.Split(',')
                Set-LenovoBiosSetting -Name $Data[0] -Value $Data[1].Trim() -Password $SupervisorPassword
            }
        }
    }

    #Set Lenovo BIOS settings - supervisor password is not set
    else
    {
        if($CsvPath)
        {
            ForEach($Setting in $Settings){
                Set-LenovoBiosSetting -Name $Setting.Name -Value $Setting.Value
            }
        }
        else
        {
            ForEach($Setting in $Settings){
                $Data = $Setting.Split(',')
                Set-LenovoBiosSetting -Name $Data[0] -Value $Data[1].Trim()
            }
        }
    }
}


#If settings were set, save the changes
if($SetSettings -and $SuccessSet -gt 0)
{
    if($SupervisorPasswordSet)
    {
        $ReturnCode = $SaveSettings.SaveBiosSettings("$($SupervisorPassword),ascii,us") | Select-Object -ExpandProperty value
    }
    else
    {
        $ReturnCode = $SaveSettings.SaveBiosSettings() | Select-Object -ExpandProperty value
    }
    
    if(($null -eq $ReturnCode) -or ($ReturnCode -eq 'Success'))
    {
        Write-Output -Value "Successfully saved BIOS settings."
        Write-LogEntry -Value "Successfully saved BIOS settings." -Severity 1
    }
    else
    {
        Write-Output "Failed to save BIOS settings. Return Code = `"$($ReturnCode)`""
        Write-LogEntry -Value "Failed to save BIOS settings. Return Code = `"$($ReturnCode)`"" -Severity 3
        Exit 1
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
Write-Output "Lenovo BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - Lenovo BIOS settings management script" -Severity 1
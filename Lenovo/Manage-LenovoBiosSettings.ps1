<#
    .DESCRIPTION
        Automatically configure Lenovo BIOS settings

    .PARAMETER SupervisorPassword
        The current BIOS password

    .EXAMPLE
        Manage-LenovoBiosSettings.ps1 -SupervisorPassword ExamplePassword

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/lenovo-bios-settings-management/
        Modified: 02/10/2020

    .CHANGELOG
        11/04/2019 - Added additional logging. Changed the default log path to $ENV:ProgramData\BiosScripts\Lenovo.
        02/10/2020 - Fixed a bug where the script would ignore the supplied Supervisior Password when attempting to change settings.
#>

#Parameters ===================================================================================================================

param ([Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SupervisorPassword)

#List of settings to be configured ============================================================================================
#==============================================================================================================================
$Settings = (
    "PXE IPV4 Network Stack;Enabled", #Enabled,Disabled
    "IPv4NetworkStack;Enable", #Enable,Disable
    "PXE IPV6 Network Stack;Enabled", #Enabled,Disabled
    "IPv6NetworkStack;Enable", #Enable,Disable
    "Intel(R) Virtualization Technology;Enabled", #Enabled,Disabled
    "VirtualizationTechnology;Enable", #Enable,Disable
    "VT-d;Enabled", #Enabled,Disabled
    "VTdFeature;Enable", #Enable,Disable
    "Enhanced Power Saving Mode;Disabled", #Enabled,Disabled
    "Wake on LAN;Primary", #Primary,Automatic,Disabled
    "Require Admin. Pass. For F12 Boot;Yes", #Yes,No
    "Physical Presence for Provisioning;Disabled", #Enabled,Disabled
    "PhysicalPresenceForTpmProvision;Disable", #Disable,Enable
    "Physical Presnce for Clear ;Disabled", #Enabled,Disabled
    "PhysicalPresenceForTpmClear;Disable", #Disable,Enable
    "Boot Up Num-Lock Status;Off" #On,Off
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

	if ($NULL -eq $TSEnv)
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

		if ($NULL -eq $SMSTSType)
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
    if ($NULL -ne $CurrentSetting)
    {
        #Check how the CurrentSetting data is formatted, then split the setting and current value
        if ($CurrentSetting -match ';')
        {
            $FormattedSetting = $CurrentSetting.Substring(0, $CurrentSetting.IndexOf(';'))
            $CurrentSettingSplit = $FormattedSetting.Split(',')
        }
        else
        {
            $CurrentSettingSplit = $CurrentSetting.Split(',')
        }

        #Setting is already set to specified value
        if ($CurrentSettingSplit[1] -eq $Value)
        {
            Write-LogEntry -Value "Setting ""$Name"" is already set to ""$Value""" -Severity 1
            $Script:AlreadySet++
        }
        #Setting is not set to specified value
        else
        {
            if (!([String]::IsNullOrEmpty($Password)))
            {
                $SettingResult = ($Interface.SetBIOSSetting("$Name,$Value,$Password,ascii,us")).Return
            }
            else
            {
                $SettingResult = ($Interface.SetBIOSSetting("$Name,$Value")).Return
            }
            

            if ($SettingResult -eq "Success")
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
	param (
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
    if (-not(Test-Path -Path 'variable:global:TimezoneBias'))
    {
        [string]$global:TimezoneBias = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalMinutes
        if ($TimezoneBias -match "^-")
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
if (Get-TaskSequenceStatus)
{
	$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
	$LogsDirectory = $TSEnv.Value("_SMSTSLogPath")
}
else
{
	$LogsDirectory = "$ENV:ProgramData\BiosScripts\Lenovo"
	if (!(Test-Path -PathType Container $LogsDirectory))
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
if (!($Error))
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
if (!($Error))
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
if (!($Error))
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
if (!($Error))
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
if (!($Error))
{
	Write-LogEntry -Value "Successfully connected to the Lenovo_SetBiosPassword WMI class" -Severity 1
}

#BIOS password checks
Write-LogEntry -Value "Check current BIOS supervisor password status" -Severity 1
$PasswordCheck = $PasswordSettings.PasswordState

if (($PasswordCheck -eq 2) -or ($PasswordCheck -eq 3) -or($PasswordCheck -eq 6) -or($PasswordCheck -eq 7))
{
    #Supervisor password set but parameter not specified
    if ([String]::IsNullOrEmpty($SupervisorPassword))
    {
        Write-LogEntry -Value "The BIOS supervisor password is set, but no password was supplied. Use the SupervisorPassword parameter when a password is set" -Severity 3
        throw "The BIOS supervisor password is set, but no password was supplied. Use the SupervisorPassword parameter when a password is set"
    }
    #Supervisor password set correctly
    if ($PasswordSet.SetBiosPassword("pap,$SupervisorPassword,$SupervisorPassword,ascii,us").Return -eq "Success")
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

#Set counters to 0
$AlreadySet = 0
$SuccessSet = 0
$FailSet = 0
$NotFound = 0

#Set Lenovo BIOS settings - supervisor password is set
if (($PasswordCheck -eq 2) -or ($PasswordCheck -eq 3) -or($PasswordCheck -eq 6) -or($PasswordCheck -eq 7))
{
    ForEach($Setting in $Settings){
        $Data = $Setting.Split(';')
        Set-LenovoBiosSetting -Name $Data[0] -Value $Data[1].Trim() -Password $SupervisorPassword
    }
}

#Set Lenovo BIOS settings - supervisor password is not set
else
{
    ForEach($Setting in $Settings){
        $Data = $Setting.Split(';')
        Set-LenovoBiosSetting -Name $Data[0] -Value $Data[1].Trim()
    }
}

#If settings were set, save the changes
if ($SuccessSet -gt 0)
{
    $SaveSettings.SaveBiosSettings() | Out-Null
}

#Display results
Write-Output "$AlreadySet settings already set correctly"
Write-LogEntry -Value "$AlreadySet settings already set correctly" -Severity 1
Write-Output "$SuccessSet settings successfully set"
Write-LogEntry -Value "$SuccessSet settings successfully set" -Severity 1
Write-Output "$FailSet settings failed to set"
Write-LogEntry -Value "$FailSet settings failed to set" -Severity 3
Write-Output "$NotFound settings not found"
Write-LogEntry -Value "$NotFound settings not found" -Severity 2
Write-Output "Lenovo BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - Lenovo BIOS settings management script" -Severity 1
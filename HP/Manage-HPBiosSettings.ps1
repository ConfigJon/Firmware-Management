<#
    .DESCRIPTION
        Automatically configure HP BIOS settings

        SetBIOSSetting Return Codes
        0 - Success
        1 - Not Supported
        2 - Unspecified Error
        3 - Timeout
        4 - Failed - (Check for typos in the setting value)
        5 - Invalid Parameter
        6 - Access Denied - (Check that the BIOS password is correct)
    
    .PARAMETER SetupPassword
        The current BIOS password

    .EXAMPLE
        Manage-HPBiosSettings.ps1 -SetupPassword ExamplePassword

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/hp-bios-settings-management/
        Modified: 7/29/2019
#>

#Parameters ===================================================================================================================

param ([Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPassword)

#List of settings to be configured ============================================================================================
#==============================================================================================================================
$Settings = (
    #Power Settings
    "After Power Loss;Off", #Off,On,Previous State
    "Deep S3;Off", #Off,On,Auto
    "Deep Sleep;Off", #Off,On,Auto
    "Power state after power loss;Power Off", #Power On,Power Off,Previous State
    "S4/S5 Max Power Savings;Disable", #Disable,Enable
    "S5 Maximum Power Savings;Disable", #Disable,Enable
    "Wake unit from sleep when lid is opened;Enable", #Disable,Enable
    "Wake when Lid is Opened;Enable", #Disable,Enable
    #Integrated Device Settings
    "Audio Alerts During Boot;Enable", #Disable,Enable
    "Audio Device;Enable", #Disable,Enable
    "Fingerprint Device;Disable", #Disable,Enable
    "Integrated Audio;Enable", #Disable,Enable
    "Integrated Camera;Enable", #Disable,Enable
    "Integrated Microphone;Enable", #Disable,Enable
    "Internal speaker;Enable", #Disable,Enable
    "Internal speakers;Enable", #Disable,Enable
    "Microphone;Enable", #Disable,Enable
    "Num Lock State at Power-On;Off", #Off,On
    "NumLock on at boot;Disable", #Disable,Enable
    "Numlock state at boot;Off", #On,Off
    "System Audio;Device available", #Device available,Device hidden
    #Virtualization Settings
    "Intel VT for Directed I/O (VT-d);Enable", #Disable,Enable
    "Intel(R) VT-d;Enable", #Disable,Enable
    "Virtualization Technology;Enable", #Disable,Enable,Reset to default
    "Virtualization Technology (VTx);Enable", #Disable,Enable,Reset to default
    "Virtualization Technology (VT-x);Enable", #Disable,Enable
    "Virtualization Technology (VTx/VTd);Enable", #Disable,Enable
    "Virtualization Technology Directed I/O (VTd);Enable", #Disable,Enable
    "Virtualization Technology Directed I/O (VT-d2);Enable", #Disable,Enable
    "Virtualization Technology for Directed I/O;Enable", #Disable,Enable,Reset to default
    "Virtualization Technology for Directed I/O (VTd);Enable", #Disable,Enable,Reset to default
    #Security Settings
    "Password prompt on F9 & F12;Enable", #Enable,Disable
    "Password prompt on F9 F11 & F12;Enable", #Enable,Disable
    "Prompt for Admin password on F9 (Boot Menu);Enable", #Disable,Enable
    "Prompt for Admin password on F11 (System Recovery);Enable", #Disable,Enable
    "Prompt for Admin password on F12 (Network Boot);Enable", #Disable,Enable
    #PXE Boot Settings
    "Network (PXE) Boot;Enable", #Disable,Enable
    "Network (PXE) Boot Configuration;IPv4 Before IPv6", #IPv4 Before IPv6,IPv6 Before IPv4,IPv4 Disabled,IPv6 Disabled
    "Network Boot;Enable", #Disable,Enable
    "Network Service Boot;Enable", #Disable,Enable
    "PXE Internal IPV4 NIC boot;Enable", #Disable,Enable
    "PXE Internal IPV6 NIC boot;Enable", #Disable,Enable
    "PXE Internal NIC boot;Enable", #Disable,Enable
    #Wake on LAN Settings
    "Remote Wakeup Boot Source;Local Hard Drive", #Remote Server,Local Hard Drive
    "S4/S5 Wake on LAN;Enable", #Disable,Enable
    "S5 Wake On LAN;Boot to Hard Drive", #Disable,Boot to Network,Boot to Hard Drive
    "Wake On LAN;Boot to Hard Drive", #Disabled,Boot to Network,Boot to Hard Drive,Boot to Normal Boot Order
    "Wake on LAN on DC mode;Enable", #Disable,Enable
    "Wake on WLAN;Enable", #Disable,Enable
    #Other Settings
    "Fast Boot;Enable", #Disable,Enable
    "LAN / WLAN Auto Switching;Enable", #Disable,Enable
    "LAN/WLAN Switching;Enable", #Disable,Enable
    "Swap Fn and Ctrl (Keys);Disable" #Disable,Enable
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

#Set a specific HP BIOS setting
Function Set-HPBiosSetting
{
    param (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Name,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Value,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Password
    )

    #Ensure the specified setting exists and get the possible values
    $CurrentSetting = $SettingList | Where-Object Name -eq $Name | Select-Object -ExpandProperty Value
    if ($NULL -ne $CurrentSetting)
    {
        #Split the current values
        $CurrentSettingSplit = $CurrentSetting.Split(',')

        #Find the currently set value
        $Count = 0
        while($Count -lt $CurrentSettingSplit.Count)
        {
            if ($CurrentSettingSplit[$Count].StartsWith('*'))
            {
                $CurrentValue = $CurrentSettingSplit[$Count]
                break
            }
            else
            {
                $Count++
            }
        }
        #Setting is already set to specified value
        if ($CurrentValue.Substring(1) -eq $Value)
        {
            Write-LogEntry -Value "Setting ""$Name"" is already set to ""$Value""" -Severity 1
            $Script:AlreadySet++
        }
        #Setting is not set to specified value
        else
        {
            if (!([String]::IsNullOrEmpty($Password)))
            {
                $SettingResult = ($Interface.SetBIOSSetting($Name,$Value,"<utf-16/>" + $Password)).Return
            }
            else
            {
                $SettingResult = ($Interface.SetBIOSSetting($Name,$Value)).Return
            }
            

            if ($SettingResult -eq 0)
            {
                Write-LogEntry -Value "Successfully set ""$Name"" to ""$Value""" -Severity 1
                $Script:SuccessSet++
            }
            else
            {
                Write-LogEntry -Value "Failed to set ""$Name"" to ""$Value"". Return code $SettingResult" -Severity 3
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
		[string]$FileName = "Manage-HPBiosSettings.log"
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
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-HPBiosSettings"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
		
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
	$LogsDirectory = "$ENV:SystemRoot\Temp\HPBiosScripts"
	if (!(Test-Path -PathType Container $LogsDirectory))
	{
		New-Item -Path $LogsDirectory -ItemType "Directory" -Force | Out-Null
	}
}
Write-Output "Log path set to $LogsDirectory\Manage-HPBiosSettings.log"
Write-LogEntry -Value "START - HP BIOS settings management script" -Severity 1

#Connect to the HP_BIOSEnumeration WMI class
try
{
    Write-LogEntry -Value "Connect to the HP_BIOSEnumeration WMI class" -Severity 1
    $SettingList = Get-WmiObject -Namespace root\HP\InstrumentedBIOS -Class HP_BIOSEnumeration
}
catch
{
    Write-LogEntry -Value "Unable to connect to the HP_BIOSEnumeration WMI class" -Severity 3
    throw "Unable to connect to the HP_BIOSEnumeration WMI class"
}

#Connect to the HP_BIOSSettingInterface WMI class
try
{
    Write-LogEntry -Value "Connect to the HP_BIOSSettingInterface WMI class" -Severity 1
    $Interface = Get-WmiObject -Namespace root\HP\InstrumentedBIOS -Class HP_BIOSSettingInterface
}
catch
{
    Write-LogEntry -Value "Unable to connect to the HP_BIOSSettingInterface WMI class" -Severity 3
    throw "Unable to connect to the HP_BIOSSettingInterface WMI class"
}

#Connect to the HP_BIOSSetting WMI class
try
{
    Write-LogEntry -Value "Connect to the HP_BIOSSetting WMI class" -Severity 1
    $HPBiosSetting = Get-WmiObject -Namespace root\HP\InstrumentedBIOS -Class HP_BIOSSetting
}
catch
{
    Write-LogEntry -Value "Unable to connect to the HP_BIOSSetting WMI class" -Severity 3
    throw "Unable to connect to the HP_BIOSSetting WMI class"
}

#BIOS password checks
Write-LogEntry -Value "Check current BIOS setup password status" -Severity 1
$PasswordCheck = ($HPBiosSetting | Where-Object Name -eq "Setup Password").IsSet

if ($PasswordCheck -eq 1)
{
    #Setup password set but parameter not specified
    if ([String]::IsNullOrEmpty($SetupPassword))
    {
        Write-LogEntry -Value "The BIOS setup password is set, but no password was supplied. Use the SetupPassword parameter when a password is set" -Severity 3
        throw "The BIOS setup password is set, but no password was supplied. Use the SetupPassword parameter when a password is set"
    }
    #Setup password set correctly
    if (($Interface.SetBIOSSetting("Setup Password","<utf-16/>" + $SetupPassword,"<utf-16/>" + $SetupPassword)).Return -eq 0)
	{
		Write-LogEntry -Value "The specified setup password matches the currently set password" -Severity 1
    }
    #Setup password not set correctly
    else
    {
        Write-LogEntry -Value "The specified setup password does not match the currently set password" -Severity 3
        throw "The specified setup password does not match the currently set password"
    }
}
else
{
    Write-LogEntry -Value "The BIOS setup password is not currently set" -Severity 1
}

#Set counters to 0
$AlreadySet = 0
$SuccessSet = 0
$FailSet = 0
$NotFound = 0

#Set HP BIOS settings - password is set
if ($PasswordCheck -eq 1)
{
    ForEach($Setting in $Settings){
        $Data = $Setting.Split(';')
        Set-HPBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim() -Password $SetupPassword
    }
}
#Set HP BIOS settings - password is not set
else
{
    ForEach($Setting in $Settings){
        $Data = $Setting.Split(';')
        Set-HPBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim()
    }
}

#Display results
Write-Output "$AlreadySet settings already set correctly"
Write-Output "$SuccessSet settings successfully set"
Write-Output "$FailSet settings failed to set"
Write-Output "$NotFound settings not found"
Write-Output "HP BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - HP BIOS settings management script" -Severity 1
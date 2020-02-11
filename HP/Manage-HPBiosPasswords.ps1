<#
	.DESCRIPTION
		Automatically configure HP BIOS passwords and prompt the user if manual intervention is required.
		
	.PARAMETER SetupSet
		Specify this switch to set a new setup password or change an existing setup password.

	.PARAMETER SetupClear
		Specify this swtich to clear an existing setup password. Must also specify the OldSetupPassword parameter.

	.PARAMETER PowerOnSet
		Specify this switch to set a new power on password or change an existing power on password.

	.PARAMETER PowerOnClear
		Specify this switch to clear an existing power on password. Must also specify the OldPowerOnPassword parameter.

	.PARAMETER SetupPassword
		Specify the new setup password to set.

	.PARAMETER OldSetupPassword
		Specify the old setup password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER PowerOnPassword
		Specify the new power on password to set.

	.PARAMETER OldPowerOnPassword
		Specify the old power on password(s) to be changed. Multiple passwords can be specified as a comma seperated list.
	
	.PARAMETER NoUserPrompt
		The script will run silently and will not prompt the user with a message box.

	.PARAMETER ContinueOnError
		The script will ignore any errors caused by changing or clearing the passwords. This will not suppress errors caused by parameter validation.

	.PARAMETER SMSTSPasswordRetry
		For use in a task sequence. If specified, the script will assume the script needs to run at least one more time. This will ignore password errors and suppress user prompts.

	.EXAMPLE
		Set a new setup password when no old passwords exist
		Manage-HPBiosPasswords.ps1 -SetupSet -SetupPassword <String>
	
		Set or change a setup password
		Manage-HPBiosPasswords.ps1 -SetupSet -SetupPassword <String> -OldSetupPassword <String1>,<String2>

		Clear existing setup password(s)
		Manage-HPBiosPasswords.ps1 -SetupClear -OldSetupPassword <String1>,<String2>

		Set a new setup password and set a new power on password when no old passwords exist
		Manage-HPBiosPasswords.ps1 -SetupSet -PowerOnSet -SetupPassword <String1> -PowerOnPassword <String1>

		Set or change an existing setup password and clear a power on password
		Manage-HPBiosPasswords.ps1 -SetupSet -SetupPassword <String> -OldSetupPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

		Clear existing Setup and power on passwords
		Manage-HPBiosPasswords.ps1 -SetupClear -OldSetupPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

		Set a new power on password when the setup password is already set
		Manage-HPBiosPasswords.ps1 -PowerOnSet -PowerOnPassword <String> -SetupPassword <String>

	.NOTES
		Created by: Jon Anderson (@ConfigJon)
		Reference: https://www.configjon.com/lenovo-bios-password-management/
		Modifed: 01/30/2020

	.CHANGELOG
		07/27/2019 - Formatting changes. Changed the SMSTSPasswordRetry parameter to be a switch instead of an integer value. Changed the SMSTSChangeSetup TS variable to HPChangeSetup.
					 Changed the SMSTSClearSetup TS variable to HPClearSetup. Changed the SMSTSChangePowerOn TS variable to HPChangePowerOn. Changed the SMSTSClearPowerOn TS variable to HPClearPowerOn.
		11/04/2019 - Added additional logging. Changed the default log path to $ENV:ProgramData\BiosScripts\HP. Modifed the parameter validation logic.
		01/30/2020 - Removed the SetupChange and PowerOnChange parameters. SetupSet and PowerOnSet now work to set or change a password. Changed the HPChangeSetup task sequence variable to HPSetSetup.
					 Changed the HPChangePowerOn task sequence variable to HPSetPowerOn. Updated the parameter validation checks.
#>

#Parameters ===================================================================================================================

param (
	[Parameter(Mandatory=$false)][Switch]$SetupSet,
	[Parameter(Mandatory=$false)][Switch]$SetupClear,
	[Parameter(Mandatory=$false)][Switch]$PowerOnSet,
	[Parameter(Mandatory=$false)][Switch]$PowerOnClear,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSetupPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$PowerOnPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPowerOnPassword,
	[Parameter(Mandatory=$false)][Switch]$NoUserPrompt,
	[Parameter(Mandatory=$false)][Switch]$ContinueOnError,
	[Parameter(Mandatory=$false)][Switch]$SMSTSPasswordRetry
)

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

#Create a user prompt with custom body and title text if the NoUserPrompt variable is not set
Function Start-UserPrompt
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True)][ValidateNotNullOrEmpty()][String[]]$BodyText,
        [Parameter(Mandatory=$True)][ValidateNotNullOrEmpty()][String[]]$TitleText
    )
	
    if (!($NoUserPrompt))
	{
		(New-Object -ComObject Wscript.Shell).Popup("$BodyText",0,"$TitleText",0x0 + 0x30) | Out-Null
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
		[string]$FileName = "Manage-HPBiosPasswords.log"
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
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-HPBiosPasswords"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
		
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
	$TSProgress = New-Object -ComObject Microsoft.SMS.TsProgressUI
	$LogsDirectory = $TSEnv.Value("_SMSTSLogPath")
}
else
{
	$LogsDirectory = "$ENV:ProgramData\BiosScripts\HP"
	if (!(Test-Path -PathType Container $LogsDirectory))
	{
		New-Item -Path $LogsDirectory -ItemType "Directory" -Force | Out-Null
	}
}
Write-Output "Log path set to $LogsDirectory\Manage-HPBiosPasswords.log"
Write-LogEntry -Value "START - HP BIOS password management script" -Severity 1

#Connect to the HP_BIOSSettingInterface WMI class
$Error.Clear()
try
{
    Write-LogEntry -Value "Connect to the HP_BIOSSettingInterface WMI class" -Severity 1
    $Interface = Get-WmiObject -Namespace root/hp/InstrumentedBIOS -Class HP_BIOSSettingInterface
}
catch
{
    Write-LogEntry -Value "Unable to connect to the HP_BIOSSettingInterface WMI class" -Severity 3
    throw "Unable to connect to the HP_BIOSSettingInterface WMI class"
}
if (!($Error))
{
	Write-LogEntry -Value "Successfully connected to the HP_BIOSSettingInterface WMI class" -Severity 1
}	

#Connect to the HP_BIOSSetting WMI class
$Error.Clear()
try
{
    Write-LogEntry -Value "Connect to the HP_BIOSSetting WMI class" -Severity 1
    $HPBiosSetting = Get-WmiObject -Namespace root/hp/InstrumentedBIOS -Class HP_BIOSSetting
}
catch
{
    Write-LogEntry -Value "Unable to connect to the HP_BIOSSetting WMI class" -Severity 3
    throw "Unable to connect to the HP_BIOSSetting WMI class"
}
if (!($Error))
{
	Write-LogEntry -Value "Successfully connected to the HP_BIOSSetting WMI class" -Severity 1
}	

#Get the current password status
Write-LogEntry -Value "Get the current password state" -Severity 1

$SetupPasswordCheck = ($HPBiosSetting | Where-Object Name -eq "Setup Password").IsSet
if ($SetupPasswordCheck -eq 1)
{
	Write-LogEntry -Value "The setup password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The setup password is not currently set" -Severity 1
}

$PowerOnPasswordCheck = ($HPBiosSetting | Where-Object Name -eq "Power-On Password").IsSet
if ($PowerOnPasswordCheck -eq 1)
{
	Write-LogEntry -Value "The power on password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The power on password is not currently set" -Severity 1
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1

if (($SetupSet) -and !($SetupPassword))
{
	$ErrorMsg = "When using the SetupSet switch, the SetupPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SetupSet -and $SetupPasswordCheck -eq 1) -and !($OldSetupPassword))
{
	$ErrorMsg = "When using the SetupSet switch where a setup password exists, the OldSetupPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SetupClear) -and !($OldSetupPassword))
{
	$ErrorMsg = "When using the SetupClear switch, the OldSetupPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($PowerOnSet) -and !($PowerOnPassword))
{
	$ErrorMsg = "When using the PowerOnSet switch, the PowerOnPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($PowerOnSet -and $PowerOnPasswordCheck -eq 1) -and !($OldPowerOnPassword))
{
	$ErrorMsg = "When using the PowerOnSet switch where a power on password exists, the OldPowerOnPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($PowerOnSet -and $SetupPasswordCheck -eq 1) -and !($SetupPassword))
{
	$ErrorMsg = "When using the PowerOnSet switch on a computer where the setup password is already set, the SetupPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($PowerOnClear) -and !($OldPowerOnPassword))
{
	$ErrorMsg = "When using the PowerOnClear switch, the OldPowerOnPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SetupSet) -and ($SetupClear))
{
	$ErrorMsg = "Cannot specify the SetupSet and SetupClear parameters simultaneously"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($PowerOnSet) -and ($PowerOnClear))
{
	$ErrorMsg = "Cannot specify the PowerOnSet and PowerOnClear parameters simultaneously"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($OldSetupPassword -or $SetupPassword) -and !($SetupSet -or $SetupClear))
{
	$ErrorMsg = "When using the OldSetupPassword or SetupPassword parameters, one of the SetupSet or SetupClear parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($OldPowerOnPassword -or $PowerOnPassword) -and !($PowerOnSet -or $PowerOnClear))
{
	$ErrorMsg = "When using the OldPowerOnPassword or PowerOnPassword parameters, one of the PowerOnSet or PowerOnClear parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if ($OldSetupPassword.Count -gt 2) #Prevents entering more than 2 old Setup passwords
{
	$ErrorMsg = "Please specify 2 or fewer old Setup passwords"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if ($OldPowerOnPassword.Count -gt 2) #Prevents entering more than 2 old power on passwords
{
	$ErrorMsg = "Please specify 2 or fewer old power on passwords"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SMSTSPasswordRetry) -and !(Get-TaskSequenceStatus))
{
	Write-LogEntry -Value "The SMSTSPasswordRetry parameter was specifed while not running in a task sequence. Setting SMSTSPasswordRetry to false." -Severity 2
	$SMSTSPasswordRetry = $False
}
Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Set variables from a previous script session
if (Get-TaskSequenceStatus)
{
	Write-LogEntry -Value "Check for existing task sequence variables" -Severity 1
	$HPSetSetup = $TSEnv.Value("HPSetSetup")
	if ($HPSetSetup -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful setup password set attempt detected" -Severity 1
	}
	$HPClearSetup = $TSEnv.Value("HPClearSetup")
	if ($HPClearSetup -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful setup password clear attempt detected" -Severity 1
	}
	$HPSetPowerOn = $TSEnv.Value("HPSetPowerOn")
	if ($HPSetPowerOn -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password set attempt detected" -Severity 1
	}
	$HPClearPowerOn = $TSEnv.Value("HPClearPowerOn")
	if ($HPClearPowerOn -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password clear attempt detected" -Severity 1
	}
}

#No setup password currently set
if ($SetupPasswordCheck -eq 0)
{
    if ($SetupClear)
    {
        Write-LogEntry -Value "No Setup password currently set. No need to clear the setup password" -Severity 2
        Clear-Variable SetupClear
    }
    if ($SetupSet)
    {
        if (($Interface.SetBIOSSetting("Setup Password","<utf-16/>" + $SetupPassword,"<utf-16/>")).Return -eq 0)
        {
            Write-LogEntry -Value "The setup password has been successfully set" -Severity 1
        }
        else
        {
            $SetupPWExists = "Failed"
            Write-LogEntry -Value "Failed to set the setup password" -Severity 3
        }
    }
}

#No power on password currently set
if ($PowerOnPasswordCheck -eq 0)
{
    if ($PowerOnClear)
    {
        Write-LogEntry -Value "No power on password currently set. No need to clear the power on password" -Severity 2
        Clear-Variable SetupClear
    }
    if (($PowerOnSet) -and ($SetupPasswordCheck -eq 0))
    {
        if (($Interface.SetBIOSSetting("Power-On Password","<utf-16/>" + $PowerOnPassword,"<utf-16/>")).Return -eq 0)
        {
            Write-LogEntry -Value "The power on password has been successfully set" -Severity 1
        }
        else
        {
            $PowerOnPWExists = "Failed"
            Write-LogEntry -Value "Failed to set the power on password" -Severity 3
        }
	}
	if (($PowerOnSet) -and ($SetupPasswordCheck -eq 1))
    {
        if (($Interface.SetBIOSSetting("Power-On Password","<utf-16/>" + $PowerOnPassword,"<utf-16/>" + $OldSetupPassword)).Return -eq 0)
        {
            Write-LogEntry -Value "The power on password has been successfully set" -Severity 1
        }
        else
        {
            $PowerOnPWExists = "Failed"
            Write-LogEntry -Value "Failed to set the power on password" -Severity 3
		}
	}
}

#If a Setup password is set, attempt to clear or change it
if ($SetupPasswordCheck -eq 1)
{
	#Change the existing Setup password
	if (($SetupSet) -and ($HPSetSetup -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to change the existing setup password" -Severity 1
		$SetupPWSet = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("HPSetSetup") = "Failed"
		}
    
		if (($Interface.SetBIOSSetting("Setup Password","<utf-16/>" + $SetupPassword,"<utf-16/>" + $SetupPassword)).Return -eq 0)
		{
			#Password already correct
			$SetupPWSet = "Success"
			if (Get-TaskSequenceStatus)
			{
				$TSEnv.Value("HPSetSetup") = "Success"
			}
			Write-LogEntry -Value "The setup password is already set correctly" -Severity 1
		}
		else
		{
			$Counter = 0
			While($Counter -lt $OldSetupPassword.Count){
                if (($Interface.SetBIOSSetting("Setup Password","<utf-16/>" + $SetupPassword,"<utf-16/>" + $OldSetupPassword[$Counter])).Return -eq 0)
				{
					#Successfully changed the password
					$SetupPWSet = "Success"
					if (Get-TaskSequenceStatus)
					{
						$TSEnv.Value("HPSetSetup") = "Success"
					}
					Write-LogEntry -Value "The setup password has been successfully changed" -Severity 1
					break
				}
				else
				{
					#Failed to change the password
					$Counter++
				}
			}
			if ($SetupPWSet -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the setup password" -Severity 3
			}
		}
	}
	
	#Clear the existing Setup password
	if (($SetupClear) -and ($HPClearSetup -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to clear the existing setup password" -Severity 1
		$SetupPWClear = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("HPClearSetup") = "Failed"
		}

		$Counter = 0
		While($Counter -lt $OldSetupPassword.Count){
			if (($Interface.SetBIOSSetting("Setup Password","<utf-16/>","<utf-16/>" + $OldSetupPassword[$Counter])).Return -eq 0)
			{
				#Successfully cleared the password
				$SetupPWClear = "Success"
				if (Get-TaskSequenceStatus)
				{
					$TSEnv.Value("HPClearSetup") = "Success"
				}
				Write-LogEntry -Value "The setup password has been successfully cleared" -Severity 1
				break
			}
			else
			{
				#Failed to clear the password
				$Counter++
			}
		}	
		if ($SetupPWClear -eq "Failed")
		{
			Write-LogEntry -Value "Failed to clear the setup password" -Severity 3
		}
	}
}

#If a power on password is set, attempt to clear or change it
if ($PowerOnPasswordCheck -eq 1)
{
	#Change the existing power on password
	if (($PowerOnSet) -and ($HPSetPowerOn -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to change the existing power on password" -Severity 1
		$PowerOnPWSet = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("HPSetPowerOn") = "Failed"
		}
    
		if (($Interface.SetBIOSSetting("Power-On Password","<utf-16/>" + $PowerOnPassword,"<utf-16/>" + $PowerOnPassword)).Return -eq 0)
		{
			#Password already correct
			$PowerOnPWSet = "Success"
			if (Get-TaskSequenceStatus)
			{
				$TSEnv.Value("HPSetPowerOn") = "Success"
			}
			Write-LogEntry -Value "The power on password is already set correctly" -Severity 1
		}
		else
		{
			$Counter = 0
			While($Counter -lt $OldPowerOnPassword.Count){
				if (($Interface.SetBIOSSetting("Power-On Password","<utf-16/>" + $PowerOnPassword,"<utf-16/>" + $OldPowerOnPassword[$Counter])).Return -eq 0)
				{
					#Successfully changed the password
					$PowerOnPWSet = "Success"
					if (Get-TaskSequenceStatus)
					{
						$TSEnv.Value("HPSetPowerOn") = "Success"
					}
					Write-LogEntry -Value "The power on password has been successfully changed" -Severity 1
					break
				}
				else
				{
					#Failed to change the password
					$Counter++
				}
			}
			if ($PowerOnPWSet -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the power on password" -Severity 3
			}
		}
	}
	
	#Clear the existing power on password
	if (($PowerOnClear) -and ($HPClearPowerOn -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to clear the existing power on password" -Severity 1
		$PowerOnPWClear = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("HPClearPowerOn") = "Failed"
		}

		$Counter = 0
		While($Counter -lt $OldPowerOnPassword.Count){
			if (($Interface.SetBIOSSetting("Power-On Password","<utf-16/>","<utf-16/>" + $OldPowerOnPassword[$Counter])).Return -eq 0)
			{
				#Successfully cleared the password
				$PowerOnPWClear = "Success"
				if (Get-TaskSequenceStatus)
				{
					$TSEnv.Value("HPClearPowerOn") = "Success"
				}
				Write-LogEntry -Value "The power on password has been successfully cleared" -Severity 1
				break
			}
			else
			{
				#Failed to clear the password
				$Counter++
			}
		}			
		if ($PowerOnPWClear -eq "Failed")
		{
			Write-LogEntry -Value "Failed to clear the power on password" -Severity 3
		}
	}
}

#Prompt the user about any failures
if ((($SetupPWExists -eq "Failed") -or ($SetupPWSet -eq "Failed") -or ($SetupPWClear -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($PowerOnPWSet -eq "Failed") -or ($PowerOnPWClear -eq "Failed")) -and (!($SMSTSPasswordRetry)))
{
	if (!($NoUserPrompt))
	{
		Write-LogEntry -Value "Failures detected, display on-screen prompts for any required manual actions" -Severity 2
		#Close the task sequence progress dialog
		if (Get-TaskSequenceStatus)
		{
			$TSProgress.CloseProgressDialog()
		}
		#Display prompts
		if ($SetupPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No setup password is set, but the script was unable to set a password. Please reboot the computer and manually set the setup password." -TitleText "HP Password Management Script"
		}
		if ($SetupPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The setup password is set, but cannot be automatically changed. Please reboot the computer and manually change the setup password." -TitleText "HP Password Management Script"
		}
		if ($SetupPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The setup password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the setup password." -TitleText "HP Password Management Script"
		}
		if ($PowerOnPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No power on password is set, but the script was unable to set a password. Please reboot the computer and manually set the power on password." -TitleText "HP Password Management Script"
		}
		if ($PowerOnPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically changed. Please reboot the computer and manually change the power on password." -TitleText "HP Password Management Script"
		}
		if ($PowerOnPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the power on password." -TitleText "HP Password Management Script"
		}
	}
	#Exit the script with an error
	if (!($ContinueOnError))
	{
		Write-LogEntry -Value "Failures detected, exiting the script" -Severity 3
		Write-Output "Password management tasks failed. Check the log file for more information"
		Write-LogEntry -Value "END - HP BIOS password management script" -Severity 1
		Exit 1
	}
	else
	{
		Write-LogEntry -Value "Failures detected, but the ContinueOnError parameter was set. Script execution will continue" -Severity 3
		Write-Output "Failures detected, but the ContinueOnError parameter was set. Script execution will continue"
	}
}
elseif ((($SetupPWExists -eq "Failed") -or ($SetupPWSet -eq "Failed") -or ($SetupPWClear -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($PowerOnPWSet -eq "Failed") -or ($PowerOnPWClear -eq "Failed")) -and ($SMSTSPasswordRetry))
{
	Write-LogEntry -Value "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed" -Severity 3
	Write-Output "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed"
}
else
{
	Write-Output "Password management tasks succeeded. Check the log file for more information"
}
Write-LogEntry -Value "END - HP BIOS password management script" -Severity 1
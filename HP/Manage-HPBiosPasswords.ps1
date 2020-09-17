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

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

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
		Modifed: 2020-09-17

	.CHANGELOG
		2019-07-27 - Formatting changes. Changed the SMSTSPasswordRetry parameter to be a switch instead of an integer value. Changed the SMSTSChangeSetup TS variable to HPChangeSetup.
					 Changed the SMSTSClearSetup TS variable to HPClearSetup. Changed the SMSTSChangePowerOn TS variable to HPChangePowerOn. Changed the SMSTSClearPowerOn TS variable to HPClearPowerOn.
		2019-11-04 - Added additional logging. Changed the default log path to $ENV:ProgramData\BiosScripts\HP. Modifed the parameter validation logic.
		2020-01-30 - Removed the SetupChange and PowerOnChange parameters. SetupSet and PowerOnSet now work to set or change a password. Changed the HPChangeSetup task sequence variable to HPSetSetup.
					 Changed the HPChangePowerOn task sequence variable to HPSetPowerOn. Updated the parameter validation checks.
		2020-09-14 - Added a LogFile parameter. Changed the default log path in full Windows to $ENV:ProgramData\ConfigJonScripts\HP.
					 Consolidated duplicate code into new functions (Stop-Script, Get-WmiData, New-HPBiosPassword, Set-HPBiosPassword, Clear-HPBiosPassword). Made a number of minor formatting and syntax changes
					 When using the SetupSet and PowerOnSet parameters, the OldPassword parameters are no longer required. There is now logic to handle and report this type of failure.
		2020-09-17 - Improved the log file path configuration

#>

#Parameters ===================================================================================================================

param(
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
	[Parameter(Mandatory=$false)][Switch]$SMSTSPasswordRetry,
	[Parameter(Mandatory=$false)][ValidateScript({
        if($_ -notmatch "(\.log)")
        {
            throw "The file specified in the LogFile paramter must be a .log file"
        }
        return $true
    })]
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\HP\Manage-HPBiosPasswords.log"
)

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

Function Get-WmiData
{
	#Gets WMI data using either the WMI or CIM cmdlets and stores the data in a variable

    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Namespace,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$ClassName,
        [Parameter(Mandatory=$true)][ValidateSet('CIM','WMI')]$CmdletType,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$Select
    )
    try
    {
        if($CmdletType -eq "CIM")
        {
            if($Select)
            {
				Write-LogEntry -Value "Get the $Classname WMI class from the $Namespace namespace and select properties: $Select" -Severity 1
                $Query = Get-CimInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction Stop | Select-Object $Select -ErrorAction Stop
            }
            else
            {
				Write-LogEntry -Value "Get the $ClassName WMI class from the $Namespace namespace" -Severity 1
                $Query = Get-CimInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction Stop
            }
        }
        elseif($CmdletType -eq "WMI")
        {
            if($Select)
            {
				Write-LogEntry -Value "Get the $Classname WMI class from the $Namespace namespace and select properties: $Select" -Severity 1
                $Query = Get-WmiObject -Namespace $Namespace -Class $ClassName -ErrorAction Stop | Select-Object $Select -ErrorAction Stop
            }
            else
            {
				Write-LogEntry -Value "Get the $ClassName WMI class from the $Namespace namespace" -Severity 1
                $Query = Get-WmiObject -Namespace $Namespace -Class $ClassName -ErrorAction Stop
            }
        }
    }
    catch
    {
		if($Select)
		{
			Stop-Script -ErrorMessage "An error occurred while attempting to get the $Select properties from the $Classname WMI class in the $Namespace namespace" -Exception $PSItem.Exception.Message
		}
		else
		{
			Stop-Script -ErrorMessage "An error occurred while connecting to the $Classname WMI class in the $Namespace namespace" -Exception $PSItem.Exception.Message	
		}
	}
	Write-LogEntry -Value "Successfully connected to the $ClassName WMI class" -Severity 1
	return $Query
}

Function New-HPBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Setup','PowerOn')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Password,
		[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPW
	)
	if($PasswordType -eq "Setup")
	{
		$PasswordName = "Setup Password"
	}
	else
	{
		$PasswordName = "Power-On Password"
	}
	#Attempt to set the power on password when the setup password is already set
	if($SetupPW)
	{
        if(($Interface.SetBIOSSetting($PasswordName,"<utf-16/>" + $Password,"<utf-16/>" + $SetupPW)).Return -eq 0)
        {
            Write-LogEntry -Value "The $PasswordType password has been successfully set" -Severity 1
        }
        else
        {
			Set-Variable -Name "$($PasswordType)PWExists" -Value "Failed" -Scope Script
            Write-LogEntry -Value "Failed to set the $PasswordType password" -Severity 3
		}
	}
	#Attempt to set the setup or power on password
	else
	{
	    if(($Interface.SetBIOSSetting($PasswordName,"<utf-16/>" + $Password,"<utf-16/>")).Return -eq 0)
		{
			Write-LogEntry -Value "The $PasswordType password has been successfully set" -Severity 1
		}
		else
		{
			Set-Variable -Name "$($PasswordType)PWExists" -Value "Failed" -Scope Script
			Write-LogEntry -Value "Failed to set the $PasswordType password" -Severity 3
		}
	}
}

Function Set-HPBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Setup','PowerOn')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Password,
		[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPassword
	)
	if($PasswordType -eq "Setup")
	{
		$PasswordName = "Setup Password"
	}
	else
	{
		$PasswordName = "Power-On Password"
	}
	Write-LogEntry -Value "Attempt to change the existing $PasswordType password" -Severity 1
	Set-Variable -Name "$($PasswordType)PWSet" -Value "Failed" -Scope Script
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("HPSet$($PasswordType)") = "Failed"
	}
	#Check if the password is already set to the correct value
	if(($Interface.SetBIOSSetting($PasswordName,"<utf-16/>" + $Password,"<utf-16/>" + $Password)).Return -eq 0)
	{
		#Password is set to correct value
		Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("HPSet$($PasswordType)") = "Success"
		}
		Write-LogEntry -Value "The $PasswordType password is already set correctly" -Severity 1
	}
	#Password is not set to correct value
	else
	{
		if($OldPassword)
		{
			$Counter = 0
			While($Counter -lt $OldPassword.Count)
			{
               	if(($Interface.SetBIOSSetting($PasswordName,"<utf-16/>" + $Password,"<utf-16/>" + $OldPassword[$Counter])).Return -eq 0)
				{
					#Successfully changed the password
					Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
					if(Get-TaskSequenceStatus)
					{
						$TSEnv.Value("HPSet$($PasswordType)") = "Success"
					}
					Write-LogEntry -Value "The $PasswordType password has been successfully changed" -Severity 1
					break
				}
				else
				{
					#Failed to change the password
					$Counter++
				}
			}
			if((Get-Variable -Name "$($PasswordType)PWSet" -ValueOnly -Scope Script) -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the $PasswordType password" -Severity 3
			}
		}
		else
		{
			Write-LogEntry -Value "The $PasswordType password is currently set to something other than then supplied value, but no old passwords were supplied. Try supplying additional values using the Old$($PasswordType)Password parameter" -Severity 3
		}
	}
}

Function Clear-HPBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Setup','PowerOn')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$OldPassword	
	)
	if($PasswordType -eq "Setup")
	{
		$PasswordName = "Setup Password"
	}
	else
	{
		$PasswordName = "Power-On Password"
	}
	Write-LogEntry -Value "Attempt to clear the existing $PasswordType password" -Severity 1
	Set-Variable -Name "$($PasswordType)PWClear" -Value "Failed" -Scope Script
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("HPClear$($PasswordType)") = "Failed"
	}
	$Counter = 0
	While($Counter -lt $OldPassword.Count)
	{
		if(($Interface.SetBIOSSetting($PasswordName,"<utf-16/>","<utf-16/>" + $OldPassword[$Counter])).Return -eq 0)
		{
			#Successfully cleared the password
			Set-Variable -Name "$($PasswordType)PWClear" -Value "Success" -Scope Script
			if(Get-TaskSequenceStatus)
			{
				$TSEnv.Value("HPClear$($PasswordType)") = "Success"
			}
			Write-LogEntry -Value "The $PasswordType password has been successfully cleared" -Severity 1
			break
		}
		else
		{
			#Failed to clear the password
			$Counter++
		}
	}	
	if((Get-Variable -Name "$($PasswordType)PWClear" -ValueOnly -Scope Script) -eq "Failed")
	{
		Write-LogEntry -Value "Failed to clear the $PasswordType password" -Severity 3
	}
}

Function Start-UserPrompt
{
	#Create a user prompt with custom body and title text if the NoUserPrompt variable is not set

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][ValidateNotNullOrEmpty()][String[]]$BodyText,
        [Parameter(Mandatory=$True)][ValidateNotNullOrEmpty()][String[]]$TitleText
    )
    if(!($NoUserPrompt))
	{
		(New-Object -ComObject Wscript.Shell).Popup("$BodyText",0,"$TitleText",0x0 + 0x30) | Out-Null
	}
}

Function Write-LogEntry
{
	#Write data to a CMTrace compatible log file. (Credit to SCConfigMgr - https://www.scconfigmgr.com/)

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
		[string]$FileName = ($script:LogFile | Split-Path -Leaf)
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
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-HPBiosPasswords"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
	$TSProgress = New-Object -ComObject Microsoft.SMS.TsProgressUI
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
Write-LogEntry -Value "START - HP BIOS password management script" -Severity 1

#Connect to the HP_BIOSSettingInterface WMI class
$Interface = Get-WmiData -Namespace root\hp\InstrumentedBIOS -ClassName HP_BIOSSettingInterface -CmdletType WMI

#Connect to the HP_BIOSSetting WMI class
$HPBiosSetting = Get-WmiData -Namespace root\hp\InstrumentedBIOS -ClassName HP_BIOSSetting -CmdletType WMI

#Get the current password status
Write-LogEntry -Value "Get the current password state" -Severity 1

$SetupPasswordCheck = ($HPBiosSetting | Where-Object Name -eq "Setup Password").IsSet
if($SetupPasswordCheck -eq 1)
{
	Write-LogEntry -Value "The setup password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The setup password is not currently set" -Severity 1
}
$PowerOnPasswordCheck = ($HPBiosSetting | Where-Object Name -eq "Power-On Password").IsSet
if($PowerOnPasswordCheck -eq 1)
{
	Write-LogEntry -Value "The power on password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The power on password is not currently set" -Severity 1
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1
if(($SetupSet) -and !($SetupPassword))
{
	Stop-Script -ErrorMessage "When using the SetupSet switch, the SetupPassword parameter must also be specified"
}
if(($SetupClear) -and !($OldSetupPassword))
{
	Stop-Script -ErrorMessage "When using the SetupClear switch, the OldSetupPassword parameter must also be specified"
}
if(($PowerOnSet) -and !($PowerOnPassword))
{
	Stop-Script -ErrorMessage "When using the PowerOnSet switch, the PowerOnPassword parameter must also be specified"
}
if(($PowerOnSet -and $SetupPasswordCheck -eq 1) -and !($SetupPassword))
{
	Stop-Script -ErrorMessage "When using the PowerOnSet switch on a computer where the setup password is already set, the SetupPassword parameter must also be specified"
}
if(($PowerOnClear) -and !($OldPowerOnPassword))
{
	Stop-Script -ErrorMessage "When using the PowerOnClear switch, the OldPowerOnPassword parameter must also be specified"
}
if(($SetupSet) -and ($SetupClear))
{
	Stop-Script -ErrorMessage "Cannot specify the SetupSet and SetupClear parameters simultaneously"
}
if(($PowerOnSet) -and ($PowerOnClear))
{
	Stop-Script -ErrorMessage "Cannot specify the PowerOnSet and PowerOnClear parameters simultaneously"
}
if(($OldSetupPassword -or $SetupPassword) -and !($SetupSet -or $SetupClear))
{
	Stop-Script -ErrorMessage "When using the OldSetupPassword or SetupPassword parameters, one of the SetupSet or SetupClear parameters must also be specified"
}
if(($OldPowerOnPassword -or $PowerOnPassword) -and !($PowerOnSet -or $PowerOnClear))
{
	Stop-Script -ErrorMessage "When using the OldPowerOnPassword or PowerOnPassword parameters, one of the PowerOnSet or PowerOnClear parameters must also be specified"
}
if($OldSetupPassword.Count -gt 2) #Prevents entering more than 2 old Setup passwords
{
	Stop-Script -ErrorMessage "Please specify 2 or fewer old Setup passwords"
}
if($OldPowerOnPassword.Count -gt 2) #Prevents entering more than 2 old power on passwords
{
	Stop-Script -ErrorMessage "Please specify 2 or fewer old power on passwords"
}
if(($SMSTSPasswordRetry) -and !(Get-TaskSequenceStatus))
{
	Write-LogEntry -Value "The SMSTSPasswordRetry parameter was specifed while not running in a task sequence. Setting SMSTSPasswordRetry to false." -Severity 2
	$SMSTSPasswordRetry = $False
}
Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Set variables from a previous script session
if(Get-TaskSequenceStatus)
{
	Write-LogEntry -Value "Check for existing task sequence variables" -Severity 1
	$HPSetSetup = $TSEnv.Value("HPSetSetup")
	if($HPSetSetup -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful setup password set attempt detected" -Severity 1
	}
	$HPClearSetup = $TSEnv.Value("HPClearSetup")
	if($HPClearSetup -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful setup password clear attempt detected" -Severity 1
	}
	$HPSetPowerOn = $TSEnv.Value("HPSetPowerOn")
	if($HPSetPowerOn -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password set attempt detected" -Severity 1
	}
	$HPClearPowerOn = $TSEnv.Value("HPClearPowerOn")
	if($HPClearPowerOn -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password clear attempt detected" -Severity 1
	}
}

#No setup password currently set
if($SetupPasswordCheck -eq 0)
{
    if($SetupClear)
    {
        Write-LogEntry -Value "No Setup password currently set. No need to clear the setup password" -Severity 2
        Clear-Variable SetupClear
    }
    if($SetupSet)
    {
		New-HPBiosPassword -PasswordType Setup -Password $SetupPassword
    }
}

#No power on password currently set
if($PowerOnPasswordCheck -eq 0)
{
    if($PowerOnClear)
    {
        Write-LogEntry -Value "No power on password currently set. No need to clear the power on password" -Severity 2
        Clear-Variable SetupClear
	}
	if($PowerOnSet)
	{
		#If the setup password is currently set, the setup password is required to set the power on password
		if(($HPBiosSetting | Where-Object Name -eq "Setup Password").IsSet -eq 1)
		{
			New-HPBiosPassword -PasswordType PowerOn -Password $PowerOnPassword -SetupPW $SetupPassword
		}
		else
		{
			New-HPBiosPassword -PasswordType PowerOn -Password $PowerOnPassword
		}
	}
}

#If a Setup password is set, attempt to clear or change it
if($SetupPasswordCheck -eq 1)
{
	#Change the existing Setup password
	if(($SetupSet) -and ($HPSetSetup -ne "Success"))
	{
		if($OldSetupPassword)
		{
			Set-HPBiosPassword -PasswordType Setup -Password $SetupPassword -OldPassword $OldSetupPassword
		}
		else
		{
			Set-HPBiosPassword -PasswordType Setup -Password $SetupPassword
		}
	}
	#Clear the existing Setup password
	if(($SetupClear) -and ($HPClearSetup -ne "Success"))
	{
		Clear-HPBiosPassword -PasswordType Setup -OldPassword $OldSetupPassword
	}
}

#If a power on password is set, attempt to clear or change it
if($PowerOnPasswordCheck -eq 1)
{
	#Change the existing power on password
	if(($PowerOnSet) -and ($HPSetPowerOn -ne "Success"))
	{
		if($OldPowerOnPassword)
		{
			Set-HPBiosPassword -PasswordType PowerOn -Password $PowerOnPassword -OldPassword $OldPowerOnPassword
		}
		else
		{
			Set-HPBiosPassword -PasswordType PowerOn -Password $PowerOnPassword
		}
	}
	#Clear the existing power on password
	if(($PowerOnClear) -and ($HPClearPowerOn -ne "Success"))
	{
		Clear-HPBiosPassword -PasswordType PowerOn -OldPassword $OldPowerOnPassword
	}
}

#Prompt the user about any failures
if((($SetupPWExists -eq "Failed") -or ($SetupPWSet -eq "Failed") -or ($SetupPWClear -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($PowerOnPWSet -eq "Failed") -or ($PowerOnPWClear -eq "Failed")) -and (!($SMSTSPasswordRetry)))
{
	if(!($NoUserPrompt))
	{
		Write-LogEntry -Value "Failures detected, display on-screen prompts for any required manual actions" -Severity 2
		#Close the task sequence progress dialog
		if(Get-TaskSequenceStatus)
		{
			$TSProgress.CloseProgressDialog()
		}
		#Display prompts
		if($SetupPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No setup password is set, but the script was unable to set a password. Please reboot the computer and manually set the setup password." -TitleText "HP Password Management Script"
		}
		if($SetupPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The setup password is set, but cannot be automatically changed. Please reboot the computer and manually change the setup password." -TitleText "HP Password Management Script"
		}
		if($SetupPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The setup password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the setup password." -TitleText "HP Password Management Script"
		}
		if($PowerOnPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No power on password is set, but the script was unable to set a password. Please reboot the computer and manually set the power on password." -TitleText "HP Password Management Script"
		}
		if($PowerOnPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically changed. Please reboot the computer and manually change the power on password." -TitleText "HP Password Management Script"
		}
		if($PowerOnPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the power on password." -TitleText "HP Password Management Script"
		}
	}
	#Exit the script with an error
	if(!($ContinueOnError))
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
elseif((($SetupPWExists -eq "Failed") -or ($SetupPWSet -eq "Failed") -or ($SetupPWClear -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($PowerOnPWSet -eq "Failed") -or ($PowerOnPWClear -eq "Failed")) -and ($SMSTSPasswordRetry))
{
	Write-LogEntry -Value "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed" -Severity 3
	Write-Output "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed"
}
else
{
	Write-Output "Password management tasks succeeded. Check the log file for more information"
}
Write-LogEntry -Value "END - HP BIOS password management script" -Severity 1
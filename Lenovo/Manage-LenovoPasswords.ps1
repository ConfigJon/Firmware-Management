<#
    .DESCRIPTION
        Automatically configure Lenovo BIOS passwords and prompt the user if manual intervention is required.

    	PASSWORD STATUS CODES
        0 - No password set
        1 - Power on password set
        2 - Supervisor password set
        3 - Power on and supervisor passwords set
        4 - Hard drive password set
        5 - Power on and hard drive passwords set
        6 - Supervisor and hard drive passwords set
        7 - Supervisor, power on, and hard drive passwords set

	.PARAMETER SupervisorChange
		Specify this switch to change an existing supervisor password. Must also specify the NewSupervisorPassword and OldSupervisorPassword parameters.

	.PARAMETER SupervisorClear
		Specify this swtich to clear an existing supervisor password. Must also specify the OldSupervisorPassword parameter.

	.PARAMETER PowerOnChange
		Specify this switch to change an existing power on password. Must also specify the NewPowerOnPassword and OldPowerOnPassword parameters.

	.PARAMETER PowerOnClear
		Specify this switch to clear an existing power on password. Must also specify the OldPowerOnPassword parameter.

	.PARAMETER HDDPasswordClear
		Specify this swtich to clear an existing master and/or user hard drive password. Must also specify the HDDMasterPassword and/or HDDUserPassword parameters.

	.PARAMETER NewSupervisorPassword
		Specify the new supervisor password to set.

	.PARAMETER OldSupervisorPassword
		Specify the old supervisor password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER NewPowerOnPassword
		Specify the new power on password to set.

	.PARAMETER OldPowerOnPassword
		Specify the old power on password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER HDDUserPassword
		Specify the current user hard drive password to clear.

	.PARAMETER HDDMasterPassword
		Specify the current master hard drive password to clear.
	
	.PARAMETER NoUserPrompt
		The script will run silently and will not prompt the user with a message box.

	.PARAMETER ContinueOnError
		The script will ignore any errors caused by changing or clearing the passwords. This will not suppress errors caused by parameter validation.

	.PARAMETER SMSTSPasswordRetry
		Specify the number of times the script needs to be re-run in a task sequence after the current run

	.EXAMPLE
		Change an existing supervisor password
		Manage-LenovoPasswords.ps1 -SupervisorChange -NewSupervisorPassword <String> -OldSupervisorPassword <String1>,<String2>

		Change an existing supervisor password and clear a power on password
		Manage-LenovoPasswords.ps1 -SupervisorChange -NewSupervisorPassword <String> -OldSupervisorPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

		Clear existing supervisor and power on passwords
		Manage-LenovoPasswords.ps1 -SupervisorClear -OldSupervisorPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

		Clear existing user and master hard drive passwords
		Manage-LenovoPasswords.ps1 -HDDPasswordClear -HDDUserPassword <String> -HDDMasterPassword <String>

		Clear an existing power on password, suppress any user prompts, and continue on error
		Manage-LenovoPasswords.ps1 -PowerOnClear -OldPowerOnPassword <String1>,<String2> -NoUserPrompt -ContinueOnError

	.NOTES
		Created by: Jon Anderson (@ConfigJon)
		Reference: https://www.configjon.com/lenovo-bios-password-management
		Modifed: 7/6/2019
#>

#Parameter declaration
param(
	[Parameter(Mandatory=$false)][Switch]$SupervisorChange,
	[Parameter(Mandatory=$false)][Switch]$SupervisorClear,
	[Parameter(Mandatory=$false)][Switch]$PowerOnChange,
	[Parameter(Mandatory=$false)][Switch]$PowerOnClear,
	[Parameter(Mandatory=$false)][Switch]$HDDPasswordClear,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$NewSupervisorPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSupervisorPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$NewPowerOnPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPowerOnPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$HDDUserPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$HDDMasterPassword,
	[Parameter(Mandatory=$false)][Switch]$NoUserPrompt,
	[Parameter(Mandatory=$false)][Switch]$ContinueOnError,
	[Parameter(Mandatory=$false)][Int]$SMSTSPasswordRetry
)

Function Get-TaskSequenceStatus
#Determine if a task sequence is currently running
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

Function Start-UserPrompt
#Create a user prompt with custom body and title text if the NoUserPrompt variable is not set
{
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

#Configure Logging and task sequence variables
if(Get-TaskSequenceStatus)
{
	$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment
	$TSProgress = New-Object -ComObject Microsoft.SMS.TsProgressUI
	$LogsDirectory = $TSEnv.Value("_SMSTSLogPath")
}
else
{
	$LogsDirectory = "$ENV:SystemRoot\Temp\LenovoPasswordScript"
	if (!(Test-Path -PathType Container $LogsDirectory))
	{
		New-Item -Path $LogsDirectory -ItemType "Directory" -Force | Out-Null
	}
}
Write-Output "Log path set to $LogsDirectory\Manage-LenovoPasswords.log"

Function Write-LogEntry
#Write data to a log file. (Credit to SCConfigMgr - https://www.scconfigmgr.com/)
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
		[string]$FileName = "Manage-LenovoPasswords.log"
	)
	# Determine log file location
	$LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName
		
	# Construct time stamp for log entry
	if(-not(Test-Path -Path 'variable:global:TimezoneBias')) {
		[string]$global:TimezoneBias = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalMinutes
		if($TimezoneBias -match "^-") {
			$TimezoneBias = $TimezoneBias.Replace('-', '+')
		}
		else {
			$TimezoneBias = '-' + $TimezoneBias
		}
	}
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff"), $TimezoneBias)
		
	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")
		
	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
		
	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-LenovoPasswords"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
		
	# Add value to log file
	try{
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
	}
	catch [System.Exception] {
		Write-Warning -Message "Unable to append log entry to $FileName file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}

Write-LogEntry -Value "START - Lenovo BIOS password management script" -Severity 1

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1

if(($SupervisorChange) -and !($NewSupervisorPassword -and $OldSupervisorPassword))
{
	$ErrorMsg = "When using the SupervisorChange switch, the NewSupervisorPassword and OldSupervisorPassword parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($SupervisorClear) -and !($OldSupervisorPassword))
{
	$ErrorMsg = "When using the SupervisorClear switch, the OldSupervisorPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($PowerOnChange) -and !($NewPowerOnPassword -and $OldPowerOnPassword))
{
	$ErrorMsg = "When using the PowerOnChange switch, the NewPowerOnPassword and OldPowerOnPassword parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($PowerOnClear) -and !($OldPowerOnPassword))
{
	$ErrorMsg = "When using the PowerOnClear switch, the OldPowerOnPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($SupervisorChange) -and ($SupervisorClear))
{
	$ErrorMsg = "Cannot specify the SupervisorChange and SupervisorClear parameters simultaneously"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($PowerOnChange) -and ($PowerOnClear))
{
	$ErrorMsg = "Cannot specify the PowerOnChange and PowerOnClear parameters simultaneously"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($HDDPasswordClear) -and !($HDDUserPassword))
{
	$ErrorMsg = "When using the HDDPasswordClear switch, the HDDUserPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($HDDMasterPassword) -and !($HDDUserPassword))
{
	$ErrorMsg = "When specifying a master hard drive password, a user hard drive password must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($HDDMasterPassword -or $HDDUserPassword) -and !($HDDPasswordClear))
{
	$ErrorMsg = "When using the HDDMasterPassword or HDDUserPassword parameters, the HDDPasswordClear parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($OldSupervisorPassword -or $NewSupervisorPassword) -and !($SupervisorChange -or $SupervisorClear))
{
	$ErrorMsg = "When using the OldSupervisorPassword or NewSupervisorPassword parameters, one of the SupervisorChange or SupervisorClear parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if(($OldPowerOnPassword -or $NewPowerOnPassword) -and !($PowerOnChange -or $PowerOnClear))
{
	$ErrorMsg = "When using the OldPowerOnPassword or NewPowerOnPassword parameters, one of the PowerOnChange or PowerOnClear parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if($OldSupervisorPassword.Count -gt 2) #Prevents entering more than 2 old supervisor passwords
{
	$ErrorMsg = "Please specify 2 or fewer old supervisor passwords"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if($OldPowerOnPassword.Count -gt 2) #Prevents entering more than 2 old power on passwords
{
	$ErrorMsg = "Please specify 2 or fewer old power on passwords"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}

#Handle the SMSTSPasswordRetry variable
if(($SMSTSPasswordRetry -gt 0) -and !(Get-TaskSequenceStatus))
{
	Write-LogEntry -Value "The SMSTSPasswordRetry was specifed while not running in a task sequence. Resetting SMSTSPasswordRetry to 0" -Severity 2
	$SMSTSPasswordRetry = 0
}
if($NUll -eq $SMSTSPasswordRetry)
{
	$SMSTSPasswordRetry = 0
}
else
{
	if(Get-TaskSequenceStatus)
	{
		Write-LogEntry -Value "Set the SMSTSPasswordRetry varaible to $SMSTSPasswordRetry" -Severity 1
		$TSEnv.Value("SMSTSPasswordRetry") = $SMSTSPasswordRetry
	}
}

#Set variables from a previous script session
if(Get-TaskSequenceStatus)
{
	Write-LogEntry -Value "Check for existing task sequence variables" -Severity 1
	$SMSTSChangeSup = $TSEnv.Value("SMSTSChangeSup")
	if($SMSTSChangeSup -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful supervisor password change attempt detected" -Severity 1
	}
	$SMSTSClearSup = $TSEnv.Value("SMSTSClearSup")
	if($SMSTSClearSup -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful supervisor password clear attempt detected" -Severity 1
	}
	$SMSTSChangePo = $TSEnv.Value("SMSTSChangePo")
	if($SMSTSChangePo -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password change attempt detected" -Severity 1
	}
	$SMSTSClearPo = $TSEnv.Value("SMSTSClearPo")
	if($SMSTSClearPo -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password clear attempt detected" -Severity 1
	}
}

#Connect to the Lenovo WMI classes
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

#Get the current password status
Write-LogEntry -Value "Get the current password state and validate the specified password is not blank" -Severity 1
$PasswordStatus = $PasswordSettings.PasswordState

#Attempting to set or clear a supervisor password when no supervisor password currently exists
if((($PasswordStatus -eq 0) -or ($PasswordStatus -eq 1) -or ($PasswordStatus -eq 4) -or ($PasswordStatus -eq 5)))
{
	if ($SupervisorChange)
	{
		$SupervisorPWExists = "Failed"
		Write-LogEntry -Value "No supervisor password currently set. Unable to change the supervisor password" -Severity 3
	}
	if ($SupervisorClear)
	{
		Write-LogEntry -Value "No supervisor password currently set. No need to clear the supervisor password" -Severity 2
		Clear-Variable SupervisorClear
	}
}

#Attempting to set or clear a power on password when no power on password currently exists
if((($PasswordStatus -eq 0) -or ($PasswordStatus -eq 2) -or ($PasswordStatus -eq 4) -or ($PasswordStatus -eq 6)))
{
	if ($PowerOnChange)
	{
		$PowerOnPWExists = "Failed"
		Write-LogEntry -Value "No power on password currently set. Unable to change the power on password" -Severity 3
	}
	if ($PowerOnClear)
	{
		Write-LogEntry -Value "No power on password currently set. No need to clear the power on password" -Severity 2
		Clear-Variable PowerOnClear
	}
}

#If a supervisor password is set, attempt to clear or change it
if(($PasswordStatus -eq 2) -or ($PasswordStatus -eq 3) -or($PasswordStatus -eq 6) -or($PasswordStatus -eq 7))
{
	#Change the existing supervisor password
	if(($SupervisorChange) -and ($SMSTSChangeSup -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to change the existing supervisor password" -Severity 1
		$SupervisorPWChange = "Failed"
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("SMSTSChangeSup") = "Failed"
		}
    
		if($PasswordSet.SetBiosPassword("pap,$NewSupervisorPassword,$NewSupervisorPassword,ascii,us").Return -eq "Success")
		{
			#Password already correct
			$SupervisorPWChange = "Success"
			if(Get-TaskSequenceStatus)
			{
				$TSEnv.Value("SMSTSChangeSup") = "Success"
			}
			Write-LogEntry -Value "The supervisor password is already set correctly" -Severity 1
		}
		else
		{
			$Counter = 0
			While($Counter -lt $OldSupervisorPassword.Count){
				if($PasswordSet.SetBiosPassword("pap,$($OldSupervisorPassword[$Counter]),$NewSupervisorPassword,ascii,us").Return -eq "Success")
				{
					#Successfully changed the password
					$SupervisorPWChange = "Success"
					if(Get-TaskSequenceStatus)
					{
						$TSEnv.Value("SMSTSChangeSup") = "Success"
					}
					Write-LogEntry -Value "The supervisor password has been successfully changed" -Severity 1
					break
				}
				else
				{
					#Failed to change the password
					$Counter++
				}
			}
			if($SupervisorPWChange -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the supervisor password" -Severity 3
			}
		}
	}
	
	#Clear the existing supervisor password
	if(($SupervisorClear) -and ($SMSTSClearSup -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to clear the existing supervisor password" -Severity 1
		$SupervisorPWClear = "Failed"
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("SMSTSClearSup") = "Failed"
		}

		$Counter = 0
		While($Counter -lt $OldSupervisorPassword.Count){
			if($PasswordSet.SetBiosPassword("pap,$($OldSupervisorPassword[$Counter]),,ascii,us").Return -eq "Success")
			{
				#Successfully cleared the password
				$SupervisorPWClear = "Success"
				if(Get-TaskSequenceStatus)
				{
					$TSEnv.Value("SMSTSClearSup") = "Success"
				}
				Write-LogEntry -Value "The supervisor password has been successfully cleared" -Severity 1
				break
			}
			else
			{
				#Failed to clear the password
				$Counter++
			}
		}
		if($SupervisorPWClear -eq "Failed")
		{
			Write-LogEntry -Value "Failed to clear the supervisor password" -Severity 3
		}
	}
}

#If a power on password is set, attempt to clear or change it
if(($PasswordStatus -eq 1) -or ($PasswordStatus -eq 3) -or($PasswordStatus -eq 5) -or($PasswordStatus -eq 7))
{
	#Change the existing supervisor password
	if(($PowerOnChange) -and ($SMSTSChangePo -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to change the existing power on password" -Severity 1
		$PowerOnPWChange = "Failed"
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("SMSTSChangePo") = "Failed"
		}
    
		if($PasswordSet.SetBiosPassword("pop,$NewPowerOnPassword,$NewPowerOnPassword,ascii,us").Return -eq "Success")
		{
			#Password already correct
			$PowerOnPWChange = "Success"
			if(Get-TaskSequenceStatus)
			{
				$TSEnv.Value("SMSTSChangePo") = "Success"
			}
			Write-LogEntry -Value "The power on password is already set correctly" -Severity 1
		}
		else
		{
			$Counter = 0
			While($Counter -lt $OldPowerOnPassword.Count){
				if($PasswordSet.SetBiosPassword("pop,$($OldPowerOnPassword[$Counter]),$NewPowerOnPassword,ascii,us").Return -eq "Success")
				{
					#Successfully changed the password
					$PowerOnPWChange = "Success"
					if(Get-TaskSequenceStatus)
					{
						$TSEnv.Value("SMSTSChangePo") = "Success"
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
			if($PowerOnPWChange -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the power on password" -Severity 3
			}
		}
	}
	
	#Clear the existing power on password
	if(($PowerOnClear) -and ($SMSTSClearPo -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to clear the existing power on password" -Severity 1
		$PowerOnPWClear = "Failed"
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("SMSTSClearPo") = "Failed"
		}

		$Counter = 0
		While($Counter -lt $OldPowerOnPassword.Count){
			if($PasswordSet.SetBiosPassword("pop,$($OldPowerOnPassword[$Counter]),,ascii,us").Return -eq "Success")
			{
				#Successfully cleared the password
				$PowerOnPWClear = "Success"
				if(Get-TaskSequenceStatus)
				{
					$TSEnv.Value("SMSTSClearPo") = "Success"
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
		if($PowerOnPWClear -eq "Failed")
		{
			Write-LogEntry -Value "Failed to clear the power on password" -Severity 3
		}
	}
}

#Attempt to clear the hard drive password(s)
if ($HDDPasswordClear)
{
	if (($HDDUserPassword) -and ($HDDMasterPassword))
	{
		Write-LogEntry -Value "Attempt to clear the existing user and master hard drive passwords" -Severity 1
		$PasswordSet.SetBiosPassword("mhdp1,$HDDMasterPassword,,ascii,us")
		$PasswordSet.SetBiosPassword("uhdp1,$HDDUserPassword,,ascii,us")
	}
	elseif (($HDDUserPassword) -and !($HDDMasterPassword))
	{
		Write-LogEntry -Value "Attempt to clear the existing user hard drive password" -Severity 1
		$PasswordSet.SetBiosPassword("uhdp1,$HDDUserPassword,,ascii,us")
	}
}

#Decrement the password retry counter
if($SMSTSPasswordRetry -gt 0)
{
	$SMSTSPasswordRetry--
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("SMSTSPasswordRetry") = $SMSTSPasswordRetry
	}
}

#Prompt the user about any failures
if((($SupervisorPWExists -eq "Failed") -or ($SupervisorPWChange -eq "Failed") -or ($SupervisorPWClear -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($PowerOnPWChange -eq "Failed") -or ($PowerOnPWClear -eq "Failed")) -and ($SMSTSPasswordRetry -eq 0))
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
		if($SupervisorPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No supervisor password is set. Please reboot the computer and manually set a supervisor password" -TitleText "Lenovo Password Management Script"
		}
		if($SupervisorPWChange -eq "Failed")
		{
			Start-UserPrompt -BodyText "The supervisor password is set, but cannot be automatically changed. Please reboot the computer and manually change the supervisor password." -TitleText "Lenovo Password Management Script"
		}
		if($SupervisorPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The supervisor password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the supervisor password." -TitleText "Lenovo Password Management Script"
		}
		if($PowerOnPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No power on password is set. Please reboot the computer and manually set a power on password." -TitleText "Lenovo Password Management Script"
		}
		if($PowerOnPWChange -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically changed. Please reboot the computer and manually change the power on password." -TitleText "Lenovo Password Management Script"
		}
		if($PowerOnPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the power on password." -TitleText "Lenovo Password Management Script"
		}
	}
	#Exit the script with an error
	if(!($ContinueOnError))
	{
		Write-LogEntry -Value "Failures detected, exiting the script" -Severity 2
		Write-Output "Password management tasks failed. Check the log file for more information"
		Exit 1
	}
}
else
{
	Write-Output "Password management tasks succeeded. Check the log file for more information"
}
Write-LogEntry -Value "END - Lenovo BIOS password management script" -Severity 1
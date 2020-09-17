<#
	.DESCRIPTION
		Automatically configure Lenovo BIOS passwords and prompt the user if manual intervention is required.

	.PARAMETER SupervisorSet
		Specify this switch to change an existing supervisor password. Must also specify the SupervisorPassword parameter.

	.PARAMETER SupervisorClear
		Specify this swtich to clear an existing supervisor password. Must also specify the OldSupervisorPassword parameter.

	.PARAMETER PowerOnSet
		Specify this switch to change an existing power on password. Must also specify the PowerOnPassword parameter.

	.PARAMETER PowerOnClear
		Specify this switch to clear an existing power on password. Must also specify the OldPowerOnPassword parameter.

	.PARAMETER SystemManagementSet
		Specify this swtich to set or change the system management password. Must also specify the SystemManagementPassword parameter.

	.PARAMETER SystemManagementClear
		Specify this swtich to clear an existing system management password. Must also specify the OldSystemManagementPassword parameter.

	.PARAMETER HDDPasswordClear
		Specify this swtich to clear an existing master and/or user hard drive password. Must also specify the HDDMasterPassword and/or HDDUserPassword parameters.

	.PARAMETER SupervisorPassword
		Specify the new supervisor password to set.

	.PARAMETER OldSupervisorPassword
		Specify the old supervisor password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER PowerOnPassword
		Specify the new power on password to set.

	.PARAMETER OldPowerOnPassword
		Specify the old power on password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER SystemManagementPassword
		Specify the new system management password to set.

	.PARAMETER OldSystemManagementPassword
		Specify the old system management password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER HDDUserPassword
		Specify the current user hard drive password to clear.

	.PARAMETER HDDMasterPassword
		Specify the current master hard drive password to clear.
	
	.PARAMETER NoUserPrompt
		The script will run silently and will not prompt the user with a message box.

	.PARAMETER ContinueOnError
		The script will ignore any errors caused by changing or clearing the passwords. This will not suppress errors caused by parameter validation.

	.PARAMETER SMSTSPasswordRetry
		For use in a task sequence. If specified, the script will assume the script needs to run at least one more time. This will ignore password errors and suppress user prompts.

	.PARAMETER LogFile
		Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

	.EXAMPLE	
		Change an existing supervisor password
		Manage-LenovoBiosPasswords.ps1 -SupervisorSet -SupervisorPassword <String> -OldSupervisorPassword <String1>,<String2>

		Change an existing supervisor password and clear a power on password
		Manage-LenovoBiosPasswords.ps1 -SupervisorSet -SupervisorPassword <String> -OldSupervisorPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

		Clear existing supervisor and power on passwords
		Manage-LenovoBiosPasswords.ps1 -SupervisorClear -OldSupervisorPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

		Clear existing user and master hard drive passwords
		Manage-LenovoBiosPasswords.ps1 -HDDPasswordClear -HDDUserPassword <String> -HDDMasterPassword <String>

		Clear an existing power on password, suppress any user prompts, and continue on error
		Manage-LenovoBiosPasswords.ps1 -PowerOnClear -OldPowerOnPassword <String1>,<String2> -NoUserPrompt -ContinueOnError

	.NOTES
		Created by: Jon Anderson (@ConfigJon)
		Reference: https://www.configjon.com/lenovo-bios-password-management
		Modifed: 2020-09-16

	.CHANGELOG
		2019-07-17 - Updated the script name to Manage-LenovoBiosPasswords. Updated the log directory name to LenovoBiosScripts. Updated the log file name to Manage-LenovoBiosPasswords
		2019-07-27 - Formatting changes. Changed the NewSupervisorPassword parameter to SupervisorPassword. Changed the NewPowerOnPassword parameter to PowerOnPassword.
					 Changed the SMSTSPasswordRetry parameter to be a switch instead of an integer value. Changed the SMSTSChangeSup TS variable to LenovoChangeSupervisor.
					 Changed the SMSTSClearSup TS variable to LenovoClearSupervisor. Changed the SMSTSChangePo TS variable to LenovoChangePowerOn. Changed the SMSTSClearPo TS variable to LenovoClearPowerOn
		2019-11-04 - Added additional logging. Changed the default log path to $ENV:ProgramData\BiosScripts\Lenovo. Modifed the parameter validation logic.
		2020-01-30 - Changed the SupervisorChange and PowerOnChange parameters to SupervisorSet and PowerOnSet. Changed the LenovoChangeSupervisor task sequence variable to LenovoSetSupervisor.
					 Changed the LenovoChangePowerOn task sequence variable to LenovoSetPowerOn. Updated the parameter validation checks.
		2020-02-10 - Added better logic for error handling when no Supervisor or Power On Passwords are set.
		2020-06-09 - Updated some Write-LogEntry lines to include missing -Severity parameters
		2020-09-16 - Added a LogFile parameter. Changed the default log path in full Windows to $ENV:ProgramData\ConfigJonScripts\Lenovo. Made a number of formatting and syntax changes
					 Consolidated duplicate code into new functions (Stop-Script, Get-WmiData, Set-LenovoBiosPassword, Clear-LenovoBiosPassword).
					 When using the SupervisorSet and SystemSet parameters, the OldPassword parameters are no longer required. There is now logic to handle and report this type of failure.
					 Added support for changing and clearing the system management password.
#>

#Parameters ===================================================================================================================

param(
	[Parameter(Mandatory=$false)][Switch]$SupervisorSet,
	[Parameter(Mandatory=$false)][Switch]$SupervisorClear,
	[Parameter(Mandatory=$false)][Switch]$PowerOnSet,
	[Parameter(Mandatory=$false)][Switch]$PowerOnClear,
	[Parameter(Mandatory=$false)][Switch]$SystemManagementSet,
	[Parameter(Mandatory=$false)][Switch]$SystemManagementClear,
	[Parameter(Mandatory=$false)][Switch]$HDDPasswordClear,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SupervisorPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSupervisorPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$PowerOnPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPowerOnPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SystemManagementPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSystemManagementPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$HDDUserPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$HDDMasterPassword,
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
	[System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\Lenovo\Manage-LenovoBiosPasswords.log"
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

<#
Function New-LenovoBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('SystemManagement')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Password
	)
	switch($PasswordType)
	{
		'SystemManagement' {$PasswordName = "smp"}
		default
		{
			Stop-Script -ErrorMessage "Unable to determine the current password type from value: $PasswordType"
		}
	}
	#Attempt to set the system management password
	if($PasswordSet.SetBiosPassword("$PasswordName,,$Password,ascii,us").Return -eq "Success")
	{
		Write-LogEntry -Value "The $PasswordType password has been successfully set" -Severity 1
	}
	else
	{
		Set-Variable -Name "$($PasswordType)PWExists" -Value "Failed" -Scope Script
		Write-LogEntry -Value "Failed to set the $PasswordType password" -Severity 3
	}
}
#>

Function Set-LenovoBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Supervisor','PowerOn','SystemManagement')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Password,
		[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPassword,
		[Parameter(Mandatory=$false)][Switch]$PwBypass
	)
	switch($PasswordType)
	{
		'Supervisor' {$PasswordName = "pap"}
		'PowerOn' {$PasswordName = "pop"}
		'SystemManagement' {$PasswordName = "smp"}
		default
		{
			Stop-Script -ErrorMessage "Unable to determine the current password type from value: $PasswordType"
		}
	}
	Write-LogEntry -Value "Attempt to change the existing $PasswordType password" -Severity 1
	Set-Variable -Name "$($PasswordType)PWSet" -Value "Failed" -Scope Script
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("LenovoSet$($PasswordType)") = "Failed"
	}
	if($PasswordSet.SetBiosPassword("$PasswordName,$Password,$Password,ascii,us").Return -eq "Success")
	{
		#Password already correct
		Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("LenovoSet$($PasswordType)") = "Success"
		}
		Write-LogEntry -Value "The $PasswordType password is already set correctly" -Severity 1
	}
	else
	{
		if($OldPassword)
		{
			$Counter = 0
			While($Counter -lt $OldPassword.Count)
			{
				if($PasswordSet.SetBiosPassword("$PasswordName,$($OldPassword[$Counter]),$Password,ascii,us").Return -eq "Success")
				{
					#Successfully changed the password
					Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
					if(Get-TaskSequenceStatus)
					{
						$TSEnv.Value("LenovoSet$($PasswordType)") = "Success"
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

Function Clear-LenovoBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Supervisor','PowerOn','SystemManagement')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$OldPassword	
	)
	switch($PasswordType)
	{
		'Supervisor' {$PasswordName = "pap"}
		'PowerOn' {$PasswordName = "pop"}
		'SystemManagement' {$PasswordName = "smp"}
		default
		{
			Stop-Script -ErrorMessage "Unable to determine the current password type from value: $PasswordType"
		}
	}
	Write-LogEntry -Value "Attempt to clear the existing $PasswordType password" -Severity 1
	Set-Variable -Name "$($PasswordType)PWClear" -Value "Failed" -Scope Script
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("LenovoClear$($PasswordType)") = "Failed"
	}
	$Counter = 0
	While($Counter -lt $OldPassword.Count)
	{
		if($PasswordSet.SetBiosPassword("$PasswordName,$($OldPassword[$Counter]),,ascii,us").Return -eq "Success")
		{
			#Successfully cleared the password
			Set-Variable -Name "$($PasswordType)PWClear" -Value "Success" -Scope Script
			if(Get-TaskSequenceStatus)
			{
				$TSEnv.Value("LenovoClear$($PasswordType)") = "Success"
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
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-LenovoBiosPasswords"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
Write-LogEntry -Value "START - Lenovo BIOS password management script" -Severity 1

#Connect to the Lenovo_BiosPasswordSettings WMI class
$PasswordSettings = Get-WmiData -Namespace root\wmi -ClassName Lenovo_BiosPasswordSettings -CmdletType WMI

#Connect to the Lenovo_SetBiosPassword WMI class
$PasswordSet = Get-WmiData -Namespace root\wmi -ClassName Lenovo_SetBiosPassword -CmdletType WMI	

#Get the current password state
Write-LogEntry -Value "Get the current password state" -Severity 1
switch($PasswordSettings.PasswordState)
{
	{$_ -eq 0}
	{
		Write-LogEntry -Value "No passwords are currently set" -Severity 1
	}
	{($_ -eq 2) -or ($_ -eq 3) -or ($_ -eq 6) -or ($_ -eq 7) -or ($_ -eq 66) -or ($_ -eq 67) -or ($_ -eq 70) -or ($_-eq 71)}
	{
		$SvpSet = $true
		Write-LogEntry -Value "The supervisor password is set" -Severity 1
	}
	{($_ -eq 1) -or ($_ -eq 3) -or ($_ -eq 5) -or ($_ -eq 7) -or ($_ -eq 65) -or ($_ -eq 67) -or ($_ -eq 69) -or ($_-eq 71)}
	{
		$PopSet = $true
		Write-LogEntry -Value "The power on password is set" -Severity 1
	}
	{($_ -eq 64) -or ($_ -eq 65) -or ($_ -eq 66) -or ($_ -eq 67) -or ($_ -eq 68) -or ($_ -eq 69) -or ($_ -eq 70) -or ($_-eq 71)}
	{
		$SmpSet = $true
		Write-LogEntry -Value "The system management password is set" -Severity 1
	}
	{($_ -eq 4) -or ($_ -eq 5) -or ($_ -eq 6) -or ($_ -eq 7) -or ($_ -eq 68) -or ($_ -eq 69) -or ($_ -eq 70) -or ($_-eq 71)}
	{
		$HdpSet = $true
		Write-LogEntry -Value "The hard drive password is set" -Severity 1
	}
	default
	{
		Stop-Script -ErrorMessage "Unable to determine the current password state from value: $($PasswordSettings.PasswordState)"
	}
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1
if(($SupervisorSet) -and !($SupervisorPassword))
{
	Stop-Script -ErrorMessage "When using the SupervisorSet switch, the SupervisorPassword parameter must also be specified"
}
if(($SupervisorClear) -and !($OldSupervisorPassword))
{
	Stop-Script -ErrorMessage "When using the SupervisorClear switch, the OldSupervisorPassword parameter must also be specified"
}
if(($PowerOnSet) -and !($PowerOnPassword))
{
	Stop-Script -ErrorMessage "When using the PowerOnSet switch, the PowerOnPassword parameter must also be specified"
}
if((($PowerOnClear) -and (($SvpSet -ne $true) -and ($SmpSet -ne $true))) -and !($OldPowerOnPassword))
{
	Stop-Script -ErrorMessage "When using the PowerOnClear switch, the OldPowerOnPassword parameter must also be specified"
}
if(($SystemManagementSet) -and !($SystemManagementPassword))
{
	Stop-Script -ErrorMessage "When using the SystemManagementSet switch, the SystemManagementPassword parameter must also be specified"
}
if((($SystemManagementClear) -and ($SvpSet -ne $true)) -and !($OldSystemManagementPassword))
{
	Stop-Script -ErrorMessage "When using the SystemManagementClear switch, the OldSystemManagementPassword parameter must also be specified"
}
if(($SupervisorSet) -and ($SupervisorClear))
{
	Stop-Script -ErrorMessage "Cannot specify the SupervisorSet and SupervisorClear parameters simultaneously"
}
if(($PowerOnSet) -and ($PowerOnClear))
{
	Stop-Script -ErrorMessage "Cannot specify the PowerOnSet and PowerOnClear parameters simultaneously"
}
if(($SystemManagementSet) -and ($SystemManagementClear))
{
	Stop-Script -ErrorMessage "Cannot specify the SystemManagementSet and SystemManagementClear parameters simultaneously"
}
if(($HDDPasswordClear) -and !($HDDUserPassword))
{
	Stop-Script -ErrorMessage "When using the HDDPasswordClear switch, the HDDUserPassword parameter must also be specified"
}
if(($HDDMasterPassword) -and !($HDDUserPassword))
{
	Stop-Script -ErrorMessage "When specifying a master hard drive password, a user hard drive password must also be specified"
}
if(($HDDMasterPassword -or $HDDUserPassword) -and !($HDDPasswordClear))
{
	Stop-Script -ErrorMessage "When using the HDDMasterPassword or HDDUserPassword parameters, the HDDPasswordClear parameter must also be specified"
}
if((($PowerOnSet -or $PowerOnClear) -and ($SvpSet -eq $true)) -and !($SupervisorPassword))
{
	Stop-Script -ErrorMessage "When attempting to change or clear the power on password while the supervisor password is set, the SupervisorPassword parameter must be specified."
}
if((($PowerOnSet -or $PowerOnClear) -and (($SvpSet -ne $true) -and ($SmpSet -eq $true)) -and !($SystemManagementPassword)))
{
	Stop-Script -ErrorMessage "When attempting to change or clear the power on password while the system management password is set, the SystemManagementPassword parameter must be specified."
}
if((($SystemManagementSet -or $SystemManagementClear) -and ($SvpSet -eq $true)) -and !($SupervisorPassword))
{
	Stop-Script -ErrorMessage "When attempting to change or clear the system management password while the supervisor password is set, the SupervisorPassword parameter must be specified."
}
if(($OldSupervisorPassword) -and !($SupervisorSet -or $SupervisorClear))
{
	Stop-Script -ErrorMessage "When using the OldSupervisorPassword parameter, one of the SupervisorSet or SupervisorClear parameters must also be specified"
}
if(($SupervisorPassword) -and !($SupervisorSet -or $SupervisorClear -or $PowerOnSet -or $PowerOnClear -or $SystemManagementSet -or $SystemManagementClear))
{
	Stop-Script -ErrorMessage "When using the SupervisorPassword parameter, one of the SupervisorSet or SupervisorClear or PowerOnSet or PowerOnClear or SystemManagementSet or SystemManagementClear parameters must also be specified"
}
if(($OldPowerOnPassword -or $PowerOnPassword) -and !($PowerOnSet -or $PowerOnClear))
{
	Stop-Script -ErrorMessage "When using the OldPowerOnPassword or PowerOnPassword parameters, one of the PowerOnSet or PowerOnClear parameters must also be specified"
}
if(($OldSystemManagementPassword -or $SystemManagementPassword) -and !($SystemManagementSet -or $SystemManagementClear -or $PowerOnSet -or $PowerOnClear))
{
	Stop-Script -ErrorMessage "When using the OldSystemManagementPassword or SystemManagementPassword parameters, one of the SystemManagementSet or SystemManagementClear parameters must also be specified"
}
if($OldSupervisorPassword.Count -gt 2) #Prevents entering more than 2 old supervisor passwords
{
	Stop-Script -ErrorMessage "Please specify 2 or fewer old supervisor passwords"
}
if($OldPowerOnPassword.Count -gt 2) #Prevents entering more than 2 old power on passwords
{
	Stop-Script -ErrorMessage "Please specify 2 or fewer old power on passwords"
}
if($OldSystemManagementPassword.Count -gt 2) #Prevents entering more than 2 old system management passwords
{
	Stop-Script -ErrorMessage "Please specify 2 or fewer old system management passwords"
}
if(($SMSTSPasswordRetry) -and !(Get-TaskSequenceStatus))
{
	Write-LogEntry -Value "The SMSTSPasswordRetry parameter was specifed while not running in a task sequence. Setting SMSTSPasswordRetry to false." -Severity 2
	$SMSTSPasswordRetry = 0
}
Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Set variables from a previous script session
if(Get-TaskSequenceStatus)
{
	Write-LogEntry -Value "Check for existing task sequence variables" -Severity 1
	$LenovoSetSupervisor = $TSEnv.Value("LenovoSetSupervisor")
	if($LenovoSetSupervisor -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful supervisor password set attempt detected" -Severity 1
	}
	$LenovoClearSupervisor = $TSEnv.Value("LenovoClearSupervisor")
	if($LenovoClearSupervisor -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful supervisor password clear attempt detected" -Severity 1
	}
	$LenovoSetPowerOn = $TSEnv.Value("LenovoSetPowerOn")
	if($LenovoSetPowerOn -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password set attempt detected" -Severity 1
	}
	$LenovoClearPowerOn = $TSEnv.Value("LenovoClearPowerOn")
	if($LenovoClearPowerOn -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful power on password clear attempt detected" -Severity 1
	}
	if($LenovoSetSystemManagement -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful system management password set attempt detected" -Severity 1
	}
	$LenovoClearSystemManagement = $TSEnv.Value("LenovoClearSystemManagement")
	if($LenovoClearSystemManagement -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful system management password clear attempt detected" -Severity 1
	}
}

#No supervisor password set
if(!$SvpSet)
{
	if($SupervisorSet)
	{
		$SupervisorPWExists = "Failed"
		Write-LogEntry -Value "No supervisor password currently set. Unable to set the supervisor password" -Severity 3
	}
	if($SupervisorClear)
	{
		Write-LogEntry -Value "No supervisor password currently set. No need to clear the supervisor password" -Severity 2
		Clear-Variable SupervisorClear
	}
}

#No power on password set
if(!$PopSet)
{
	if($PowerOnSet)
	{
		$PowerOnPWExists = "Failed"
		Write-LogEntry -Value "No power on password currently set. Unable to set the power on password" -Severity 3
	}
	if($PowerOnClear)
	{
		Write-LogEntry -Value "No power on password currently set. No need to clear the power on password" -Severity 2
		Clear-Variable PowerOnClear
	}
}

#No system management password set
if(!$SmpSet)
{
	if($SystemManagementSet)
	{
		$SystemManagementPWExists = "Failed"
		Write-LogEntry -Value "No system management password currently set. Unable to set the system management password" -Severity 3
		#New-LenovoBiosPassword -PasswordType SystemManagement -Password $SystemManagementPassword
	}
	if($SystemManagementClear)
	{
		Write-LogEntry -Value "No system management password currently set. No need to clear the system management password" -Severity 2
		Clear-Variable SystemManagementClear
	}
}

#If a supervisor password is set, attempt to clear or change it
if($SvpSet)
{
	#Change the existing supervisor password
	if(($SupervisorSet) -and ($LenovoSetSupervisor -ne "Success"))
	{
		if($OldSupervisorPassword)
		{
			Set-LenovoBiosPassword -PasswordType Supervisor -Password $SupervisorPassword -OldPassword $OldSupervisorPassword
		}
		else
		{
			Set-LenovoBiosPassword -PasswordType Supervisor -Password $SupervisorPassword
		}
	}
	#Clear the existing supervisor password
	if(($SupervisorClear) -and ($LenovoClearSupervisor -ne "Success"))
	{
		Clear-LenovoBiosPassword -PasswordType Supervisor -OldPassword $OldSupervisorPassword
	}
}

#If a power on password is set, attempt to clear or change it
if($PopSet)
{
	#Change the existing supervisor password
	if(($PowerOnSet) -and ($LenovoSetPowerOn -ne "Success"))
	{
		if($SvpSet)
		{
			Set-LenovoBiosPassword -PasswordType PowerOn -Password $PowerOnPassword -OldPassword $SupervisorPassword
		}
		elseif($SmpSet)
		{
			Set-LenovoBiosPassword -PasswordType PowerOn -Password $PowerOnPassword -OldPassword $SystemManagementPassword
		}
		else
		{
			if($OldPowerOnPassword)
			{
				Set-LenovoBiosPassword -PasswordType PowerOn -Password $PowerOnPassword -OldPassword $OldPowerOnPassword
			}
			else
			{
				Set-LenovoBiosPassword -PasswordType PowerOn -Password $PowerOnPassword
			}
		}
	}
	#Clear the existing power on password
	if(($PowerOnClear) -and ($LenovoClearPowerOn -ne "Success"))
	{
		if($SvpSet)
		{
			Clear-LenovoBiosPassword -PasswordType PowerOn -OldPassword $SupervisorPassword
		}
		elseif($SmpSet)
		{
			Clear-LenovoBiosPassword -PasswordType PowerOn -OldPassword $SystemManagementPassword
		}
		else
		{
			Clear-LenovoBiosPassword -PasswordType PowerOn -OldPassword $OldPowerOnPassword
		}
	}
}

#If a system management password is set, attempt to clear or change it
if($SmpSet)
{
	#Change the existing supervisor password
	if(($SystemManagementSet) -and ($LenovoSetSystemManagement -ne "Success"))
	{
		if($SvpSet)
		{
			Set-LenovoBiosPassword -PasswordType SystemManagement -Password $SystemManagementPassword -OldPassword $SupervisorPassword
		}
		else
		{
			if($OldSystemManagementPassword)
			{
				Set-LenovoBiosPassword -PasswordType SystemManagement -Password $SystemManagementPassword -OldPassword $OldSystemManagementPassword
			}
			else
			{
				Set-LenovoBiosPassword -PasswordType SystemManagement -Password $SystemManagementPassword
			}
		}
	}
	#Clear the existing system management password
	if(($SystemManagementClear) -and ($LenovoClearSystemManagement -ne "Success"))
	{
		if($SvpSet)
		{
			Clear-LenovoBiosPassword -PasswordType SystemManagement -OldPassword $SupervisorPassword
		}
		else
		{
			Clear-LenovoBiosPassword -PasswordType SystemManagement -OldPassword $OldSystemManagementPassword
		}
	}
}

#Attempt to clear the hard drive password(s)
if($HdpSet)
{
	if($HDDPasswordClear)
	{
		if(($HDDUserPassword) -and ($HDDMasterPassword))
		{
			Write-LogEntry -Value "Attempt to clear the existing user and master hard drive passwords" -Severity 1
			$PasswordSet.SetBiosPassword("mhdp1,$HDDMasterPassword,,ascii,us")
			$PasswordSet.SetBiosPassword("uhdp1,$HDDUserPassword,,ascii,us")
		}
		elseif(($HDDUserPassword) -and !($HDDMasterPassword))
		{
			Write-LogEntry -Value "Attempt to clear the existing user hard drive password" -Severity 1
			$PasswordSet.SetBiosPassword("uhdp1,$HDDUserPassword,,ascii,us")
		}
	}
}

#Prompt the user about any password set failures
if(($SupervisorPWExists -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($SystemManagementPWExists -eq "Failed"))
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
		if($PowerOnPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No power on password is set. Please reboot the computer and manually set a power on password." -TitleText "Lenovo Password Management Script"
		}
		if($SystemManagementPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No system management password is set. Please reboot the computer and manually set a system management password." -TitleText "Lenovo Password Management Script"
		}
	}
	#Exit the script with an error
	if(!($ContinueOnError))
	{
		Write-LogEntry -Value "Failures detected, exiting the script" -Severity 3
		Write-Output "Password management tasks failed. Check the log file for more information"
		Write-LogEntry -Value "END - Lenovo BIOS password management script" -Severity 1
		Exit 1
	}
	else
	{
		Write-LogEntry -Value "Failures detected, but the ContinueOnError parameter was set. Script execution will continue" -Severity 3
		Write-Output "Failures detected, but the ContinueOnError parameter was set. Script execution will continue"
	}
}

#Prompt the user about any password change or clear failures
if((($SupervisorPWSet -eq "Failed") -or ($SupervisorPWClear -eq "Failed") -or ($PowerOnPWSet -eq "Failed") -or ($PowerOnPWClear -eq "Failed") -or ($SystemManagementPWSet -eq "Failed") -or ($SystemManagementPWClear -eq "Failed")) -and (!($SMSTSPasswordRetry)))
{
	if(!($NoUserPrompt))
	{
		Write-LogEntry -Value "Failures detected, display on-screen prompts for any required manual actions" -Severity 2
		#Close the task sequence progress dialog
		if(Get-TaskSequenceStatus)
		{
			$TSProgress.CloseProgressDialog()
		}
		if($SupervisorPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The supervisor password is set, but cannot be automatically changed. Please reboot the computer and manually change the supervisor password." -TitleText "Lenovo Password Management Script"
		}
		if($SupervisorPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The supervisor password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the supervisor password." -TitleText "Lenovo Password Management Script"
		}
		if($PowerOnPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically changed. Please reboot the computer and manually change the power on password." -TitleText "Lenovo Password Management Script"
		}
		if($PowerOnPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The power on password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the power on password." -TitleText "Lenovo Password Management Script"
		}
		if($SystemManagementPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The system management password is set, but cannot be automatically changed. Please reboot the computer and manually change the system management password." -TitleText "Lenovo Password Management Script"
		}
		if($SystemManagementPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The system management password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the system management password." -TitleText "Lenovo Password Management Script"
		}
	}
	#Exit the script with an error
	if(!($ContinueOnError))
	{
		Write-LogEntry -Value "Failures detected, exiting the script" -Severity 3
		Write-Output "Password management tasks failed. Check the log file for more information"
		Write-LogEntry -Value "END - Lenovo BIOS password management script" -Severity 1
		Exit 1
	}
	else
	{
		Write-LogEntry -Value "Failures detected, but the ContinueOnError parameter was set. Script execution will continue" -Severity 3
		Write-Output "Failures detected, but the ContinueOnError parameter was set. Script execution will continue"
	}
}
elseif((($SupervisorPWExists -eq "Failed") -or ($SupervisorPWSet -eq "Failed") -or ($SupervisorPWClear -eq "Failed") -or ($PowerOnPWExists -eq "Failed") -or ($PowerOnPWSet -eq "Failed") -or ($PowerOnPWClear -eq "Failed") -or ($SystemManagementPWSet -eq "Failed") -or ($SystemManagementPWClear -eq "Failed")) -and ($SMSTSPasswordRetry))
{
	Write-LogEntry -Value "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed" -Severity 3
	Write-Output "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed"
}
else
{
	Write-Output "Password management tasks succeeded. Check the log file for more information"
}
Write-LogEntry -Value "END - Lenovo BIOS password management script" -Severity 1
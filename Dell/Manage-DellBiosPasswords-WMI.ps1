<#
	.DESCRIPTION
		Automatically configure Dell BIOS passwords and prompt the user if manual intervention is required.

	.PARAMETER AdminSet
		Specify this switch to set a new admin password or change an existing admin password.

	.PARAMETER AdminClear
		Specify this swtich to clear an existing admin password. Must also specify the OldAdminPassword parameter.

	.PARAMETER SystemSet
		Specify this switch to set a new system password or change an existing system password.

	.PARAMETER SystemClear
		Specify this switch to clear an existing system password. Must also specify the OldSystemPassword parameter.

	.PARAMETER AdminPassword
		Specify the new admin password to set.

	.PARAMETER OldAdminPassword
		Specify the old admin password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

	.PARAMETER SystemPassword
		Specify the new system password to set.

	.PARAMETER OldSystemPassword
		Specify the old system password(s) to be changed. Multiple passwords can be specified as a comma seperated list.
	
	.PARAMETER NoUserPrompt
		The script will run silently and will not prompt the user with a message box.

	.PARAMETER ContinueOnError
		The script will ignore any errors caused by changing or clearing the passwords. This will not suppress errors caused by parameter validation.

	.PARAMETER SMSTSPasswordRetry
		For use in a task sequence. If specified, the script will assume the script needs to run at least one more time. This will ignore password errors and suppress user prompts.

	.PARAMETER LogFile
		Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

	.EXAMPLE
		Set a new admin password
		Manage-DellBiosPasswords-WMI.ps1 -AdminSet -AdminPassword <String>
	
		Set or change a admin password
		Manage-DellBiosPasswords-WMI.ps1 -AdminSet -AdminPassword <String> -OldAdminPassword <String1>,<String2>,<String3>

		Clear existing admin password(s)
		Manage-DellBiosPasswords-WMI.ps1 -AdminClear -OldAdminPassword <String1>,<String2>,<String3>

		Set a new admin password and set a new system password
		Manage-DellBiosPasswords-WMI.ps1 -AdminSet -SystemSet -AdminPassword <String> -SystemPassword <String>

	.NOTES
		Created by: Jon Anderson (@ConfigJon)
		Reference: https://www.configjon.com/dell-bios-password-management-wmi/
		Modified: 2020-09-17

	.CHANGELOG
		2020-09-14 - When using the AdminSet and SystemSet parameters, the OldPassword parameters are no longer required. There is now logic to handle and report this type of failure.
		2020-09-17 - Improved the log file path configuration

#>

#Parameters ===================================================================================================================

param(
	[Parameter(Mandatory=$false)][Switch]$AdminSet,
	[Parameter(Mandatory=$false)][Switch]$AdminClear,
	[Parameter(Mandatory=$false)][Switch]$SystemSet,
	[Parameter(Mandatory=$false)][Switch]$SystemClear,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$AdminPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldAdminPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SystemPassword,
	[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSystemPassword,
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
	[System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\Dell\Manage-DellBiosPasswords-WMI.log"
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
	$Counter = 0
	while($Counter -lt 6)
	{
		if($CmdletType -eq "CIM")
		{
			if($Select)
			{
				Write-LogEntry -Value "Get the $Classname WMI class from the $Namespace namespace and select properties: $Select" -Severity 1
				$Query = Get-CimInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction SilentlyContinue | Select-Object $Select -ErrorAction SilentlyContinue
			}
			else
			{
				Write-LogEntry -Value "Get the $ClassName WMI class from the $Namespace namespace" -Severity 1
				$Query = Get-CimInstance -Namespace $Namespace -ClassName $ClassName -ErrorAction SilentlyContinue
			}
		}
		elseif($CmdletType -eq "WMI")
		{
			if($Select)
			{
				Write-LogEntry -Value "Get the $Classname WMI class from the $Namespace namespace and select properties: $Select" -Severity 1
				$Query = Get-WmiObject -Namespace $Namespace -Class $ClassName -ErrorAction SilentlyContinue | Select-Object $Select -ErrorAction SilentlyContinue
			}
			else
			{
				Write-LogEntry -Value "Get the $ClassName WMI class from the $Namespace namespace" -Severity 1
				$Query = Get-WmiObject -Namespace $Namespace -Class $ClassName -ErrorAction SilentlyContinue
			}
		}
		if($Query -eq $NULL)
		{
			if($Select)
			{
				Write-LogEntry -Value "An error occurred while attempting to get the $Select properties from the $Classname WMI class in the $Namespace namespace. Retry in 30 seconds" -Severity 2
			}
			else
			{
				Write-LogEntry -Value "An error occurred while connecting to the $Classname WMI class in the $Namespace namespace. Retry in 30 seconds" -Severity 2
			}
			Start-Sleep -Seconds 30
			$Counter++
		}
		else
		{
			break
		}
	}
	if($Query -eq $NULL)
	{
		if($Select)
		{
			Stop-Script -ErrorMessage "An error occurred while attempting to get the $Select properties from the $Classname WMI class in the $Namespace namespace"
		}
		else
		{
			Stop-Script -ErrorMessage "An error occurred while connecting to the $Classname WMI class in the $Namespace namespace"
		}
	}
	Write-LogEntry -Value "Successfully connected to the $ClassName WMI class" -Severity 1
	return $Query
}

Function New-DellBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Admin','System')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Password,
		[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$AdminPW
	)
	#Attempt to set the system password when the admin password is already set
	if($AdminPW)
	{
		#Encode the AdminPassword
		$AdminBytes = $Script:Encoder.GetBytes($AdminPW)

		if(($SecurityInterface.SetNewPassword(1,$AdminBytes.Length,$AdminBytes,$PasswordType,$AdminPW,$Password)).Status -eq 0)
		{
			Write-LogEntry -Value "The $PasswordType password has been successfully set" -Severity 1
		}
		else
		{
			Set-Variable -Name "$($PasswordType)PWExists" -Value "Failed" -Scope Script
			Write-LogEntry -Value "Failed to set the $PasswordType password" -Severity 3
		}
	}
	#Attempt to set the admin or system password
	else
	{
		if(($SecurityInterface.SetNewPassword(0,0,0,$PasswordType,"",$Password)).Status -eq 0)
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

Function Set-DellBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Admin','System')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Password,
		[Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPassword
	)
	#Encode the password
	$PasswordBytes = $Script:Encoder.GetBytes($Password)
	Write-LogEntry -Value "Attempt to change the existing $PasswordType password" -Severity 1
	Set-Variable -Name "$($PasswordType)PWSet" -Value "Failed" -Scope Script
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("DellSet$($PasswordType)") = "Failed"
	}
	#Check if the password is already set to the correct value
	if(($SecurityInterface.SetNewPassword(1,$PasswordBytes.Length,$PasswordBytes,$PasswordType,$Password,$Password)).Status -eq 0)
	{
		#Password is set to correct value
		Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
		if(Get-TaskSequenceStatus)
		{
			$TSEnv.Value("DellSet$($PasswordType)") = "Success"
		}
		Write-LogEntry -Value "The $PasswordType password is already set correctly" -Severity 1
	}
	#Password is not set to correct value
	else
	{
		if($OldPassword)
		{
			$Counter = 0
			while($Counter -lt $OldPassword.Count)
			{
				#Encode the old password
				$OldBytes = $Script:Encoder.GetBytes($OldPassword[$Counter])
				#Attempt to change the password
				if(($SecurityInterface.SetNewPassword(1,$OldBytes.Length,$OldBytes,$PasswordType,$OldPassword[$Counter],$Password)).Status -eq 0)
				{
					#Successfully changed the password
					Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
					if(Get-TaskSequenceStatus)
					{
						$TSEnv.Value("DellSet$($PasswordType)") = "Success"
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
			#Report password change failure
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

Function Clear-DellBiosPassword
{
	param(
		[Parameter(Mandatory=$true)][ValidateSet('Admin','System')]$PasswordType,
		[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$OldPassword	
	)
	Write-LogEntry -Value "Attempt to clear the existing $PasswordType password" -Severity 1
	Set-Variable -Name "$($PasswordType)PWClear" -Value "Failed" -Scope Script
	if(Get-TaskSequenceStatus)
	{
		$TSEnv.Value("DellClear$($PasswordType)") = "Failed"
	}
	$Counter = 0
	while($Counter -lt $OldPassword.Count)
	{
		#Encode the old password
		$OldBytes = $Script:Encoder.GetBytes($OldPassword[$Counter])
		#Attempt to clear the password
		if(($SecurityInterface.SetNewPassword(1,$OldBytes.Length,$OldBytes,$PasswordType,$OldPassword[$Counter],"")).Status -eq 0)
		{
			#Successfully cleared the password
			Set-Variable -Name "$($PasswordType)PWClear" -Value "Success" -Scope Script
			if(Get-TaskSequenceStatus)
			{
				$TSEnv.Value("DellClear$($PasswordType)") = "Success"
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
	#Report password clear failure
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
	# Determine log file location
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
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-DellBiosPasswords-WMI"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
Write-LogEntry -Value "START - Dell BIOS password management script" -Severity 1

#Connect to the SecurityInterface WMI class
$SecurityInterface = Get-WmiData -Namespace root\dcim\sysman\wmisecurity -ClassName SecurityInterface -CmdletType WMI

#Connect to the PasswordObject CIM instance
$PasswordObject = Get-WmiData -Namespace root\dcim\sysman\wmisecurity -ClassName PasswordObject -CmdletType CIM

#Create the encoding object used for working with the passwords
$Encoder = New-Object System.Text.UTF8Encoding

#Get the current password status
Write-LogEntry -Value "Get the current password state" -Severity 1

$AdminPasswordCheck = ($PasswordObject | Where-Object NameId -EQ "Admin").IsPasswordSet
if($AdminPasswordCheck -eq 1)
{
	Write-LogEntry -Value "The Admin password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The Admin password is not currently set" -Severity 1
}
$SystemPasswordCheck = ($PasswordObject | Where-Object NameId -EQ "System").IsPasswordSet
if($SystemPasswordCheck -eq 1)
{
	Write-LogEntry -Value "The System password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The System password is not currently set" -Severity 1
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1
if(($AdminSet) -and !($AdminPassword))
{
	Stop-Script -ErrorMessage "When using the AdminSet switch, the AdminPassword parameter must also be specified"
}
if(($AdminClear) -and !($OldAdminPassword))
{
	Stop-Script -ErrorMessage "When using the AdminClear switch, the OldAdminPassword parameter must also be specified"
}
if(($SystemSet) -and !($SystemPassword))
{
	Stop-Script -ErrorMessage "When using the SystemSet switch, the SystemPassword parameter must also be specified"
}
if(($SystemSet -and $AdminPasswordCheck -eq 1) -and !($AdminPassword))
{
	Stop-Script -ErrorMessage "When attempting to set a system password while the admin password is already set, the AdminPassword parameter must be specified"
}
if(($SystemClear) -and !($OldSystemPassword))
{
	Stop-Script -ErrorMessage "When using the SystemClear switch, the OldSystemPassword parameter must also be specified"
}
if(($AdminSet) -and ($AdminClear))
{
	Stop-Script -ErrorMessage "Cannot specify the AdminSet and AdminClear parameters simultaneously"
}
if(($SystemSet) -and ($SystemClear))
{
	Stop-Script -ErrorMessage "Cannot specify the SystemSet and SystemClear parameters simultaneously"
}
if(($OldAdminPassword) -and !($AdminSet -or $AdminClear))
{
	Stop-Script -ErrorMessage "When using the OldAdminPassword parameter, one of the AdminSet or AdminClear parameters must also be specified"
}
if(($AdminPassword) -and !($AdminSet -or $AdminClear -or $SystemSet))
{
	Stop-Script -ErrorMessage "When using the AdminPassword parameter, one of the AdminSet or AdminClear or SystemSet parameters must also be specified"
}
if(($OldSystemPassword -or $SystemPassword) -and !($SystemSet -or $SystemClear))
{
	Stop-Script -ErrorMessage "When using the OldSystemPassword or SystemPassword parameters, one of the SystemSet or SystemClear parameters must also be specified"
}
if(($AdminClear) -and ($SystemPasswordCheck -eq 1))
{
	Write-LogEntry -Value "Warning: The the AdminClear parameter has been specified and the system password is set. Clearing the admin password will also clear the system password." -Severity 2
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
	$DellSetAdmin = $TSEnv.Value("DellSetAdmin")
	if($DellSetAdmin -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful admin password set attempt detected" -Severity 1
	}
	$DellClearAdmin = $TSEnv.Value("DellClearAdmin")
	if($DellClearAdmin -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful admin password clear attempt detected" -Severity 1
	}
	$DellSetSystem = $TSEnv.Value("DellSetSystem")
	if($DellSetSystem -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful system password set attempt detected" -Severity 1
	}
	$DellClearSystem = $TSEnv.Value("DellClearSystem")
	if($DellClearSystem -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful system password clear attempt detected" -Severity 1
	}
}

#No admin password currently set
if($AdminPasswordCheck -eq 0)
{
    if($AdminClear)
    {
        Write-LogEntry -Value "No admin password currently set. No need to clear the admin password" -Severity 2
        Clear-Variable AdminClear
    }
    if($AdminSet)
    {
		if($SystemPasswordCheck -eq 1)
		{
			$SystemAlreadySet = "Failed"
			Write-LogEntry -Value "Failed to set the admin password. The system password is already set." -Severity 3			
		}
		else
		{
			New-DellBiosPassword -PasswordType Admin -Password $AdminPassword
		}
	}
}

#No system password currently set
if($SystemPasswordCheck -eq 0)
{
    if($SystemClear)
    {
        Write-LogEntry -Value "No system password currently set. No need to clear the system password" -Severity 2
        Clear-Variable SystemClear
    }
    if($SystemSet)
	{
		#If the admin password is currently set, the admin password is required to set the system password
		if(($PasswordObject | Where-Object NameId -EQ "Admin").IsPasswordSet -eq 1)
		{
			New-DellBiosPassword -PasswordType System -Password $SystemPassword -AdminPW $AdminPassword
		}
		else
		{
			New-DellBiosPassword -PasswordType System -Password $SystemPassword
		}
	}
}

#If a admin password is set, attempt to clear or change it
if($AdminPasswordCheck -eq 1)
{
	#Change the existing admin password
	if(($AdminSet) -and ($DellSetAdmin -ne "Success"))
	{
		if($OldAdminPassword)
		{
			Set-DellBiosPassword -PasswordType Admin -Password $AdminPassword -OldPassword $OldAdminPassword
		}
		else
		{
			Set-DellBiosPassword -PasswordType Admin -Password $AdminPassword
		}
	}
	#Clear the existing admin password
	if(($AdminClear) -and ($DellClearAdmin -ne "Success"))
	{
		Clear-DellBiosPassword -PasswordType Admin -OldPassword $OldAdminPassword
	}
}

#If a system password is set, attempt to clear or change it
if($SystemPasswordCheck -eq 1)
{
	#Change the existing system password
	if(($SystemSet) -and ($DellSetSystem -ne "Success"))
	{
		if($OldSystemPassword)
		{
			Set-DellBiosPassword -PasswordType System -Password $SystemPassword -OldPassword $OldSystemPassword
		}
		else
		{
			Set-DellBiosPassword -PasswordType System -Password $SystemPassword
		}
	}
	#Clear the existing system password
	if(($SystemClear) -and ($DellClearSystem -ne "Success"))
	{
		Clear-DellBiosPassword -PasswordType System -OldPassword $OldSystemPassword
	}
}

#Prompt the user about any failures
if((($AdminPWExists -eq "Failed") -or ($AdminPWSet -eq "Failed") -or ($AdminPWClear -eq "Failed") -or ($SystemPWExists -eq "Failed") -or ($SystemPWSet -eq "Failed") -or ($SystemPWClear -eq "Failed") -or ($SystemAlreadySet -eq "Failed")) -and (!($SMSTSPasswordRetry)))
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
		if($AdminPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No admin password is set, but the script was unable to set a password. Please reboot the computer and manually set the admin password." -TitleText "Dell Password Management Script"
		}
		if($AdminPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The admin password is set, but cannot be automatically changed. Please reboot the computer and manually change the admin password." -TitleText "Dell Password Management Script"
		}
		if($AdminPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The admin password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the admin password." -TitleText "Dell Password Management Script"
		}
		if($SystemPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No system password is set, but the script was unable to set a password. Please reboot the computer and manually set the system password." -TitleText "Dell Password Management Script"
		}
		if($SystemPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The system password is set, but cannot be automatically changed. Please reboot the computer and manually change the system password." -TitleText "Dell Password Management Script"
		}
		if($SystemPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The system password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the system password." -TitleText "Dell Password Management Script"
		}
		if($SystemAlreadySet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The admin password cannot be set while the system password is set. Please reboot the computer and manually clear the system password." -TitleText "Dell Password Management Script"
		}
	}
	#Exit the script with an error
	if(!($ContinueOnError))
	{
		Write-LogEntry -Value "Failures detected, exiting the script" -Severity 3
		Write-Output "Password management tasks failed. Check the log file for more information"
		Write-LogEntry -Value "END - Dell BIOS password management script" -Severity 1
		Exit 1
	}
	else
	{
		Write-LogEntry -Value "Failures detected, but the ContinueOnError parameter was set. Script execution will continue" -Severity 3
		Write-Output "Failures detected, but the ContinueOnError parameter was set. Script execution will continue"
	}
}
elseif((($AdminPWExists -eq "Failed") -or ($AdminPWSet -eq "Failed") -or ($AdminPWClear -eq "Failed") -or ($SystemPWExists -eq "Failed") -or ($SystemPWSet -eq "Failed") -or ($SystemPWClear -eq "Failed") -or ($SystemAlreadySet -eq "Failed")) -and ($SMSTSPasswordRetry))
{
	Write-LogEntry -Value "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed" -Severity 3
	Write-Output "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed"
}
else
{
	Write-Output "Password management tasks succeeded. Check the log file for more information"
}
Write-LogEntry -Value "END - Dell BIOS password management script" -Severity 1
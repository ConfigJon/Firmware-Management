<#
	.DESCRIPTION
		Automatically configure Dell BIOS passwords and prompt the user if manual intervention is required.

	.PARAMETER AdminSet
		Specify this switch to set a new admin password or change an existing admin password.

	.PARAMETER AdminClear
		Specify this swtich to clear an existing admin password. Must also specify the OldAdminPassword parameter.

	.PARAMETER SystemSet
		Specify this switch to set a new system password or change an existing setup password.

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

	.EXAMPLE
		Set a new admin password
		Manage-DellBiosPasswords.ps1 -AdminSet -AdminPassword <String>

		Set or change a admin password
		Manage-DellBiosPasswords.ps1 -AdminSet -AdminPassword <String> -OldAdminPassword <String1>,<String2>,<String3>

		Clear existing admin password(s)
		Manage-DellBiosPasswords.ps1 -AdminClear -OldAdminPassword <String1>,<String2>,<String3>

		Set a new admin password and set a new system password
		Manage-DellBiosPasswords.ps1 -AdminSet -SystemSet -AdminPassword <String> -SystemPassword <String>

	.NOTES
		Created by: Jon Anderson (@ConfigJon)
		Reference: https://www.configjon.com/dell-bios-password-management/
		Modified: 01/30/2020

	.CHANGELOG
		01/30/2020 - Removed the AdminChange and SystemChange parameters. AdminSet and SystemSet now work to set or change a password. Changed the DellChangeAdmin task sequence variable to DellSetAdmin.
					 Changed the DellChangeSystem task sequence variable to DellSetSystem. Updated the parameter validation checks.
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
		[string]$FileName = "Manage-DellBiosPasswords.log"
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
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-DellBiosPasswords"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"

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
	$LogsDirectory = "$ENV:ProgramData\BiosScripts\Dell"
	if (!(Test-Path -PathType Container $LogsDirectory))
	{
		New-Item -Path $LogsDirectory -ItemType "Directory" -Force | Out-Null
	}
}
Write-Output "Log path set to $LogsDirectory\Manage-DellBiosPasswords.log"
Write-LogEntry -Value "START - Dell BIOS password management script" -Severity 1

#Check if 32 or 64 bit architecture
if ([System.Environment]::Is64BitOperatingSystem)
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
	$LocalVersion = Get-Module -ListAvailable -Name 'DellBIOSProvider' -ErrorAction Stop | Select-Object -ExpandProperty Version
}
catch
{
    $Local = $true
    if (Test-Path "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider")
    {
        $LocalVersion = Get-Content "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider\DellBIOSProvider.psd1" | Select-String "ModuleVersion ="
        $LocalVersion = (([regex]".*'(.*)'").Matches($LocalVersion))[0].Groups[1].Value
        if ($NULL -ne $LocalVersion)
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
if (($NULL -ne $LocalVersion) -and (!($Local)))
{
    Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
}

#Verify the DellBIOSProvider module is imported
Write-LogEntry -Value "Verify the DellBIOSProvider module is imported" -Severity 1
$ModuleCheck = Get-Module DellBIOSProvider
if ($ModuleCheck)
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
    if (!($Error))
    {
        Write-LogEntry -Value "Successfully imported the DellBIOSProvider module" -Severity 1
    }
}

#Get the current password status
Write-LogEntry -Value "Get the current password state" -Severity 1

$AdminPasswordCheck = Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet | Select-Object -ExpandProperty CurrentValue
if ($AdminPasswordCheck -eq "True")
{
	Write-LogEntry -Value "The Admin password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The Admin password is not currently set" -Severity 1
}

$SystemPasswordCheck = Get-Item -Path DellSmbios:\Security\IsSystemPasswordSet | Select-Object -ExpandProperty CurrentValue
if ($SystemPasswordCheck -eq "True")
{
	Write-LogEntry -Value "The System password is currently set" -Severity 1
}
else
{
	Write-LogEntry -Value "The System password is not currently set" -Severity 1
}

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1

if (($AdminSet) -and !($AdminPassword))
{
	$ErrorMsg = "When using the AdminSet switch, the AdminPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($AdminSet -and $AdminPasswordCheck -eq 1) -and !($OldAdminPassword))
{
	$ErrorMsg = "When using the AdminSet switch where a admin password exists, the OldAdminPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($AdminClear) -and !($OldAdminPassword))
{
	$ErrorMsg = "When using the AdminClear switch, the OldAdminPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SystemSet) -and !($SystemPassword))
{
	$ErrorMsg = "When using the SystemSet switch, the SystemPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SystemSet -and $SystemPasswordCheck -eq 1) -and !($OldSystemPassword))
{
	$ErrorMsg = "When using the SystemSet switch where a system password exists, the OldSystemPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SystemSet -and $AdminPasswordCheck -eq "True") -and !($AdminPassword))
{
	$ErrorMsg = "When attempting to set a system password while the admin password is already set, the AdminPassword parameter must be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SystemClear) -and !($OldSystemPassword))
{
	$ErrorMsg = "When using the SystemClear switch, the OldSystemPassword parameter must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($AdminSet) -and ($AdminClear))
{
	$ErrorMsg = "Cannot specify the AdminSet and AdminClear parameters simultaneously"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($SystemSet) -and ($SystemClear))
{
	$ErrorMsg = "Cannot specify the SystemSet and SystemClear parameters simultaneously"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($OldAdminPassword) -and !($AdminSet -or $AdminClear))
{
	$ErrorMsg = "When using the OldAdminPassword parameter, one of the AdminSet or AdminClear parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($AdminPassword) -and !($AdminSet -or $AdminClear -or $SystemSet))
{
	$ErrorMsg = "When using the AdminPassword parameter, one of the AdminSet or AdminClear or SystemSet parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($OldSystemPassword -or $SystemPassword) -and !($SystemSet -or $SystemClear))
{
	$ErrorMsg = "When using the OldSystemPassword or SystemPassword parameters, one of the SystemSet or SystemClear parameters must also be specified"
	Write-LogEntry -Value $ErrorMsg -Severity 3
	throw $ErrorMsg
}
if (($AdminClear) -and ($SystemPasswordCheck -eq "True"))
{
	Write-LogEntry -Value "Warning: The the AdminClear parameter has been specified and the system password is set. Clearing the admin password will also clear the system password." -Severity 2
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
	$DellSetAdmin = $TSEnv.Value("DellSetAdmin")
	if ($DellSetAdmin -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful admin password set attempt detected" -Severity 1
	}
	$DellClearAdmin = $TSEnv.Value("DellClearAdmin")
	if ($DellClearAdmin -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful admin password clear attempt detected" -Severity 1
	}
	$DellSetSystem = $TSEnv.Value("DellSetSystem")
	if ($DellSetSystem -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful system password set attempt detected" -Severity 1
	}
	$DellClearSystem = $TSEnv.Value("DellClearSystem")
	if ($DellClearSystem -eq "Failed")
	{
		Write-LogEntry -Value "Previous unsuccessful system password clear attempt detected" -Severity 1
	}
}

#No admin password currently set
if ($AdminPasswordCheck -eq "False")
{
    if ($AdminClear)
    {
        Write-LogEntry -Value "No admin password currently set. No need to clear the admin password" -Severity 2
        Clear-Variable AdminClear
    }
    if ($AdminSet)
    {
		if ($SystemPasswordCheck -eq "True")
		{
			$SystemAlreadySet = "Failed"
			Write-LogEntry -Value "Failed to set the admin password. The system password is already set." -Severity 3
		}
		else
		{
			$Error.Clear()
			try
			{
				Set-Item -Path DellSmbios:\Security\AdminPassword $AdminPassword -ErrorAction Stop
			}
			catch
			{
				$AdminPWExists = "Failed"
				Write-LogEntry -Value "Failed to set the admin password" -Severity 3
			}
			if (!($Error))
			{
				Write-LogEntry -Value "The admin password has been successfully set" -Severity 1
			}
		}
	}
}

#No system password currently set
if ($SystemPasswordCheck -eq "False")
{
    if ($SystemClear)
    {
        Write-LogEntry -Value "No system password currently set. No need to clear the system password" -Severity 2
        Clear-Variable SystemClear
    }
    if ($SystemSet)
	{
		if ((Get-Item -Path DellSmbios:\Security\IsAdminPasswordSet | Select-Object -ExpandProperty CurrentValue) -eq "True")
		{
			$Error.Clear()
			try
			{
				Set-Item -Path DellSmbios:\Security\SystemPassword $SystemPassword -Password $AdminPassword -ErrorAction Stop
			}
			catch
			{
				$SystemPWExists = "Failed"
				Write-LogEntry -Value "Failed to set the system password" -Severity 3
			}
			if (!($Error))
			{
				Write-LogEntry -Value "The system password has been successfully set" -Severity 1
			}
		}
		else
		{
			$Error.Clear()
			try
			{
				Set-Item -Path DellSmbios:\Security\SystemPassword $SystemPassword -ErrorAction Stop
			}
			catch
			{
				$SystemPWExists = "Failed"
				Write-LogEntry -Value "Failed to set the system password" -Severity 3
			}
			if (!($Error))
			{
				Write-LogEntry -Value "The system password has been successfully set" -Severity 1
			}
		}
	}
}

#If a admin password is set, attempt to clear or change it
if ($AdminPasswordCheck -eq "True")
{
	#Change the existing admin password
	if (($AdminSet) -and ($DellSetAdmin -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to change the existing admin password" -Severity 1
		$AdminPWSet = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("DellSetAdmin") = "Failed"
		}

		try
		{
			Set-Item -Path DellSmbios:\Security\AdminPassword $AdminPassword -Password $AdminPassword -ErrorAction Stop
		}
		catch
		{
			$AdminSetFail = $true
			$Counter = 0
			While($Counter -lt $OldAdminPassword.Count){
				$Error.Clear()
				try
				{
					Set-Item -Path DellSmbios:\Security\AdminPassword $AdminPassword -Password $OldAdminPassword[$Counter] -ErrorAction Stop
				}
				catch
				{
					#Failed to change the password
					$Counter++
				}
				if (!($Error))
				{
					#Successfully changed the password
					$AdminPWSet = "Success"
					if (Get-TaskSequenceStatus)
					{
						$TSEnv.Value("DellSetAdmin") = "Success"
					}
					Write-LogEntry -Value "The admin password has been successfully changed" -Severity 1
					break
				}
			}
			if ($AdminPWSet -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the admin password" -Severity 3
			}
		}
		if (!($AdminSetFail))
		{
			#Password already correct
			$AdminPWSet = "Success"
			if (Get-TaskSequenceStatus)
			{
				$TSEnv.Value("DellSetAdmin") = "Success"
			}
			Write-LogEntry -Value "The admin password is already set correctly" -Severity 1
		}
	}

	#Clear the existing admin password
	if (($AdminClear) -and ($DellClearAdmin -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to clear the existing admin password" -Severity 1
		$AdminPWClear = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("DellClearAdmin") = "Failed"
		}

		$Counter = 0
		While($Counter -lt $OldAdminPassword.Count){
			$Error.Clear()
			try
			{
				Set-Item -Path DellSmbios:\Security\AdminPassword "" -Password $OldAdminPassword[$Counter] -ErrorAction Stop
			}
			catch
			{
				#Failed to clear the password
				$Counter++
			}
			if (!($Error))
			{
				#Successfully cleared the password
				$AdminPWClear = "Success"
				if (Get-TaskSequenceStatus)
				{
					$TSEnv.Value("DellClearAdmin") = "Success"
				}
				if ($SystemPasswordCheck -eq "True")
				{
					Write-LogEntry -Value "The admin password and system password have been successfully cleared" -Severity 1
					break
				}
				else
				{
					Write-LogEntry -Value "The admin password has been successfully cleared" -Severity 1
					break
				}
			}
			if ($AdminPWClear -eq "Failed")
			{
				Write-LogEntry -Value "Failed to clear the admin password" -Severity 3
			}
		}
	}
}

#If a system password is set, attempt to clear or change it
if ($SystemPasswordCheck -eq "True")
{
	#Change the existing system password
	if (($SystemSet) -and ($DellSetSystem -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to change the existing system password" -Severity 1
		$SystemPWSet = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("DellSetSystem") = "Failed"
		}

		try
		{
			Set-Item -Path DellSmbios:\Security\SystemPassword $SystemPassword -Password $SystemPassword -ErrorAction Stop
		}
		catch
		{
			$SystemSetFail = $true
			$Counter = 0
			While($Counter -lt $OldSystemPassword.Count){
				$Error.Clear()
				try
				{
					Set-Item -Path DellSmbios:\Security\SystemPassword $SystemPassword -Password $OldSystemPassword[$Counter] -ErrorAction Stop
				}
				catch
				{
					#Failed to change the password
					$Counter++
				}
				if (!($Error))
				{
				#Successfully changed the password
					$SystemPWSet = "Success"
					if (Get-TaskSequenceStatus)
					{
						$TSEnv.Value("DellSetSystem") = "Success"
					}
					Write-LogEntry -Value "The system password has been successfully changed" -Severity 1
					break
				}
			}
			if ($SystemPWSet -eq "Failed")
			{
				Write-LogEntry -Value "Failed to change the system password" -Severity 3
			}
		}
		if (!($SystemSetFail))
		{
			#Password already correct
			$SystemPWSet = "Success"
			if (Get-TaskSequenceStatus)
			{
				$TSEnv.Value("DellSetSystem") = "Success"
			}
			Write-LogEntry -Value "The system password is already set correctly" -Severity 1
		}
	}

	#Clear the existing system password
	if (($SystemClear) -and ($DellClearSystem -ne "Success"))
	{
		Write-LogEntry -Value "Attempt to clear the existing system password" -Severity 1
		$SystemPWClear = "Failed"
		if (Get-TaskSequenceStatus)
		{
			$TSEnv.Value("DellClearSystem") = "Failed"
		}

		$Counter = 0
		While($Counter -lt $OldSystemPassword.Count){
			$Error.Clear()
			try
			{
				Set-Item -Path DellSmbios:\Security\SystemPassword "" -Password $OldSystemPassword[$Counter] -ErrorAction Stop
			}
			catch
			{
				#Failed to clear the password
				$Counter++
			}
			if (!($Error))
			{
				#Successfully cleared the password
				$SystemPWClear = "Success"
				if (Get-TaskSequenceStatus)
				{
					$TSEnv.Value("DellClearSystem") = "Success"
				}
				Write-LogEntry -Value "The system password has been successfully cleared" -Severity 1
				break
			}
		}
		if ($SystemPWClear -eq "Failed")
		{
			Write-LogEntry -Value "Failed to clear the system password" -Severity 3
		}
	}
}

#Prompt the user about any failures
if ((($AdminPWExists -eq "Failed") -or ($AdminPWSet -eq "Failed") -or ($AdminPWClear -eq "Failed") -or ($SystemPWExists -eq "Failed") -or ($SystemPWSet -eq "Failed") -or ($SystemPWClear -eq "Failed") -or ($SystemAlreadySet -eq "Failed")) -and (!($SMSTSPasswordRetry)))
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
		if ($AdminPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No admin password is set, but the script was unable to set a password. Please reboot the computer and manually set the admin password." -TitleText "Dell Password Management Script"
		}
		if ($AdminPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The admin password is set, but cannot be automatically changed. Please reboot the computer and manually change the admin password." -TitleText "Dell Password Management Script"
		}
		if ($AdminPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The admin password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the admin password." -TitleText "Dell Password Management Script"
		}
		if ($SystemPWExists -eq "Failed")
		{
			Start-UserPrompt -BodyText "No system password is set, but the script was unable to set a password. Please reboot the computer and manually set the system password." -TitleText "Dell Password Management Script"
		}
		if ($SystemPWSet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The system password is set, but cannot be automatically changed. Please reboot the computer and manually change the system password." -TitleText "Dell Password Management Script"
		}
		if ($SystemPWClear -eq "Failed")
		{
			Start-UserPrompt -BodyText "The system password is set, but cannot be automatically cleared. Please reboot the computer and manually clear the system password." -TitleText "Dell Password Management Script"
		}
		if ($SystemAlreadySet -eq "Failed")
		{
			Start-UserPrompt -BodyText "The admin password cannot be set while the system password is set. Please reboot the computer and manually clear the system password." -TitleText "Dell Password Management Script"
		}
	}
	#Exit the script with an error
	if (!($ContinueOnError))
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
elseif ((($AdminPWExists -eq "Failed") -or ($AdminPWSet -eq "Failed") -or ($AdminPWClear -eq "Failed") -or ($SystemPWExists -eq "Failed") -or ($SystemPWSet -eq "Failed") -or ($SystemPWClear -eq "Failed") -or ($SystemAlreadySet -eq "Failed")) -and ($SMSTSPasswordRetry))
{
	Write-LogEntry -Value "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed" -Severity 3
	Write-Output "Failures detected, but the SMSTSPasswordRetry parameter was set. No user prompts will be displayed"
}
else
{
	Write-Output "Password management tasks succeeded. Check the log file for more information"
}
Write-LogEntry -Value "END - Dell BIOS password management script" -Severity 1
<#
    .DESCRIPTION
        Automatically configure HP BIOS passwords and prompt the user if manual intervention is required.

    .PARAMETER SetupSet
        Specify this switch to set a new setup password or change an existing setup password.

    .PARAMETER SetupClear
        Specify this swtich to clear an existing setup password. Must also specify the OldSetupPassword parameter.

    .PARAMETER PowerOnSet
        Specify this switch to set a new power on password or change an existing power on password. HP firmware requires a setup password to manage the power on password, so a setup password must already be set (specify SetupPassword) or be set in the same run (specify SetupSet and SetupPassword).

    .PARAMETER PowerOnClear
        Specify this switch to clear an existing power on password. When a setup password is set, also specify the SetupPassword parameter (the setup password authorizes the clear). When no setup password is set, also specify the OldPowerOnPassword parameter.

    .PARAMETER SetupPassword
        Specify the new setup password to set.

    .PARAMETER OldSetupPassword
        Specify the old setup password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

    .PARAMETER PowerOnPassword
        Specify the new power on password to set.

    .PARAMETER OldPowerOnPassword
        Specify the old power on password(s) to be changed. Multiple passwords can be specified as a comma seperated list.

    .PARAMETER SetupPasswordCmsFile
        Specify the path to a CMS-encrypted file containing the new setup password. The file is decrypted in memory at runtime using the device's document-encryption certificate. Use this instead of SetupPassword to keep the password off the command line. Cannot be combined with SetupPassword.

    .PARAMETER OldSetupPasswordCmsFile
        Specify the path(s) to CMS-encrypted file(s) containing the old setup password(s) to be changed. Multiple paths can be specified as a comma separated list. Use this instead of OldSetupPassword. Cannot be combined with OldSetupPassword.

    .PARAMETER PowerOnPasswordCmsFile
        Specify the path to a CMS-encrypted file containing the new power on password. The file is decrypted in memory at runtime using the device's document-encryption certificate. Use this instead of PowerOnPassword to keep the password off the command line. Cannot be combined with PowerOnPassword.

    .PARAMETER OldPowerOnPasswordCmsFile
        Specify the path(s) to CMS-encrypted file(s) containing the old power on password(s) to be changed. Multiple paths can be specified as a comma separated list. Use this instead of OldPowerOnPassword. Cannot be combined with OldPowerOnPassword.

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
        Manage-HPBiosPasswords-WMI.ps1 -SetupSet -SetupPassword <String>

        Set or change a setup password
        Manage-HPBiosPasswords-WMI.ps1 -SetupSet -SetupPassword <String> -OldSetupPassword <String1>,<String2>

        Clear existing setup password(s)
        Manage-HPBiosPasswords-WMI.ps1 -SetupClear -OldSetupPassword <String1>,<String2>

        Set a new setup password and set a new power on password when no old passwords exist
        Manage-HPBiosPasswords-WMI.ps1 -SetupSet -PowerOnSet -SetupPassword <String1> -PowerOnPassword <String1>

        Set or change an existing setup password and clear a power on password
        Manage-HPBiosPasswords-WMI.ps1 -SetupSet -SetupPassword <String> -OldSetupPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

        Clear existing Setup and power on passwords
        Manage-HPBiosPasswords-WMI.ps1 -SetupClear -OldSetupPassword <String1>,<String2> -PowerOnClear -OldPowerOnPassword <String1>,<String2>

        Set a new power on password when the setup password is already set
        Manage-HPBiosPasswords-WMI.ps1 -PowerOnSet -PowerOnPassword <String> -SetupPassword <String>

        Clear a power on password when the setup password is set (the setup password authorizes the clear)
        Manage-HPBiosPasswords-WMI.ps1 -PowerOnClear -SetupPassword <String>

        Clear both the setup and power on passwords in a single run (the power on password is cleared first, authorized by the setup password)
        Manage-HPBiosPasswords-WMI.ps1 -SetupClear -OldSetupPassword <String1>,<String2> -PowerOnClear

        Set a new setup password sourced from a CMS-encrypted file
        Manage-HPBiosPasswords-WMI.ps1 -SetupSet -SetupPasswordCmsFile <String>

    .NOTES
        Created by: Jon Anderson
        Reference: https://www.configjon.com/hp-bios-password-management/
        Version: 2.3.0
        Modified: 2026-05-24

    .CHANGELOG
        See .NOTES Reference for additional detail on each release.

        2.3.0 (2026-05-24)
            - Added secure password sourcing. New optional CmsFile parameters mirror each existing password parameter and source the password from a CMS-encrypted file, decrypted
              in memory at runtime, so the password never appears on the command line. The plain-text parameters are unchanged; specifying both variants of the same password is rejected.

        2.2.0 (2026-05-23)
            - Added detection for HP Sure Admin (Enhanced BIOS Authentication Mode). When Sure Admin is enabled the script now logs a clear message and exits cleanly without
              attempting password actions, since Sure Admin requires signed payloads or a local access key rather than a BIOS password.
            - Changing or clearing the power on password now uses the setup password to authorize the operation when a setup password is set, matching how setting a new power on
              password already works, since HP firmware requires the setup password to manage the power on password. OldPowerOnPassword is still used to authorize the change or
              clear on computers that have no setup password set.
            - Added the ability to clear the setup and power on passwords in a single run (SetupClear with PowerOnClear): the power on password is cleared first under setup
              authority, then the setup password itself is cleared.
            - Fixed a parameter validation contradiction that blocked using SetupPassword to authorize PowerOnSet when a setup password is already set (SetupPassword is now valid
              with SetupSet, PowerOnSet, or PowerOnClear; OldSetupPassword still requires SetupSet or SetupClear).
            - Added a validation check that stops the script with a clear message when PowerOnSet is used on a computer with no setup password set without also specifying SetupSet,
              since a setup password must exist to manage the power on password.

        2.1.0 (2026-05-23)
            - Fixed a bug where attempting to clear a non-existent Power-On password cleared the SetupClear variable instead of PowerOnClear.

        2.0.0 (2026-05-22)
            - Renamed to Manage-HPBiosPasswords-WMI.ps1 to distinguish it from the new HPCMSL variant.
            - Migrated all WMI access from Get-WmiObject to Get-CimInstance, and HP_BIOSSettingInterface.SetBIOSSetting calls to Invoke-CimMethod, so the script runs on both Windows
              PowerShell 5.1 and PowerShell 7.
            - Added [CmdletBinding(PositionalBinding=$false)] so all arguments must be named.

        Pre-2.0.0 (legacy)
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

[CmdletBinding(PositionalBinding=$false)]
param(
    [Parameter(Mandatory=$false)][Switch]$SetupSet,
    [Parameter(Mandatory=$false)][Switch]$SetupClear,
    [Parameter(Mandatory=$false)][Switch]$PowerOnSet,
    [Parameter(Mandatory=$false)][Switch]$PowerOnClear,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSetupPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$PowerOnPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPowerOnPassword,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$SetupPasswordCmsFile,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldSetupPasswordCmsFile,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$PowerOnPasswordCmsFile,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPowerOnPasswordCmsFile,
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
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\HP\Manage-HPBiosPasswords-WMI.log"
)

#Script version
$Version = '2.3.0'

#Log component name
$Component = 'Manage-HPBiosPasswords-WMI'

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
        if([string]::IsNullOrEmpty($SMSTSType))
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

Function Get-CmsPassword
{
    #Decrypt one or more CMS-encrypted password files and return the plain-text value(s).
    #Decryption uses the device's store-resident document-encryption certificate (the matching private key must be present in Cert:\LocalMachine\My or Cert:\CurrentUser\My).
    #One plain-text value is returned per file, preserving the input order so [String[]] password slots round-trip.
    #The decrypted value is never written to the log, only the file path and a generic failure reason.

    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String[]]$CmsFile
    )
    $Result = foreach($File in $CmsFile)
    {
        if(!(Test-Path -LiteralPath $File))
        {
            Stop-Script -ErrorMessage "CMS password file not found: $File"
        }
        try
        {
            Unprotect-CmsMessage -LiteralPath $File -ErrorAction Stop
        }
        catch
        {
            Stop-Script -ErrorMessage "Failed to decrypt the CMS password file: $File. Ensure the document-encryption certificate's private key is present in the certificate store" -Exception $_.Exception.Message
        }
    }
    return $Result
}

Function Get-WmiData
{
    #Gets WMI data using the CIM cmdlets and stores the data in a variable

    param(
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$Namespace,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][String]$ClassName,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$Select
    )
    $Counter = 0
    while($Counter -lt 6)
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
        if($null -eq $Query)
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
    if($null -eq $Query)
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
        if((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value=("<utf-16/>" + $Password); Password=("<utf-16/>" + $SetupPW)}).Return -eq 0)
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
        if((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value=("<utf-16/>" + $Password); Password="<utf-16/>"}).Return -eq 0)
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
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPassword,
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
    Write-LogEntry -Value "Attempt to change the existing $PasswordType password" -Severity 1
    Set-Variable -Name "$($PasswordType)PWSet" -Value "Failed" -Scope Script
    if(Get-TaskSequenceStatus)
    {
        $TSEnv.Value("HPSet$($PasswordType)") = "Failed"
    }
    #When a setup password authorizes the change (power on password management), set the new value directly using the setup password
    if($SetupPW)
    {
        if((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value=("<utf-16/>" + $Password); Password=("<utf-16/>" + $SetupPW)}).Return -eq 0)
        {
            Set-Variable -Name "$($PasswordType)PWSet" -Value "Success" -Scope Script
            if(Get-TaskSequenceStatus)
            {
                $TSEnv.Value("HPSet$($PasswordType)") = "Success"
            }
            Write-LogEntry -Value "The $PasswordType password has been successfully changed" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Failed to change the $PasswordType password" -Severity 3
        }
    }
    #Check if the password is already set to the correct value
    elseif((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value=("<utf-16/>" + $Password); Password=("<utf-16/>" + $Password)}).Return -eq 0)
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
                   if((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value=("<utf-16/>" + $Password); Password=("<utf-16/>" + $OldPassword[$Counter])}).Return -eq 0)
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
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$OldPassword,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$SetupPW
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
    #When a setup password authorizes the clear (power on password management), clear using the setup password
    if($SetupPW)
    {
        $Counter = 0
        While($Counter -lt $SetupPW.Count)
        {
            if((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value="<utf-16/>"; Password=("<utf-16/>" + $SetupPW[$Counter])}).Return -eq 0)
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
    }
    else
    {
        $Counter = 0
        While($Counter -lt $OldPassword.Count)
        {
            if((Invoke-CimMethod -InputObject $Interface -MethodName SetBIOSSetting -Arguments @{Name=$PasswordName; Value="<utf-16/>"; Password=("<utf-16/>" + $OldPassword[$Counter])}).Return -eq 0)
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
    #Write data to a CMTrace compatible log file. (Credit to MSEndpointMgr - https://www.msendpointmgr.com/)

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
        [string]$FileName = ($script:LogFile | Split-Path -Leaf),
        [parameter(Mandatory = $false, HelpMessage = "Name of the component that the log entry will be associated with.")]
        [ValidateNotNullOrEmpty()]
        [string]$Component = $script:Component
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
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($Component)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
Write-LogEntry -Value "START - HP BIOS password management script (version $Version)" -Severity 1

#Resolve secure password sources (decrypt any CMS-encrypted password files into the standard password variables)
if($SetupPassword -and $SetupPasswordCmsFile)
{
    Stop-Script -ErrorMessage "Specify either the SetupPassword or the SetupPasswordCmsFile parameter, not both"
}
if($OldSetupPassword -and $OldSetupPasswordCmsFile)
{
    Stop-Script -ErrorMessage "Specify either the OldSetupPassword or the OldSetupPasswordCmsFile parameter, not both"
}
if($PowerOnPassword -and $PowerOnPasswordCmsFile)
{
    Stop-Script -ErrorMessage "Specify either the PowerOnPassword or the PowerOnPasswordCmsFile parameter, not both"
}
if($OldPowerOnPassword -and $OldPowerOnPasswordCmsFile)
{
    Stop-Script -ErrorMessage "Specify either the OldPowerOnPassword or the OldPowerOnPasswordCmsFile parameter, not both"
}
if($SetupPasswordCmsFile)
{
    Write-LogEntry -Value "Decrypting the setup password from the supplied CMS file" -Severity 1
    $SetupPassword = Get-CmsPassword -CmsFile $SetupPasswordCmsFile
}
if($OldSetupPasswordCmsFile)
{
    Write-LogEntry -Value "Decrypting the old setup password(s) from the supplied CMS file(s)" -Severity 1
    $OldSetupPassword = Get-CmsPassword -CmsFile $OldSetupPasswordCmsFile
}
if($PowerOnPasswordCmsFile)
{
    Write-LogEntry -Value "Decrypting the power on password from the supplied CMS file" -Severity 1
    $PowerOnPassword = Get-CmsPassword -CmsFile $PowerOnPasswordCmsFile
}
if($OldPowerOnPasswordCmsFile)
{
    Write-LogEntry -Value "Decrypting the old power on password(s) from the supplied CMS file(s)" -Severity 1
    $OldPowerOnPassword = Get-CmsPassword -CmsFile $OldPowerOnPasswordCmsFile
}

#Connect to the HP_BIOSSettingInterface WMI class
$Interface = Get-WmiData -Namespace root\hp\InstrumentedBIOS -ClassName HP_BIOSSettingInterface

#Connect to the HP_BIOSSetting WMI class
$HPBiosSetting = Get-WmiData -Namespace root\hp\InstrumentedBIOS -ClassName HP_BIOSSetting

#Check for HP Sure Admin (Enhanced BIOS Authentication Mode). When enabled, BIOS changes require signed payloads or a local access key rather than a password, which this script does not perform, so exit cleanly without attempting any password actions.
$SureAdminCurrent = $null
$SureAdminSetting = ($HPBiosSetting | Where-Object Name -eq "Enhanced BIOS Authentication Mode").Value
if($SureAdminSetting)
{
    foreach($SureAdminValue in $SureAdminSetting.Split(','))
    {
        if($SureAdminValue.StartsWith('*'))
        {
            $SureAdminCurrent = $SureAdminValue.Substring(1)
            break
        }
    }
}
if($SureAdminCurrent -eq "Enable")
{
    Write-LogEntry -Value "HP Sure Admin (Enhanced BIOS Authentication Mode) is enabled. This script manages password-based BIOS authentication only. Use HP's Sure Admin tooling (signed payloads or the local access key) on Sure Admin enabled devices." -Severity 2
    Write-Output "HP Sure Admin is enabled. No password actions were taken."
    Write-LogEntry -Value "END - HP BIOS password management script" -Severity 1
    Exit 0
}

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
if(($PowerOnSet -and $SetupPasswordCheck -ne 1) -and !($SetupSet))
{
    Stop-Script -ErrorMessage "When using the PowerOnSet switch on a computer with no setup password set, a setup password must also be set in the same run using the SetupSet and SetupPassword parameters, since HP firmware requires a setup password to manage the power on password"
}
if(($PowerOnClear -and $SetupPasswordCheck -eq 1) -and !($SetupClear) -and !($SetupPassword))
{
    Stop-Script -ErrorMessage "When using the PowerOnClear switch on a computer where the setup password is set, the SetupPassword parameter must also be specified"
}
if(($PowerOnClear -and $SetupPasswordCheck -ne 1) -and !($OldPowerOnPassword))
{
    Stop-Script -ErrorMessage "When using the PowerOnClear switch on a computer with no setup password set, the OldPowerOnPassword parameter must also be specified"
}
if(($SetupSet) -and ($SetupClear))
{
    Stop-Script -ErrorMessage "Cannot specify the SetupSet and SetupClear parameters simultaneously"
}
if(($PowerOnSet) -and ($PowerOnClear))
{
    Stop-Script -ErrorMessage "Cannot specify the PowerOnSet and PowerOnClear parameters simultaneously"
}
if(($SetupPassword) -and !($SetupSet -or $PowerOnSet -or $PowerOnClear))
{
    Stop-Script -ErrorMessage "When using the SetupPassword parameter, one of the SetupSet, PowerOnSet, or PowerOnClear parameters must also be specified"
}
if(($OldSetupPassword) -and !($SetupSet -or $SetupClear))
{
    Stop-Script -ErrorMessage "When using the OldSetupPassword parameter, one of the SetupSet or SetupClear parameters must also be specified"
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
        Clear-Variable PowerOnClear
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

#If a Setup password is set, attempt to change it (clearing the Setup password is handled after the power on password operations below, since the Setup password authorizes those operations and must be removed last)
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
}

#If a power on password is set, attempt to clear or change it
if($PowerOnPasswordCheck -eq 1)
{
    #Change the existing power on password
    if(($PowerOnSet) -and ($HPSetPowerOn -ne "Success"))
    {
        if($SetupPasswordCheck -eq 1 -and $SetupPassword)
        {
            #When the setup password is set, it authorizes changing the power on password
            Set-HPBiosPassword -PasswordType PowerOn -Password $PowerOnPassword -SetupPW $SetupPassword
        }
        elseif($OldPowerOnPassword)
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
        if($SetupPasswordCheck -eq 1 -and $SetupPassword)
        {
            #When the setup password is set and retained, it authorizes clearing the power on password
            Clear-HPBiosPassword -PasswordType PowerOn -SetupPW $SetupPassword
        }
        elseif($SetupPasswordCheck -eq 1 -and $SetupClear -and $OldSetupPassword)
        {
            #When clearing both passwords, the existing setup password authorizes clearing the power on password before the setup password itself is removed
            Clear-HPBiosPassword -PasswordType PowerOn -SetupPW $OldSetupPassword
        }
        else
        {
            Clear-HPBiosPassword -PasswordType PowerOn -OldPassword $OldPowerOnPassword
        }
    }
}

#Clear the existing Setup password (done after the power on password operations, since the Setup password authorizes those operations and must be removed last)
if(($SetupPasswordCheck -eq 1) -and ($SetupClear) -and ($HPClearSetup -ne "Success"))
{
    Clear-HPBiosPassword -PasswordType Setup -OldPassword $OldSetupPassword
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

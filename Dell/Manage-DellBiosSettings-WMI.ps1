<#
    .DESCRIPTION
        Automatically configure Dell BIOS settings

        WMI Status Codes
            0 - Success
            1 - Failed
            2 - Invalid Parameter
            3 - Access Denied
            4 - Not Supported
            5 - Memory Error
            6 - Protocol Error
    
    .PARAMETER GetSettings
        Instruct the script to get a list of current BIOS settings

    .PARAMETER SetSettings
        Instruct the script to set BIOS settings

    .PARAMETER CsvPath
        The path to the CSV file to be imported or exported

    .PARAMETER AdminPassword
        The current BIOS password

    .PARAMETER SetDefaults
        Instructs the script to set all BIOS settings to a default value. Accptable values are (BuiltInSafeDefaults,LastKnownGood,Factory,UserConf1,UserConf2)

    .PARAMETER SetBootOrder
        The desired boot order to be set on the system. Values should be specified in a comma separated list

    .PARAMETER BootMode
        Used with the SetBootOrder switch. Specifies the boot mode the boot order should be set for. Accptable values are (UEFI or Legacy)

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        #Set BIOS settings supplied in the script
        Manage-DellBiosSettings-WMI.ps1 -SetSettings -AdminPassword ExamplePassword

        #Set BIOS settings supplied in a CSV file
        Manage-DellBiosSettings-WMI.ps1 -SetSettings -CsvPath C:\Temp\Settings.csv -AdminPassword ExamplePassword

        #Set all BIOS settings to factory default values
        Manage-DellBiosSettings-WMI.ps1 -SetDefaults Factory -AdminPassword ExamplePassword

        #Set the UEFI boot order
        Manage-DellBiosSettings-WMI.ps1 -SetBootOrder "Windows Boot Manager","Onboard NIC(IPV4)","Onboard NIC(IPV6)" -BootMode UEFI -AdminPassword ExamplePassword

        #Set BIOS settings supplied in the script and set the UEFI boot order
        Manage-DellBiosSettings-WMI.ps1 -SetSettings -SetBootOrder "Windows Boot Manager","Onboard NIC(IPV4)","Onboard NIC(IPV6)" -BootMode UEFI -AdminPassword ExamplePassword

        #Output a list of current BIOS settings to the screen
        Manage-DellBiosSettings-WMI.ps1 -GetSettings

        #Output a list of current BIOS settings to a CSV file
        Manage-DellBiosSettings-WMI.ps1 -GetSettings -CsvPath C:\Temp\Settings.csv

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/dell-bios-settings-management-wmi/
        Modified: 2020-09-17

    .CHANGELOG
        2020-09-17 - Improved the log file path configuration
        
#>

#Parameters ===================================================================================================================

param(
    [Parameter(Mandatory=$false)][Switch]$GetSettings,
    [Parameter(Mandatory=$false)][Switch]$SetSettings,
    [Parameter(Mandatory=$false)][ValidateSet('BuiltInSafeDefaults','LastKnownGood','Factory','UserConf1','UserConf2')]$SetDefaults,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$SetBootOrder,
    [Parameter(Mandatory=$false)][ValidateSet('UEFI','Legacy')]$BootMode,
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$AdminPassword,
    [ValidateScript({
        if($_ -notmatch "(\.csv)")
        {
            throw "The specified file must be a .csv file"
        }
        return $true 
    })]
    [System.IO.FileInfo]$CsvPath,
    [Parameter(Mandatory=$false)][ValidateScript({
        if($_ -notmatch "(\.log)")
        {
            throw "The file specified in the LogFile paramter must be a .log file"
        }
        return $true
    })]
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\Dell\Manage-DellBiosSettings-WMI.log"
)

#List of settings to be configured ============================================================================================
#==============================================================================================================================
$Settings = (
    "NumLock,Enabled",
    "WakeOnLan,LanOnly",
    "Virtualization,Enabled",
    "VtForDirectIo,Enabled"
)
#==============================================================================================================================
#==============================================================================================================================

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

Function Set-DellBiosSetting
{
    #Set a specific Dell BIOS setting

    param(
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Name,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Value,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Password,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String[]]$NewBootOrder,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$BootMode,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][String]$Defaults
    )
    #Set the boot order
    if($BootOrder)
    {
        [String]$CurrentValue = $BootOrder | Where-Object BootListType -eq $BootMode | Select-Object -ExpandProperty BootOrder

        if($CurrentValue -eq $NewBootOrder)
        {
                Write-LogEntry -Value "The ""$BootMode"" boot order is already set to ""$NewBootOrder""" -Severity 1
                $Script:AlreadySet++
        }
        else
        {
            if(!([String]::IsNullOrEmpty($Password)))
            {
                $SettingResult = ($BootOrderInterface.Set(1,$Bytes.Length,$Bytes,$BootMode,$NewBootOrder.Count,$NewBootOrder)).Status
            }
            else
            {
                $SettingResult = ($BootOrderInterface.Set(0,0,0,$BootMode,$NewBootOrder.Count,$NewBootOrder)).Status
            }
            if($SettingResult -eq 0)
            {
                Write-LogEntry -Value "Successfully set the ""$BootMode"" boot order to ""$NewBootOrder""" -Severity 1
                $Script:SuccessSet++
            }
            else
            {
                Write-LogEntry -Value "Failed to set the ""$BootMode"" boot order to ""$NewBootOrder"". Return code: $SettingResult" -Severity 3
                $Script:FailSet++
            }
        }
    }
    #Load default BIOS settings
    elseif($Defaults)
    {
        if(!([String]::IsNullOrEmpty($Password)))
        {
            $SettingResult = ($AttributeInterface.SetBIOSDefaults(1,$Bytes.Length,$Bytes,$Defaults)).Status
        }
        else
        {
            $SettingResult = ($AttributeInterface.SetBIOSDefaults(0,0,0,$Defaults)).Status
        }
        if($SettingResult -eq 0)
        {
            Write-LogEntry -Value "Successfully loaded default BIOS settings" -Severity 1
            $Script:DefaultSet = $True
        }
        else
        {
            Write-LogEntry -Value "Failed to load default BIOS settings. Return code: $SettingResult" -Severity 3
            $Script:DefaultSet = $False
        }
    }
    #Set all other settings
    else
    {
        #Ensure the specified setting exists and get the possible values
        $CurrentValue = $SettingList | Where-Object AttributeName -eq $Name | Select-Object -ExpandProperty CurrentValue
        if($NULL -ne $CurrentValue)
        {
            #Setting is already set to specified value
            if($CurrentValue -eq $Value)
            {
                Write-LogEntry -Value "Setting ""$Name"" is already set to ""$Value""" -Severity 1
                $Script:AlreadySet++
            }
            #Setting is not set to specified value
            else
            {
                if(!([String]::IsNullOrEmpty($Password)))
                {
                    $SettingResult = ($AttributeInterface.SetAttribute(1,$Bytes.Length,$Bytes,$Name,$Value)).Status
                }
                else
                {
                    $SettingResult = ($AttributeInterface.SetAttribute(0,0,0,$Name,$Value)).Status
                }
            
                if($SettingResult -eq 0)
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
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Manage-DellBiosSettings-WMI"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
Write-LogEntry -Value "START - Dell BIOS settings management script" -Severity 1

#Parameter validation
Write-LogEntry -Value "Begin parameter validation" -Severity 1
if($GetSettings -and ($SetSettings -or $SetBootOrder -or $SetDefaults))
{
    Stop-Script -ErrorMessage "Cannot specify the GetSettings and SetSettings or SetBootOrder or SetDefaults parameters at the same time"
}
if(!($GetSettings -or $SetSettings -or $SetDefaults -or $SetBootOrder))
{
    Stop-Script -ErrorMessage "One of the GetSettings or SetSettings or SetDefaults or SetBootOrder parameters must be specified when running this script"
}
if($SetSettings -and !($Settings -or $CsvPath))
{
    Stop-Script -ErrorMessage "Settings must be specified using either the Settings variable in the script or the CsvPath parameter"
}
if($SetBootOrder -and !($BootMode))
{
    Stop-Script -ErrorMessage "When using the SetBootOrder parameter, the BootMode parameter must also be specified"
}
if($SetSettings -and $SetDefaults)
{
	$ErrorMsg = "Both the SetSettings and SetDefaults parameters have been used. The SetDefaults parameter will override any other settings"
    Write-LogEntry -Value $ErrorMsg -Severity 2
}
if($SetBootOrder -and $SetDefaults)
{
	$ErrorMsg = "Both the SetBootOrder and SetDefaults parameters have been used. The SetDefaults parameter will override any other settings"
    Write-LogEntry -Value $ErrorMsg -Severity 2
}
if(($SetBootOrder -or $SetDefaults) -and $CsvPath -and !($SetSettings))
{
	$ErrorMsg = "The CsvPath parameter has been specified without the SetSettings paramter. The CSV file will be ignored"
    Write-LogEntry -Value $ErrorMsg -Severity 2
}
Write-LogEntry -Value "Parameter validation completed" -Severity 1

#Connect to the BIOSAttributeInterface WMI class
$AttributeInterface = Get-WmiData -Namespace root\dcim\sysman\biosattributes -ClassName BIOSAttributeInterface -CmdletType WMI

#Connect to the SecurityInterface WMI class
$SecurityInterface = Get-WmiData -Namespace root\dcim\sysman\wmisecurity -ClassName SecurityInterface -CmdletType WMI

#Connect to the EnumerationAttribute WMI class
$Enumeration = Get-WmiData -Namespace root\dcim\sysman\biosattributes -ClassName EnumerationAttribute -CmdletType CIM -Select "AttributeName","CurrentValue","PossibleValue"

#Connect to the IntegerAttribute WMI class
$Integer = Get-WmiData -Namespace root\dcim\sysman\biosattributes -ClassName IntegerAttribute -CmdletType CIM -Select "AttributeName","CurrentValue","PossibleValue"

#Connect to the StringAttribute WMI class
$String = Get-WmiData -Namespace root\dcim\sysman\biosattributes -ClassName StringAttribute -CmdletType CIM -Select "AttributeName","CurrentValue","PossibleValue"

#Connect to the BootOrder WMI class
if($SetBootOrder -or $GetSettings)
{
    $BootOrder = Get-WmiData -Namespace root\dcim\sysman\biosattributes -ClassName BootOrder -CmdletType CIM
    $BootOrderInterface = Get-WmiData -Namespace root\dcim\sysman\biosattributes -ClassName SetBootOrder -CmdletType WMI
}

#Combine the setting lists into a single object
$SettingList = $Enumeration + $Integer + $String | Sort-Object AttributeName

#Format the password if set
if($AdminPassword)
{
    $Encoder = New-Object System.Text.UTF8Encoding
    $Bytes = $Encoder.GetBytes($AdminPassword)
}

#Convert the SetDefaults parameter to correct value if set
<#
    Default settings value mappings
        0 - BuiltInSafeDefaults
        1 - LastKnownGood
        2 - Factory
        3 - UserConf1
        4 - UserConf2
#>
if($SetDefaults)
{
    switch($SetDefaults)
    {
        BuiltInSafeDefaults {$DefaultValue = 0}
        LastKnownGood {$DefaultValue = 1}
        Factory {$DefaultValue = 2}
        UserConf1 {$DefaultValue = 3}
        UserConf2 {$DefaultValue = 4}
        default
        {
            Stop-Script -ErrorMessage "Failed to match the SetDefaults parameter ($SetDefaults) to a value. Use one of the 5 supported values with the SetDefaults parameter (BuiltInSafeDefaults, LastKnownGood, Factory, UserConf1, UserConf2"
        }
    }
}

#Set counters to 0
if($SetSettings -or $SetBootOrder)
{
    $AlreadySet = 0
    $SuccessSet = 0
    $FailSet = 0
    $NotFound = 0
    $DefaultSet = $Null
}

#Get the current password status
if($SetSettings -or $SetDefaults -or $SetBootOrder)
{
    Write-LogEntry -Value "Get the current password state" -Severity 1
    $AdminPasswordCheck = Get-CimInstance -Namespace root/dcim/sysman/wmisecurity -ClassName PasswordObject | Where-Object NameId -EQ "Admin" | Select-Object -ExpandProperty IsPasswordSet
    if($AdminPasswordCheck -eq 1)
    {
        Write-LogEntry -Value "The admin password is currently set" -Severity 1
        #Admin password set but parameter not specified
        if([String]::IsNullOrEmpty($AdminPassword))
        {
            Stop-Script -ErrorMessage "The admin password is set, but no password was supplied. Use the AdminPassword parameter when a password is set"
        }
        #Admin password set correctly
        if(($SecurityInterface.SetNewPassword(1,$Bytes.Length,$Bytes,"Admin",$AdminPassword,$AdminPassword)).Status -eq 0)
	    {
		    Write-LogEntry -Value "The specified admin password matches the currently set password" -Severity 1
        }
        #Supervisor password not set correctly
        else
        {
            Stop-Script -ErrorMessage "The specified admin password does not match the currently set password"
        }
    }
    else
    {
        Write-LogEntry -Value "The admin password is not currently set" -Severity 1
    }
}

#Get the current settings
if($GetSettings)
{
    Write-LogEntry -Value "Getting a list of current BIOS settings" -Severity 1
    #Write the current boot order to the log file
    $BootListObject = $BootOrder | Where-Object IsActive -eq  1 | Select-Object BootListType,BootOrder
    Write-LogEntry -Value "The current boot order is: $($BootListObject.BootOrder)" -Severity 1
    #Get all other settings
    $SettingObject = foreach($Setting in $SettingList){
        $PossibleValue = [String]$Setting.PossibleValue
        [PSCustomObject]@{
            Name = $Setting.AttributeName
            CurrentValue = $Setting.CurrentValue
            PossibleValue = $PossibleValue
        }
    }
    if($CsvPath)
    {
        Write-LogEntry -Value "Exporting settings to $CsvPath" -Severity 1
        $SettingObject | Export-Csv -Path $CsvPath -NoTypeInformation
        (Get-Content $CsvPath) | ForEach-Object {$_ -Replace '"',""} | Out-File $CsvPath -Force -Encoding ascii
    }
    else
    {
        Write-Output $SettingObject
    }
}

if($SetSettings -or $SetDefaults -or $SetBootOrder)
{    
    #Import settings from CSV
    if($CsvPath)
    {
        Clear-Variable Settings -ErrorAction SilentlyContinue
        $Settings = Import-Csv -Path $CsvPath
    }
    #Set Dell BIOS settings - password is set
    if($AdminPasswordCheck -eq 1)
    {
        #Set the boot order
        if($SetBootOrder)
        {
            Set-DellBiosSetting -NewBootOrder $SetBootOrder -BootMode $BootMode -Password $AdminPassword
        }
        #Set all other settings
        if($SetSettings)
        {
            if($CsvPath)
            {
                foreach($Setting in $Settings){
                    Set-DellBiosSetting -Name $Setting.Name -Value $Setting.Value -Password $AdminPassword
                }
            }
            else
            {
                foreach($Setting in $Settings){
                    $Data = $Setting.Split(',')
                    Set-DellBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim() -Password $AdminPassword
                }
            }
        }
        #Set defaults
        if($SetDefaults)
        {
            Set-DellBiosSetting -Defaults $DefaultValue -Password $AdminPassword
        }   
    }
    #Set Dell BIOS settings - password is not set
    else
    {
        #Set the boot order
        if($SetBootOrder)
        {
            Set-DellBiosSetting -NewBootOrder $SetBootOrder -BootMode $BootMode
        }
        #Set all other settings
        if($SetSettings)
        {
            if($CsvPath)
            {
                foreach($Setting in $Settings){
                    Set-DellBiosSetting -Name $Setting.Name -Value $Setting.Value
                }   
            }
            else
            {
                foreach($Setting in $Settings){
                    $Data = $Setting.Split(',')
                    Set-DellBiosSetting -Name $Data[0].Trim() -Value $Data[1].Trim()
                }   
            }
        }
        #Set defaults
        if($SetDefaults)
        {
            Set-DellBiosSetting -Defaults $DefaultValue
        }
    }
}

#Display results
if($SetSettings -or $SetBootOrder)
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
if($SetDefaults)
{
    if($DefaultSet -eq $True)
    {
        Write-Output "Successfully loaded ""$SetDefaults"" BIOS settings"    
    }
    else
    {
        Write-Output "Failed to load ""$SetDefaults"" BIOS settings"
    }
}
Write-Output "Dell BIOS settings Management completed. Check the log file for more information"
Write-LogEntry -Value "END - Dell BIOS settings management script" -Severity 1
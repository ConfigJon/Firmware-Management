<#
    .DESCRIPTION
        Import the Dell Command | PowerShell Provider module (Version 2.4.0)
    
    .PARAMETER ModulePath
        Specify the location of the DellBIOSProvider module source files. This parameter should be specified when the script is running in WinPE or when the system does not have internet access

    .PARAMETER DllPath
        Specify the location of the .dll files required to run the DellBIOSProvider module. This parameter should be specified when the script is running in WinPE or when the system does not have the required Visual C++ Redistributables installed

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        Running in a full Windows OS and installing from the internet    
            Install-DellBiosProvider.ps1

        Running in WinPE
            Install-DellBiosProvider.ps1 -ModulePath DellBIOSProvider -DllPath DllFiles

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/working-with-the-dell-command-powershell-provider/
        Modified: 2021-03-24

	.CHANGELOG
        2020-09-07 - Added a LogFile parameter. Changed the default log path in full Windows to $ENV:ProgramData\ConfigJonScripts\Dell.
                     Created a new function (Stop-Script) to consolidate some duplicate code and improve error reporting. Made a number of minor formatting and syntax changes
        2020-09-17 - Improved the log file path configuration
        2021-03-24 - Updated the required .dll files to support version 2.4.0 of the DellBIOSProvider module
        2022-02-20 - This script works with the DellBIOSProvider module version 2.4.0. This script will no longer be updated.
                     See the latest version of "Install-DellBiosProvider.ps1 on my GitHub for current version support"

#>

#Parameters ===================================================================================================================

param(
    [ValidateScript({
        if(!($_ | Test-Path))
        {
            throw "The ModulePath folder path does not exist"
        }
        if(!($_ | Test-Path -PathType Container))
        {
            throw "The ModulePath argument must be a folder path"
        }
        return $true 
    })]
    [Parameter(Mandatory=$false)][System.IO.DirectoryInfo]$ModulePath,
    [ValidateScript({
        if(!($_ | Test-Path))
        {
            throw "The DllPath folder path does not exist"
        }
        if(!($_ | Test-Path -PathType Container))
        {
            throw "The DllPath argument must be a folder path"
        }
        return $true 
    })]
    [Parameter(Mandatory=$false)][System.IO.DirectoryInfo]$DllPath,
    [Parameter(Mandatory=$false)][ValidateScript({
        if($_ -notmatch "(\.log)")
        {
            throw "The file specified in the LogFile paramter must be a .log file"
        }
        return $true
    })]
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\Dell\Install-DellBiosProvider.log"
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

Function Uninstall-DellBIOSProvider
{
    #Uninstall the DellBIOSProvider module

    Write-LogEntry -Value "Uninstall previous versions of the DellBIOSProvider module" -Severity 1
    $Module = Get-Package DellBIOSProvider -ErrorAction SilentlyContinue
    $ModuleFile = Test-Path "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider"
    if(($NULL -eq $Module) -and !($ModuleFile))
    {
        Write-LogEntry -Value "The DellBIOSProvider module is not currently installed" -Severity 1
    }
    else
    {
        if($NULL -ne $Module)
        {
            while($NULL -ne $Module)
            {
                $Version = Get-Package DellBIOSProvider -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
                Write-LogEntry -Value "Uninstalling DellBIOSProvider module version $Version" -Severity 1
                $Error.Clear()
                try
                {
                    $Module | Uninstall-Module -Force -ErrorAction Stop | Out-Null
                }
                catch
                {
                    Write-LogEntry -Value "Failed to uninstall DellBIOSProvider module version $Version" -Severity 3
                }
                if(!($Error))
                {
                    Write-LogEntry -Value "Successfully uninstalled DellBIOSProvider module version $Version" -Severity 1
                }
                Clear-Variable Module,Version
                $Module = Get-Package DellBIOSProvider -ErrorAction SilentlyContinue
            }
        }
        else
        {
            Remove-Item "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider" -Recurse -Force
            Write-LogEntry -Value "Successfully uninstalled the existing DellBIOSProvider module" -Severity 1
        }   
    }
}

Function Install-DellBIOSProviderLocal
{
    #Install the DellBIOSProvider from local source files

    param(
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$InstallPath,
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$ModulePath,
        [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Version
    )
    Write-LogEntry -Value "Install the DellBIOSProvider module from $ModulePath" -Severity 1
    $Error.Clear()
    try
    {
        Copy-Item $ModulePath -Destination "$InstallPath\WindowsPowerShell\Modules" -Recurse -Force -ErrorAction Stop | Out-Null
    }
    catch
    {
        Write-LogEntry -Value "Failed to copy the DellBIOSProvider module from $ModulePath to $InstallPath\WindowsPowerShell\Modules" -Severity 3
    }
    if(!($Error))
    {
        if($NULL -ne $Version)
        {
            Write-LogEntry -Value "Successfully installed DellBIOSProvider module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the DellBIOSProvider module" -Severity 1
        }
    }
}

Function Install-DellBIOSProviderRemote
{
    #Install the DellBIOSProvider from the PowerShell Gallery

    param(
        [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Version
    )
    Write-LogEntry -Value "Install the DellBIOSProvider module from the PowerShell Gallery" -Severity 1
    $Error.Clear()
    try
    {
        Install-Module -Name DellBIOSProvider -Force -ErrorAction Stop | Out-Null
    }
    catch
    {
        Stop-Script -ErrorMessage "Unable to install the DellBIOSProvider module from the PowerShell Gallery" -Exception $PSItem.Exception.Message
    }
    if(!($Error))
    {
        if($NULL -ne $Version)
        {
            Write-LogEntry -Value "Successfully installed DellBIOSProvider module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the DellBIOSProvider module" -Severity 1
        }
    }
}

Function Update-NuGet
{
    #Update the NuGet package provider

    #Check if the NuGet package provider is installed
    $Nuget = Get-PackageProvider | Where-Object Name -eq "NuGet"
    #If NuGet is installed, ensure it is the current version
    if($Nuget)
    {
        $Major = $Nuget.Version | Select-Object -ExpandProperty Major
        $Minor = $Nuget.Version | Select-Object -ExpandProperty Minor
        $Build = $Nuget.Version | Select-Object -ExpandProperty Build
        $Revision = $Nuget.Version | Select-Object -ExpandProperty Revision
        $NugetLocalVersion = "$($Major)." + "$($Minor)." + "$($Build)." + "$($Revision)"
        $NugetWebVersion = Find-PackageProvider NuGet | Select-Object -ExpandProperty Version
        if($NugetLocalVersion -ge $NugetWebVersion)
        {
            Write-LogEntry -Value "The latest version of the NuGet package provider ($NugetLocalVersion) is already installed" -Severity 1
        }
        #If the currently installed version of NuGet is outdated, update it from the internet
        else
        {
            Write-LogEntry -Value "Updating the NuGet package provider" -Severity 1
            $Error.Clear()
            try
            {
                Install-PackageProvider -Name "NuGet" -Force -Confirm:$False -ErrorAction Stop | Out-Null
            }
            catch
            {
                Write-LogEntry -Value "Unable to update the NuGet package provider" -Severity 3
            }
            if(!($Error))
            {
                Write-LogEntry -Value "Successfully updated the NuGet package provider to version $NugetWebVersion" -Severity 1
            }
        }
    }
    #If NuGet is not installed, install it from the internet
    else
    {
        Write-LogEntry -Value "Update the NuGet package provider" -Severity 1
        $Error.Clear()
        try
        {
            Install-PackageProvider -Name "NuGet" -Force -Confirm:$False -ErrorAction Stop | Out-Null
        }
        catch
        {
            Write-LogEntry -Value "Unable to update the NuGet package provider" -Severity 3
        }
        if(!($Error))
        {
            Write-LogEntry -Value "Successfully updated the NuGet package provider" -Severity 1
        }
    }
}

Function Copy-Dll
{
    #Copy .dll files

    param(
        [ValidateScript({
            if($_ -notmatch "(\.dll)")
            {
                throw "The specified file must be a .dll file"
            }
            return $true 
        })]
        [Parameter(Mandatory=$true)][System.IO.FileInfo]$DllFile,
        [ValidateScript({
            if(!($_ | Test-Path))
            {
                throw "The DllTargetPath folder path does not exist"
            }
            if(!($_ | Test-Path -PathType Container))
            {
                throw "The DllTargetPath argument must be a folder path"
            }
            return $true 
        })]
        [Parameter(Mandatory=$true)][System.IO.DirectoryInfo]$DllTargetPath,
        [ValidateScript({
            if(!($_ | Test-Path))
            {
                throw "The DllSourcePath folder path does not exist"
            }
            if(!($_ | Test-Path -PathType Container))
            {
                throw "The DllSourcePath argument must be a folder path"
            }
            return $true 
        })]
        [Parameter(Mandatory=$false)][System.IO.DirectoryInfo]$DllSourcePath
    )
    if(!(Test-Path "$DllTargetPath\$DllFile"))
    {
        Write-LogEntry -Value "Could not find $DllTargetPath\$DllFile" -Severity 2
        Write-LogEntry -Value "Copying $DllFile to $DllTargetPath" -Severity 1
        if($DllSourcePath)
        {
            $Error.Clear()
            try
            {
                Copy-Item -Path "$DllSourcePath\$DllFile" -Destination "$DllTargetPath\$DllFile" -Force -ErrorAction Stop
            }
            catch
            {
                $Script:DllFailure = $true
                Write-LogEntry -Value "Failed to copy $DllFile to $DllTargetPath" -Severity 2
            }
            if(!($Error))
            {
                Write-LogEntry -Value "Successfully copied $DllFile to $DllTargetPath" -Severity 1
            }
        }
        else
        {
            $Error.Clear()
            try
            {
                Copy-Item -Path "$DllFile" -Destination "$DllTargetPath\$DllFile" -Force -ErrorAction Stop
            }
            catch
            {
                $Script:DllFailure = $true
                Write-LogEntry -Value "Failed to copy $DllFile to $DllTargetPath" -Severity 2
            }
            if(!($Error))
            {
                Write-LogEntry -Value "Successfully copied $DllFile to $DllTargetPath" -Severity 1
            }
        }
    }
    else
    {
        Write-LogEntry -Value "Found $DllFile" -Severity 1    
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
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Install-DellBiosProvider"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
Write-LogEntry -Value "START - Dell BIOS provider installation script" -Severity 1

#Check the PowerShell version
Write-LogEntry -Value "Checking the installed PowerShell version" -Severity 1
$PsVer = $PSVersionTable.PSVersion | Select-Object -ExpandProperty Major

if($PsVer -ge 3)
{
    Write-LogEntry -Value "The current PowerShell version is $PsVer" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The current PowerShell version is $PsVer. The mininum supported PowerShell version is 3"
}

#Check the SMBIOS version
Write-LogEntry -Value "Checking the SMBIOS version of the system" -Severity 1
$BiosVerMajor = Get-WmiObject -Class Win32_Bios | Select-Object -ExpandProperty SMBIOSMajorVersion
$BiosVerMinor = Get-WmiObject -Class Win32_Bios | Select-Object -ExpandProperty SMBIOSMinorVersion
$BiosVerFull = "$($BiosVerMajor)." + "$($BiosVerMinor)"

if($BiosVerFull -ge 2.4)
{
    Write-LogEntry -Value "The current SMBIOS version is $BiosVerFull" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The current SMBIOS version is $BiosVerFull. The mininum supported SMBIOS version is 2.4"
}

#Verify the required .dll files exist
Write-LogEntry -Value "Verify Visual C++ DLL files exist" -Severity 1

if($DllPath)
{
    Copy-Dll -DllSourcePath "$DllPath" -DllFile "msvcp140.dll" -DllTargetPath "$ENV:windir\System32"
    Copy-Dll -DllSourcePath "$DllPath" -DllFile "vcruntime140.dll" -DllTargetPath "$ENV:windir\System32"
    Copy-Dll -DllSourcePath "$DllPath" -DllFile "vcruntime140_1.dll" -DllTargetPath "$ENV:windir\System32"
}
else
{
    Copy-Dll -DllFile "msvcp140.dll" -DllTargetPath "$ENV:windir\System32"
    Copy-Dll -DllFile "vcruntime140.dll" -DllTargetPath "$ENV:windir\System32"
    Copy-Dll -DllFile "vcruntime140_1.dll" -DllTargetPath "$ENV:windir\System32"
}
if($DllFailure)
{
    Stop-Script -ErrorMessage "One or more .dll files were not found or failed to copy"
}

#Check if 32 or 64 bit architecture
if([System.Environment]::Is64BitOperatingSystem)
{
    $ModuleInstallPath = $env:ProgramFiles
}
else
{
    $ModuleInstallPath = ${env:ProgramFiles(x86)}    
}

#Get the version of the currently installed DellBIOSProvider module
Write-LogEntry -Value "Checking the version of the currently installed DellBIOSProvider module" -Severity 1
try
{
    $LocalVersion = Get-Package DellBIOSProvider -ErrorAction Stop | Select-Object -ExpandProperty Version -ErrorAction Stop
}
catch
{
    $Local = $true
    if(Test-Path "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider")
    {
        $LocalVersion = Get-Content "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider\DellBIOSProvider.psd1" | Select-String "ModuleVersion ="
        $LocalVersion = (([regex]".*'(.*)'").Matches($LocalVersion))[0].Groups[1].Value
        if($NULL -ne $LocalVersion)
        {
            Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "DellBIOSProvider module not found on the local machine" -Severity 2
        }
    }
    else
    {
        Write-LogEntry -Value "DellBIOSProvider module not found on the local machine" -Severity 2
    }
}
if(($NULL -ne $LocalVersion) -and (!($Local)))
{
    Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
}

#Attempt to install the DellBIOSProvider module from local source files
if($ModulePath)
{
    #Get the version of the DellBIOSProvider in the ModulePath
    if(Test-Path "$ModulePath\DellBIOSProvider.psd1")
    {
        try
        {
            Write-LogEntry -Value "Checking the version of the DellBIOSProvider module in $ModulePath" -Severity 1
            $SourceVersion = Get-Content "$ModulePath\DellBIOSProvider.psd1" -ErrorAction Stop | Select-String "ModuleVersion =" -ErrorAction Stop
            $SourceVersion = (([regex]".*'(.*)'").Matches($SourceVersion))[0].Groups[1].Value
            if($NULL -ne $SourceVersion)
            {
                Write-LogEntry -Value "The version of the DellBIOSProvider module in $ModulePath is $SourceVersion" -Severity 1
            }
        }
        catch
        {
            Write-LogEntry -Value "Failed to check the version of the DellBIOSProvider module in $ModulePath" -Severity 3
        }
    }
    else
    {
        Write-LogEntry -Value "DellBIOSProvider.psd1 not found in $ModulePath" -Severity 3
    }
    if(($NULL -ne $SourceVersion) -and ($NULL -ne $LocalVersion))
    {
        if($SourceVersion -eq $LocalVersion)
        {
            Write-LogEntry -Value "The latest version of the DellBIOSProvider module is already installed" -Severity 1
        }
        else
        {
            Uninstall-DellBIOSProvider
            Install-DellBIOSProviderLocal -InstallPath $ModuleInstallPath -ModulePath $ModulePath -Version $SourceVersion
        }
    }
    elseif(($NULL -ne $SourceVersion) -and ($NULL -eq $LocalVersion))
    {
        Install-DellBIOSProviderLocal -InstallPath $ModuleInstallPath -ModulePath $ModulePath -Version $SourceVersion
    }
    else
    {
        Uninstall-DellBIOSProvider
        Install-DellBIOSProviderLocal -InstallPath $ModuleInstallPath -ModulePath $ModulePath
    }
}

#Attempt to install the DellBIOSProvider module from the PowerShell Gallery
else
{
    #Ensure the NuGet package provider is installed an updated
    Write-LogEntry -Value "Checking the version of the NuGet package provider" -Severity 1
    Update-NuGet
    #Get the version of the DellBIOSProvider module in the PowerShell Gallery
    Write-LogEntry -Value "Checking the version of the DellBIOSProvider module in the PowerShell Gallery" -Severity 1
    try
    {
        $WebVersion = Find-Package DellBIOSProvider -ErrorAction Stop | Select-Object -ExpandProperty Version -ErrorAction Stop
        if($NULL -ne $WebVersion)
        {
            Write-LogEntry -Value "The version of the DellBIOSProvider module in the PowerShell Gallery is $WebVersion" -Severity 1
        }
    }
    catch
    {
        Write-LogEntry -Value "Failed to check the version of the DellBIOSProvider module in the PowerShell Gallery" -Severity 3
    }
    if(($NULL -ne $WebVersion) -and ($NULL -ne $LocalVersion))
    {
        if($WebVersion -eq $LocalVersion)
        {
            Write-LogEntry -Value "The latest version of the DellBIOSProvider module is already installed" -Severity 1
        }
        else
        {
            Uninstall-DellBIOSProvider
            Install-DellBIOSProviderRemote -Version $WebVersion
        }
    }
    elseif(($NULL -ne $WebVersion) -and ($NULL -eq $LocalVersion))
    {
        Install-DellBIOSProviderRemote -Version $WebVersion
    }
    else
    {
        Uninstall-DellBIOSProvider
        Install-DellBIOSProviderRemote
    }
}

#Import the DellBIOSProvider module
if($DllFailure)
{
    Write-LogEntry -Value "Unable to import the DellBIOSProvider module. One or more DLL files are missing" -Severity 3
}
else
{
    Write-LogEntry -Value "Import the DellBIOSProvider module" -Severity 1
    $Error.Clear()
    try
    {
        Import-Module DellBIOSProvider -Force -ErrorAction Stop
    }
    catch 
    {
        Stop-Script -ErrorMessage "Failed to import the DellBIOSProvider module" -Exception $PSItem.Exception.Message
    }
    if(!($Error))
    {
        Write-LogEntry -Value "Successfully imported the DellBIOSProvider module" -Severity 1
    }
}
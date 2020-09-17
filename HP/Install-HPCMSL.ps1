<#
    .DESCRIPTION
        Install the HP Client Management Script Library PowerShell modules
    
    .PARAMETER ModulePath
        Specify the location of the HPCMSL modules source files. This parameter should be specified when the script is running in WinPE or when the system does not have internet access

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        Running in a full Windows OS and installing from the internet
            Install-HPCMSL.ps1

        Running in WinPE or offline
            Install-HPCMSL.ps1 -ModulePath HPCMSL

    .NOTES
        Created by: Jon Anderson (@ConfigJon)
        Reference: https://www.configjon.com/installing-the-hp-client-management-script-library\
        Modified: 2020-09-17

    .CHANGELOG
        2020-09-14 - Added a LogFile parameter. Changed the default log path in full Windows to $ENV:ProgramData\ConfigJonScripts\HP.
                     Created a new function (Stop-Script) to consolidate some duplicate code and improve error reporting. Made a number of minor formatting and syntax changes
        2020-09-17 - Improved the log file path configuration
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
    [Parameter(DontShow)][Switch]$Rerun,
    [Parameter(Mandatory=$false)][ValidateScript({
        if($_ -notmatch "(\.log)")
        {
            throw "The file specified in the LogFile paramter must be a .log file"
        }
        return $true
    })]
    [System.IO.FileInfo]$LogFile = "$ENV:ProgramData\ConfigJonScripts\HP\Install-HPCMSL.log"
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

Function Install-HPCMSLLocal
{
    #Install the HPCMSL from local source files

    param(
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$InstallPath,
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$ModulePath,
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$ModuleName,
        [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Version
    )
    Write-LogEntry -Value "Install the $ModuleName module from $ModulePath" -Severity 1
    $Error.Clear()
    try
    {
        Copy-Item $ModulePath -Destination "$InstallPath\WindowsPowerShell\Modules" -Recurse -Force | Out-Null
    }
    catch
    {
        Write-LogEntry -Value "Failed to copy the $ModuleName module from $ModulePath to $InstallPath\WindowsPowerShell\Modules" -Severity 3
    }
    if(!($Error))
    {
        if($NULL -ne $Version)
        {
            Write-LogEntry -Value "Successfully installed $ModuleName module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the $ModuleName module" -Severity 1
        }
    }
}

Function Install-HPCMSLRemote
{
    #Install the HPCMSL from the PowerShell Gallery

    param(
        [parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][String]$Version
    )
    Write-LogEntry -Value "Install the HPCMSL module from the PowerShell Gallery" -Severity 1
    $Error.Clear()
    try
    {
        Install-Module -Name HPCMSL -Force -AcceptLicense
    }
    catch
    {
        Stop-Script -ErrorMessage "Unable to install the HPCMSL module from the PowerShell Gallery"
    }
    if(!($Error))
    {
        if($NULL -ne $Version)
        {
            Write-LogEntry -Value "Successfully installed HPCMSL module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the HPCMSL module" -Severity 1
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
                Install-PackageProvider -Name "NuGet" -Force -Confirm:$False | Out-Null
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
            Install-PackageProvider -Name "NuGet" -Force -Confirm:$False | Out-Null
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

Function Update-PowerShellGet
{
    #Update the PowerShellGet module

    Import-Module -Name PowerShellGet -Force
    $PsGetVersion = Get-Module PowerShellGet | Select-Object -ExpandProperty Version
    $PsGetVersionFull = "$($PsGetVersion.Major)." + "$($PsGetVersion.Minor)." + "$($PsGetVersion.Build)"
    $PsGetVersionWeb = Find-Package -Name PowerShellGet | Select-Object -ExpandProperty Version
    if($PsGetVersionFull -ge $PsGetVersionWeb)
    {
        Write-LogEntry -Value "The latest version of the PowerShellGet module ($PsGetVersionFull) is already installed" -Severity 1
    }
    #If the currently installed version of the PowerShellGet module is outdated, update it from the internet
    else
    {
        Write-LogEntry -Value "Updating the PowerShellGet module" -Severity 1
        $Error.Clear()
        try
        {
            Remove-Module -Name PowerShellGet -Force #Unload the current version of the PowerShellGet module
            Install-Module -Name PowerShellGet -Force #Install the latest version of the PowerShellGet module
        }
        catch
        {
            Write-LogEntry -Value "Unable to update the PowerShellGet module" -Severity 3
        }
        if(!($Error))
        {
            Write-LogEntry -Value "Successfully updated the PowerShellGet module to version $PsGetVersionWeb" -Severity 1
            #Re-launch the script in a new session to detect the new PowerShellGet module version
            Start-Process -FilePath "$Env:WinDir\system32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File $ScriptPath -Rerun" -Wait -PassThru
            exit
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
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Install-HPCMSL"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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

#Set the names of the HPCMSL modules
$HPModules = ("HP.ClientManagement","HP.Firmware","HP.Private","HP.Repo","HP.Sinks","HP.Softpaq","HP.Utility","HPCMSL")

#Get the path to the script (Used if the script needs to be re-launched)
$ScriptPath = $MyInvocation.MyCommand

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

if(!($Rerun)){
    Write-Output "Log path set to $LogFile"
    Write-LogEntry -Value "START - HP Client Management Script Library installation script" -Severity 1

    #Make sure the folder names in the ModulePath match the HPCMSL folder names
    if($ModulePath)
    {
        Write-LogEntry -Value "Validate the folder names in $ModulePath" -Severity 1
        #Validate the top level folders
        $ModulePathFolders = Get-ChildItem $ModulePath -Directory | Select-Object -ExpandProperty Name
        ForEach($Folder in $ModulePathFolders){
            if($HPModules -notcontains $Folder)
            {
                Write-LogEntry -Value "$Folder is not a valid HPCMSL module folder name" -Severity 3
                $InvalidFolder = $True
            }       
        }
        if($InvalidFolder)
        {
            Stop-Script -ErrorMessage "Invalid folder names found in $ModulePath. Valid folder names are: ""HP.ClientManagement"" ""HP.Firmware"" ""HP.Private"" ""HP.Repo"" ""HP.Sinks"" ""HP.Softpaq"" ""HP.Utility"" ""HPCMSL"""
        }
        else
        {
            #Validate the subfolders
            ForEach($Folder in $ModulePathFolders){
                $Subfolder = Get-ChildItem "$ModulePath\$Folder" -Directory | Select-Object -ExpandProperty Name
                if($NULL -eq $Subfolder)
                {
                    Write-LogEntry -Value "No subfolders detected under $Folder. There should be 1 first-level subfolder." -Severity 3
                    $InvalidSubfolder = $True
                }
                elseif($Subfolder.Count -gt 1)
                {
                    Write-LogEntry -Value "Multiple first-level subfolders detected under $Folder. There should only be 1 first-level subfolder." -Severity 3
                    $InvalidSubfolder = $True
                }
                else
                {
                    $PatternCheck = (([regex]"^(\*|\d+(\.\d+){1,3}(\.\*)?)$").Matches($Subfolder)).Success
                    if($PatternCheck -ne "True")
                    {
                        Write-LogEntry -Value "$Folder\$Subfolder is not a valid subfolder" -Severity 3
                        $InvalidSubfolder = $True
                    }
                }
            }
            if($InvalidSubfolder)
            {
                Write-LogEntry -Value "Each module folder should contain a single first-level subfolder. The folder should be named the version of the module." -Severity 2
                Write-LogEntry -Value 'Valid first-level subfolder names should be in the format of "1.2" or "1.2.3" or "1.2.3.4"' -Severity 2
                throw "Invalid subfolder structure found in $ModulePath. See the log file for more details"
            }
            else
            {
                Write-LogEntry -Value "Successfully validated the folder names" -Severity 1
            }
        }
    }

    #Check the PowerShell version
    Write-LogEntry -Value "Checking the installed PowerShell version" -Severity 1
    $PsVerMajor = $PSVersionTable.PSVersion | Select-Object -ExpandProperty Major
    $PsVerMinor = $PSVersionTable.PSVersion | Select-Object -ExpandProperty Minor
    $PsVerFull = "$($PsVerMajor)." + "$($PsVerMinor)."
    if($PsVerFull -ge 5.1)
    {
        Write-LogEntry -Value "The current PowerShell version is $PsVerFull" -Severity 1
    }
    else
    {
        Stop-Script -ErrorMessage "The current PowerShell version is $PsVerFull. The mininum supported PowerShell version is 5.1"
    }

    #Set the PowerShell Module insatll path
    $ModuleInstallPath = $env:ProgramFiles

    #Get the versions of the currently installed HPCMSL modules
    Write-LogEntry -Value "Checking the versions of the currently installed HPCMSL modules" -Severity 1
    $LocalModuleVersions = [PSCustomObject]@{}
    ForEach($HPModule in $HPModules){
        if(Test-Path "$ModuleInstallPath\WindowsPowerShell\Modules\$HPModule")
        {
            $LocalVersionList = Get-ChildItem "$ModuleInstallPath\WindowsPowerShell\Modules\$HPModule" -Directory | Select-Object -ExpandProperty Name
            if($LocalVersionList.Count -gt 1)
            {
                $LocalVersion = "0.0"
                ForEach($Version in $LocalVersionList){
                    if(([Version]$Version).CompareTo([Version]$LocalVersion) -eq 1)
                    {
                        $LocalVersion = $Version
                    }
                }
            }
            else
            {
                $LocalVersion = $LocalVersionList
            }
             
            if($NULL -ne $LocalVersion)
            {
                Write-LogEntry -Value "The version of the currently installed $HPModule module is $LocalVersion" -Severity 1
                $LocalModuleVersions | Add-Member -NotePropertyName $HPModule -NotePropertyValue $LocalVersion
            }
            else
            {
                Write-LogEntry -Value "$HPModule module not found on the local machine" -Severity 2
                $LocalModuleVersions | Add-Member -NotePropertyName $HPModule -NotePropertyValue '0.0'
            }
        }
        else
        {
            Write-LogEntry -Value "$HPModule module not found on the local machine" -Severity 2
            $LocalModuleVersions | Add-Member -NotePropertyName $HPModule -NotePropertyValue '0.0'
        }
    }
}

#Attempt to install the HPCMSL from local source files
if($ModulePath)
{
    ForEach($HPModule in $HPModules){
        #Get the version of the module
        try
        {
            $SourceVersion = Get-ChildItem "$ModulePath\$HPModule" -Directory | Select-Object -ExpandProperty Name
            if($NULL -ne $SourceVersion)
            {
                Write-LogEntry -Value "The version of the $HPModule module in $ModulePath is $SourceVersion" -Severity 1
            }
        }
        catch
        {
            Write-LogEntry -Value "Failed to check the version of the $HPModule module in $ModulePath" -Severity 3
        }
        if($NULL -ne $SourceVersion)
        {
            $LocalVersionCompare = ([Version]$SourceVersion).CompareTo([Version]$LocalModuleVersions.$HPModule)
            if($LocalVersionCompare -eq 0)
            {
                Write-LogEntry -Value "The latest version of the $HPModule module is already installed" -Severity 1
            }
            elseif($LocalVersionCompare -eq -1)
            {
                Write-LogEntry -Value "A newer version of $HPModule is already installed" -Severity 1
            }
            else
            {
                Install-HPCMSLLocal -InstallPath $ModuleInstallPath -ModuleName $HPModule -ModulePath "$ModulePath\$HPModule" -Version $SourceVersion
            }
        }
        else
        {
            Install-HPCMSLLocal -InstallPath $ModuleInstallPath -ModuleName $HPModule -ModulePath "$ModulePath\$HPModule"
        }
    }
}
#Attempt to install the HPCMSL module from the PowerShell Gallery
else
{
    if(!($Rerun)){
        #Ensure the NuGet package provider is installed and updated
        Write-LogEntry -Value "Checking the version of the NuGet package provider" -Severity 1
        Update-NuGet

        #Ensure the PowerShellGet module is updated
        Write-LogEntry -Value "Checking the version of the PowerShellGet module" -Severity 1
        Update-PowerShellGet
    }
    #Get the version of the HPCMSL module in the PowerShell Gallery
    Write-LogEntry -Value "Checking the version of the HPCMSL module in the PowerShell Gallery" -Severity 1
    try
    {
        $WebVersion = Find-Package HPCMSL | Select-Object -ExpandProperty Version
        if($NULL -ne $WebVersion)
        {
            Write-LogEntry -Value "The version of the HPCMSL module in the PowerShell Gallery is $WebVersion" -Severity 1
        }
    }
    catch
    {
        Write-LogEntry -Value "Failed to check the version of the HPCMSL module in the PowerShell Gallery" -Severity 3
    }
    if($NULL -ne $WebVersion)
    {
        $WebVersionCompare = ([Version]$WebVersion).CompareTo([Version]$LocalModuleVersions.HPCMSL)
        if($WebVersionCompare -eq 0)
        {
            Write-LogEntry -Value "The latest version of the HPCMSL module is already installed" -Severity 1
        }
        elseif($WebVersionCompare -eq -1)
        {
            Write-LogEntry -Value "A newer version of the HPCMSL module is already installed" -Severity 1
        }
        else
        {
            Install-HPCMSLRemote -Version $WebVersion
        }
    }
    elseif(($NULL -ne $WebVersion) -and ($NULL -eq $LocalVersion))
    {
        Install-HPCMSLRemote -Version $WebVersion
    }
    else
    {
        Install-HPCMSLRemote
    }
}

#Import the HPCMSL module
Write-LogEntry -Value "Import the HPCMSL module" -Severity 1
$Error.Clear()
try
{
    Import-Module HPCMSL -Force -ErrorAction Stop
}
catch 
{
    Stop-Script -ErrorMessage "Failed to import the HPCMSL module" -Exception $PSItem.Exception.Message
}
if(!($Error))
{
    Write-LogEntry -Value "Successfully imported the HPCMSL module" -Severity 1
}

Write-LogEntry -Value "END - HP Client Management Script Library installation script" -Severity 1
<#
    .DESCRIPTION
        Install the HP Client Management Script Library PowerShell modules

    .PARAMETER ModulePath
        Specify the location of the HPCMSL modules source files. This parameter should be specified when the script is running in WinPE or when the system does not have internet access

    .PARAMETER AllEditions
        When specified, installs the HPCMSL modules for both Windows PowerShell 5.1 (Program Files\WindowsPowerShell\Modules) and PowerShell 7+ (Program Files\PowerShell\Modules).
        By default, modules are only installed for the PowerShell edition currently running the script.

    .PARAMETER Import
        When specified, imports the HPCMSL module into the current session after installation. Off by default.

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        Running in a full Windows OS and installing from the internet
            Install-HPCMSL.ps1

        Running in WinPE or offline
            Install-HPCMSL.ps1 -ModulePath HPCMSL

        Installing for both Windows PowerShell 5.1 and PowerShell 7
            Install-HPCMSL.ps1 -AllEditions

        Installing and importing the module into the current session
            Install-HPCMSL.ps1 -Import

    .NOTES
        Created by: Jon Anderson
        Reference: https://www.configjon.com/installing-the-hp-client-management-script-library
        Modified: 2026-05-18

    .CHANGELOG
        2020-09-14 - Added a LogFile parameter. Changed the default log path in full Windows to $env:ProgramData\ConfigJonScripts\HP.
                     Created a new function (Stop-Script) to consolidate some duplicate code and improve error reporting. Made a number of minor formatting and syntax changes
        2020-09-17 - Improved the log file path configuration
        2026-05-18 - Replaced the hard-coded HPCMSL module list with dynamic discovery
                     Added support for installing for both PowerShell 5.1 and PowerShell 7
                     -Added -AllEditions switch
                     -Made the install path edition-aware (PowerShell\Modules vs WindowsPowerShell\Modules)
                     Added TLS 1.2 enforcement before connecting to the PowerShell Gallery
                     Removed the automatic import of the module at the end of the script
                     -Added a discovery check using Get-Module -ListAvailable to verify successful install
                     -Module can still be imported via optional -Import swtich
                     Gated the NuGet/PowerShellGet bootstrap to PowerShell 5.1. Under PowerShell 7 NuGet and PowerShellGet are already current
                     Fixed string-based version comparison bugs in Update-NuGet and Update-PowerShellGet
                     The rerun launched by Update-PowerShellGet now uses the running host's executable instead of hard-coding powershell.exe
                     Normalized formatting and style throughout the script
                     Several smaller bug fixes and improvements
#>

#Parameters ===================================================================================================================

param(
    [ValidateScript({
        if (-not ($_ | Test-Path))
        {
            throw "The ModulePath folder path does not exist"
        }
        if (-not ($_ | Test-Path -PathType Container))
        {
            throw "The ModulePath argument must be a folder path"
        }
        return $true
    })]
    [Parameter(Mandatory = $false)][System.IO.DirectoryInfo]$ModulePath,
    [Parameter(Mandatory = $false)][switch]$AllEditions,
    [Parameter(Mandatory = $false)][switch]$Import,
    [Parameter(DontShow)][switch]$Rerun,
    [Parameter(Mandatory = $false)][ValidateScript({
        if ($_ -notmatch '\.log$')
        {
            throw "The file specified in the LogFile parameter must be a .log file"
        }
        return $true
    })]
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\ConfigJonScripts\HP\Install-HPCMSL.log"
)

#Functions ====================================================================================================================

function Get-TaskSequenceStatus
{
    #Determine if a task sequence is currently running
    try
    {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    }
    catch { }
    if ($null -eq $TSEnv)
    {
        return $false
    }
    else
    {
        try
        {
            $SMSTSType = $TSEnv.Value("_SMSTSType")
        }
        catch { }
        if ($null -eq $SMSTSType)
        {
            return $false
        }
        else
        {
            return $true
        }
    }
}

function Test-WinPE
{
    #Determine if the script is running in the Windows Preinstallation Environment (WinPE)

    return (Test-Path -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\MiniNT')
}

function Stop-Script
{
    #Write an error to the log file and terminate the script

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ErrorMessage,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$Exception
    )
    Write-LogEntry -Value $ErrorMessage -Severity 3
    if ($Exception)
    {
        Write-LogEntry -Value "Exception Message: $Exception" -Severity 3
    }
    throw $ErrorMessage
}

function Get-ModuleInstallPaths
{
    #Return the module install root path

    param(
        [Parameter(Mandatory = $false)][switch]$AllEditions
    )
    if ($AllEditions)
    {
        return @(
            (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'),
            (Join-Path $env:ProgramFiles 'PowerShell\Modules')
        )
    }
    $Subpath = if ($PSVersionTable.PSEdition -eq 'Core') { 'PowerShell\Modules' } else { 'WindowsPowerShell\Modules' }
    return @((Join-Path $env:ProgramFiles $Subpath))
}

function Find-PwshExe
{
    #Locate pwsh.exe across the common install methods. Returns the Source (full path) and InstallType (Path, Msi, or Msix)

    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')

    $Cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($Cmd)
    {
        return [PSCustomObject]@{ Source = $Cmd.Source; InstallType = 'Path' }
    }

    $MsiPath = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if (Test-Path -LiteralPath $MsiPath)
    {
        return [PSCustomObject]@{ Source = $MsiPath; InstallType = 'Msi' }
    }

    $WindowsAppsRoot = Join-Path $env:ProgramFiles 'WindowsApps'
    if (Test-Path -LiteralPath $WindowsAppsRoot)
    {
        $MsixCandidates = Get-ChildItem -LiteralPath $WindowsAppsRoot -Filter 'Microsoft.PowerShell_*' -Directory -ErrorAction SilentlyContinue
        foreach ($Pkg in $MsixCandidates)
        {
            $Candidate = Join-Path $Pkg.FullName 'pwsh.exe'
            if (Test-Path -LiteralPath $Candidate)
            {
                return [PSCustomObject]@{ Source = $Candidate; InstallType = 'Msix' }
            }
        }
    }

    return $null
}

function Get-HPCMSLOnlineModules
{
    #Discover the HPCMSL module from the PowerShell Gallery
    #Fall back to a static list (for inventory only) when the gallery is unreachable

    try
    {
        $Discovered = Find-Module -Name HPCMSL -IncludeDependencies -Repository PSGallery -ErrorAction Stop
        $Names = $Discovered | Select-Object -ExpandProperty Name | Sort-Object -Unique
        Write-LogEntry -Value "Discovered $($Names.Count) HPCMSL module(s) in the PowerShell Gallery" -Severity 1
        return $Names
    }
    catch
    {
        Write-LogEntry -Value "Failed to query the PowerShell Gallery for HPCMSL dependencies: $($_.Exception.Message)" -Severity 2
        Write-LogEntry -Value "Falling back to built-in module list for inventory purposes only" -Severity 2
        #Snapshot of the HPCMSL 1.8.6 module set (April 2026). Used for local inventory if discovery fails, never used to reject anything
        return @(
            'HP.ClientManagement', 'HP.Consent',  'HP.Displays',          'HP.Docks',
            'HP.Firmware',         'HP.Notifications', 'HP.Private',      'HP.Repo',
            'HP.Retail',           'HP.Security', 'HP.Sinks',             'HP.SmartExperiences',
            'HP.Softpaq',          'HP.Utility',  'HPCMSL'
        )
    }
}

function Get-HPCMSLOfflineModules
{
    #Discover modules from a local folder by enumerating $ModulePath. Folder name = module name

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ModulePath
    )
    $Valid = New-Object System.Collections.Generic.List[string]
    $Folders = Get-ChildItem -Path $ModulePath -Directory
    foreach ($Folder in $Folders)
    {
        $Subs = Get-ChildItem -Path $Folder.FullName -Directory
        if (-not $Subs)
        {
            Write-LogEntry -Value "No version subfolder found under $($Folder.Name) - skipping" -Severity 2
            continue
        }
        if ($Subs.Count -gt 1)
        {
            Write-LogEntry -Value "Multiple version subfolders detected under $($Folder.Name); expected exactly one - skipping" -Severity 2
            continue
        }
        $Sub = $Subs[0]
        $PatternCheck = (([regex]"^(\*|\d+(\.\d+){1,3}(\.\*)?)$").Matches($Sub.Name)).Success
        if (-not $PatternCheck)
        {
            Write-LogEntry -Value "$($Folder.Name)\$($Sub.Name) is not a valid version subfolder name - skipping" -Severity 2
            continue
        }
        $Psd1 = Join-Path $Sub.FullName "$($Folder.Name).psd1"
        if (-not (Test-Path $Psd1))
        {
            Write-LogEntry -Value "$($Folder.Name) is missing manifest $($Folder.Name).psd1 at $($Sub.FullName) - skipping" -Severity 2
            continue
        }
        $Valid.Add($Folder.Name) | Out-Null
    }
    if ($Valid.Count -eq 0)
    {
        Stop-Script -ErrorMessage "No valid HPCMSL module folders were discovered in $ModulePath"
    }
    Write-LogEntry -Value "Discovered $($Valid.Count) valid HPCMSL module folder(s) in $ModulePath" -Severity 1
    return $Valid.ToArray()
}

function Install-HPCMSLLocal
{
    #Install a single HPCMSL module from local source files

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$InstallPath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ModulePath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ModuleName,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$Version
    )
    Write-LogEntry -Value "Install the $ModuleName module from $ModulePath to $InstallPath" -Severity 1
    if (-not (Test-Path $InstallPath))
    {
        try
        {
            New-Item -Path $InstallPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch
        {
            Stop-Script -ErrorMessage "Failed to create the module destination directory $InstallPath" -Exception $_.Exception.Message
        }
    }
    try
    {
        Copy-Item -Path $ModulePath -Destination $InstallPath -Recurse -Force -ErrorAction Stop | Out-Null
        if ($Version)
        {
            Write-LogEntry -Value "Successfully installed $ModuleName module version $Version to $InstallPath" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the $ModuleName module to $InstallPath" -Severity 1
        }
    }
    catch
    {
        Stop-Script -ErrorMessage "Failed to copy the $ModuleName module from $ModulePath to $InstallPath" -Exception $_.Exception.Message
    }
}

function Install-HPCMSLRemote
{
    #Install the HPCMSL module from the PowerShell Gallery for AllUsers scope

    param(
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$Version
    )
    Write-LogEntry -Value "Install the HPCMSL module from the PowerShell Gallery" -Severity 1
    try
    {
        Install-Module -Name HPCMSL -Force -AcceptLicense -Scope AllUsers -AllowClobber -ErrorAction Stop
        if ($Version)
        {
            Write-LogEntry -Value "Successfully installed HPCMSL module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the HPCMSL module" -Severity 1
        }
    }
    catch
    {
        Stop-Script -ErrorMessage "Unable to install the HPCMSL module from the PowerShell Gallery" -Exception $_.Exception.Message
    }
}

function Update-NuGet
{
    #Update the NuGet package provider

    $Nuget = Get-PackageProvider -ListAvailable -ErrorAction SilentlyContinue | Where-Object Name -eq 'NuGet' | Select-Object -First 1
    if ($Nuget)
    {
        $NugetLocalVersion = [Version]$Nuget.Version.ToString()
        try
        {
            $NugetWebVersion = [Version]((Find-PackageProvider -Name NuGet -ErrorAction Stop | Select-Object -First 1 -ExpandProperty Version).ToString())
        }
        catch
        {
            Write-LogEntry -Value "Unable to query the latest NuGet package provider version: $($_.Exception.Message)" -Severity 2
            return
        }
        if ($NugetLocalVersion -ge $NugetWebVersion)
        {
            Write-LogEntry -Value "The latest version of the NuGet package provider ($NugetLocalVersion) is already installed" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Updating the NuGet package provider from $NugetLocalVersion to $NugetWebVersion" -Severity 1
            try
            {
                Install-PackageProvider -Name 'NuGet' -Force -Confirm:$false -ErrorAction Stop | Out-Null
                Write-LogEntry -Value "Successfully updated the NuGet package provider to version $NugetWebVersion" -Severity 1
            }
            catch
            {
                Write-LogEntry -Value "Unable to update the NuGet package provider: $($_.Exception.Message)" -Severity 3
            }
        }
    }
    else
    {
        Write-LogEntry -Value "Installing the NuGet package provider" -Severity 1
        try
        {
            Install-PackageProvider -Name 'NuGet' -Force -Confirm:$false -ErrorAction Stop | Out-Null
            Write-LogEntry -Value "Successfully installed the NuGet package provider" -Severity 1
        }
        catch
        {
            Write-LogEntry -Value "Unable to install the NuGet package provider: $($_.Exception.Message)" -Severity 3
        }
    }
}

function Update-PowerShellGet
{
    #Update the PowerShellGet module

    Import-Module -Name PowerShellGet -Force
    $PsGetLocalVersion = [Version](Get-Module PowerShellGet | Select-Object -First 1 -ExpandProperty Version).ToString()
    try
    {
        $PsGetWebVersion = [Version]((Find-Module -Name PowerShellGet -Repository PSGallery -ErrorAction Stop | Select-Object -First 1 -ExpandProperty Version).ToString())
    }
    catch
    {
        Write-LogEntry -Value "Unable to query the latest PowerShellGet version: $($_.Exception.Message)" -Severity 2
        return
    }
    if ($PsGetLocalVersion -ge $PsGetWebVersion)
    {
        Write-LogEntry -Value "The latest version of the PowerShellGet module ($PsGetLocalVersion) is already installed" -Severity 1
        return
    }
    Write-LogEntry -Value "Updating the PowerShellGet module from $PsGetLocalVersion to $PsGetWebVersion" -Severity 1
    try
    {
        Remove-Module -Name PowerShellGet -Force
        Install-Module -Name PowerShellGet -Force -Scope AllUsers -AllowClobber -ErrorAction Stop
        Write-LogEntry -Value "Successfully updated the PowerShellGet module to version $PsGetWebVersion" -Severity 1
    }
    catch
    {
        Write-LogEntry -Value "Unable to update the PowerShellGet module: $($_.Exception.Message)" -Severity 3
        return
    }
    #Re-launch the script in a new session under the same edition so the updated PowerShellGet is loaded
    $HostExe = (Get-Process -Id $PID).Path
    Write-LogEntry -Value "Re-launching the script under $HostExe to pick up the updated PowerShellGet module" -Severity 1
    $RelaunchArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath, '-Rerun')
    if ($AllEditions) { $RelaunchArgs += '-AllEditions' }
    if ($Import)      { $RelaunchArgs += '-Import' }
    if ($LogFile)     { $RelaunchArgs += @('-LogFile', $LogFile) }
    $RerunProc = Start-Process -FilePath $HostExe -ArgumentList $RelaunchArgs -Wait -PassThru
    exit $RerunProc.ExitCode
}

function Write-LogEntry
{
    #Write data to a CMTrace compatible log file. (Credit to MSEndpointMgr - https://www.msendpointmgr.com/)

    param(
        [Parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [Parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("1", "2", "3")]
        [string]$Severity,
        [Parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = ($script:LogFile | Split-Path -Leaf)
    )
    #Determine log file location
    $LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName
    #Construct time stamp for log entry
    if (-not (Test-Path -Path 'variable:global:TimezoneBias'))
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
    #Construct date for log entry
    $Date = (Get-Date -Format "MM-dd-yyyy")
    #Construct context for log entry
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    #Construct final log entry
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Install-HPCMSL"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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

#Get the path to the script (Used if the script needs to be re-launched)
$ScriptPath = $MyInvocation.MyCommand.Path

#Configure Logging and task sequence variables
if (Get-TaskSequenceStatus)
{
    $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    $LogsDirectory = $TSEnv.Value("_SMSTSLogPath")
}
else
{
    $LogsDirectory = ($LogFile | Split-Path)
    if ([string]::IsNullOrEmpty($LogsDirectory))
    {
        $LogsDirectory = $PSScriptRoot
    }
    else
    {
        if (-not (Test-Path -PathType Container $LogsDirectory))
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

if (-not $Rerun)
{
    Write-Output "Log path set to $LogFile"
    Write-LogEntry -Value "START - HP Client Management Script Library installation script" -Severity 1
    if ($AllEditions)
    {
        Write-LogEntry -Value "-AllEditions specified: modules will be installed for both Windows PowerShell 5.1 and PowerShell 7" -Severity 1
    }
}
else
{
    Write-LogEntry -Value "Script re-launched with -Rerun after PowerShellGet update" -Severity 1
}

#Check the PowerShell version
Write-LogEntry -Value "Checking the installed PowerShell version" -Severity 1
$PsVer = $PSVersionTable.PSVersion
if (($PsVer.Major -gt 5) -or (($PsVer.Major -eq 5) -and ($PsVer.Minor -ge 1)))
{
    Write-LogEntry -Value "The current PowerShell version is $PsVer ($($PSVersionTable.PSEdition))" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The current PowerShell version is $PsVer. The minimum supported PowerShell version is 5.1"
}

#Determine the module install path
$ModuleInstallPaths = @(Get-ModuleInstallPaths -AllEditions:$AllEditions)
foreach ($Path in $ModuleInstallPaths)
{
    Write-LogEntry -Value "Module install path: $Path" -Severity 1
}

#Check for NuGet and PowerShellGet updates
if (-not $ModulePath)
{
    try
    {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Write-LogEntry -Value "Enforced TLS 1.2 for outbound connections" -Severity 1
    }
    catch
    {
        Write-LogEntry -Value "Could not enforce TLS 1.2: $($_.Exception.Message)" -Severity 2
    }

    if (-not $Rerun)
    {
        #Skip the update for PowerShell 7
        if ($PSVersionTable.PSEdition -eq 'Core')
        {
            Write-LogEntry -Value "Skipping NuGet/PowerShellGet bootstrap under PowerShell 7 (current versions ship in-box)" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Checking the version of the NuGet package provider" -Severity 1
            Update-NuGet

            Write-LogEntry -Value "Checking the version of the PowerShellGet module" -Severity 1
            Update-PowerShellGet
        }
    }
}

#Discover the HPCMSL module set
if ($ModulePath)
{
    $HPModules = Get-HPCMSLOfflineModules -ModulePath $ModulePath
}
else
{
    $HPModules = Get-HPCMSLOnlineModules
}
Write-LogEntry -Value "HPCMSL module set: $($HPModules -join ', ')" -Severity 1

#Inventory the locally installed module versions across each install root
Write-LogEntry -Value "Checking the versions of the currently installed HPCMSL modules" -Severity 1
$Inventory = @{}
foreach ($Root in $ModuleInstallPaths)
{
    $Inventory[$Root] = @{}
    foreach ($HPModule in $HPModules)
    {
        $LocalVersion = $null
        $ModuleRootPath = Join-Path $Root $HPModule
        if (Test-Path $ModuleRootPath)
        {
            $VersionList = Get-ChildItem -Path $ModuleRootPath -Directory | Select-Object -ExpandProperty Name
            foreach ($V in $VersionList)
            {
                try
                {
                    $Parsed = [Version]$V
                    if (($null -eq $LocalVersion) -or ($Parsed.CompareTo([Version]$LocalVersion) -eq 1))
                    {
                        $LocalVersion = $V
                    }
                }
                catch { }
            }
        }
        if ($LocalVersion)
        {
            Write-LogEntry -Value "Installed: $HPModule $LocalVersion (under $Root)" -Severity 1
            $Inventory[$Root][$HPModule] = $LocalVersion
        }
        else
        {
            Write-LogEntry -Value "Not installed: $HPModule (under $Root)" -Severity 2
            $Inventory[$Root][$HPModule] = '0.0'
        }
    }
}

#Attempt to install the HPCMSL from local source files
if ($ModulePath)
{
    foreach ($HPModule in $HPModules)
    {
        $SourceVersion = $null
        try
        {
            $SourceVersion = Get-ChildItem -Path (Join-Path $ModulePath $HPModule) -Directory -ErrorAction Stop | Select-Object -ExpandProperty Name -First 1
            if ($SourceVersion)
            {
                Write-LogEntry -Value "The version of the $HPModule module in $ModulePath is $SourceVersion" -Severity 1
            }
        }
        catch
        {
            Write-LogEntry -Value "Failed to check the version of the $HPModule module in ${ModulePath}: $($_.Exception.Message)" -Severity 3
        }
        foreach ($Root in $ModuleInstallPaths)
        {
            if ($SourceVersion)
            {
                $Cmp = ([Version]$SourceVersion).CompareTo([Version]$Inventory[$Root][$HPModule])
                if ($Cmp -eq 0)
                {
                    Write-LogEntry -Value "The latest version of the $HPModule module is already installed under $Root" -Severity 1
                }
                elseif ($Cmp -eq -1)
                {
                    Write-LogEntry -Value "A newer version of $HPModule is already installed under $Root" -Severity 1
                }
                else
                {
                    Install-HPCMSLLocal -InstallPath $Root -ModuleName $HPModule -ModulePath (Join-Path $ModulePath $HPModule) -Version $SourceVersion
                }
            }
            else
            {
                Install-HPCMSLLocal -InstallPath $Root -ModuleName $HPModule -ModulePath (Join-Path $ModulePath $HPModule)
            }
        }
    }
}
#Attempt to install the HPCMSL module from the PowerShell Gallery
else
{
    Write-LogEntry -Value "Checking the version of the HPCMSL module in the PowerShell Gallery" -Severity 1
    $WebVersion = $null
    try
    {
        $WebVersion = Find-Module -Name HPCMSL -Repository PSGallery -ErrorAction Stop | Select-Object -First 1 -ExpandProperty Version
        if ($WebVersion)
        {
            Write-LogEntry -Value "The version of the HPCMSL module in the PowerShell Gallery is $WebVersion" -Severity 1
        }
    }
    catch
    {
        Write-LogEntry -Value "Failed to check the version of the HPCMSL module in the PowerShell Gallery: $($_.Exception.Message)" -Severity 3
    }

    $IsCore = $PSVersionTable.PSEdition -eq 'Core'
    $CurrentEditionRoot = if ($IsCore) { Join-Path $env:ProgramFiles 'PowerShell\Modules' } else { Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules' }
    $LocalHPCMSL = $Inventory[$CurrentEditionRoot].HPCMSL
    if (-not $LocalHPCMSL) { $LocalHPCMSL = '0.0' }
    if ($WebVersion)
    {
        $Cmp = ([Version]$WebVersion).CompareTo([Version]$LocalHPCMSL)
        if ($Cmp -eq 0)
        {
            Write-LogEntry -Value "The latest version of the HPCMSL module is already installed under $CurrentEditionRoot" -Severity 1
        }
        elseif ($Cmp -lt 0)
        {
            Write-LogEntry -Value "A newer version of the HPCMSL module is already installed under $CurrentEditionRoot" -Severity 1
        }
        else
        {
            Install-HPCMSLRemote -Version $WebVersion
        }
    }
    else
    {
        Install-HPCMSLRemote
    }

    #When AllEditions is set, also run Install-Module under the other PowerShell edition
    if ($AllEditions)
    {
        $OtherEditionRoot = if ($IsCore) { Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules' } else { Join-Path $env:ProgramFiles 'PowerShell\Modules' }
        $OtherLocalHPCMSL = $Inventory[$OtherEditionRoot].HPCMSL
        if (-not $OtherLocalHPCMSL) { $OtherLocalHPCMSL = '0.0' }

        $NeedsOtherInstall = $true
        if ($WebVersion)
        {
            $OtherCmp = ([Version]$WebVersion).CompareTo([Version]$OtherLocalHPCMSL)
            if ($OtherCmp -eq 0)
            {
                Write-LogEntry -Value "The latest version of the HPCMSL module is already installed under $OtherEditionRoot" -Severity 1
                $NeedsOtherInstall = $false
            }
            elseif ($OtherCmp -lt 0)
            {
                Write-LogEntry -Value "A newer version of the HPCMSL module is already installed under $OtherEditionRoot" -Severity 1
                $NeedsOtherInstall = $false
            }
        }

        if ($NeedsOtherInstall)
        {
            if ($IsCore)
            {
                $OtherHostName = 'powershell.exe'
                $OtherHostCmd = Get-Command $OtherHostName -ErrorAction SilentlyContinue
                $OtherHostSource = if ($OtherHostCmd) { $OtherHostCmd.Source } else { $null }
                $OtherHostInstallType = 'Path'
            }
            else
            {
                $OtherHostName = 'pwsh.exe'
                $PwshInfo = Find-PwshExe
                $OtherHostSource = if ($PwshInfo) { $PwshInfo.Source } else { $null }
                $OtherHostInstallType = if ($PwshInfo) { $PwshInfo.InstallType } else { $null }
            }

            if ($OtherHostSource)
            {
                Write-LogEntry -Value "Re-running Install-HPCMSL.ps1 under $OtherHostSource for -AllEditions coverage" -Severity 1
                if ($OtherHostInstallType -eq 'Msix' -and [System.Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)
                {
                    Write-LogEntry -Value "Detected MSIX-installed PowerShell 7 while running as SYSTEM. MSIX installs are per-user-registered. For reliable -AllEditions support in unattended deployments, install PowerShell 7 via the MSI package." -Severity 1
                }
                #Invoke the script so the child host can run its own NuGet/PowerShellGet bootstrap if needed
                $ChildArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ScriptPath)
                if ($Import)  { $ChildArgs += '-Import' }
                if ($LogFile) { $ChildArgs += @('-LogFile', $LogFile) }
                #Reset PSModulePath for the child to the machine-level registry value before launching
                $SavedPSModulePath = $env:PSModulePath
                try
                {
                    $env:PSModulePath = [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')
                    $Proc = Start-Process -FilePath $OtherHostSource -ArgumentList $ChildArgs -Wait -PassThru -ErrorAction Stop
                    if ($Proc.ExitCode -ne 0)
                    {
                        Write-LogEntry -Value "$OtherHostName Install-HPCMSL.ps1 exited with code $($Proc.ExitCode)" -Severity 3
                    }
                    else
                    {
                        Write-LogEntry -Value "$OtherHostName Install-HPCMSL.ps1 completed successfully" -Severity 1
                    }
                }
                catch
                {
                    Write-LogEntry -Value "Failed to install HPCMSL under ${OtherHostName}: $($_.Exception.Message)" -Severity 3
                }
                finally
                {
                    $env:PSModulePath = $SavedPSModulePath
                }
            }
            else
            {
                Write-LogEntry -Value "$OtherHostName not found on this system - skipping the second-edition install" -Severity 2
            }
        }
    }
}

#Verify the HPCMSL module is installed and discoverable
Write-LogEntry -Value "Verifying the HPCMSL module is installed and discoverable" -Severity 1
$DiscoveredHPCMSL = Get-Module -Name HPCMSL -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
if ($DiscoveredHPCMSL)
{
    Write-LogEntry -Value "Verified the HPCMSL module is discoverable (version $($DiscoveredHPCMSL.Version) at $($DiscoveredHPCMSL.ModuleBase))" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The HPCMSL module is not discoverable after installation. Confirm the module files were installed to a path in PSModulePath."
}

#Optionally import the HPCMSL module into the current session
if ($Import)
{
    Write-LogEntry -Value "Import the HPCMSL module" -Severity 1
    try
    {
        Import-Module HPCMSL -Force -ErrorAction Stop
        Write-LogEntry -Value "Successfully imported the HPCMSL module" -Severity 1
    }
    catch
    {
        #In WinPE the module import fails because some sub-modules (e.g. HP.Notifications) depend on the WinRT
        if (Test-WinPE)
        {
            Write-LogEntry -Value "Could not import the HPCMSL module in WinPE. This is expected: HP.Notifications needs resources which WinPE does not provide. Other individual modules can still be imported. Exception: $($_.Exception.Message)" -Severity 2
        }
        else
        {
            Stop-Script -ErrorMessage "Failed to import the HPCMSL module" -Exception $_.Exception.Message
        }
    }
}

Write-LogEntry -Value "END - HP Client Management Script Library installation script" -Severity 1
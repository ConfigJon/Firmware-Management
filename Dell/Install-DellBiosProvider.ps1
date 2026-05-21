<#
    .DESCRIPTION
        Install the Dell Command | PowerShell Provider (DellBIOSProvider) module

    .PARAMETER ModulePath
        Specify the location of the DellBIOSProvider module source files. This parameter should be specified when the script is running in WinPE or when the system does not have internet access

    .PARAMETER DllPath
        Specify the location of the .dll files required to run the DellBIOSProvider module.
        This parameter should be specified when the script is running in WinPE or when the system does not have the required Visual C++ Redistributables installed.
        If this parameter is not specified and the .dll files are missing on the system, the script looks for them in the script root directory

    .PARAMETER Import
        When specified, imports the DellBIOSProvider module into the current session after installation. Off by default.

    .PARAMETER LogFile
        Specify the name of the log file along with the full path where it will be stored. The file must have a .log extension. During a task sequence the path will always be set to _SMSTSLogPath

    .EXAMPLE
        Running in a full Windows OS and installing from the internet
            Install-DellBiosProvider.ps1

        Running in WinPE
            Install-DellBiosProvider.ps1 -ModulePath DellBIOSProvider -DllPath DllFiles

        Installing and importing the module into the current session
            Install-DellBiosProvider.ps1 -Import

    .NOTES
        Created by: Jon Anderson
        Reference: https://www.configjon.com/working-with-the-dell-command-powershell-provider/
        Modified: 2026-05-20

    .CHANGELOG
        2020-09-07 - Added a LogFile parameter. Changed the default log path in full Windows to $ENV:ProgramData\ConfigJonScripts\Dell.
                     Created a new function (Stop-Script) to consolidate some duplicate code and improve error reporting. Made a number of minor formatting and syntax changes
        2020-09-17 - Improved the log file path configuration
        2022-02-20 - Updated the required .dll files to support version 2.6.0 of the DellBIOSProvider module
        2026-05-20 - Added -Scope AllUsers -AllowClobber to the gallery install so the install location matches where the script inventories the module
                     Added a PowerShellGet bootstrap (Update-PowerShellGet) plus a -Rerun self-relaunch so a stock Windows PowerShell 5.1 image can reach the gallery after the update
                     Fixed the string-based version comparison in Update-NuGet
                     Added TLS 1.2 enforcement before connecting to the PowerShell Gallery
                     Install-DellBIOSProviderLocal copy failures are now fatal. The local install now also creates the destination Modules folder if it does not exist
                     Replaced the deprecated Get-WmiObject with Get-CimInstance
                     Reduced the required Visual C++ runtime DLL set to the three files DellBIOSProvider 2.10.1 actually links (msvcp140, vcruntime140, vcruntime140_1)
                     Removed the automatic import of the module at the end of the script
                     -Added a discovery check using Get-Module -ListAvailable to verify successful install
                     -Module can still be imported via optional -Import swtich
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
    [ValidateScript({
        if (-not ($_ | Test-Path))
        {
            throw "The DllPath folder path does not exist"
        }
        if (-not ($_ | Test-Path -PathType Container))
        {
            throw "The DllPath argument must be a folder path"
        }
        return $true
    })]
    [Parameter(Mandatory = $false)][System.IO.DirectoryInfo]$DllPath,
    [Parameter(Mandatory = $false)][switch]$Import,
    [Parameter(DontShow)][switch]$Rerun,
    [Parameter(Mandatory = $false)][ValidateScript({
        if ($_ -notmatch '\.log$')
        {
            throw "The file specified in the LogFile parameter must be a .log file"
        }
        return $true
    })]
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\ConfigJonScripts\Dell\Install-DellBiosProvider.log"
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
    #Determine if the script is running in the Windows Preinstallation Environment (WinPE).

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

function Uninstall-DellBIOSProvider
{
    #Uninstall the DellBIOSProvider module

    Write-LogEntry -Value "Uninstall previous versions of the DellBIOSProvider module" -Severity 1
    $Module = Get-Package DellBIOSProvider -ErrorAction SilentlyContinue
    $ModuleFile = Test-Path "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider"
    if (($null -eq $Module) -and (-not $ModuleFile))
    {
        Write-LogEntry -Value "The DellBIOSProvider module is not currently installed" -Severity 1
    }
    elseif ($null -ne $Module)
    {
        while ($null -ne $Module)
        {
            $Version = Get-Package DellBIOSProvider -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Version
            Write-LogEntry -Value "Uninstalling DellBIOSProvider module version $Version" -Severity 1
            try
            {
                $Module | Uninstall-Module -Force -ErrorAction Stop | Out-Null
                Write-LogEntry -Value "Successfully uninstalled DellBIOSProvider module version $Version" -Severity 1
            }
            catch
            {
                Write-LogEntry -Value "Failed to uninstall DellBIOSProvider module version $Version" -Severity 3
                break
            }
            Clear-Variable Module, Version
            $Module = Get-Package DellBIOSProvider -ErrorAction SilentlyContinue
        }
    }
    else
    {
        Remove-Item "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider" -Recurse -Force
        Write-LogEntry -Value "Successfully uninstalled the existing DellBIOSProvider module" -Severity 1
    }
}

function Install-DellBIOSProviderLocal
{
    #Install the DellBIOSProvider from local source files

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$InstallPath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ModulePath,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$Version
    )
    $Destination = Join-Path $InstallPath 'WindowsPowerShell\Modules'
    Write-LogEntry -Value "Install the DellBIOSProvider module from $ModulePath to $Destination" -Severity 1
    if (-not (Test-Path $Destination))
    {
        try
        {
            New-Item -Path $Destination -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch
        {
            Stop-Script -ErrorMessage "Failed to create the module destination directory $Destination" -Exception $_.Exception.Message
        }
    }
    try
    {
        Copy-Item $ModulePath -Destination $Destination -Recurse -Force -ErrorAction Stop | Out-Null
        if ($Version)
        {
            Write-LogEntry -Value "Successfully installed DellBIOSProvider module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the DellBIOSProvider module" -Severity 1
        }
    }
    catch
    {
        Stop-Script -ErrorMessage "Failed to copy the DellBIOSProvider module from $ModulePath to $Destination" -Exception $_.Exception.Message
    }
}

function Install-DellBIOSProviderRemote
{
    #Install the DellBIOSProvider from the PowerShell Gallery

    param(
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$Version
    )
    Write-LogEntry -Value "Install the DellBIOSProvider module from the PowerShell Gallery" -Severity 1
    try
    {
        Install-Module -Name DellBIOSProvider -Force -Scope AllUsers -AllowClobber -ErrorAction Stop
        if ($Version)
        {
            Write-LogEntry -Value "Successfully installed DellBIOSProvider module version $Version" -Severity 1
        }
        else
        {
            Write-LogEntry -Value "Successfully installed the DellBIOSProvider module" -Severity 1
        }
    }
    catch
    {
        Stop-Script -ErrorMessage "Unable to install the DellBIOSProvider module from the PowerShell Gallery" -Exception $_.Exception.Message
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
    $RelaunchArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$ScriptPath`"", '-Rerun')
    if ($Import)  { $RelaunchArgs += '-Import' }
    if ($DllPath) { $RelaunchArgs += @('-DllPath', "`"$DllPath`"") }
    if ($LogFile) { $RelaunchArgs += @('-LogFile', "`"$LogFile`"") }
    $RerunProc = Start-Process -FilePath $HostExe -ArgumentList $RelaunchArgs -Wait -PassThru
    exit $RerunProc.ExitCode
}

function Copy-Dll
{
    #Copy .dll files

    param(
        [ValidateScript({
            if ($_ -notmatch '\.dll$')
            {
                throw "The specified file must be a .dll file"
            }
            return $true
        })]
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$DllFile,
        [ValidateScript({
            if (-not ($_ | Test-Path))
            {
                throw "The DllTargetPath folder path does not exist"
            }
            if (-not ($_ | Test-Path -PathType Container))
            {
                throw "The DllTargetPath argument must be a folder path"
            }
            return $true
        })]
        [Parameter(Mandatory = $true)][System.IO.DirectoryInfo]$DllTargetPath,
        [ValidateScript({
            if (-not ($_ | Test-Path))
            {
                throw "The DllSourcePath folder path does not exist"
            }
            if (-not ($_ | Test-Path -PathType Container))
            {
                throw "The DllSourcePath argument must be a folder path"
            }
            return $true
        })]
        [Parameter(Mandatory = $false)][System.IO.DirectoryInfo]$DllSourcePath
    )
    if (-not (Test-Path "$DllTargetPath\$DllFile"))
    {
        Write-LogEntry -Value "Could not find $DllTargetPath\$DllFile" -Severity 2
        Write-LogEntry -Value "Copying $DllFile to $DllTargetPath" -Severity 1
        $Source = if ($DllSourcePath) { Join-Path $DllSourcePath $DllFile } elseif ($PSScriptRoot) { Join-Path $PSScriptRoot $DllFile } else { "$DllFile" }
        try
        {
            Copy-Item -Path $Source -Destination "$DllTargetPath\$DllFile" -Force -ErrorAction Stop
            Write-LogEntry -Value "Successfully copied $DllFile to $DllTargetPath" -Severity 1
        }
        catch
        {
            $Script:DllFailure = $true
            Write-LogEntry -Value "Failed to copy $DllFile to $DllTargetPath" -Severity 2
        }
    }
    else
    {
        Write-LogEntry -Value "Found $DllFile" -Severity 1
    }
}

function Write-LogEntry
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
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""Install-DellBiosProvider"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
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
    Write-LogEntry -Value "START - Dell BIOS provider installation script" -Severity 1
}
else
{
    Write-LogEntry -Value "Script re-launched with -Rerun after PowerShellGet update" -Severity 1
}

#Check the PowerShell version
Write-LogEntry -Value "Checking the installed PowerShell version" -Severity 1
$PsVer = $PSVersionTable.PSVersion | Select-Object -ExpandProperty Major
if ($PsVer -ge 3)
{
    Write-LogEntry -Value "The current PowerShell version is $PsVer" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The current PowerShell version is $PsVer. The minimum supported PowerShell version is 3"
}

#Check the SMBIOS version
Write-LogEntry -Value "Checking the SMBIOS version of the system" -Severity 1
$Bios = Get-CimInstance -ClassName Win32_BIOS
$BiosVerMajor = $Bios | Select-Object -ExpandProperty SMBIOSMajorVersion
$BiosVerMinor = $Bios | Select-Object -ExpandProperty SMBIOSMinorVersion
$BiosVerFull = "$($BiosVerMajor)." + "$($BiosVerMinor)"
$BiosVer = $null
if (($null -ne $BiosVerMajor) -and ($null -ne $BiosVerMinor))
{
    try { $BiosVer = [Version]$BiosVerFull } catch { $BiosVer = $null }
}
if (($null -ne $BiosVer) -and ($BiosVer -ge [Version]'2.4'))
{
    Write-LogEntry -Value "The current SMBIOS version is $BiosVerFull" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The current SMBIOS version is $BiosVerFull. The minimum supported SMBIOS version is 2.4"
}

#Verify the required Visual C++ runtime DLLs exist (copied to System32 when missing)
Write-LogEntry -Value "Verify Visual C++ DLL files exist" -Severity 1
$RequiredDlls = @('msvcp140.dll', 'vcruntime140.dll', 'vcruntime140_1.dll')
foreach ($Dll in $RequiredDlls)
{
    if ($DllPath)
    {
        Copy-Dll -DllSourcePath "$DllPath" -DllFile $Dll -DllTargetPath "$env:windir\System32"
    }
    else
    {
        Copy-Dll -DllFile $Dll -DllTargetPath "$env:windir\System32"
    }
}
if ($DllFailure)
{
    Stop-Script -ErrorMessage "One or more .dll files were not found or failed to copy"
}

#Check if 32 or 64 bit architecture
if ([System.Environment]::Is64BitOperatingSystem)
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
    $LocalVersion = Get-Package DellBIOSProvider -ErrorAction Stop |
        Select-Object -ExpandProperty Version -ErrorAction Stop |
        Sort-Object { [Version]$_ } -Descending |
        Select-Object -First 1
}
catch
{
    $Local = $true
    $LocalModuleRoot = "$ModuleInstallPath\WindowsPowerShell\Modules\DellBIOSProvider"
    $LocalPsd1 = if (Test-Path $LocalModuleRoot) { Get-ChildItem -Path $LocalModuleRoot -Filter 'DellBIOSProvider.psd1' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1 } else { $null }
    if ($LocalPsd1)
    {
        $LocalVersion = Get-Content $LocalPsd1.FullName | Select-String "ModuleVersion ="
        $LocalVersion = (([regex]".*'(.*)'").Matches($LocalVersion))[0].Groups[1].Value
        if ($null -ne $LocalVersion)
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
if (($null -ne $LocalVersion) -and (-not $Local))
{
    Write-LogEntry -Value "The version of the currently installed DellBIOSProvider module is $LocalVersion" -Severity 1
}

#Attempt to install the DellBIOSProvider module from local source files
if ($ModulePath)
{
    $SourcePsd1 = Get-ChildItem -Path $ModulePath -Filter 'DellBIOSProvider.psd1' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($SourcePsd1)
    {
        try
        {
            Write-LogEntry -Value "Checking the version of the DellBIOSProvider module in $ModulePath" -Severity 1
            $SourceVersion = Get-Content $SourcePsd1.FullName -ErrorAction Stop | Select-String "ModuleVersion =" -ErrorAction Stop
            $SourceVersion = (([regex]".*'(.*)'").Matches($SourceVersion))[0].Groups[1].Value
            if ($null -ne $SourceVersion)
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
    if (($null -ne $SourceVersion) -and ($null -ne $LocalVersion))
    {
        if ($SourceVersion -eq $LocalVersion)
        {
            Write-LogEntry -Value "The latest version of the DellBIOSProvider module is already installed" -Severity 1
        }
        else
        {
            Uninstall-DellBIOSProvider
            Install-DellBIOSProviderLocal -InstallPath $ModuleInstallPath -ModulePath $ModulePath -Version $SourceVersion
        }
    }
    elseif (($null -ne $SourceVersion) -and ($null -eq $LocalVersion))
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
    try
    {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Write-LogEntry -Value "Enforced TLS 1.2 for outbound connections" -Severity 1
    }
    catch
    {
        Write-LogEntry -Value "Could not enforce TLS 1.2: $($_.Exception.Message)" -Severity 2
    }

    #Bootstrap NuGet + PowerShellGet so a stock Windows PowerShell 5.1 image can reach the gallery
    if (-not $Rerun)
    {
        Write-LogEntry -Value "Checking the version of the NuGet package provider" -Severity 1
        Update-NuGet
        Write-LogEntry -Value "Checking the version of the PowerShellGet module" -Severity 1
        Update-PowerShellGet
    }

    #Get the version of the DellBIOSProvider module in the PowerShell Gallery
    Write-LogEntry -Value "Checking the version of the DellBIOSProvider module in the PowerShell Gallery" -Severity 1
    try
    {
        $WebVersion = Find-Package DellBIOSProvider -ErrorAction Stop | Select-Object -ExpandProperty Version -ErrorAction Stop
        if ($null -ne $WebVersion)
        {
            Write-LogEntry -Value "The version of the DellBIOSProvider module in the PowerShell Gallery is $WebVersion" -Severity 1
        }
    }
    catch
    {
        Write-LogEntry -Value "Failed to check the version of the DellBIOSProvider module in the PowerShell Gallery" -Severity 3
    }
    if ($null -eq $WebVersion)
    {
        if ($null -ne $LocalVersion)
        {
            Write-LogEntry -Value "Could not determine the gallery version of the DellBIOSProvider module. Keeping the currently installed version ($LocalVersion) rather than risk removing a working module." -Severity 2
        }
        else
        {
            Write-LogEntry -Value "Could not determine the gallery version, and no DellBIOSProvider module is installed. Attempting a best-effort install of the latest version from the gallery." -Severity 2
            Install-DellBIOSProviderRemote
        }
    }
    elseif ($null -ne $LocalVersion)
    {
        if ($WebVersion -eq $LocalVersion)
        {
            Write-LogEntry -Value "The latest version of the DellBIOSProvider module is already installed" -Severity 1
        }
        else
        {
            Uninstall-DellBIOSProvider
            Install-DellBIOSProviderRemote -Version $WebVersion
        }
    }
    else
    {
        Install-DellBIOSProviderRemote -Version $WebVersion
    }
}

#Verify the DellBIOSProvider module is installed and discoverable
Write-LogEntry -Value "Verifying the DellBIOSProvider module is installed and discoverable" -Severity 1
$Discovered = Get-Module -Name DellBIOSProvider -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
if ($Discovered)
{
    Write-LogEntry -Value "Verified the DellBIOSProvider module is discoverable (version $($Discovered.Version) at $($Discovered.ModuleBase))" -Severity 1
}
else
{
    Stop-Script -ErrorMessage "The DellBIOSProvider module is not discoverable after installation. Confirm the module files were installed to a path in PSModulePath."
}

#Optionally import the DellBIOSProvider module into the current session
if ($Import)
{
    Write-LogEntry -Value "Import the DellBIOSProvider module" -Severity 1
    try
    {
        Import-Module DellBIOSProvider -Force -ErrorAction Stop
        Write-LogEntry -Value "Successfully imported the DellBIOSProvider module" -Severity 1
    }
    catch
    {
        if (Test-WinPE)
        {
            Stop-Script -ErrorMessage "Failed to import the DellBIOSProvider module in WinPE. Confirm the required Visual C++ DLLs are present (see the -DllPath parameter) and that the WinPE-NetFx, WinPE-Scripting, WinPE-WMI, and WinPE-PowerShell components are included in the boot image." -Exception $_.Exception.Message
        }
        else
        {
            Stop-Script -ErrorMessage "Failed to import the DellBIOSProvider module" -Exception $_.Exception.Message
        }
    }
}

Write-LogEntry -Value "END - Dell BIOS provider installation script" -Severity 1
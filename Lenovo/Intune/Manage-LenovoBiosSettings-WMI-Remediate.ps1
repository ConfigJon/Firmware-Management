<#
    .DESCRIPTION
        Intune remediation script for Lenovo BIOS settings management. Compares the admin-edited
        $DesiredSettings hashtable to the BIOS-reported state, decrypts the password named by the pw marker
        (CMS material in $IntunePayload), stages each drifted setting (SetBIOSSetting) and commits them with
        a single SaveBiosSettings. Settings that fail are cached so they do not retry forever on hardware
        that cannot accept them. Skips on non-Lenovo hardware reporting (SKIPPED: settings out-of-scope).
        Settings are committed as one batch and apply at next reboot, so the Save status is authoritative
        (no readback - the next Detect cycle verifies). Cert-auth devices (PasswordState=128) short-circuit
        (SKIPPED) and the -Detect script flags them INCOMPATIBLE.

        Build the payload with Tools\Build-IntunePayload.ps1; the certificate it references must be in
        Cert:\LocalMachine\My on the device before this runs. Full walkthrough in the blog posts below.

    .PARAMETER Profile
        Name of the desired-state profile, stamped into the marker. Must match the value in the paired script.

    .PARAMETER RetryFailedAfterDays
        Days after which a per-setting Failed cache is retried. 0 (default) = never retry automatically.

    .PARAMETER NoPassword
        For devices with NO BIOS supervisor password. Default OFF. When set, the password-marker dependency is dropped
        and the run is tagged 'mode=nopassword'. MUST match the value in the paired script. Do not combine with
        BIOS password management on the same device.

    .PARAMETER LogFile
        CMTrace-compatible diagnostics log. Defaults to the IME log directory. Must have a .log extension.

    .LINK
        https://www.configjon.com/intune-bios-settings-management/
        https://www.configjon.com/bios-management-with-intune/

    .NOTES
        Created by: Jon Anderson
        Version: 1.0.0
        Modified: 2026-06-06
#>

#Parameters ===================================================================================================================

[CmdletBinding(SupportsShouldProcess = $true, PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Profile = 'baseline-v1',

    #Internal: marker registry root.
    [Parameter(DontShow)]
    [ValidateNotNullOrEmpty()]
    [string]$MarkerBasePath = 'HKLM:\SOFTWARE\ConfigJonScripts\FirmwareManagement\BIOSSettings',

    #Internal: password marker location (read-only).
    [Parameter(DontShow)]
    [ValidateNotNullOrEmpty()]
    [string]$PwMarkerPath = 'HKLM:\SOFTWARE\ConfigJonScripts\FirmwareManagement\BIOSPassword',

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$RetryFailedAfterDays = 0,

    #Opt-in: BIOS has no supervisor password. Settings are read/applied unauthenticated. Must match the paired -Detect script.
    [Parameter(Mandatory = $false)]
    [switch]$NoPassword,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
            if ($_ -notmatch '\.log$')
            {
                throw "The file specified in the LogFile parameter must be a .log file"
            }
            return $true
        })]
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Manage-LenovoBiosSettings-WMI-Remediate.log",

    [Parameter(DontShow)]
    [switch]$SkipManufacturerCheck
)

$Version = '1.0.0'
$Component = 'Manage-LenovoBiosSettings-WMI-Remediate'

#Desired state ===============================================================================================================
#Edit this hashtable to match your desired Lenovo BIOS configuration.
#Names match the first comma-delimited field of Lenovo_BiosSetting.CurrentSetting exactly.
#Values match the parsed CurrentSetting value exactly. ('Enable' / 'Disable' / 'Auto').
#Use the existing Manage-LenovoBiosSettings.ps1 GetSettings mode to list what the device exposes.

$DesiredSettings = @{
    # Examples - replace with the settings your devices should standardize on:
    # 'WakeOnLAN'                     = 'AutomaticEnable'
    # 'VirtualizationTechnology'      = 'Enable'
    # 'FingerprintReader'             = 'Enable'
    # 'PasswordCountExceededError'    = 'Disable'
}

#Payload =====================================================================================================================
#Replaced by Tools\Build-IntunePayload.ps1 before deployment.

# <Payload>
$IntunePayload = $null
# </Payload>

#Functions ====================================================================================================================

function Write-LogEntry
{
    #CMTrace-compatible logger.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Value,
        [Parameter(Mandatory = $true)][ValidateSet("1", "2", "3")][string]$Severity,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()][string]$FileName = ($script:LogFile | Split-Path -Leaf)
    )
    $LogFilePath = Join-Path -Path $LogsDirectory -ChildPath $FileName
    [string]$Bias = [System.TimeZoneInfo]::Local.GetUtcOffset((Get-Date)).TotalMinutes
    if ($Bias -match "^-")
    {
        $TimezoneBias = $Bias.Replace('-', '+')
    }
    else
    {
        $TimezoneBias = '-' + $Bias
    }
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), $TimezoneBias)
    $Date = (Get-Date -Format "MM-dd-yyyy")
    $Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    $LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$Component"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
    try
    {
        Out-File -InputObject $LogText -Append -Encoding Default -FilePath $LogFilePath -ErrorAction Stop -WhatIf:$false
    }
    catch [System.Exception]
    {
        Write-Warning -Message "Unable to append log entry to $FileName file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

function Get-PasswordMarker
{
    #Read the BIOS password marker from the registry.

    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath)
    if (-not (Test-Path -LiteralPath $BasePath)) { return $null }
    try { $Item = Get-ItemProperty -LiteralPath $BasePath -ErrorAction Stop } catch { return $null }
    return @{
        Version        = [int]$Item.Version
        SetDate        = [string]$Item.SetDate
        CertThumbprint = [string]$Item.CertThumbprint
    }
}

function Get-SettingsMarker
{
    #Read the top-level BIOS settings marker.

    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath)
    if (-not (Test-Path -LiteralPath $BasePath)) { return $null }
    try { $Item = Get-ItemProperty -LiteralPath $BasePath -ErrorAction Stop } catch { return $null }
    return @{
        Profile          = [string]$Item.Profile
        LastFullRun      = [string]$Item.LastFullRun
        DesiredStateHash = [string]$Item.DesiredStateHash
    }
}

function Set-SettingsMarker
{
    #Write the top-level settings marker (Profile, LastFullRun, DesiredStateHash).

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Profile,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$DesiredStateHash
    )
    if ($WhatIfPreference)
    {
        Write-LogEntry -Value "WhatIf: would write top-level settings marker Profile=$Profile DesiredStateHash=$DesiredStateHash at $BasePath" -Severity 1
        return
    }
    if (-not (Test-Path -LiteralPath $BasePath))
    {
        New-Item -Path $BasePath -Force | Out-Null
    }
    $Now = (Get-Date).ToString('o')
    Set-ItemProperty -LiteralPath $BasePath -Name 'Profile' -Value $Profile -Type String -Force
    Set-ItemProperty -LiteralPath $BasePath -Name 'LastFullRun' -Value $Now -Type String -Force
    Set-ItemProperty -LiteralPath $BasePath -Name 'DesiredStateHash' -Value $DesiredStateHash -Type String -Force
}

function Get-SettingMarker
{
    #Read a per-setting marker.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Name
    )
    $SettingPath = Join-Path -Path $BasePath -ChildPath "Settings\$Name"
    if (-not (Test-Path -LiteralPath $SettingPath)) { return $null }
    try { $Item = Get-ItemProperty -LiteralPath $SettingPath -ErrorAction Stop } catch { return $null }
    return @{
        Name              = $Name
        DesiredValue      = [string]$Item.DesiredValue
        LastVerifiedValue = [string]$Item.LastVerifiedValue
        SetDate           = [string]$Item.SetDate
        Status            = [string]$Item.Status
        FailReason        = [string]$Item.FailReason
    }
}

function Set-SettingMarker
{
    #Write a per-setting marker.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Name,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$DesiredValue,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$LastVerifiedValue,
        [Parameter(Mandatory = $true)][ValidateSet('Applied', 'Failed')][string]$Status,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$FailReason = ''
    )
    if ($WhatIfPreference)
    {
        Write-LogEntry -Value "WhatIf: would write per-setting marker [$Name] DesiredValue=$DesiredValue LastVerifiedValue=$LastVerifiedValue Status=$Status FailReason=$FailReason" -Severity 1
        return
    }
    $SettingPath = Join-Path -Path $BasePath -ChildPath "Settings\$Name"
    if (-not (Test-Path -LiteralPath $SettingPath))
    {
        New-Item -Path $SettingPath -Force | Out-Null
    }
    $Now = (Get-Date).ToString('o')
    Set-ItemProperty -LiteralPath $SettingPath -Name 'DesiredValue' -Value $DesiredValue -Type String -Force
    Set-ItemProperty -LiteralPath $SettingPath -Name 'LastVerifiedValue' -Value $LastVerifiedValue -Type String -Force
    Set-ItemProperty -LiteralPath $SettingPath -Name 'SetDate' -Value $Now -Type String -Force
    Set-ItemProperty -LiteralPath $SettingPath -Name 'Status' -Value $Status -Type String -Force
    Set-ItemProperty -LiteralPath $SettingPath -Name 'FailReason' -Value $FailReason -Type String -Force
}

function Get-AllSettingMarkers
{
    #Enumerate per-setting markers beneath <BasePath>\Settings.

    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath)
    $SettingsPath = Join-Path -Path $BasePath -ChildPath 'Settings'
    $Result = [ordered]@{}
    if (-not (Test-Path -LiteralPath $SettingsPath)) { return $Result }
    $Children = Get-ChildItem -LiteralPath $SettingsPath -ErrorAction SilentlyContinue
    foreach ($Child in $Children)
    {
        $Marker = Get-SettingMarker -BasePath $BasePath -Name $Child.PSChildName
        if ($Marker)
        {
            $Result[$Child.PSChildName] = $Marker
        }
    }
    return $Result
}

function Get-LenovoCertAuthEnabled
{
    #Return $true if PasswordState = 128 (Lenovo cert-based BIOS auth).

    try
    {
        $PasswordSettings = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_BiosPasswordSettings' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        Write-LogEntry -Value "Lenovo_BiosPasswordSettings WMI query (cert-auth) failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    if ($null -eq $PasswordSettings) { return $false }
    return ([int]$PasswordSettings.PasswordState -eq 128)
}

function Get-LenovoPasswordState
{
    #Return the raw Lenovo_BiosPasswordSettings.PasswordState (int) or $null on failure.

    try
    {
        $PasswordSettings = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_BiosPasswordSettings' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        Write-LogEntry -Value "Lenovo_BiosPasswordSettings WMI query (PasswordState) failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    if ($null -eq $PasswordSettings) { return $null }
    return [int]$PasswordSettings.PasswordState
}

function Get-LenovoBiosSettingCurrentValue
{
    #Parse a Lenovo_BiosSetting CurrentSetting ("Name,Value") and return the value.

    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$RawSetting)
    if ([string]::IsNullOrEmpty($RawSetting)) { return '' }
    $Formatted = if ($RawSetting -match ';') { $RawSetting.Substring(0, $RawSetting.IndexOf(';')) } else { $RawSetting }
    $Parts = $Formatted.Split(',', 2)
    if ($Parts.Count -lt 2) { return '' }
    return $Parts[1].Trim()
}

function Get-LenovoBiosSettingName
{
    #Parse the Name portion of a Lenovo_BiosSetting CurrentSetting string. Returns '' on unparseable input.

    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$RawSetting)
    if ([string]::IsNullOrEmpty($RawSetting)) { return '' }
    $Formatted = if ($RawSetting -match ';') { $RawSetting.Substring(0, $RawSetting.IndexOf(';')) } else { $RawSetting }
    $Parts = $Formatted.Split(',', 2)
    if ($Parts.Count -lt 1) { return '' }
    return $Parts[0].Trim()
}

function Get-LenovoBiosSettings
{
    #Read all Lenovo BIOS settings via Lenovo_BiosSetting and return as Name -> current value.

    try
    {
        $Items = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_BiosSetting' -ErrorAction Stop
    }
    catch
    {
        Write-LogEntry -Value "Lenovo_BiosSetting WMI query failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    $Result = @{}
    foreach ($Item in $Items)
    {
        $Raw = [string]$Item.CurrentSetting
        if ([string]::IsNullOrWhiteSpace($Raw)) { continue }
        $Name = Get-LenovoBiosSettingName -RawSetting $Raw
        if ([string]::IsNullOrWhiteSpace($Name)) { continue }
        $Result[$Name] = Get-LenovoBiosSettingCurrentValue -RawSetting $Raw
    }
    return $Result
}

function Test-RemediationCert
{
    #Verify the CMS certificate is installed, has a private key, and isn't expired.

    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Thumbprint)
    $CertPath = "Cert:\LocalMachine\My\$Thumbprint"
    if (-not (Test-Path -LiteralPath $CertPath))
    {
        return @{ Ok = $false; Reason = "cert-missing thumbprint=$Thumbprint (certificate not deployed to LocalMachine\My?)" }
    }
    $Cert = Get-Item -LiteralPath $CertPath
    if (-not $Cert.HasPrivateKey)
    {
        return @{ Ok = $false; Reason = "cert-no-private-key thumbprint=$Thumbprint" }
    }
    if ($Cert.NotAfter -lt (Get-Date))
    {
        return @{ Ok = $false; Reason = "cert-expired thumbprint=$Thumbprint expired=$($Cert.NotAfter.ToString('o'))" }
    }
    return @{ Ok = $true; Reason = '' }
}

function Get-CmsPlaintextFromPayload
{
    #Decrypt the CMS password material for the given version and return the plaintext.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)][ValidateRange(1, [int]::MaxValue)][int]$Version
    )
    $Key = "$Version"
    if (-not $Payload.Files.Contains($Key))
    {
        throw "Payload does not contain a CMS file for version $Version"
    }
    $Base64 = [string]$Payload.Files[$Key]
    if ([string]::IsNullOrEmpty($Base64))
    {
        throw "Payload entry for version $Version is empty"
    }
    $Bytes = [Convert]::FromBase64String($Base64)
    $PemText = [System.Text.Encoding]::UTF8.GetString($Bytes)
    try
    {
        $Plaintext = Unprotect-CmsMessage -Content $PemText -ErrorAction Stop
    }
    catch
    {
        throw "Failed to decrypt CMS for version $Version (cert thumbprint $($Payload.CertThumbprint) likely missing or no private key): $($_.Exception.Message)"
    }
    return $Plaintext
}

function Get-DesiredStateHash
{
    #SHA-256 over $DesiredSettings ("name=value" lines, sorted).

    param([Parameter(Mandatory = $true)][hashtable]$DesiredSettings)
    if ($DesiredSettings.Count -eq 0) { return '' }
    $Canonical = ($DesiredSettings.Keys | Sort-Object | ForEach-Object { "$_=$($DesiredSettings[$_])" }) -join "`n"
    $Sha = [System.Security.Cryptography.SHA256]::Create()
    try { $HashBytes = $Sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Canonical)) }
    finally { $Sha.Dispose() }
    return (([System.BitConverter]::ToString($HashBytes)) -replace '-', '').ToLowerInvariant()
}

function Format-DriftNames
{
    #Join drifted setting names into a comma-separated list, capped at MaxLength chars with '+more' suffix so STDOUT stays within the Intune 2048-byte budget.

    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Names,
        [Parameter(Mandatory = $false)][int]$MaxLength = 1500
    )
    if ($Names.Count -eq 0) { return '' }
    $List = $Names -join ','
    if ($List.Length -le $MaxLength) { return $List }
    $Acc = ''
    foreach ($N in $Names)
    {
        $Candidate = if ($Acc -eq '') { $N } else { "$Acc,$N" }
        if (($Candidate + ',+more').Length -gt $MaxLength) { break }
        $Acc = $Candidate
    }
    return "$Acc,+more"
}

function Get-SettingsRemediationPlan
{
    #Determine remediation plan based on desired state + observed BIOS state + markers + payload coverage.

    param(
        [Parameter(Mandatory = $true)][hashtable]$DesiredSettings,
        [Parameter(Mandatory = $false)][AllowNull()]$CurrentBiosValues,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$PerSettingMarkers,
        [Parameter(Mandatory = $true)][bool]$PasswordMarkerPresent,
        [Parameter(Mandatory = $false)][int]$PasswordMarkerVersion = 0,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][int[]]$PayloadVersions,
        [Parameter(Mandatory = $false)][bool]$NoPassword = $false,
        [Parameter(Mandatory = $false)][ValidateRange(0, [int]::MaxValue)][int]$RetryFailedAfterDays = 0,
        [Parameter(Mandatory = $false)][datetime]$Now = (Get-Date)
    )

    if (-not $NoPassword -and -not $PasswordMarkerPresent)
    {
        return @{
            Verdict           = 'EarlyExitFail'
            EarlyExitStdout   = 'FAILED: settings pw-marker-missing (cannot remediate without password)'
            EarlyExitCode     = 1
            Reason            = 'pw marker missing - cannot determine which CMS to decrypt for BIOS authorize'
            PerSettingPlan    = @()
            SetCount          = 0
            CompliantCount    = 0
            UnsupportedCount  = 0
            FailedCachedCount = 0
        }
    }
    if ($null -eq $CurrentBiosValues)
    {
        return @{
            Verdict           = 'EarlyExitFail'
            EarlyExitStdout   = 'DEGRADED: settings bios-query-failed (Lenovo BIOS setting namespace unreachable)'
            EarlyExitCode     = 1
            Reason            = 'CIM query for Lenovo_BiosSetting failed'
            PerSettingPlan    = @()
            SetCount          = 0
            CompliantCount    = 0
            UnsupportedCount  = 0
            FailedCachedCount = 0
        }
    }
    if (-not $NoPassword -and $PasswordMarkerVersion -le 0)
    {
        return @{
            Verdict           = 'EarlyExitFail'
            EarlyExitStdout   = 'FAILED: settings pw-marker-cleared (Version=0 - clear pw before applying authenticated settings)'
            EarlyExitCode     = 1
            Reason            = 'pw marker indicates cleared password (Version=0); cannot authenticate SetBIOSSetting'
            PerSettingPlan    = @()
            SetCount          = 0
            CompliantCount    = 0
            UnsupportedCount  = 0
            FailedCachedCount = 0
        }
    }
    if (-not $NoPassword -and ($PayloadVersions -notcontains $PasswordMarkerVersion))
    {
        return @{
            Verdict           = 'EarlyExitFail'
            EarlyExitStdout   = "FAILED: settings pw-version-not-in-payload v=$PasswordMarkerVersion"
            EarlyExitCode     = 1
            Reason            = "pw marker says v=$PasswordMarkerVersion but settings payload does not contain it (re-run Build-IntunePayload after adding the CMS)"
            PerSettingPlan    = @()
            SetCount          = 0
            CompliantCount    = 0
            UnsupportedCount  = 0
            FailedCachedCount = 0
        }
    }

    $Plan = New-Object System.Collections.Generic.List[hashtable]
    $UnsupportedCount = 0
    $FailedCachedCount = 0
    $CompliantCount = 0
    $SetCount = 0

    foreach ($Name in ($DesiredSettings.Keys | Sort-Object))
    {
        $DesiredValue = [string]$DesiredSettings[$Name]
        $Current = $CurrentBiosValues[$Name]

        if ($null -eq $Current)
        {
            [void]$Plan.Add(@{
                    Name         = $Name
                    Action       = 'Skip-Unsupported'
                    DesiredValue = $DesiredValue
                    CurrentValue = $null
                })
            $UnsupportedCount++
            continue
        }
        $Current = [string]$Current

        $Marker = $PerSettingMarkers[$Name]
        if ($null -ne $Marker -and $Marker.Status -eq 'Failed' -and $Marker.DesiredValue -eq $DesiredValue)
        {
            $RetryNow = $false
            if ($RetryFailedAfterDays -gt 0 -and -not [string]::IsNullOrEmpty($Marker.SetDate))
            {
                $LastTry = [datetime]::MinValue
                if ([datetime]::TryParse($Marker.SetDate, [ref]$LastTry))
                {
                    if (($Now - $LastTry).TotalDays -ge $RetryFailedAfterDays) { $RetryNow = $true }
                }
            }
            if (-not $RetryNow)
            {
                [void]$Plan.Add(@{
                        Name         = $Name
                        Action       = 'Skip-FailedCached'
                        DesiredValue = $DesiredValue
                        CurrentValue = $Current
                    })
                $FailedCachedCount++
                continue
            }
        }

        if ($Current -eq $DesiredValue)
        {
            [void]$Plan.Add(@{
                    Name         = $Name
                    Action       = 'Skip-Compliant'
                    DesiredValue = $DesiredValue
                    CurrentValue = $Current
                })
            $CompliantCount++
        }
        else
        {
            [void]$Plan.Add(@{
                    Name         = $Name
                    Action       = 'Set'
                    DesiredValue = $DesiredValue
                    CurrentValue = $Current
                })
            $SetCount++
        }
    }

    return @{
        Verdict           = 'CanProceed'
        EarlyExitStdout   = ''
        EarlyExitCode     = 0
        Reason            = ''
        PerSettingPlan    = $Plan.ToArray()
        SetCount          = $SetCount
        CompliantCount    = $CompliantCount
        UnsupportedCount  = $UnsupportedCount
        FailedCachedCount = $FailedCachedCount
    }
}

function Invoke-LenovoSetBiosSetting
{
    #Stage a single Lenovo BIOS setting.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$SettingName,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$SettingValue,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$Password,
        [Parameter(Mandatory = $false)][switch]$NoPassword
    )
    if ($WhatIfPreference)
    {
        Write-LogEntry -Value "WhatIf: would call Lenovo SetBIOSSetting name=$SettingName value=$SettingValue" -Severity 1
        return @{ Status = 'Success'; Threw = $false; ExceptionMessage = '' }
    }
    if ($NoPassword)
    {
        #NoPassword: write the Name,Value pair directly, no opcode authorization.
        try
        {
            $Interface = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_SetBiosSetting' -ErrorAction Stop | Select-Object -First 1
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_SetBiosSetting failed: $($_.Exception.Message)" }
        }
        try
        {
            $Result = Invoke-CimMethod -InputObject $Interface -MethodName 'SetBIOSSetting' -Arguments @{ parameter = "$SettingName,$SettingValue" } -ErrorAction Stop
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Unauthenticated SetBIOSSetting threw: $($_.Exception.Message)" }
        }
        return @{ Status = [string]$Result.Return; Threw = $false; ExceptionMessage = '' }
    }
    try
    {
        $OpcodeInterface = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_WmiOpcodeInterface' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_WmiOpcodeInterface failed: $($_.Exception.Message)" }
    }
    try
    {
        $Interface = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_SetBiosSetting' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_SetBiosSetting failed: $($_.Exception.Message)" }
    }
    if (($null -ne $OpcodeInterface) -and ($OpcodeInterface.Active -eq $true))
    {
        #Opcode-authorized path: WmiOpcodePasswordAdmin authorizes, then SetBIOSSetting carries Name,Value (Admin opcode is non-poisoning for settings methods).
        try
        {
            [void](Invoke-CimMethod -InputObject $OpcodeInterface -MethodName 'WmiOpcodeInterface' -Arguments @{ Parameter = "WmiOpcodePasswordAdmin:$Password;" } -ErrorAction Stop)
            $Result = Invoke-CimMethod -InputObject $Interface -MethodName 'SetBIOSSetting' -Arguments @{ parameter = "$SettingName,$SettingValue" } -ErrorAction Stop
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Opcode-authorized SetBIOSSetting threw: $($_.Exception.Message)" }
        }
        return @{ Status = [string]$Result.Return; Threw = $false; ExceptionMessage = '' }
    }
    #Legacy fallback: password embedded in the parameter string (complex characters may fail).
    try
    {
        $Result = Invoke-CimMethod -InputObject $Interface -MethodName 'SetBIOSSetting' -Arguments @{ parameter = "$SettingName,$SettingValue,$Password,ascii,us" } -ErrorAction Stop
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Legacy SetBIOSSetting threw: $($_.Exception.Message)" }
    }
    return @{ Status = [string]$Result.Return; Threw = $false; ExceptionMessage = '' }
}

function Invoke-LenovoSaveBiosSettings
{
    #Commit the staged Lenovo BIOS settings.
    #SaveBiosSettings returns its status via .Value (not .Return); empty/null Value means success on many
    #firmware versions, so it is normalized to 'Success' here. Returns @{ Status; Threw; ExceptionMessage }.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$Password,
        [Parameter(Mandatory = $false)][switch]$NoPassword
    )
    if ($WhatIfPreference)
    {
        Write-LogEntry -Value "WhatIf: would call Lenovo SaveBiosSettings" -Severity 1
        return @{ Status = 'Success'; Threw = $false; ExceptionMessage = '' }
    }
    if ($NoPassword)
    {
        #NoPassword: bare Save, no opcode authorization.
        try
        {
            $SaveSettings = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_SaveBiosSettings' -ErrorAction Stop | Select-Object -First 1
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_SaveBiosSettings failed: $($_.Exception.Message)" }
        }
        try
        {
            $Result = Invoke-CimMethod -InputObject $SaveSettings -MethodName 'SaveBiosSettings' -ErrorAction Stop
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Unauthenticated SaveBiosSettings threw: $($_.Exception.Message)" }
        }
        #Same empty/null-on-success normalization as the authenticated paths (see function header).
        $RawNp = [string]$Result.Value
        $StatusNp = if ([string]::IsNullOrEmpty($RawNp)) { 'Success' } else { $RawNp }
        return @{ Status = $StatusNp; Threw = $false; ExceptionMessage = '' }
    }
    try
    {
        $OpcodeInterface = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_WmiOpcodeInterface' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_WmiOpcodeInterface failed: $($_.Exception.Message)" }
    }
    try
    {
        $SaveSettings = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_SaveBiosSettings' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_SaveBiosSettings failed: $($_.Exception.Message)" }
    }
    $RawStatus = $null
    if (($null -ne $OpcodeInterface) -and ($OpcodeInterface.Active -eq $true))
    {
        try
        {
            [void](Invoke-CimMethod -InputObject $OpcodeInterface -MethodName 'WmiOpcodeInterface' -Arguments @{ Parameter = "WmiOpcodePasswordAdmin:$Password;" } -ErrorAction Stop)
            $Result = Invoke-CimMethod -InputObject $SaveSettings -MethodName 'SaveBiosSettings' -ErrorAction Stop
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Opcode-authorized SaveBiosSettings threw: $($_.Exception.Message)" }
        }
        $RawStatus = [string]$Result.Value
    }
    else
    {
        try
        {
            $Result = Invoke-CimMethod -InputObject $SaveSettings -MethodName 'SaveBiosSettings' -Arguments @{ parameter = "$Password,ascii,us" } -ErrorAction Stop
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Legacy SaveBiosSettings threw: $($_.Exception.Message)" }
        }
        $RawStatus = [string]$Result.Value
    }
    #Lenovo firmware quirk: empty/null Value on success - normalize to 'Success' (see function header).
    $Status = if ([string]::IsNullOrEmpty($RawStatus)) { 'Success' } else { $RawStatus }
    return @{ Status = $Status; Threw = $false; ExceptionMessage = '' }
}

function Get-LenovoSetFailReason
{
    #Translate a Lenovo write result into a FailReason persisted on the per-setting marker.

    param([Parameter(Mandatory = $true)][hashtable]$OpResult)
    if ($OpResult.Threw) { return 'write-threw' }
    switch -Regex ([string]$OpResult.Status)
    {
        'Access Denied'     { 'access-denied'; break }
        'Invalid Parameter' { 'invalid-param'; break }
        'Not Supported'     { 'not-supported'; break }
        default             { 'set-failed' }
    }
}

function Invoke-RemediateSettings
{
    #Apply the drifted settings: decrypt the unlock password (unless NoPassword), stage each via SetBIOSSetting, commit with one SaveBiosSettings, then write per-setting + top-level markers (applies at next reboot).

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Profile,
        [Parameter(Mandatory = $false)][AllowNull()][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $false)][ValidateRange(0, [int]::MaxValue)][int]$PasswordVersion = 0,
        [Parameter(Mandatory = $false)][switch]$NoPassword,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][hashtable[]]$SettingsToApply,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$MarkerBasePath,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$DesiredStateHash,
        [Parameter(Mandatory = $true)][int]$CompliantCount,
        [Parameter(Mandatory = $true)][int]$UnsupportedCount,
        [Parameter(Mandatory = $true)][int]$FailedCachedCount
    )
    $Password = $null
    $Applied = New-Object System.Collections.Generic.List[string]
    $Failed = New-Object System.Collections.Generic.List[string]
    try
    {
        if ($NoPassword)
        {
            $Password = ''
            Write-LogEntry -Value "NoPassword mode: applying settings unauthenticated (no opcode authorize)" -Severity 1
        }
        else
        {
            try
            {
                $Password = Get-CmsPlaintextFromPayload -Payload $Payload -Version $PasswordVersion
            }
            catch
            {
                Write-LogEntry -Value "CMS decrypt failed for pw v=$PasswordVersion`: $($_.Exception.Message)" -Severity 3
                return @{ Stdout = "FAILED: settings cms-decrypt-failed v=$PasswordVersion"; ExitCode = 1 }
            }
        }

        #Phase 1: stage each SetBIOSSetting. Capture per-setting result for use after Save + post-query.
        $WriteResults = @{}
        foreach ($Item in $SettingsToApply)
        {
            $Name = $Item.Name
            $DesiredValue = $Item.DesiredValue
            Write-LogEntry -Value "Staging setting: $Name = '$DesiredValue' (was '$($Item.CurrentValue)')" -Severity 1
            $Op = Invoke-LenovoSetBiosSetting -SettingName $Name -SettingValue $DesiredValue -Password $Password -NoPassword:$NoPassword
            $WriteResults[$Name] = $Op
            if ($Op.Threw)
            {
                Write-LogEntry -Value "SetBIOSSetting threw for $Name`: $($Op.ExceptionMessage)" -Severity 3
            }
            elseif ($Op.Status -ne 'Success')
            {
                Write-LogEntry -Value "SetBIOSSetting returned status='$($Op.Status)' for $Name" -Severity 2
            }
            else
            {
                Write-LogEntry -Value "SetBIOSSetting returned status=Success for $Name - staged for save" -Severity 1
            }
        }

        #Phase 2: SaveBiosSettings commits all staged changes atomically; non-Success loses the whole batch.
        $SaveOp = Invoke-LenovoSaveBiosSettings -Password $Password -NoPassword:$NoPassword
        if ($SaveOp.Threw)
        {
            Write-LogEntry -Value "SaveBiosSettings threw: $($SaveOp.ExceptionMessage)" -Severity 3
            foreach ($Item in $SettingsToApply)
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Item.Name -DesiredValue $Item.DesiredValue -LastVerifiedValue ([string]$Item.CurrentValue) -Status 'Failed' -FailReason 'save-threw'
                [void]$Failed.Add($Item.Name)
            }
            Set-SettingsMarker -BasePath $MarkerBasePath -Profile $Profile -DesiredStateHash $DesiredStateHash
            return @{
                Stdout   = "FAILED: settings save-threw profile=$Profile staged=$($SettingsToApply.Count) ($($SaveOp.ExceptionMessage -replace '[\r\n]+', ' '))"
                ExitCode = 1
            }
        }
        if ($SaveOp.Status -ne 'Success')
        {
            Write-LogEntry -Value "SaveBiosSettings returned status='$($SaveOp.Status)' - all staged settings lost" -Severity 3
            foreach ($Item in $SettingsToApply)
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Item.Name -DesiredValue $Item.DesiredValue -LastVerifiedValue ([string]$Item.CurrentValue) -Status 'Failed' -FailReason 'save-failed'
                [void]$Failed.Add($Item.Name)
            }
            Set-SettingsMarker -BasePath $MarkerBasePath -Profile $Profile -DesiredStateHash $DesiredStateHash
            return @{
                Stdout   = "FAILED: settings save-failed status=$($SaveOp.Status) profile=$Profile staged=$($SettingsToApply.Count)"
                ExitCode = 1
            }
        }
        Write-LogEntry -Value "SaveBiosSettings returned Success - staged settings committed; will apply at next reboot" -Severity 1

        #Phase 3: classify each queued setting and write its marker. Lenovo applies at next reboot (no readback);
        #SetBIOSSetting + SaveBiosSettings Success is authoritative, so mark Applied with LastVerifiedValue=Desired.
        foreach ($Item in $SettingsToApply)
        {
            $Name = $Item.Name
            $Desired = $Item.DesiredValue
            $OpResult = $WriteResults[$Name]
            $PreValue = [string]$Item.CurrentValue
            if ($OpResult.Threw)
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PreValue -Status 'Failed' -FailReason (Get-LenovoSetFailReason -OpResult $OpResult)
                [void]$Failed.Add($Name)
                continue
            }
            if ($OpResult.Status -ne 'Success')
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PreValue -Status 'Failed' -FailReason (Get-LenovoSetFailReason -OpResult $OpResult)
                [void]$Failed.Add($Name)
                continue
            }
            Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $Desired -Status 'Applied'
            [void]$Applied.Add($Name)
        }

        #Always refresh the top-level marker so Profile / LastFullRun / DesiredStateHash reflect the current run.
        Set-SettingsMarker -BasePath $MarkerBasePath -Profile $Profile -DesiredStateHash $DesiredStateHash

        if ($Failed.Count -eq 0)
        {
            return @{
                Stdout   = "REMEDIATED: settings profile=$Profile applied=$($Applied.Count) compliant=$CompliantCount unsupported=$UnsupportedCount failed-cached=$FailedCachedCount"
                ExitCode = 0
            }
        }
        $FailedNames = Format-DriftNames -Names ($Failed.ToArray())
        return @{
            Stdout   = "FAILED: settings profile=$Profile applied=$($Applied.Count) failed=$($Failed.Count) failed-names=$FailedNames"
            ExitCode = 1
        }
    }
    finally
    {
        if ($null -ne $Password) { $Password = $null }
    }
}

#Main =========================================================================================================================

#Configure the logging directory.
$LogsDirectory = ($LogFile | Split-Path)
if ([string]::IsNullOrEmpty($LogsDirectory))
{
    $LogsDirectory = $PSScriptRoot
}
elseif (-not (Test-Path -PathType Container $LogsDirectory))
{
    try
    {
        New-Item -Path $LogsDirectory -ItemType 'Directory' -Force -ErrorAction Stop -WhatIf:$false | Out-Null
    }
    catch
    {
        $LogsDirectory = $env:TEMP
    }
}

Write-LogEntry -Value "START - Lenovo BIOS settings remediation (Intune) v$Version" -Severity 1
Write-LogEntry -Value "Profile=$Profile  MarkerBasePath=$MarkerBasePath  PwMarkerPath=$PwMarkerPath  RetryFailedAfterDays=$RetryFailedAfterDays  NoPassword=$NoPassword  WhatIf=$WhatIfPreference" -Severity 1
Write-LogEntry -Value "Managed settings count: $($DesiredSettings.Count)" -Severity 1

#NoPassword mode tag.
$ModeNote = if ($NoPassword) { ' mode=nopassword' } else { '' }

#Manufacturer check.
if (-not $SkipManufacturerCheck)
{
    try
    {
        $Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).Manufacturer
    }
    catch
    {
        $Manufacturer = $null
        Write-LogEntry -Value "Win32_ComputerSystem query failed: $($_.Exception.Message)" -Severity 2
    }
    if (-not [string]::IsNullOrWhiteSpace($Manufacturer) -and $Manufacturer -notlike 'LENOVO*')
    {
        Write-LogEntry -Value "Manufacturer is '$Manufacturer' (not Lenovo) - script does not apply" -Severity 1
        Write-LogEntry -Value "END - Lenovo BIOS settings remediation (out of scope)" -Severity 1
        Write-Output 'SKIPPED: settings out-of-scope (not Lenovo)'
        exit 0
    }
    if ([string]::IsNullOrWhiteSpace($Manufacturer))
    {
        Write-LogEntry -Value "Manufacturer indeterminate - proceeding (BIOS calls will fail if not Lenovo)" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "Manufacturer = '$Manufacturer' - Lenovo hardware confirmed, proceeding" -Severity 1
    }
}

#DesiredSettings check.
if ($DesiredSettings.Count -eq 0)
{
    Write-LogEntry -Value '$DesiredSettings is empty - nothing to manage. Edit the script body before deploying.' -Severity 2
    Write-LogEntry -Value 'END - Lenovo BIOS settings remediation (no work)' -Severity 1
    Write-Output 'SKIPPED: settings no-managed-settings ($DesiredSettings empty - edit the script body)'
    exit 0
}

#Cert-auth short-circuit: signed payloads required, settings remediation does not apply. Exit 0 so Intune does not retry.
$CertAuth = Get-LenovoCertAuthEnabled
if ($CertAuth -eq $true)
{
    Write-LogEntry -Value "Lenovo cert-auth (PasswordState=128) is enabled - settings remediation does not apply. Skipping." -Severity 2
    Write-LogEntry -Value "END - Lenovo BIOS settings remediation (cert-auth incompatible)" -Severity 1
    Write-Output 'SKIPPED: settings cert-auth-incompatible (signed payloads required)'
    exit 0
}
if ($null -eq $CertAuth)
{
    Write-LogEntry -Value "Cert-auth query failed - proceeding (BIOS authorize may surface the failure)" -Severity 2
}

#NoPassword safety guard: refuse to write unauthenticated into a password-protected BIOS.
if ($NoPassword)
{
    $PwState = Get-LenovoPasswordState
    if ($null -eq $PwState)
    {
        Write-LogEntry -Value "NoPassword: PasswordState query failed - cannot confirm absence of a BIOS password; proceeding" -Severity 2
    }
    elseif ($PwState -ne 0 -and $PwState -ne 128)
    {
        Write-LogEntry -Value "NoPassword mode but PasswordState=$PwState (a BIOS password is set). Lenovo unauthenticated writes return Success but silently do not apply - refusing to proceed." -Severity 3
        Write-Output "FAILED: settings nopassword-password-set (BIOS password present; unauthenticated writes silently fail)$ModeNote"
        exit 1
    }
    else
    {
        Write-LogEntry -Value "NoPassword: PasswordState=$PwState - no blocking BIOS password, safe to write unauthenticated" -Severity 1
    }
}

#Password checks.
if (-not $NoPassword)
{
    #Payload check.
    if ($null -eq $IntunePayload -or $null -eq $IntunePayload.Files -or $IntunePayload.Files.Count -eq 0)
    {
        Write-LogEntry -Value '$IntunePayload is null or empty. Build-IntunePayload.ps1 was not run before deployment.' -Severity 3
        Write-Output 'FAILED: settings payload-not-built ($IntunePayload is null - run Tools\Build-IntunePayload.ps1)'
        exit 1
    }

    #Installed cert check.
    $CertCheck = Test-RemediationCert -Thumbprint $IntunePayload.CertThumbprint
    if (-not $CertCheck.Ok)
    {
        Write-LogEntry -Value "Cert check failed: $($CertCheck.Reason)" -Severity 3
        Write-Output "FAILED: settings $($CertCheck.Reason)"
        exit 1
    }
    Write-LogEntry -Value "Cert thumbprint=$($IntunePayload.CertThumbprint) found in LocalMachine\My with usable private key" -Severity 1
}

#Password marker check.
$PwMarker = Get-PasswordMarker -BasePath $PwMarkerPath
$PwMarkerPresent = ($null -ne $PwMarker)
$PwMarkerVersion = if ($PwMarkerPresent) { [int]$PwMarker.Version } else { 0 }
if ($PwMarkerPresent)
{
    Write-LogEntry -Value "PW marker: Version=$PwMarkerVersion SetDate=$($PwMarker.SetDate) CertThumbprint=$($PwMarker.CertThumbprint)" -Severity 1
}
else
{
    Write-LogEntry -Value "PW marker not present at $PwMarkerPath" -Severity 2
}

#Settings marker check.
$TopLevelMarker = Get-SettingsMarker -BasePath $MarkerBasePath
if ($null -eq $TopLevelMarker)
{
    Write-LogEntry -Value "Top-level settings marker not present at $MarkerBasePath" -Severity 1
}
else
{
    Write-LogEntry -Value "Top-level marker: Profile=$($TopLevelMarker.Profile) LastFullRun=$($TopLevelMarker.LastFullRun) DesiredStateHash=$($TopLevelMarker.DesiredStateHash)" -Severity 1
}

#Per-Setting marker check.
$PerSettingMarkers = Get-AllSettingMarkers -BasePath $MarkerBasePath
Write-LogEntry -Value "Per-setting markers found: $($PerSettingMarkers.Count)" -Severity 1

#Get current BIOS settings.
$CurrentBiosValues = Get-LenovoBiosSettings
if ($null -eq $CurrentBiosValues)
{
    Write-LogEntry -Value "BIOS setting query returned null" -Severity 2
}
else
{
    Write-LogEntry -Value "BIOS setting query returned $($CurrentBiosValues.Count) settings" -Severity 1
}

$PayloadVersions = @(if (-not $NoPassword) { $IntunePayload.Files.Keys | ForEach-Object { [int]$_ } | Sort-Object })
Write-LogEntry -Value "Payload versions available: $($PayloadVersions -join ', ')" -Severity 1

$DesiredStateHash = Get-DesiredStateHash -DesiredSettings $DesiredSettings
Write-LogEntry -Value "DesiredStateHash=$DesiredStateHash" -Severity 1

$Plan = Get-SettingsRemediationPlan `
    -DesiredSettings $DesiredSettings `
    -CurrentBiosValues $CurrentBiosValues `
    -PerSettingMarkers $PerSettingMarkers `
    -PasswordMarkerPresent $PwMarkerPresent `
    -PasswordMarkerVersion $PwMarkerVersion `
    -PayloadVersions $PayloadVersions `
    -NoPassword $NoPassword `
    -RetryFailedAfterDays $RetryFailedAfterDays

Write-LogEntry -Value "Plan: Verdict=$($Plan.Verdict)  Set=$($Plan.SetCount)  Compliant=$($Plan.CompliantCount)  Unsupported=$($Plan.UnsupportedCount)  FailedCached=$($Plan.FailedCachedCount)" -Severity 1
if (-not [string]::IsNullOrEmpty($Plan.Reason))
{
    Write-LogEntry -Value "Plan reason: $($Plan.Reason)" -Severity 1
}

if ($Plan.Verdict -eq 'EarlyExitFail')
{
    Write-LogEntry -Value "Outcome: $($Plan.EarlyExitStdout)" -Severity 3
    Write-LogEntry -Value "END - Lenovo BIOS settings remediation (exit $($Plan.EarlyExitCode))" -Severity 1
    Write-Output "$($Plan.EarlyExitStdout)$ModeNote"
    exit $Plan.EarlyExitCode
}

#Filter to settings needing a write; handle the already-compliant case.
$SettingsToApply = @($Plan.PerSettingPlan | Where-Object { $_.Action -eq 'Set' })
if ($SettingsToApply.Count -eq 0)
{
    Set-SettingsMarker -BasePath $MarkerBasePath -Profile $Profile -DesiredStateHash $DesiredStateHash
    $Stdout = "SKIPPED: settings profile=$Profile already-compliant (managed=$($DesiredSettings.Count) compliant=$($Plan.CompliantCount) unsupported=$($Plan.UnsupportedCount) failed-cached=$($Plan.FailedCachedCount))$ModeNote"
    Write-LogEntry -Value "Outcome: $Stdout" -Severity 1
    Write-LogEntry -Value "END - Lenovo BIOS settings remediation (exit 0, no drift)" -Severity 1
    Write-Output $Stdout
    exit 0
}

Write-LogEntry -Value "$($SettingsToApply.Count) setting(s) to apply: $((($SettingsToApply | ForEach-Object { $_.Name }) -join ', '))" -Severity 1

$Stdout = ''
$ExitCode = 1

#Remediate settings.
try
{
    $r = Invoke-RemediateSettings `
        -Profile $Profile `
        -Payload $IntunePayload `
        -PasswordVersion $PwMarkerVersion `
        -NoPassword:$NoPassword `
        -SettingsToApply $SettingsToApply `
        -MarkerBasePath $MarkerBasePath `
        -DesiredStateHash $DesiredStateHash `
        -CompliantCount $Plan.CompliantCount `
        -UnsupportedCount $Plan.UnsupportedCount `
        -FailedCachedCount $Plan.FailedCachedCount

    $Stdout = $r.Stdout
    $ExitCode = $r.ExitCode
}
catch
{
    Write-LogEntry -Value "Unhandled exception during remediation dispatch: $($_.Exception.Message)" -Severity 3
    $Stdout = "FAILED: settings internal-error ($($_.Exception.Message -replace '[\r\n]+', ' '))"
    $ExitCode = 1
}

$Stdout = "$Stdout$ModeNote"
Write-LogEntry -Value "Outcome: $Stdout" -Severity 1
Write-LogEntry -Value "END - Lenovo BIOS settings remediation (exit $ExitCode)" -Severity 1
#Write STDOUT for Intune reporting.
Write-Output $Stdout
exit $ExitCode

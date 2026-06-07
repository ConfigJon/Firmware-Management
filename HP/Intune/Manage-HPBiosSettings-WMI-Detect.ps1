<#
    .DESCRIPTION
        Intune detection script for HP BIOS settings management. Compares each setting in the
        admin-edited $DesiredSettings hashtable to the BIOS-reported current value and exits 0 (compliant)
        or 1 (non-compliant / degraded). Detection only reads the BIOS (HP settings reads are
        unauthenticated). Per-setting Unsupported and cached-Failed states are excluded from the drift
        count so the dashboard does not show permanent non-compliance on hardware that cannot accept a
        setting. Skips on non-HP hardware reporting (COMPLIANT: settings out-of-scope). When HP Sure Admin
        (Enhanced BIOS Authentication Mode) is enabled the device is reported NONCOMPLIANT INCOMPATIBLE.

    .PARAMETER Profile
        Name of the desired-state profile, stamped into the marker. Must match the value in the paired script.

    .PARAMETER RetryFailedAfterDays
        Days after which a cached-Failed setting is re-surfaced as drift rather than skipped. 0 (default) = never.

    .PARAMETER NoPassword
        For devices with NO BIOS setup password. Default OFF. When set, the password-marker dependency is dropped
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

[CmdletBinding(PositionalBinding = $false)]
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

    #Internal: reporting emit-on-change marker root.
    [Parameter(DontShow)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportingMarkerPath = 'HKLM:\SOFTWARE\ConfigJonScripts\FirmwareManagement\Reporting',

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$RetryFailedAfterDays = 0,

    #Opt-in: BIOS has no setup password. Settings are read unauthenticated. Must match the paired -Remediate script.
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
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Manage-HPBiosSettings-WMI-Detect.log",

    [Parameter(DontShow)]
    [switch]$SkipManufacturerCheck
)

$Version = '1.0.0'
$Component = 'Manage-HPBiosSettings-WMI-Detect'

#Desired state ===============================================================================================================
#Edit this hashtable to match your desired HP BIOS configuration.
#Names match the HP_BIOSSetting Name exactly.
#Values match the BIOS current value exactly. ('Enable' / 'Disable').
#Use the existing Manage-HPBiosSettings-WMI.ps1 GetSettings mode to list what the device exposes.

$DesiredSettings = @{
    # Examples - replace with the settings your devices should standardize on:
    # 'Wake On LAN'                                        = 'Boot to Hard Drive'
    # 'Virtualization Technology (VTx)'                    = 'Enable'
    # 'Fingerprint Device'                                 = 'Disable'
    # 'Prompt for Admin password on F9 (Boot Menu)'        = 'Enable'
    # 'Prompt for Admin password on F11 (System Recovery)' = 'Enable'
    # 'Prompt for Admin password on F12 (Network Boot)'    = 'Enable'
}

#Reporting (Log Analytics) ===================================================================================================
#Optional: push each detection result to a Log Analytics workspace via the Logs Ingestion API.
#Disabled by default. Set $ReportingEnabled = $true and fill in $ReportingConfig to enable.
#A reporting failure does not affect the compliance result or exit code.
#The client-auth certificate's private key must be in Cert:\LocalMachine\My (readable by SYSTEM)
#Data is only sent when the reportable state changes, or once per HeartbeatDays.
#All values below are non-secret.
#Full setup, data model, and KQL: https://www.configjon.com/intune-bios-reporting/
$ReportingEnabled       = $false
$ReportingDepth         = 'Managed'   # Summary | Managed | Full
$ReportingHeartbeatDays = 7
$ReportingConfig = @{
    TenantId              = ''
    ClientId              = ''
    CertThumbprint        = ''
    DceUri                = ''
    SummaryDcrImmutableId = ''
    SummaryStream         = 'Custom-BiosManagementRun_CL'
    DetailDcrImmutableId  = ''
    DetailStream          = 'Custom-BiosSettingsDetail_CL'
}

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
        Out-File -InputObject $LogText -Append -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
    }
    catch [System.Exception]
    {
        Write-Warning -Message "Unable to append log entry to $FileName file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
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

function Test-PwMarkerPresent
{
    #Return whether the password marker exists.

    param([Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Path)
    return (Test-Path -LiteralPath $Path)
}

function Get-HPSureAdminEnrolled
{
    #Query HP_BIOSSetting for Enhanced BIOS Authentication Mode. Returns $true if Sure Admin is enabled.

    try
    {
        $Settings = Get-CimInstance -Namespace 'root\hp\InstrumentedBIOS' -ClassName 'HP_BIOSSetting' -ErrorAction Stop
    }
    catch
    {
        Write-LogEntry -Value "HP_BIOSSetting WMI query (Sure Admin) failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    $SureAdmin = $Settings | Where-Object { $_.Name -eq 'Enhanced BIOS Authentication Mode' } | Select-Object -First 1
    if ($null -eq $SureAdmin -or [string]::IsNullOrEmpty($SureAdmin.Value))
    {
        return $false
    }
    foreach ($Candidate in ($SureAdmin.Value -split ',' | ForEach-Object { $_.Trim() }))
    {
        if ($Candidate.StartsWith('*'))
        {
            return ($Candidate.Substring(1) -eq 'Enable')
        }
    }
    return $false
}

function Get-HPBiosSettingCurrentValue
{
    #Parse an HP_BIOSSetting Value string and return the current value.

    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$RawValue)
    if ([string]::IsNullOrEmpty($RawValue)) { return '' }
    foreach ($Candidate in ($RawValue -split ',' | ForEach-Object { $_.Trim() }))
    {
        if ($Candidate.StartsWith('*'))
        {
            return $Candidate.Substring(1)
        }
    }
    return $RawValue.Trim()
}

function Get-HPBiosSettings
{
    #Read all HP BIOS settings via HP_BIOSSetting and return as Name -> current value. Excludes the password settings and empty-named noise rows.

    try
    {
        $Items = Get-CimInstance -Namespace 'root\hp\InstrumentedBIOS' -ClassName 'HP_BIOSSetting' -ErrorAction Stop
    }
    catch
    {
        Write-LogEntry -Value "HP_BIOSSetting WMI query failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    $Result = @{}
    foreach ($Item in $Items)
    {
        if ([string]::IsNullOrWhiteSpace($Item.Name)) { continue }
        if ($Item.Name -eq 'Setup Password' -or $Item.Name -eq 'Power-On Password') { continue }
        $Result[$Item.Name] = Get-HPBiosSettingCurrentValue -RawValue ([string]$Item.Value)
    }
    return $Result
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

function Get-SettingsClassification
{
    #Given the desired state + observed BIOS state + markers, return the verdict.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Profile,
        [Parameter(Mandatory = $true)][hashtable]$DesiredSettings,
        [Parameter(Mandatory = $false)]$CurrentBiosValues, # hashtable Name->Value, or $null if query failed
        [Parameter(Mandatory = $false)]$TopLevelMarker, # hashtable or $null
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$PerSettingMarkers,
        [Parameter(Mandatory = $true)][bool]$PasswordMarkerPresent,
        [Parameter(Mandatory = $false)][bool]$NoPassword = $false,
        [Parameter(Mandatory = $false)][AllowNull()][object]$SureAdmin, # $true/$false/$null (query failed)
        [Parameter(Mandatory = $false)][ValidateRange(0, [int]::MaxValue)][int]$RetryFailedAfterDays = 0,
        [Parameter(Mandatory = $false)][datetime]$Now = (Get-Date)
    )

    #NONCOMPLIANT INCOMPATIBLE - Sure Admin enrolled, settings cannot be remediated via password.
    if ($SureAdmin -eq $true)
    {
        return @{
            Status       = 'NONCOMPLIANT'
            Stdout       = 'NONCOMPLIANT: settings sure-admin-enrolled (INCOMPATIBLE - signed payloads required)'
            Reason       = 'HP Sure Admin (Enhanced BIOS Authentication Mode) is enabled - settings remediation does not apply'
            DriftedNames = @()
        }
    }

    #DEGRADED - pw marker missing (settings remediation cannot unlock)
    if (-not $NoPassword -and -not $PasswordMarkerPresent)
    {
        return @{
            Status = 'DEGRADED'
            Stdout = 'DEGRADED: settings pw-marker-missing (cannot remediate without password)'
            Reason = 'pw marker missing - settings remediation has no way to unlock the BIOS'
            DriftedNames = @()
        }
    }

    #DEGRADED - BIOS query failed
    if ($null -eq $CurrentBiosValues)
    {
        return @{
            Status = 'DEGRADED'
            Stdout = 'DEGRADED: settings bios-query-failed (HP BIOS setting namespace unreachable)'
            Reason = 'CIM query for HP_BIOSSetting failed'
            DriftedNames = @()
        }
    }

    #NONCOMPLIANT - profile mismatch (admin redeployed with a new profile name)
    if ($null -ne $TopLevelMarker -and -not [string]::IsNullOrEmpty($TopLevelMarker.Profile) -and $TopLevelMarker.Profile -ne $Profile)
    {
        return @{
            Status = 'NONCOMPLIANT'
            Stdout = "NONCOMPLIANT: settings profile-mismatch expected=$Profile actual=$($TopLevelMarker.Profile)"
            Reason = "marker says profile=$($TopLevelMarker.Profile) but script targets $Profile"
            DriftedNames = @()
        }
    }

    #Per-setting classification
    #SettingDetail collects the per-setting breakdown (name/desired/current/status) for the reporting feature.
    $DriftedNames = New-Object System.Collections.Generic.List[string]
    $SettingDetail = New-Object System.Collections.Generic.List[object]
    $UnsupportedCount = 0
    $FailedCount = 0
    $BlockedCount = 0
    $BlockedReasons = New-Object System.Collections.Generic.List[string]
    $BlockedReasonSet = @('access-denied', 'write-threw', 'save-threw', 'save-failed', 'query-failed')
    $CompliantCount = 0

    foreach ($Name in ($DesiredSettings.Keys | Sort-Object))
    {
        $DesiredValue = [string]$DesiredSettings[$Name]
        $Current = $CurrentBiosValues[$Name]
        $Marker = $PerSettingMarkers[$Name]
        $MarkerSetDate = if ($null -ne $Marker) { [string]$Marker.SetDate } else { '' }

        if ($null -eq $Current)
        {
            $UnsupportedCount++
            [void]$SettingDetail.Add(@{ SettingName = $Name; DesiredValue = $DesiredValue; CurrentValue = ''; PerSettingStatus = 'Unsupported'; Managed = $true; SetDate = $MarkerSetDate })
            continue
        }

        $Current = [string]$Current

        #Skip cached-Failed settings (Remediate already tried + failed).
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
                $FailedCount++
                if ($BlockedReasonSet -contains $Marker.FailReason)
                {
                    $BlockedCount++
                    if (-not $BlockedReasons.Contains($Marker.FailReason)) { [void]$BlockedReasons.Add($Marker.FailReason) }
                }
                [void]$SettingDetail.Add(@{ SettingName = $Name; DesiredValue = $DesiredValue; CurrentValue = $Current; PerSettingStatus = 'Failed'; FailReason = $Marker.FailReason; Managed = $true; SetDate = $MarkerSetDate })
                continue
            }
        }

        if ($Current -eq $DesiredValue)
        {
            $CompliantCount++
            [void]$SettingDetail.Add(@{ SettingName = $Name; DesiredValue = $DesiredValue; CurrentValue = $Current; PerSettingStatus = 'Compliant'; Managed = $true; SetDate = $MarkerSetDate })
        }
        else
        {
            [void]$DriftedNames.Add($Name)
            [void]$SettingDetail.Add(@{ SettingName = $Name; DesiredValue = $DesiredValue; CurrentValue = $Current; PerSettingStatus = 'Drift'; Managed = $true; SetDate = $MarkerSetDate })
        }
    }

    $Managed = $DesiredSettings.Count - $UnsupportedCount
    $DriftCount = $DriftedNames.Count

    #Surface blocked failures in STDOUT.
    $BlockedReasonList = ($BlockedReasons | Sort-Object) -join ','
    $BlockedNote = if ($BlockedCount -gt 0) { " blocked=$BlockedCount blocked-reasons=$BlockedReasonList" } else { '' }

    if ($DriftCount -eq 0)
    {
        return @{
            Status           = 'COMPLIANT'
            Stdout           = "COMPLIANT: settings profile=$Profile managed=$Managed unsupported=$UnsupportedCount failed=$FailedCount drift=0$BlockedNote"
            Reason           = 'all managed non-failed settings match desired'
            DriftedNames     = @()
            ManagedCount     = $Managed
            CompliantCount   = $CompliantCount
            DriftCount       = 0
            UnsupportedCount = $UnsupportedCount
            FailedCount      = $FailedCount
            BlockedCount     = $BlockedCount
            BlockedReasons   = $BlockedReasons.ToArray()
            SettingDetail    = $SettingDetail.ToArray()
        }
    }

    $NameList = Format-DriftNames -Names ($DriftedNames.ToArray())
    return @{
        Status           = 'NONCOMPLIANT'
        Stdout           = "NONCOMPLIANT: settings profile=$Profile managed=$Managed unsupported=$UnsupportedCount failed=$FailedCount drift=$DriftCount names=$NameList$BlockedNote"
        Reason           = "$DriftCount setting(s) drifted from desired state"
        DriftedNames     = $DriftedNames.ToArray()
        ManagedCount     = $Managed
        CompliantCount   = $CompliantCount
        DriftCount       = $DriftCount
        UnsupportedCount = $UnsupportedCount
        FailedCount      = $FailedCount
        BlockedCount     = $BlockedCount
        BlockedReasons   = $BlockedReasons.ToArray()
        SettingDetail    = $SettingDetail.ToArray()
    }
}

#Reporting functions ========================================================================================================
#Shared, vendor-agnostic Log Analytics reporting

function Get-DeviceFacts
{
    #Device identity: computer name, SMBIOS UUID as DeviceId, manufacturer/model.

    $Name = $env:COMPUTERNAME
    $Id = ''; $Manufacturer = ''; $Model = ''
    try
    {
        $Cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $Manufacturer = [string]$Cs.Manufacturer
        $Model = [string]$Cs.Model
    }
    catch { }
    try { $Id = [string](Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction Stop).UUID } catch { }
    return @{ DeviceName = $Name; DeviceId = $Id; Manufacturer = $Manufacturer; Model = $Model }
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

function Get-BiosReportPayload
{
    #Assemble the summary record and per-setting detail rows.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Settings', 'Password')][string]$Component,
        [Parameter(Mandatory = $true)][ValidateSet('Summary', 'Managed', 'Full')][string]$Depth,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RunId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TimeGenerated,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DeviceId,
        [string]$DeviceName = '',
        [string]$Manufacturer = '',
        [string]$Model = '',
        [string]$Profile = '',
        [string]$DesiredStateHash = '',
        [string]$ScriptVersion = '',
        [ValidateSet('normal', 'nopassword')][string]$Mode = 'normal',
        [string]$OverallTag = '',
        [int]$ManagedCount = 0,
        [int]$CompliantCount = 0,
        [int]$DriftCount = 0,
        [int]$UnsupportedCount = 0,
        [int]$FailedCount = 0,
        [int]$BlockedCount = 0,
        [string[]]$DriftNames = @(),
        [int]$PasswordVersion = 0,
        [string]$CertThumbprint = '',
        [AllowEmptyCollection()][hashtable[]]$SettingDetail = @()
    )

    $Summary = [ordered]@{
        TimeGenerated    = $TimeGenerated
        DeviceName       = $DeviceName
        DeviceId         = $DeviceId
        Manufacturer     = $Manufacturer
        Model            = $Model
        Component        = $Component
        Profile          = $Profile
        DesiredStateHash = $DesiredStateHash
        ScriptVersion    = $ScriptVersion
        Mode             = $Mode
        OverallTag       = $OverallTag
        ManagedCount     = $ManagedCount
        CompliantCount   = $CompliantCount
        DriftCount       = $DriftCount
        UnsupportedCount = $UnsupportedCount
        FailedCount      = $FailedCount
        BlockedCount     = $BlockedCount
        DriftNames       = (($DriftNames | Sort-Object) -join ',')
        PasswordVersion  = $PasswordVersion
        CertThumbprint   = $CertThumbprint
        RunId            = $RunId
    }

    $Detail = New-Object System.Collections.Generic.List[object]
    if ($Depth -ne 'Summary')
    {
        foreach ($Item in ($SettingDetail | Sort-Object { [string]$_.SettingName }))
        {
            if ($Depth -ne 'Full' -and -not [bool]$Item.Managed) { continue }
            [void]$Detail.Add([ordered]@{
                    TimeGenerated    = $TimeGenerated
                    DeviceId         = $DeviceId
                    RunId            = $RunId
                    SettingName      = [string]$Item.SettingName
                    DesiredValue     = [string]$Item.DesiredValue
                    CurrentValue     = [string]$Item.CurrentValue
                    PerSettingStatus = [string]$Item.PerSettingStatus
                    FailReason       = [string]$Item.FailReason
                    Managed          = [bool]$Item.Managed
                    SetDate          = [string]$Item.SetDate
                })
        }
    }

    return [ordered]@{ Summary = $Summary; Detail = $Detail.ToArray() }
}

function Get-BiosReportStateHash
{
    #SHA-256 over the payload's stable fields (excludes TimeGenerated, RunId, SetDate) so the hash changes only when something worth re-reporting changes.

    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateNotNull()][System.Collections.IDictionary]$Payload)
    $S = $Payload.Summary
    $Builder = [System.Text.StringBuilder]::new()
    [void]$Builder.AppendLine("C=$($S.Component)|P=$($S.Profile)|H=$($S.DesiredStateHash)|M=$($S.Mode)|T=$($S.OverallTag)")
    [void]$Builder.AppendLine("mc=$($S.ManagedCount)|cc=$($S.CompliantCount)|dc=$($S.DriftCount)|uc=$($S.UnsupportedCount)|fc=$($S.FailedCount)|bk=$($S.BlockedCount)")
    [void]$Builder.AppendLine("dn=$($S.DriftNames)|pv=$($S.PasswordVersion)|ct=$($S.CertThumbprint)")
    [void]$Builder.AppendLine("mf=$($S.Manufacturer)|md=$($S.Model)|sv=$($S.ScriptVersion)")
    foreach ($D in ($Payload.Detail | Sort-Object { [string]$_.SettingName }))
    {
        [void]$Builder.AppendLine("$($D.SettingName)=$($D.DesiredValue)|$($D.CurrentValue)|$($D.PerSettingStatus)|$($D.FailReason)|$($D.Managed)")
    }
    $Sha = [System.Security.Cryptography.SHA256]::Create()
    try { $HashBytes = $Sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Builder.ToString())) }
    finally { $Sha.Dispose() }
    return (([System.BitConverter]::ToString($HashBytes)) -replace '-', '').ToLowerInvariant()
}

function Test-ShouldEmitReport
{
    #Emit when: never emitted, hash changed, recorded date unparseable, or last emit older than the heartbeat.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$CurrentHash,
        [AllowEmptyString()][string]$LastHash = '',
        [AllowEmptyString()][string]$LastEmitDate = '',
        [ValidateRange(0, [int]::MaxValue)][int]$HeartbeatDays = 7,
        [datetime]$Now = (Get-Date)
    )
    if ([string]::IsNullOrEmpty($LastHash)) { return @{ Emit = $true; Reason = 'never-emitted' } }
    if ($CurrentHash -ne $LastHash) { return @{ Emit = $true; Reason = 'state-changed' } }
    if ($HeartbeatDays -gt 0)
    {
        $Last = [datetime]::MinValue
        if (-not [datetime]::TryParse($LastEmitDate, [ref]$Last)) { return @{ Emit = $true; Reason = 'bad-last-date' } }
        if (($Now - $Last).TotalDays -ge $HeartbeatDays) { return @{ Emit = $true; Reason = 'heartbeat' } }
    }
    return @{ Emit = $false; Reason = 'unchanged' }
}

function Get-ReportingMarker
{
    #Read the per-Component emit-on-change marker. Empty strings when absent (first run always emits).

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath,
        [Parameter(Mandatory = $true)][ValidateSet('Settings', 'Password')][string]$Component
    )
    $Path = Join-Path -Path $BasePath -ChildPath $Component
    if (-not (Test-Path -LiteralPath $Path)) { return @{ LastEmittedHash = ''; LastEmittedDate = '' } }
    try { $Item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop } catch { return @{ LastEmittedHash = ''; LastEmittedDate = '' } }
    return @{ LastEmittedHash = [string]$Item.LastEmittedHash; LastEmittedDate = [string]$Item.LastEmittedDate }
}

function Set-ReportingMarker
{
    #Write the per-Component emit-on-change marker.

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath,
        [Parameter(Mandatory = $true)][ValidateSet('Settings', 'Password')][string]$Component,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Hash,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Date
    )
    if ($WhatIfPreference) { return }
    $Path = Join-Path -Path $BasePath -ChildPath $Component
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -LiteralPath $Path -Name 'LastEmittedHash' -Value $Hash -Type String -Force
    Set-ItemProperty -LiteralPath $Path -Name 'LastEmittedDate' -Value $Date -Type String -Force
}

function ConvertTo-Base64Url
{
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)
    ([Convert]::ToBase64String($Bytes)).TrimEnd('=') -replace '\+', '-' -replace '/', '_'
}

function ConvertTo-BiosReportJson
{
    #Serialize records to a JSON array.

    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Records)
    if ($Records.Count -eq 0) { return '[]' }
    $Json = ConvertTo-Json -InputObject $Records -Depth 10 -Compress
    if (-not $Json.StartsWith('[')) { $Json = "[$Json]" }
    return $Json
}

function Get-BiosReportAuthToken
{
    #Cert-based client-credentials token for the Logs Ingestion scope. Builds + signs the client-assertion JWT with the cert private key (RS256/PKCS1, x5t header). Returns @{ Ok; Token; Reason }.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ClientId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$CertThumbprint,
        [string]$CertStoreLocation = 'Cert:\LocalMachine\My',
        [int]$TimeoutSec = 30
    )
    try
    {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
        $CertPath = Join-Path -Path $CertStoreLocation -ChildPath $CertThumbprint
        if (-not (Test-Path -LiteralPath $CertPath)) { return @{ Ok = $false; Token = $null; Reason = "cert-not-found thumbprint=$CertThumbprint" } }
        $Cert = Get-Item -LiteralPath $CertPath
        $Rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Cert)
        if ($null -eq $Rsa) { return @{ Ok = $false; Token = $null; Reason = 'cert-no-private-key (or not readable by current identity)' } }

        $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $X5t = ConvertTo-Base64Url ($Cert.GetCertHash())
        $Header = @{ alg = 'RS256'; typ = 'JWT'; x5t = $X5t } | ConvertTo-Json -Compress
        $Now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
        $Claims = @{ aud = $TokenEndpoint; iss = $ClientId; sub = $ClientId; jti = [guid]::NewGuid().ToString(); nbf = $Now; exp = $Now + 300; iat = $Now } | ConvertTo-Json -Compress
        $SigningInput = (ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($Header))) + '.' + (ConvertTo-Base64Url ([Text.Encoding]::UTF8.GetBytes($Claims)))
        $Signature = $Rsa.SignData([Text.Encoding]::ASCII.GetBytes($SigningInput), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $Assertion = "$SigningInput." + (ConvertTo-Base64Url $Signature)

        $Body = @{
            grant_type            = 'client_credentials'
            client_id             = $ClientId
            scope                 = 'https://monitor.azure.com/.default'
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $Assertion
        }
        try { $Response = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -Body $Body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec $TimeoutSec -ErrorAction Stop }
        catch { return @{ Ok = $false; Token = $null; Reason = "token-request-failed: $($_.Exception.Message)" } }
        if ([string]::IsNullOrEmpty($Response.access_token)) { return @{ Ok = $false; Token = $null; Reason = 'token-response-missing-access_token' } }
        return @{ Ok = $true; Token = $Response.access_token; Reason = '' }
    }
    catch { return @{ Ok = $false; Token = $null; Reason = "token-unhandled: $($_.Exception.Message)" } }
}

function Send-BiosReportRecords
{
    #POST a record array to one Logs Ingestion stream. Returns @{ Ok; Reason; Count }.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Token,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DceUri,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DcrImmutableId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$StreamName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Records,
        [string]$ApiVersion = '2023-01-01',
        [int]$TimeoutSec = 30
    )
    if ($Records.Count -eq 0) { return @{ Ok = $true; Reason = 'no-records'; Count = 0 } }
    $Uri = "$($DceUri.TrimEnd('/'))/dataCollectionRules/$DcrImmutableId/streams/$StreamName`?api-version=$ApiVersion"
    $Json = ConvertTo-BiosReportJson -Records $Records
    $Headers = @{ Authorization = "Bearer $Token" }
    try { Invoke-RestMethod -Method Post -Uri $Uri -Headers $Headers -Body $Json -ContentType 'application/json' -TimeoutSec $TimeoutSec -ErrorAction Stop | Out-Null }
    catch
    {
        $Code = $null
        try { $Code = [int]$_.Exception.Response.StatusCode } catch { }
        $Reason = if ($Code) { "ingest-post-failed http=${Code}: $($_.Exception.Message)" } else { "ingest-post-failed: $($_.Exception.Message)" }
        return @{ Ok = $false; Reason = $Reason; Count = $Records.Count }
    }
    return @{ Ok = $true; Reason = ''; Count = $Records.Count }
}

function Invoke-BiosReportEmit
{
    #Top-level emit: one token, then POST summary + (if any) detail.

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNull()][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)][ValidateNotNull()][hashtable]$Config
    )
    try
    {
        $Auth = Get-BiosReportAuthToken -TenantId $Config.TenantId -ClientId $Config.ClientId -CertThumbprint $Config.CertThumbprint
        if (-not $Auth.Ok) { return @{ Ok = $false; Reason = $Auth.Reason } }

        $SummaryResult = Send-BiosReportRecords -Token $Auth.Token -DceUri $Config.DceUri -DcrImmutableId $Config.SummaryDcrImmutableId -StreamName $Config.SummaryStream -Records @($Payload.Summary)
        if (-not $SummaryResult.Ok) { return @{ Ok = $false; Reason = "summary $($SummaryResult.Reason)" } }

        $DetailResult = @{ Ok = $true; Count = 0 }
        if ($Payload.Detail.Count -gt 0)
        {
            $DetailResult = Send-BiosReportRecords -Token $Auth.Token -DceUri $Config.DceUri -DcrImmutableId $Config.DetailDcrImmutableId -StreamName $Config.DetailStream -Records $Payload.Detail
            if (-not $DetailResult.Ok) { return @{ Ok = $false; Reason = "detail $($DetailResult.Reason)"; SummarySent = $true } }
        }
        return @{ Ok = $true; Reason = ''; SummaryCount = $SummaryResult.Count; DetailCount = $DetailResult.Count }
    }
    catch { return @{ Ok = $false; Reason = "emit-unhandled: $($_.Exception.Message)" } }
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
        New-Item -Path $LogsDirectory -ItemType 'Directory' -Force -ErrorAction Stop | Out-Null
    }
    catch
    {
        $LogsDirectory = $env:TEMP
    }
}

Write-LogEntry -Value "START - HP BIOS settings detection (Intune) v$Version" -Severity 1
Write-LogEntry -Value "Profile=$Profile  MarkerBasePath=$MarkerBasePath  PwMarkerPath=$PwMarkerPath  RetryFailedAfterDays=$RetryFailedAfterDays  NoPassword=$NoPassword  Reporting=$ReportingEnabled" -Severity 1
Write-LogEntry -Value "Managed settings count: $($DesiredSettings.Count)" -Severity 1

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
    if (-not [string]::IsNullOrWhiteSpace($Manufacturer) -and $Manufacturer -notlike 'HP*' -and $Manufacturer -notlike 'Hewlett*')
    {
        Write-LogEntry -Value "Manufacturer is '$Manufacturer' (not HP) - script does not apply" -Severity 1
        Write-LogEntry -Value "END - HP BIOS settings detection (out of scope)" -Severity 1
        Write-Output 'COMPLIANT: settings out-of-scope (not HP)'
        exit 0
    }
    if ([string]::IsNullOrWhiteSpace($Manufacturer))
    {
        Write-LogEntry -Value "Manufacturer indeterminate - proceeding with normal detection" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "Manufacturer = '$Manufacturer' - HP hardware confirmed, proceeding" -Severity 1
    }
}

#DesiredSettings check.
if ($DesiredSettings.Count -eq 0)
{
    Write-LogEntry -Value '$DesiredSettings is empty - nothing to manage. Edit the script body before deploying.' -Severity 2
    Write-LogEntry -Value 'END - HP BIOS settings detection (no work)' -Severity 1
    Write-Output "COMPLIANT: settings profile=$Profile managed=0 (no settings configured in script body)"
    exit 0
}

#Password detection.
try
{
    $PwMarkerPresent = Test-PwMarkerPresent -Path $PwMarkerPath
    Write-LogEntry -Value "Password marker present: $PwMarkerPresent (path=$PwMarkerPath)" -Severity 1

    $TopLevelMarker = Get-SettingsMarker -BasePath $MarkerBasePath
    if ($null -eq $TopLevelMarker)
    {
        Write-LogEntry -Value "Top-level settings marker not present at $MarkerBasePath" -Severity 1
    }
    else
    {
        Write-LogEntry -Value "Top-level marker: Profile=$($TopLevelMarker.Profile) LastFullRun=$($TopLevelMarker.LastFullRun) DesiredStateHash=$($TopLevelMarker.DesiredStateHash)" -Severity 1
    }

    $PerSettingMarkers = Get-AllSettingMarkers -BasePath $MarkerBasePath
    Write-LogEntry -Value "Per-setting markers found: $($PerSettingMarkers.Count)" -Severity 1

    $SureAdmin = Get-HPSureAdminEnrolled
    if ($null -eq $SureAdmin)
    {
        Write-LogEntry -Value "Sure Admin query failed - proceeding with normal settings classification" -Severity 2
    }
    elseif ($SureAdmin)
    {
        Write-LogEntry -Value "Sure Admin (Enhanced BIOS Authentication Mode) is ENABLED - device is INCOMPATIBLE with settings remediation" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "Sure Admin is not enabled - normal settings remediation applies" -Severity 1
    }

    $CurrentBiosValues = Get-HPBiosSettings
    if ($null -eq $CurrentBiosValues)
    {
        Write-LogEntry -Value "HP BIOS setting query returned null (DEGRADED)" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "HP BIOS setting query returned $($CurrentBiosValues.Count) settings" -Severity 1
    }

    $Result = Get-SettingsClassification `
        -Profile $Profile `
        -DesiredSettings $DesiredSettings `
        -CurrentBiosValues $CurrentBiosValues `
        -TopLevelMarker $TopLevelMarker `
        -PerSettingMarkers $PerSettingMarkers `
        -PasswordMarkerPresent $PwMarkerPresent `
        -NoPassword $NoPassword `
        -SureAdmin $SureAdmin `
        -RetryFailedAfterDays $RetryFailedAfterDays
}
catch
{
    Write-LogEntry -Value "Unhandled exception during classification: $($_.Exception.Message)" -Severity 3
    $Result = @{
        Status = 'DEGRADED'
        Stdout = "DEGRADED: settings classification-failed ($($_.Exception.Message -replace '[\r\n]+', ' '))"
        Reason = 'unhandled exception'
    }
}

#NoPassword mode tag.
if ($NoPassword)
{
    $Result.Stdout = "$($Result.Stdout) mode=nopassword"
}

Write-LogEntry -Value "Classification: $($Result.Status) - $($Result.Reason)" -Severity 1
Write-LogEntry -Value "STDOUT: $($Result.Stdout)" -Severity 1

#Optional Log Analytics reporting.
if ($ReportingEnabled)
{
    try
    {
        $Facts = Get-DeviceFacts
        $RptMode = if ($NoPassword) { 'nopassword' } else { 'normal' }
        $RptDriftNames = if ($null -ne $Result.DriftedNames) { [string[]]$Result.DriftedNames } else { @() }

        $RptDetail = New-Object System.Collections.Generic.List[hashtable]
        if ($null -ne $Result.SettingDetail) { foreach ($D in $Result.SettingDetail) { [void]$RptDetail.Add([hashtable]$D) } }
        if ($ReportingDepth -eq 'Full' -and $null -ne $CurrentBiosValues)
        {
            foreach ($Bn in $CurrentBiosValues.Keys)
            {
                if (-not $DesiredSettings.ContainsKey($Bn))
                {
                    [void]$RptDetail.Add(@{ SettingName = $Bn; DesiredValue = ''; CurrentValue = [string]$CurrentBiosValues[$Bn]; PerSettingStatus = ''; FailReason = ''; Managed = $false; SetDate = '' })
                }
            }
        }

        $RptDesiredHash = Get-DesiredStateHash -DesiredSettings $DesiredSettings
        $Payload = Get-BiosReportPayload -Component 'Settings' -Depth $ReportingDepth `
            -RunId ([guid]::NewGuid().ToString()) -TimeGenerated ([DateTime]::UtcNow.ToString('o')) `
            -DeviceId $Facts.DeviceId -DeviceName $Facts.DeviceName -Manufacturer $Facts.Manufacturer -Model $Facts.Model `
            -Profile $Profile -DesiredStateHash $RptDesiredHash -ScriptVersion $Version -Mode $RptMode `
            -OverallTag $Result.Status `
            -ManagedCount ([int]$Result.ManagedCount) -CompliantCount ([int]$Result.CompliantCount) -DriftCount ([int]$Result.DriftCount) `
            -UnsupportedCount ([int]$Result.UnsupportedCount) -FailedCount ([int]$Result.FailedCount) -BlockedCount ([int]$Result.BlockedCount) `
            -DriftNames $RptDriftNames -SettingDetail ($RptDetail.ToArray())

        $StateHash = Get-BiosReportStateHash -Payload $Payload
        $RptMarker = Get-ReportingMarker -BasePath $ReportingMarkerPath -Component 'Settings'
        $EmitDecision = Test-ShouldEmitReport -CurrentHash $StateHash -LastHash $RptMarker.LastEmittedHash -LastEmitDate $RptMarker.LastEmittedDate -HeartbeatDays $ReportingHeartbeatDays
        if ($EmitDecision.Emit)
        {
            Write-LogEntry -Value "Reporting: emitting ($($EmitDecision.Reason)) depth=$ReportingDepth detail=$($Payload.Detail.Count)" -Severity 1
            $EmitResult = Invoke-BiosReportEmit -Payload $Payload -Config $ReportingConfig
            if ($EmitResult.Ok)
            {
                Set-ReportingMarker -BasePath $ReportingMarkerPath -Component 'Settings' -Hash $StateHash -Date ([DateTime]::UtcNow.ToString('o'))
                Write-LogEntry -Value "Reporting: sent OK (summary=$($EmitResult.SummaryCount) detail=$($EmitResult.DetailCount))" -Severity 1
            }
            else
            {
                Write-LogEntry -Value "Reporting: emit failed - $($EmitResult.Reason) (compliance verdict unaffected)" -Severity 2
            }
        }
        else
        {
            Write-LogEntry -Value "Reporting: skipped ($($EmitDecision.Reason))" -Severity 1
        }
    }
    catch
    {
        Write-LogEntry -Value "Reporting: unhandled exception - $($_.Exception.Message) (compliance verdict unaffected)" -Severity 2
    }
}

Write-LogEntry -Value "END - HP BIOS settings detection" -Severity 1

#Write STDOUT for Intune reporting.
Write-Output $Result.Stdout

if ($Result.Status -eq 'COMPLIANT')
{
    exit 0
}
exit 1

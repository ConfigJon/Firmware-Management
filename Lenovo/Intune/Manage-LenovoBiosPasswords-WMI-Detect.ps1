<#
    .DESCRIPTION
        Intune detection script for Lenovo BIOS supervisor password management. Compares the on-device
        state (the HKLM marker + Lenovo's PasswordState bitmask) to the target version and exits 0
        (compliant) or 1 (non-compliant). Detection never attempts to unlock the BIOS. Scope is rotate +
        clear only (initial supervisor-password set is not supported). Skips on non-Lenovo hardware
        reporting (COMPLIANT: pw out-of-scope). Cert-auth devices (PasswordState=128) require signed
        commands, so they are reported INCOMPATIBLE (non-compliant).

    .PARAMETER TargetVersion
        Integer version this device should be at. 0 clears the password. Must match the paired -Remediate script.

    .PARAMETER LogFile
        CMTrace-compatible diagnostics log. Defaults to the IME log directory. Must have a .log extension.

    .LINK
        https://www.configjon.com/intune-bios-password-management/
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
    [ValidateRange(0, [int]::MaxValue)]
    [int]$TargetVersion = 1,

    #Internal: marker registry root.
    [Parameter(DontShow)]
    [ValidateNotNullOrEmpty()]
    [string]$MarkerBasePath = 'HKLM:\SOFTWARE\ConfigJonScripts\FirmwareManagement\BIOSPassword',

    #Internal: reporting emit-on-change marker root.
    [Parameter(DontShow)]
    [ValidateNotNullOrEmpty()]
    [string]$ReportingMarkerPath = 'HKLM:\SOFTWARE\ConfigJonScripts\FirmwareManagement\Reporting',

    [Parameter(Mandatory = $false)]
    [ValidateScript({
            if ($_ -notmatch '\.log$')
            {
                throw "The file specified in the LogFile parameter must be a .log file"
            }
            return $true
        })]
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Manage-LenovoBiosPasswords-WMI-Detect.log",

    #Internal: bypasses the manufacturer guard so the script can be tested on non-Lenovo hardware.
    [Parameter(DontShow)]
    [switch]$SkipManufacturerCheck
)

$Version = '1.0.0'
$Component = 'Manage-LenovoBiosPasswords-WMI-Detect'

#Reporting (Log Analytics) ===================================================================================================
#Optional: push each detection result to a Log Analytics workspace via the Logs Ingestion API.
#Disabled by default. Set $ReportingEnabled = $true and fill in $ReportingConfig to enable.
#A reporting failure does not affect the compliance result or exit code.
#The client-auth certificate's private key must be in Cert:\LocalMachine\My (readable by SYSTEM)
#Data is only sent when the reportable state changes, or once per HeartbeatDays.
#All values below are non-secret.
#Full setup, data model, and KQL: https://www.configjon.com/intune-bios-reporting/
$ReportingEnabled       = $false
$ReportingHeartbeatDays = 7
$ReportingConfig = @{
    TenantId              = ''
    ClientId              = ''
    CertThumbprint        = ''
    DceUri                = ''
    SummaryDcrImmutableId = ''
    SummaryStream         = 'Custom-BiosManagementRun_CL'
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

function Get-LenovoPasswordState
{
    #Query Lenovo_BiosPasswordSettings.PasswordState.

    try
    {
        $PasswordSettings = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_BiosPasswordSettings' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        Write-LogEntry -Value "Lenovo_BiosPasswordSettings WMI query failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    if ($null -eq $PasswordSettings)
    {
        Write-LogEntry -Value "Lenovo_BiosPasswordSettings query returned no rows (provider available but unexpected payload)" -Severity 2
        return $null
    }
    $Raw = [int]$PasswordSettings.PasswordState
    $CertAuth = ($Raw -eq 128)
    $SupervisorSet = if ($CertAuth) { $false } else { (($Raw -band 2) -ne 0) }
    return @{
        Raw           = $Raw
        SupervisorSet = $SupervisorSet
        CertAuth      = $CertAuth
    }
}

function Get-PasswordClassification
{
    #Based on the inputs, return @{ Status='COMPLIANT|NONCOMPLIANT|DEGRADED'; Stdout='...'; Reason='...' }.

    param(
        [Parameter(Mandatory = $true)][ValidateRange(0, [int]::MaxValue)][int]$TargetVersion,
        [Parameter(Mandatory = $false)]$Marker, #$null or hashtable returned by Get-PasswordMarker
        [Parameter(Mandatory = $false)][AllowNull()][object]$BiosSet, #$true, $false, or $null (query failed)
        [Parameter(Mandatory = $false)][AllowNull()][object]$CertAuth  #$true, $false, or $null (query failed)
    )
    $MarkerVersionStr = if ($null -ne $Marker) { "$($Marker.Version)" } else { '-' }

    #Cert-auth coexistence: distinct INCOMPATIBLE state, exit non-compliant so the admin sees the device.
    if ($CertAuth -eq $true)
    {
        return @{
            Status = 'NONCOMPLIANT'
            Stdout = 'NONCOMPLIANT: pw cert-auth-enabled (INCOMPATIBLE - signed payloads required)'
            Reason = 'Lenovo BIOS certificate-based authentication is in use (PasswordState 128) - password lifecycle does not apply'
        }
    }

    #DEGRADED - BIOS query failed, fall back to marker-only
    if ($null -eq $BiosSet)
    {
        if ($null -eq $Marker)
        {
            return @{
                Status = 'DEGRADED'
                Stdout = "DEGRADED: pw v=- marker-missing bios-query-failed (marker-only)"
                Reason = 'bios-query-failed and marker missing'
            }
        }
        if ($Marker.Version -eq $TargetVersion)
        {
            return @{
                Status = 'DEGRADED'
                Stdout = "DEGRADED: pw v=$($Marker.Version) marker-ok bios-query-failed (marker-only)"
                Reason = 'bios-query-failed but marker matches target'
            }
        }
        return @{
            Status = 'DEGRADED'
            Stdout = "DEGRADED: pw v=$($Marker.Version) marker-version-drift expected=$TargetVersion bios-query-failed (marker-only)"
            Reason = 'bios-query-failed and marker does not match target'
        }
    }

    $BiosSetStr = if ($BiosSet) { 'true' } else { 'false' }

    #TargetVersion = 0 (cleared)
    if ($TargetVersion -eq 0)
    {
        if (-not $BiosSet)
        {
            $Tag = if ($null -ne $Marker -and $Marker.Version -eq 0) { 'marker-ok-cleared' } else { 'marker-clear-no-bios-pw' }
            return @{
                Status = 'COMPLIANT'
                Stdout = "COMPLIANT: pw v=$MarkerVersionStr set=false $Tag"
                Reason = 'cleared as desired'
            }
        }
        return @{
            Status = 'NONCOMPLIANT'
            Stdout = "NONCOMPLIANT: pw v=$MarkerVersionStr set=true marker-needs-clear expected=0"
            Reason = 'TargetVersion=0 but BIOS pw still set'
        }
    }

    #TargetVersion > 0
    if ($null -eq $Marker)
    {
        return @{
            Status = 'NONCOMPLIANT'
            Stdout = "NONCOMPLIANT: pw v=- set=$BiosSetStr marker-missing"
            Reason = 'marker missing'
        }
    }
    if (-not $BiosSet)
    {
        return @{
            Status = 'NONCOMPLIANT'
            Stdout = "NONCOMPLIANT: pw v=$($Marker.Version) set=false bios-reset-detected"
            Reason = 'marker says password should be set but BIOS reports not set (BIOS reset, CMOS clear, RTC battery)'
        }
    }
    if ($Marker.Version -eq $TargetVersion)
    {
        return @{
            Status = 'COMPLIANT'
            Stdout = "COMPLIANT: pw v=$($Marker.Version) set=true marker-ok"
            Reason = 'marker matches target and BIOS is set'
        }
    }
    if ($Marker.Version -lt $TargetVersion)
    {
        return @{
            Status = 'NONCOMPLIANT'
            Stdout = "NONCOMPLIANT: pw v=$($Marker.Version) set=true marker-version-drift expected=$TargetVersion"
            Reason = 'marker version is older than target (needs rotate)'
        }
    }
    return @{
        Status = 'NONCOMPLIANT'
        Stdout = "NONCOMPLIANT: pw v=$($Marker.Version) set=true marker-version-future expected=$TargetVersion"
        Reason = 'marker version is newer than target (misconfigured deployment or rollback)'
    }
}

#Reporting functions ========================================================================================================
#Shared, vendor-agnostic Log Analytics reporting (password variant omits per-setting FailReason/BlockedCount)

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
    [void]$Builder.AppendLine("mc=$($S.ManagedCount)|cc=$($S.CompliantCount)|dc=$($S.DriftCount)|uc=$($S.UnsupportedCount)|fc=$($S.FailedCount)")
    [void]$Builder.AppendLine("dn=$($S.DriftNames)|pv=$($S.PasswordVersion)|ct=$($S.CertThumbprint)")
    [void]$Builder.AppendLine("mf=$($S.Manufacturer)|md=$($S.Model)|sv=$($S.ScriptVersion)")
    foreach ($D in ($Payload.Detail | Sort-Object { [string]$_.SettingName }))
    {
        [void]$Builder.AppendLine("$($D.SettingName)=$($D.DesiredValue)|$($D.CurrentValue)|$($D.PerSettingStatus)|$($D.Managed)")
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

Write-LogEntry -Value "START - Lenovo BIOS password detection (Intune) v$Version" -Severity 1
Write-LogEntry -Value "TargetVersion=$TargetVersion  MarkerBasePath=$MarkerBasePath  Reporting=$ReportingEnabled" -Severity 1

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
        Write-LogEntry -Value "END - Lenovo BIOS password detection (out of scope)" -Severity 1
        Write-Output "COMPLIANT: pw out-of-scope (not Lenovo)"
        exit 0
    }
    if ([string]::IsNullOrWhiteSpace($Manufacturer))
    {
        Write-LogEntry -Value "Manufacturer indeterminate - proceeding with normal detection (will fall through to BIOS-query result)" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "Manufacturer = '$Manufacturer' - Lenovo hardware confirmed, proceeding" -Severity 1
    }
}

#Password detection.
try
{
    $Marker = Get-PasswordMarker -BasePath $MarkerBasePath
    if ($null -eq $Marker)
    {
        Write-LogEntry -Value "Marker not present at $MarkerBasePath" -Severity 1
    }
    else
    {
        Write-LogEntry -Value "Marker present: Version=$($Marker.Version) SetDate=$($Marker.SetDate) CertThumbprint=$($Marker.CertThumbprint)" -Severity 1
    }

    $PwState = Get-LenovoPasswordState
    if ($null -eq $PwState)
    {
        Write-LogEntry -Value "Lenovo_BiosPasswordSettings query failed - DEGRADED classification" -Severity 2
        $BiosSet = $null
        $CertAuth = $null
    }
    else
    {
        Write-LogEntry -Value "Lenovo_BiosPasswordSettings.PasswordState=$($PwState.Raw)  SupervisorSet=$($PwState.SupervisorSet)  CertAuth=$($PwState.CertAuth)" -Severity 1
        $BiosSet = $PwState.SupervisorSet
        $CertAuth = $PwState.CertAuth
        if ($CertAuth)
        {
            Write-LogEntry -Value "PasswordState=128 - cert-based BIOS authentication is in use (INCOMPATIBLE with password lifecycle)" -Severity 2
        }
    }

    $Result = Get-PasswordClassification -TargetVersion $TargetVersion -Marker $Marker -BiosSet $BiosSet -CertAuth $CertAuth
}
catch
{
    Write-LogEntry -Value "Unhandled exception during classification: $($_.Exception.Message)" -Severity 3
    $Result = @{
        Status = 'DEGRADED'
        Stdout = "DEGRADED: pw classification-failed ($($_.Exception.Message -replace '[\r\n]+', ' '))"
        Reason = 'unhandled exception'
    }
}

Write-LogEntry -Value "Classification: $($Result.Status) - $($Result.Reason)" -Severity 1
Write-LogEntry -Value "STDOUT: $($Result.Stdout)" -Severity 1

#Optional Log Analytics reporting.
if ($ReportingEnabled)
{
    try
    {
        $Facts = Get-DeviceFacts
        $RptPwVersion = if ($null -ne $Marker) { [int]$Marker.Version } else { 0 }
        $RptCertThumbprint = if ($null -ne $Marker) { [string]$Marker.CertThumbprint } else { '' }

        $Payload = Get-BiosReportPayload -Component 'Password' -Depth 'Summary' `
            -RunId ([guid]::NewGuid().ToString()) -TimeGenerated ([DateTime]::UtcNow.ToString('o')) `
            -DeviceId $Facts.DeviceId -DeviceName $Facts.DeviceName -Manufacturer $Facts.Manufacturer -Model $Facts.Model `
            -ScriptVersion $Version -OverallTag $Result.Status `
            -PasswordVersion $RptPwVersion -CertThumbprint $RptCertThumbprint

        $StateHash = Get-BiosReportStateHash -Payload $Payload
        $RptMarker = Get-ReportingMarker -BasePath $ReportingMarkerPath -Component 'Password'
        $EmitDecision = Test-ShouldEmitReport -CurrentHash $StateHash -LastHash $RptMarker.LastEmittedHash -LastEmitDate $RptMarker.LastEmittedDate -HeartbeatDays $ReportingHeartbeatDays
        if ($EmitDecision.Emit)
        {
            Write-LogEntry -Value "Reporting: emitting ($($EmitDecision.Reason)) component=Password tag=$($Result.Status)" -Severity 1
            $EmitResult = Invoke-BiosReportEmit -Payload $Payload -Config $ReportingConfig
            if ($EmitResult.Ok)
            {
                Set-ReportingMarker -BasePath $ReportingMarkerPath -Component 'Password' -Hash $StateHash -Date ([DateTime]::UtcNow.ToString('o'))
                Write-LogEntry -Value "Reporting: sent OK (summary=$($EmitResult.SummaryCount))" -Severity 1
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

Write-LogEntry -Value "END - Lenovo BIOS password detection" -Severity 1

#Write STDOUT for Intune reporting.
Write-Output $Result.Stdout

if ($Result.Status -eq 'COMPLIANT')
{
    exit 0
}
exit 1

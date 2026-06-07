<#
    .DESCRIPTION
        Intune remediation script for Dell BIOS settings management. Compares the admin-edited
        $DesiredSettings hashtable to the BIOS-reported state, decrypts the password named by the pw marker
        (CMS material in $IntunePayload), and applies drifted settings via SetAttribute, verifying each by
        readback before writing its per-setting marker. Settings that fail are cached so they do not retry
        forever on hardware that cannot accept them. Skips on non-Dell hardware reporting (SKIPPED: settings
        out-of-scope).

        Build the payload with Tools\Build-IntunePayload.ps1; the certificate it references must be in
        Cert:\LocalMachine\My on the device before this runs. Full walkthrough in the blog posts below.

    .PARAMETER Profile
        Name of the desired-state profile, stamped into the marker. Must match the value in the paired script.

    .PARAMETER RetryFailedAfterDays
        Days after which a per-setting Failed cache is retried. 0 (default) = never retry automatically.

    .PARAMETER NoPassword
        For devices with NO BIOS admin password. Default OFF. When set, the password-marker dependency is dropped
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

    #Opt-in: BIOS has no admin password. Settings are read/applied unauthenticated. Must match the paired -Detect script.
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
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Manage-DellBiosSettings-WMI-Remediate.log",

    [Parameter(DontShow)]
    [switch]$SkipManufacturerCheck
)

$Version = '1.0.0'
$Component = 'Manage-DellBiosSettings-WMI-Remediate'

#Desired state ===============================================================================================================
#Edit this hashtable to match your desired Dell BIOS configuration.
#Names match the BIOS AttributeName exactly.
#Values match the BIOS CurrentValue exactly. ('Enabled' / 'Disabled').
#Use the existing Manage-DellBiosSettings-WMI.ps1 GetSettings mode to list what the device exposes.

$DesiredSettings = @{
    # Examples - replace with the settings your devices should standardize on:
    # 'WakeOnLan'      = 'LanOnly'
    # 'Virtualization' = 'Enabled'
    # 'BootSequence'   = 'UEFI'
    # 'SecureBoot'     = 'Enabled'
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

function Get-DellBiosSettings
{
    #Read all Dell BIOS attributes (Enumeration + Integer + String) and return as Name -> CurrentValue.

    $Result = @{}
    foreach ($ClassName in 'EnumerationAttribute', 'IntegerAttribute', 'StringAttribute')
    {
        try
        {
            $Items = Get-CimInstance -Namespace 'root\dcim\sysman\biosattributes' -ClassName $ClassName -ErrorAction Stop
        }
        catch
        {
            Write-LogEntry -Value "Dell $ClassName WMI query failed: $($_.Exception.Message)" -Severity 2
            return $null
        }
        foreach ($Item in $Items)
        {
            if (-not [string]::IsNullOrEmpty($Item.AttributeName))
            {
                $Result[$Item.AttributeName] = [string]$Item.CurrentValue
            }
        }
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
        [Parameter(Mandatory = $false)][AllowNull()]$CurrentBiosValues, # hashtable Name->Value, or $null if query failed
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
            Verdict          = 'EarlyExitFail'
            EarlyExitStdout  = 'FAILED: settings pw-marker-missing (cannot remediate without password)'
            EarlyExitCode    = 1
            Reason           = 'pw marker missing - cannot determine which CMS to decrypt for BIOS unlock'
            PerSettingPlan   = @()
            SetCount         = 0
            CompliantCount   = 0
            UnsupportedCount = 0
            FailedCachedCount = 0
        }
    }
    if ($null -eq $CurrentBiosValues)
    {
        return @{
            Verdict          = 'EarlyExitFail'
            EarlyExitStdout  = 'DEGRADED: settings bios-query-failed (Dell BIOS attribute namespace unreachable)'
            EarlyExitCode    = 1
            Reason           = 'CIM query for EnumerationAttribute / IntegerAttribute / StringAttribute failed'
            PerSettingPlan   = @()
            SetCount         = 0
            CompliantCount   = 0
            UnsupportedCount = 0
            FailedCachedCount = 0
        }
    }
    if (-not $NoPassword -and $PasswordMarkerVersion -le 0)
    {
        return @{
            Verdict          = 'EarlyExitFail'
            EarlyExitStdout  = 'FAILED: settings pw-marker-cleared (Version=0 - clear pw before applying authenticated settings)'
            EarlyExitCode    = 1
            Reason           = 'pw marker indicates cleared password (Version=0); cannot authenticate SetAttribute'
            PerSettingPlan   = @()
            SetCount         = 0
            CompliantCount   = 0
            UnsupportedCount = 0
            FailedCachedCount = 0
        }
    }
    if (-not $NoPassword -and ($PayloadVersions -notcontains $PasswordMarkerVersion))
    {
        return @{
            Verdict          = 'EarlyExitFail'
            EarlyExitStdout  = "FAILED: settings pw-version-not-in-payload v=$PasswordMarkerVersion"
            EarlyExitCode    = 1
            Reason           = "pw marker says v=$PasswordMarkerVersion but settings payload does not contain it (re-run Build-IntunePayload after adding the CMS)"
            PerSettingPlan   = @()
            SetCount         = 0
            CompliantCount   = 0
            UnsupportedCount = 0
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
            #Cached-Failed for the same desired value. Skip unless the retry window has elapsed.
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

function Invoke-DellSetBiosAttribute
{
    #Call Dell SetAttribute for one BIOS setting (authenticated when a password is supplied). Returns @{ Status; Threw; ExceptionMessage }.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$AttributeName,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$AttributeValue,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$Password
    )
    if ($WhatIfPreference)
    {
        $AuthLabel = if ([string]::IsNullOrEmpty($Password)) { 'unauthenticated' } else { 'authenticated' }
        Write-LogEntry -Value "WhatIf: would call Dell SetAttribute name=$AttributeName value=$AttributeValue ($AuthLabel)" -Severity 1
        return @{ Status = 0; Threw = $false; ExceptionMessage = '' }
    }
    try
    {
        $AttributeInterface = Get-CimInstance -Namespace 'root\dcim\sysman\biosattributes' -ClassName 'BIOSAttributeInterface' -ErrorAction Stop
    }
    catch
    {
        return @{ Status = -1; Threw = $true; ExceptionMessage = "Get-CimInstance BIOSAttributeInterface failed: $($_.Exception.Message)" }
    }
    $Encoder = New-Object System.Text.UTF8Encoding
    if ([string]::IsNullOrEmpty($Password))
    {
        $CimArgs = @{
            SecType        = [uint32]0
            SecHndCount    = [uint32]0
            SecHandle      = [byte[]]@()
            AttributeName  = $AttributeName
            AttributeValue = $AttributeValue
        }
    }
    else
    {
        $Bytes = $Encoder.GetBytes($Password)
        $CimArgs = @{
            SecType        = [uint32]1
            SecHndCount    = [uint32]$Bytes.Length
            SecHandle      = [byte[]]$Bytes
            AttributeName  = $AttributeName
            AttributeValue = $AttributeValue
        }
    }
    try
    {
        $Result = Invoke-CimMethod -InputObject $AttributeInterface -MethodName 'SetAttribute' -Arguments $CimArgs -ErrorAction Stop
    }
    catch
    {
        return @{ Status = -1; Threw = $true; ExceptionMessage = "Invoke-CimMethod SetAttribute threw: $($_.Exception.Message)" }
    }
    return @{ Status = [int]$Result.Status; Threw = $false; ExceptionMessage = '' }
}

function Get-DellSetFailReason
{
    #Translate a Dell write result into a FailReason persisted on the per-setting marker.

    param([Parameter(Mandatory = $true)][hashtable]$OpResult)
    if ($OpResult.Threw) { return 'write-threw' }
    switch ([int]$OpResult.Status)
    {
        3       { 'access-denied' }
        4       { 'access-denied' }
        1       { 'invalid-param' }
        default { 'set-failed' }
    }
}

function Invoke-RemediateSettings
{
    #Apply the drifted settings: decrypt the unlock password (unless NoPassword), SetAttribute each, verify via a single post-apply readback, then write per-setting + top-level markers.

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
            Write-LogEntry -Value "NoPassword mode: applying settings unauthenticated (SetAttribute SecType=0)" -Severity 1
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

        #Phase 1: queue each SetAttribute. Capture per-setting result for use after the post-query.
        $WriteResults = @{}
        foreach ($Item in $SettingsToApply)
        {
            $Name = $Item.Name
            $DesiredValue = $Item.DesiredValue
            Write-LogEntry -Value "Applying setting: $Name = '$DesiredValue' (was '$($Item.CurrentValue)')" -Severity 1
            $Op = Invoke-DellSetBiosAttribute -AttributeName $Name -AttributeValue $DesiredValue -Password $Password
            $WriteResults[$Name] = $Op
            if ($Op.Threw)
            {
                Write-LogEntry -Value "SetAttribute threw for $Name`: $($Op.ExceptionMessage)" -Severity 3
            }
            elseif ($Op.Status -ne 0)
            {
                Write-LogEntry -Value "SetAttribute returned status=$($Op.Status) for $Name" -Severity 2
            }
            else
            {
                Write-LogEntry -Value "SetAttribute returned status=0 for $Name - will verify via post-apply readback" -Severity 1
            }
        }

        #Phase 2: single post-apply BIOS re-query for verify-by-readback.
        $PostBiosValues = $null
        if (-not $WhatIfPreference)
        {
            $PostBiosValues = Get-DellBiosSettings
            if ($null -eq $PostBiosValues)
            {
                Write-LogEntry -Value "Post-apply BIOS query failed - cannot verify any setting; marking all queued settings as Failed" -Severity 3
            }
        }

        #Phase 3: classify each queued setting and write its marker.
        foreach ($Item in $SettingsToApply)
        {
            $Name = $Item.Name
            $Desired = $Item.DesiredValue
            $OpResult = $WriteResults[$Name]
            $PreValue = [string]$Item.CurrentValue
            if ($OpResult.Threw)
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PreValue -Status 'Failed' -FailReason (Get-DellSetFailReason -OpResult $OpResult)
                [void]$Failed.Add($Name)
                continue
            }
            if ($OpResult.Status -ne 0)
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PreValue -Status 'Failed' -FailReason (Get-DellSetFailReason -OpResult $OpResult)
                [void]$Failed.Add($Name)
                continue
            }
            if ($WhatIfPreference)
            {
                [void]$Applied.Add($Name)
                continue
            }
            if ($null -eq $PostBiosValues)
            {
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PreValue -Status 'Failed' -FailReason 'query-failed'
                [void]$Failed.Add($Name)
                continue
            }
            $PostValue = [string]$PostBiosValues[$Name]
            if ($null -eq $PostBiosValues[$Name])
            {
                Write-LogEntry -Value "Post-apply BIOS does not report attribute $Name - marking Failed" -Severity 3
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue '' -Status 'Failed' -FailReason 'readback-mismatch'
                [void]$Failed.Add($Name)
                continue
            }
            if ($PostValue -ne $Desired)
            {
                Write-LogEntry -Value "Readback mismatch for $Name`: BIOS now '$PostValue', expected '$Desired'" -Severity 3
                Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PostValue -Status 'Failed' -FailReason 'readback-mismatch'
                [void]$Failed.Add($Name)
                continue
            }
            Set-SettingMarker -BasePath $MarkerBasePath -Name $Name -DesiredValue $Desired -LastVerifiedValue $PostValue -Status 'Applied'
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

Write-LogEntry -Value "START - Dell BIOS settings remediation (Intune) v$Version" -Severity 1
Write-LogEntry -Value "Profile=$Profile  MarkerBasePath=$MarkerBasePath  PwMarkerPath=$PwMarkerPath  RetryFailedAfterDays=$RetryFailedAfterDays  NoPassword=$NoPassword  WhatIf=$WhatIfPreference" -Severity 1

#NoPassword mode tag.
$ModeNote = if ($NoPassword) { ' mode=nopassword' } else { '' }
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
    if (-not [string]::IsNullOrWhiteSpace($Manufacturer) -and $Manufacturer -notlike 'Dell*')
    {
        Write-LogEntry -Value "Manufacturer is '$Manufacturer' (not Dell) - script does not apply" -Severity 1
        Write-LogEntry -Value "END - Dell BIOS settings remediation (out of scope)" -Severity 1
        Write-Output 'SKIPPED: settings out-of-scope (not Dell)'
        exit 0
    }
    if ([string]::IsNullOrWhiteSpace($Manufacturer))
    {
        Write-LogEntry -Value "Manufacturer indeterminate - proceeding (BIOS calls will fail if not Dell)" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "Manufacturer = '$Manufacturer' - Dell hardware confirmed, proceeding" -Severity 1
    }
}

#DesiredSettings check.
if ($DesiredSettings.Count -eq 0)
{
    Write-LogEntry -Value '$DesiredSettings is empty - nothing to manage. Edit the script body before deploying.' -Severity 2
    Write-LogEntry -Value 'END - Dell BIOS settings remediation (no work)' -Severity 1
    Write-Output 'SKIPPED: settings no-managed-settings ($DesiredSettings empty - edit the script body)'
    exit 0
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
$CurrentBiosValues = Get-DellBiosSettings
if ($null -eq $CurrentBiosValues)
{
    Write-LogEntry -Value "BIOS attribute query returned null" -Severity 2
}
else
{
    Write-LogEntry -Value "BIOS attribute query returned $($CurrentBiosValues.Count) settings" -Severity 1
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
    Write-LogEntry -Value "END - Dell BIOS settings remediation (exit $($Plan.EarlyExitCode))" -Severity 1
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
    Write-LogEntry -Value "END - Dell BIOS settings remediation (exit 0, no drift)" -Severity 1
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
Write-LogEntry -Value "END - Dell BIOS settings remediation (exit $ExitCode)" -Severity 1
#Write STDOUT for Intune reporting.
Write-Output $Stdout
exit $ExitCode

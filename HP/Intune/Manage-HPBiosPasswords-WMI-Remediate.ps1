<#
    .DESCRIPTION
        Intune remediation script for HP BIOS setup password management (set / reapply / rotate / clear).
        Reads the HKLM marker + HP BIOS state, decides the action, and applies it using the CMS-encrypted
        password material in this script's $IntunePayload block. On any failure the marker is not advanced,
        so the next detection cycle re-classifies and the fallback list recovers the device. Skips on
        non-HP hardware reporting (SKIPPED: pw out-of-scope). When HP Sure Admin (Enhanced BIOS
        Authentication Mode) is enabled the password lifecycle does not apply, so this script short-circuits
        (SKIPPED, exit 0) and the -Detect script flags the device INCOMPATIBLE.

        Build the payload with Tools\Build-IntunePayload.ps1; the certificate it references must be in
        Cert:\LocalMachine\My on the device before this runs. Full walkthrough in the blog posts below.

    .PARAMETER TargetVersion
        Integer version this device should be at. 0 clears the password. Must match the paired -Detect script.

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

[CmdletBinding(SupportsShouldProcess = $true, PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$TargetVersion = 1,

    #Internal: marker registry root.
    [Parameter(DontShow)]
    [ValidateNotNullOrEmpty()]
    [string]$MarkerBasePath = 'HKLM:\SOFTWARE\ConfigJonScripts\FirmwareManagement\BIOSPassword',

    [Parameter(Mandatory = $false)]
    [ValidateScript({
            if ($_ -notmatch '\.log$')
            {
                throw "The file specified in the LogFile parameter must be a .log file"
            }
            return $true
        })]
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Manage-HPBiosPasswords-WMI-Remediate.log",

    [Parameter(DontShow)]
    [switch]$SkipManufacturerCheck
)

$Version = '1.0.0'
$Component = 'Manage-HPBiosPasswords-WMI-Remediate'

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

function Set-PasswordMarker
{
    #Write the BIOS password marker to the registry.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$BasePath,
        [Parameter(Mandatory = $true)][ValidateRange(0, [int]::MaxValue)][int]$Version,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$CertThumbprint
    )
    if ($WhatIfPreference)
    {
        Write-LogEntry -Value "WhatIf: would write marker Version=$Version CertThumbprint=$CertThumbprint at $BasePath" -Severity 1
        return
    }
    if (-not (Test-Path -LiteralPath $BasePath))
    {
        New-Item -Path $BasePath -Force | Out-Null
    }
    $Now = (Get-Date).ToString('o')
    Set-ItemProperty -LiteralPath $BasePath -Name 'Version' -Value $Version -Type DWord -Force
    Set-ItemProperty -LiteralPath $BasePath -Name 'SetDate' -Value $Now -Type String -Force
    Set-ItemProperty -LiteralPath $BasePath -Name 'CertThumbprint' -Value $CertThumbprint -Type String -Force
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

function Get-HPSetupPasswordSet
{
    #Query HP_BIOSSetting for the 'Setup Password' row and read the IsSet flag. Returns $true/$false/$null.

    try
    {
        $Setting = Get-CimInstance -Namespace 'root\hp\InstrumentedBIOS' -ClassName 'HP_BIOSSetting' -ErrorAction Stop |
            Where-Object { $_.Name -eq 'Setup Password' } | Select-Object -First 1
    }
    catch
    {
        Write-LogEntry -Value "HP_BIOSSetting WMI query (Setup Password) failed: $($_.Exception.Message)" -Severity 2
        return $null
    }
    if ($null -eq $Setting)
    {
        Write-LogEntry -Value "HP_BIOSSetting query returned no row for Name='Setup Password'" -Severity 2
        return $null
    }
    return ([int]$Setting.IsSet -eq 1)
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

function Get-RemediationPlan
{
    #Decide the remediation action (set/reapply/rotate/clear/drifted/skip/fail) from target version, marker, BIOS state, and payload coverage.

    param(
        [Parameter(Mandatory = $true)][ValidateRange(0, [int]::MaxValue)][int]$TargetVersion,
        [Parameter(Mandatory = $false)]$Marker,
        [Parameter(Mandatory = $false)][AllowNull()][object]$BiosSet,
        [Parameter(Mandatory = $true)][int[]]$PayloadVersions
    )

    if ($null -eq $BiosSet)
    {
        return @{ Action = 'FailDegraded'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = 'bios-query-failed' }
    }
    if ($TargetVersion -gt 0 -and ($PayloadVersions -notcontains $TargetVersion))
    {
        return @{ Action = 'FailTargetNotInPayload'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = "target-version-not-in-payload v=$TargetVersion" }
    }
    if ($TargetVersion -eq 0)
    {
        if (-not $BiosSet)
        {
            return @{ Action = 'SkipAlreadyCurrent'; FallbackOrder = @(); TargetVersion = 0; Reason = 'pw already cleared' }
        }
        $Ordered = @()
        if ($null -ne $Marker -and $Marker.Version -gt 0 -and ($PayloadVersions -contains $Marker.Version))
        {
            $Ordered += $Marker.Version
        }
        foreach ($v in ($PayloadVersions | Sort-Object -Descending))
        {
            if ($Ordered -notcontains $v) { $Ordered += $v }
        }
        return @{ Action = 'Clear'; FallbackOrder = $Ordered; TargetVersion = 0; Reason = 'clear pw' }
    }
    if ($null -eq $Marker)
    {
        if (-not $BiosSet)
        {
            return @{ Action = 'Set'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = 'Fresh - no current pw, set target' }
        }
        $Ordered = @($PayloadVersions | Sort-Object -Descending)
        return @{ Action = 'Drifted'; FallbackOrder = $Ordered; TargetVersion = $TargetVersion; Reason = 'marker missing but BIOS pw set - walk fallback list' }
    }
    if ($Marker.Version -eq $TargetVersion)
    {
        if ($BiosSet)
        {
            return @{ Action = 'SkipAlreadyCurrent'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = "pw already at v=$TargetVersion" }
        }
        return @{ Action = 'Reapply'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = 'BIOS reset detected - reapply target' }
    }
    if ($Marker.Version -gt $TargetVersion)
    {
        return @{ Action = 'FailMisconfigured'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = "marker v=$($Marker.Version) > target v=$TargetVersion (rollback not supported)" }
    }
    if (-not $BiosSet)
    {
        return @{ Action = 'Reapply'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = 'BIOS reset detected during rotate - reapply target without unlock' }
    }
    $Ordered = @()
    if ($PayloadVersions -contains $Marker.Version) { $Ordered += $Marker.Version }
    foreach ($v in ($PayloadVersions | Sort-Object -Descending))
    {
        if ($v -eq $TargetVersion) { continue }
        if ($Ordered -notcontains $v) { $Ordered += $v }
    }
    return @{ Action = 'Rotate'; FallbackOrder = $Ordered; TargetVersion = $TargetVersion; Reason = "rotate v=$($Marker.Version) -> v=$TargetVersion" }
}

function Invoke-HPSetSetupPassword
{
    #Call HP SetBIOSSetting for the 'Setup Password' attribute.
    #Return codes: 0=Success, 1=Not-Supported, 2=Error, 3=Timeout, 4=Failed, 5=Invalid-Param, 6=Access-Denied.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$OldPassword,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$NewPassword
    )
    if ($WhatIfPreference)
    {
        $ShapeLabel = if ([string]::IsNullOrEmpty($OldPassword)) { 'no-old-pw (Set/Reapply)' } elseif ([string]::IsNullOrEmpty($NewPassword)) { 'clear (old provided)' } else { 'change (old+new)' }
        Write-LogEntry -Value "WhatIf: would call HP SetBIOSSetting [Setup Password] - $ShapeLabel" -Severity 1
        return @{ Status = 0; Threw = $false; ExceptionMessage = '' }
    }
    try
    {
        $Interface = Get-CimInstance -Namespace 'root\hp\InstrumentedBIOS' -ClassName 'HP_BIOSSettingInterface' -ErrorAction Stop
    }
    catch
    {
        return @{ Status = -1; Threw = $true; ExceptionMessage = "Get-CimInstance HP_BIOSSettingInterface failed: $($_.Exception.Message)" }
    }
    $CimArgs = @{
        Name     = 'Setup Password'
        Value    = '<utf-16/>' + $NewPassword
        Password = if ([string]::IsNullOrEmpty($OldPassword)) { '<utf-16/>' } else { '<utf-16/>' + $OldPassword }
    }
    try
    {
        $Result = Invoke-CimMethod -InputObject $Interface -MethodName 'SetBIOSSetting' -Arguments $CimArgs -ErrorAction Stop
    }
    catch
    {
        return @{ Status = -1; Threw = $true; ExceptionMessage = "Invoke-CimMethod SetBIOSSetting threw: $($_.Exception.Message)" }
    }
    return @{ Status = [int]$Result.Return; Threw = $false; ExceptionMessage = '' }
}

function Invoke-RemediateSetOrReapply
{
    #Set or reapply the target password (no unlock needed), verify by readback, then advance the marker.

    param(
        [Parameter(Mandatory = $true)][ValidateSet('Set', 'Reapply')][string]$Action,
        [Parameter(Mandatory = $true)][int]$TargetVersion,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)][string]$MarkerBasePath
    )
    $NewPw = $null
    try
    {
        $NewPw = Get-CmsPlaintextFromPayload -Payload $Payload -Version $TargetVersion
        $Op = Invoke-HPSetSetupPassword -OldPassword '' -NewPassword $NewPw
        if ($Op.Threw)
        {
            Write-LogEntry -Value "HP WMI threw on $Action`: $($Op.ExceptionMessage)" -Severity 3
            return @{ Stdout = "FAILED: pw $($Action.ToLower()) v=$TargetVersion (WMI threw)"; ExitCode = 1 }
        }
        if ($Op.Status -ne 0)
        {
            Write-LogEntry -Value "HP SetBIOSSetting returned status=$($Op.Status) on $Action" -Severity 3
            return @{ Stdout = "FAILED: pw bios-set-failed status=$($Op.Status) (action=$($Action.ToLower()) v=$TargetVersion)"; ExitCode = 1 }
        }
        if (-not $WhatIfPreference)
        {
            $PostSet = Get-HPSetupPasswordSet
            if ($PostSet -ne $true)
            {
                Write-LogEntry -Value "$Action`: SetBIOSSetting reported success but IsSet=$PostSet" -Severity 3
                return @{ Stdout = "FAILED: pw set-readback-mismatch v=$TargetVersion"; ExitCode = 1 }
            }
        }
        Set-PasswordMarker -BasePath $MarkerBasePath -Version $TargetVersion -CertThumbprint $Payload.CertThumbprint
        $Label = if ($Action -eq 'Set') { 'set' } else { 'reapplied' }
        return @{ Stdout = "REMEDIATED: pw $Label v=$TargetVersion"; ExitCode = 0 }
    }
    finally
    {
        if ($null -ne $NewPw) { $NewPw = $null }
    }
}

function Invoke-RemediateRotate
{
    #Rotate to the target password: use the fallback list to unlock, set the new password, verify by readback, then advance the marker.

    param(
        [Parameter(Mandatory = $true)][hashtable]$Plan,
        [Parameter(Mandatory = $true)]$Marker,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)][int]$TargetVersion,
        [Parameter(Mandatory = $true)][string]$MarkerBasePath
    )
    $NewPw = $null
    $OldPw = $null
    $Unlocked = $false
    $UsedVersion = -1
    try
    {
        $NewPw = Get-CmsPlaintextFromPayload -Payload $Payload -Version $TargetVersion
        foreach ($CandidateVer in $Plan.FallbackOrder)
        {
            if ($null -ne $OldPw) { $OldPw = $null }
            try
            {
                $OldPw = Get-CmsPlaintextFromPayload -Payload $Payload -Version $CandidateVer
            }
            catch
            {
                Write-LogEntry -Value "Failed to decrypt CMS for candidate v=$CandidateVer - skipping: $($_.Exception.Message)" -Severity 2
                continue
            }
            Write-LogEntry -Value "Attempting unlock with v=$CandidateVer" -Severity 1
            $Op = Invoke-HPSetSetupPassword -OldPassword $OldPw -NewPassword $NewPw
            if ($Op.Threw)
            {
                Write-LogEntry -Value "HP WMI threw during rotate: $($Op.ExceptionMessage)" -Severity 3
                return @{ Stdout = "FAILED: pw rotate v=$CandidateVer->v=$TargetVersion (WMI threw)"; ExitCode = 1 }
            }
            if ($Op.Status -eq 0)
            {
                $Unlocked = $true
                $UsedVersion = $CandidateVer
                Write-LogEntry -Value "Unlock succeeded with v=$CandidateVer (HP status=0)" -Severity 1
                break
            }
            Write-LogEntry -Value "Unlock with v=$CandidateVer returned status=$($Op.Status) - trying next candidate" -Severity 2
        }
        if (-not $Unlocked)
        {
            Write-LogEntry -Value "All fallback candidates failed - drifted/recovery state" -Severity 3
            return @{ Stdout = "FAILED: pw drifted - no fallback password unlocked the BIOS (tried v=[$($Plan.FallbackOrder -join ',')])"; ExitCode = 1 }
        }
        if (-not $WhatIfPreference)
        {
            $PostSet = Get-HPSetupPasswordSet
            if ($PostSet -ne $true)
            {
                Write-LogEntry -Value "Rotate: SetBIOSSetting reported success but IsSet=$PostSet" -Severity 3
                return @{ Stdout = "FAILED: pw set-readback-mismatch v=$TargetVersion"; ExitCode = 1 }
            }
        }
        Set-PasswordMarker -BasePath $MarkerBasePath -Version $TargetVersion -CertThumbprint $Payload.CertThumbprint
        $MarkerVer = if ($null -ne $Marker) { "$($Marker.Version)" } else { '-' }
        if ($null -ne $Marker -and $UsedVersion -eq $Marker.Version)
        {
            return @{ Stdout = "REMEDIATED: pw rotated v=$UsedVersion->v=$TargetVersion"; ExitCode = 0 }
        }
        return @{ Stdout = "REMEDIATED: pw rotated v=$UsedVersion->v=$TargetVersion via fallback (marker said v=$MarkerVer)"; ExitCode = 0 }
    }
    finally
    {
        if ($null -ne $NewPw) { $NewPw = $null }
        if ($null -ne $OldPw) { $OldPw = $null }
    }
}

function Invoke-RemediateClear
{
    #Clear the password: use the fallback list to unlock, set an empty password, verify by readback, then write marker Version=0.

    param(
        [Parameter(Mandatory = $true)][hashtable]$Plan,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)][string]$MarkerBasePath
    )
    $OldPw = $null
    $Unlocked = $false
    $UsedVersion = -1
    try
    {
        foreach ($CandidateVer in $Plan.FallbackOrder)
        {
            if ($null -ne $OldPw) { $OldPw = $null }
            try
            {
                $OldPw = Get-CmsPlaintextFromPayload -Payload $Payload -Version $CandidateVer
            }
            catch
            {
                Write-LogEntry -Value "Failed to decrypt CMS for candidate v=$CandidateVer - skipping: $($_.Exception.Message)" -Severity 2
                continue
            }
            Write-LogEntry -Value "Attempting clear-via-unlock with v=$CandidateVer" -Severity 1
            $Op = Invoke-HPSetSetupPassword -OldPassword $OldPw -NewPassword ''
            if ($Op.Threw)
            {
                Write-LogEntry -Value "HP WMI threw during clear: $($Op.ExceptionMessage)" -Severity 3
                return @{ Stdout = "FAILED: pw clear (WMI threw)"; ExitCode = 1 }
            }
            if ($Op.Status -eq 0)
            {
                $Unlocked = $true
                $UsedVersion = $CandidateVer
                Write-LogEntry -Value "Clear succeeded with unlock v=$CandidateVer (HP status=0)" -Severity 1
                break
            }
            Write-LogEntry -Value "Clear-with-v=$CandidateVer returned status=$($Op.Status) - trying next candidate" -Severity 2
        }
        if (-not $Unlocked)
        {
            Write-LogEntry -Value "All clear candidates failed - drifted/recovery state" -Severity 3
            return @{ Stdout = "FAILED: pw clear - no fallback password unlocked the BIOS (tried v=[$($Plan.FallbackOrder -join ',')])"; ExitCode = 1 }
        }
        if (-not $WhatIfPreference)
        {
            $PostSet = Get-HPSetupPasswordSet
            if ($PostSet -ne $false)
            {
                Write-LogEntry -Value "Clear: SetBIOSSetting reported success but IsSet=$PostSet" -Severity 3
                return @{ Stdout = "FAILED: pw clear-readback-mismatch (still IsSet=true)"; ExitCode = 1 }
            }
        }
        Set-PasswordMarker -BasePath $MarkerBasePath -Version 0 -CertThumbprint $Payload.CertThumbprint
        return @{ Stdout = "REMEDIATED: pw cleared (used v=$UsedVersion to unlock)"; ExitCode = 0 }
    }
    finally
    {
        if ($null -ne $OldPw) { $OldPw = $null }
    }
}

function Invoke-RemediateDrifted
{
    #Recover a drifted device (marker missing but BIOS pw set): use the fallback list to unlock, set the target, verify by readback, then write the marker.

    param(
        [Parameter(Mandatory = $true)][hashtable]$Plan,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Payload,
        [Parameter(Mandatory = $true)][int]$TargetVersion,
        [Parameter(Mandatory = $true)][string]$MarkerBasePath
    )
    $NewPw = $null
    $OldPw = $null
    $Unlocked = $false
    $UsedVersion = -1
    try
    {
        $NewPw = Get-CmsPlaintextFromPayload -Payload $Payload -Version $TargetVersion
        foreach ($CandidateVer in $Plan.FallbackOrder)
        {
            if ($null -ne $OldPw) { $OldPw = $null }
            try
            {
                $OldPw = Get-CmsPlaintextFromPayload -Payload $Payload -Version $CandidateVer
            }
            catch
            {
                Write-LogEntry -Value "Failed to decrypt CMS for candidate v=$CandidateVer - skipping: $($_.Exception.Message)" -Severity 2
                continue
            }
            Write-LogEntry -Value "Drifted: attempting unlock with v=$CandidateVer" -Severity 1
            $Op = Invoke-HPSetSetupPassword -OldPassword $OldPw -NewPassword $NewPw
            if ($Op.Threw)
            {
                Write-LogEntry -Value "HP WMI threw during drifted-rotate: $($Op.ExceptionMessage)" -Severity 3
                return @{ Stdout = "FAILED: pw drifted v=$CandidateVer->v=$TargetVersion (WMI threw)"; ExitCode = 1 }
            }
            if ($Op.Status -eq 0)
            {
                $Unlocked = $true
                $UsedVersion = $CandidateVer
                Write-LogEntry -Value "Drifted-unlock succeeded with v=$CandidateVer (HP status=0)" -Severity 1
                break
            }
            Write-LogEntry -Value "Drifted-unlock with v=$CandidateVer returned status=$($Op.Status) - trying next candidate" -Severity 2
        }
        if (-not $Unlocked)
        {
            Write-LogEntry -Value "All drifted-fallback candidates failed - recovery state" -Severity 3
            return @{ Stdout = "FAILED: pw drifted - no fallback password unlocked the BIOS (tried v=[$($Plan.FallbackOrder -join ',')])"; ExitCode = 1 }
        }
        if (-not $WhatIfPreference)
        {
            $PostSet = Get-HPSetupPasswordSet
            if ($PostSet -ne $true)
            {
                Write-LogEntry -Value "Drifted: SetBIOSSetting reported success but IsSet=$PostSet" -Severity 3
                return @{ Stdout = "FAILED: pw set-readback-mismatch v=$TargetVersion"; ExitCode = 1 }
            }
        }
        Set-PasswordMarker -BasePath $MarkerBasePath -Version $TargetVersion -CertThumbprint $Payload.CertThumbprint
        return @{ Stdout = "REMEDIATED: pw rotated v=$UsedVersion->v=$TargetVersion via fallback (marker was missing)"; ExitCode = 0 }
    }
    finally
    {
        if ($null -ne $NewPw) { $NewPw = $null }
        if ($null -ne $OldPw) { $OldPw = $null }
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

Write-LogEntry -Value "START - HP BIOS password remediation (Intune) v$Version" -Severity 1
Write-LogEntry -Value "TargetVersion=$TargetVersion  MarkerBasePath=$MarkerBasePath  WhatIf=$WhatIfPreference" -Severity 1

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
        Write-LogEntry -Value "END - HP BIOS password remediation (out of scope)" -Severity 1
        Write-Output "SKIPPED: pw out-of-scope (not HP)"
        exit 0
    }
    if ([string]::IsNullOrWhiteSpace($Manufacturer))
    {
        Write-LogEntry -Value "Manufacturer indeterminate - proceeding (BIOS calls will fail if not HP)" -Severity 2
    }
    else
    {
        Write-LogEntry -Value "Manufacturer = '$Manufacturer' - HP hardware confirmed, proceeding" -Severity 1
    }
}

#Sure Admin short-circuit: signed payloads required, password lifecycle does not apply. Exit 0 so Intune does not retry.
$SureAdmin = Get-HPSureAdminEnrolled
if ($SureAdmin -eq $true)
{
    Write-LogEntry -Value "HP Sure Admin (Enhanced BIOS Authentication Mode) is enabled - password lifecycle does not apply. Skipping remediation." -Severity 2
    Write-LogEntry -Value "END - HP BIOS password remediation (Sure Admin incompatible)" -Severity 1
    Write-Output "SKIPPED: pw sure-admin-incompatible (signed payloads required)"
    exit 0
}
if ($null -eq $SureAdmin)
{
    Write-LogEntry -Value "Sure Admin query failed - proceeding (subsequent BIOS write may surface the failure)" -Severity 2
}

#Payload check.
if ($null -eq $IntunePayload -or $null -eq $IntunePayload.Files -or $IntunePayload.Files.Count -eq 0)
{
    Write-LogEntry -Value '$IntunePayload is null or empty. Build-IntunePayload.ps1 was not run before deployment.' -Severity 3
    Write-Output 'FAILED: pw payload-not-built ($IntunePayload is null - run Tools\Build-IntunePayload.ps1)'
    exit 1
}

#Installed cert check.
$CertCheck = Test-RemediationCert -Thumbprint $IntunePayload.CertThumbprint
if (-not $CertCheck.Ok)
{
    Write-LogEntry -Value "Cert check failed: $($CertCheck.Reason)" -Severity 3
    Write-Output "FAILED: pw $($CertCheck.Reason)"
    exit 1
}
Write-LogEntry -Value "Cert thumbprint=$($IntunePayload.CertThumbprint) found in LocalMachine\My with usable private key" -Severity 1

#Marker check.
$Marker = Get-PasswordMarker -BasePath $MarkerBasePath
if ($null -eq $Marker)
{
    Write-LogEntry -Value "Marker not present at $MarkerBasePath" -Severity 1
}
else
{
    Write-LogEntry -Value "Marker present: Version=$($Marker.Version) SetDate=$($Marker.SetDate) CertThumbprint=$($Marker.CertThumbprint)" -Severity 1
}

#Current state check.
$BiosSet = Get-HPSetupPasswordSet
Write-LogEntry -Value "BIOS Setup Password IsSet=$BiosSet" -Severity 1

$PayloadVersions = @($IntunePayload.Files.Keys | ForEach-Object { [int]$_ } | Sort-Object)
Write-LogEntry -Value "Payload versions available: $($PayloadVersions -join ', ')" -Severity 1

$Plan = Get-RemediationPlan -TargetVersion $TargetVersion -Marker $Marker -BiosSet $BiosSet -PayloadVersions $PayloadVersions
Write-LogEntry -Value "Plan: Action=$($Plan.Action)  Reason=$($Plan.Reason)  FallbackOrder=[$($Plan.FallbackOrder -join ',')]" -Severity 1

$Stdout = ''
$ExitCode = 1

#Take action based on the current state.
try
{
    switch ($Plan.Action)
    {
        'SkipAlreadyCurrent'
        {
            $verStr = if ($null -eq $Marker) { '-' } else { "$($Marker.Version)" }
            $Stdout = "SKIPPED: pw v=$verStr already current"
            $ExitCode = 0
            break
        }
        'FailDegraded'
        {
            $Stdout = "DEGRADED: pw bios-query-failed - cannot determine current state"
            $ExitCode = 1
            break
        }
        'FailMisconfigured'
        {
            $Stdout = "FAILED: pw misconfigured marker v=$($Marker.Version) > target v=$TargetVersion"
            $ExitCode = 1
            break
        }
        'FailTargetNotInPayload'
        {
            $Stdout = "FAILED: pw target-version-not-in-payload v=$TargetVersion (re-run Build-IntunePayload after adding the CMS)"
            $ExitCode = 1
            break
        }
        { $_ -in 'Set', 'Reapply' }
        {
            $r = Invoke-RemediateSetOrReapply -Action $Plan.Action -TargetVersion $TargetVersion -Payload $IntunePayload -MarkerBasePath $MarkerBasePath
            $Stdout = $r.Stdout
            $ExitCode = $r.ExitCode
            break
        }
        'Rotate'
        {
            $r = Invoke-RemediateRotate -Plan $Plan -Marker $Marker -Payload $IntunePayload -TargetVersion $TargetVersion -MarkerBasePath $MarkerBasePath
            $Stdout = $r.Stdout
            $ExitCode = $r.ExitCode
            break
        }
        'Clear'
        {
            $r = Invoke-RemediateClear -Plan $Plan -Payload $IntunePayload -MarkerBasePath $MarkerBasePath
            $Stdout = $r.Stdout
            $ExitCode = $r.ExitCode
            break
        }
        'Drifted'
        {
            $r = Invoke-RemediateDrifted -Plan $Plan -Payload $IntunePayload -TargetVersion $TargetVersion -MarkerBasePath $MarkerBasePath
            $Stdout = $r.Stdout
            $ExitCode = $r.ExitCode
            break
        }
        default
        {
            Write-LogEntry -Value "Internal error: unhandled plan action '$($Plan.Action)'" -Severity 3
            $Stdout = "FAILED: pw internal-error unhandled-action=$($Plan.Action)"
            $ExitCode = 1
        }
    }
}
catch
{
    Write-LogEntry -Value "Unhandled exception during remediation dispatch: $($_.Exception.Message)" -Severity 3
    $Stdout = "FAILED: pw internal-error ($($_.Exception.Message -replace '[\r\n]+', ' '))"
    $ExitCode = 1
}

Write-LogEntry -Value "Outcome: $Stdout" -Severity 1
Write-LogEntry -Value "END - HP BIOS password remediation (exit $ExitCode)" -Severity 1
#Write STDOUT for Intune reporting.
Write-Output $Stdout
exit $ExitCode

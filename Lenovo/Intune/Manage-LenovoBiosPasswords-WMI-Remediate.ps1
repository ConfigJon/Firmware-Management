<#
    .DESCRIPTION
        Intune remediation script for Lenovo BIOS supervisor password management (rotate / clear only).
        Reads the HKLM marker + Lenovo BIOS state, decides the action, and applies it using the
        CMS-encrypted password material in this script's $IntunePayload block. On any failure the marker
        is not advanced, so the next detection cycle re-classifies and the fallback list recovers the
        device. Skips on non-Lenovo hardware reporting (SKIPPED: pw out-of-scope). Setting an initial
        supervisor password is not supported (Lenovo needs the SDBM workflow), so that path returns FAILED.
        Cert-auth devices (PasswordState=128) short-circuit (SKIPPED) and the -Detect script flags them
        INCOMPATIBLE. Password changes apply at next reboot, so the SetUpdate return is authoritative (no readback).

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
    [System.IO.FileInfo]$LogFile = "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\Manage-LenovoBiosPasswords-WMI-Remediate.log",

    [Parameter(DontShow)]
    [switch]$SkipManufacturerCheck
)

$Version = '1.0.0'
$Component = 'Manage-LenovoBiosPasswords-WMI-Remediate'

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
    #Decide the remediation action (rotate/clear/drifted/skip/fail) from target version, marker, BIOS state, and payload coverage.
    #Initial supervisor set is unsupported, so those paths map to FailInitNotSupportedV1.

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
            return @{ Action = 'FailInitNotSupportedV1'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = 'Fresh state: no current pw - initial supervisor set requires SDBM' }
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
        return @{ Action = 'FailInitNotSupportedV1'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = "BIOS reset detected at marker.V=$TargetVersion - reapply would require initial-set" }
    }
    if ($Marker.Version -gt $TargetVersion)
    {
        return @{ Action = 'FailMisconfigured'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = "marker v=$($Marker.Version) > target v=$TargetVersion (rollback not supported)" }
    }
    if (-not $BiosSet)
    {
        return @{ Action = 'FailInitNotSupportedV1'; FallbackOrder = @(); TargetVersion = $TargetVersion; Reason = "BIOS reset detected during rotate marker.V=$($Marker.Version) -> target v=$TargetVersion - reapply would require initial-set" }
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

function Invoke-LenovoSetSupervisorPassword
{
    #Change the Lenovo supervisor (pap) password.

    param(
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$OldPassword,
        [Parameter(Mandatory = $true)][ValidateNotNull()][AllowEmptyString()][string]$NewPassword
    )
    if ($WhatIfPreference)
    {
        $ShapeLabel = if ([string]::IsNullOrEmpty($NewPassword)) { 'clear (old provided)' } else { 'change (old+new)' }
        Write-LogEntry -Value "WhatIf: would call Lenovo SetUpdate [pap] - $ShapeLabel" -Severity 1
        return @{ Status = 'Success'; Threw = $false; ExceptionMessage = '' }
    }
    try
    {
        $OpcodeInterface = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_WmiOpcodeInterface' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_WmiOpcodeInterface failed: $($_.Exception.Message)" }
    }
    if (($null -ne $OpcodeInterface) -and ($OpcodeInterface.Active -eq $true))
    {
        #Preferred opcode change sequence; SetUpdate commits and returns the authoritative status string.
        try
        {
            [void](Invoke-CimMethod -InputObject $OpcodeInterface -MethodName 'WmiOpcodeInterface' -Arguments @{ Parameter = "WmiOpcodePasswordType:pap;" } -ErrorAction Stop)
            [void](Invoke-CimMethod -InputObject $OpcodeInterface -MethodName 'WmiOpcodeInterface' -Arguments @{ Parameter = "WmiOpcodePasswordCurrent01:$OldPassword;" } -ErrorAction Stop)
            [void](Invoke-CimMethod -InputObject $OpcodeInterface -MethodName 'WmiOpcodeInterface' -Arguments @{ Parameter = "WmiOpcodePasswordNew01:$NewPassword;" } -ErrorAction Stop)
            $Result = Invoke-CimMethod -InputObject $OpcodeInterface -MethodName 'WmiOpcodeInterface' -Arguments @{ Parameter = "WmiOpcodePasswordSetUpdate;" } -ErrorAction Stop
        }
        catch
        {
            return @{ Status = ''; Threw = $true; ExceptionMessage = "Opcode SetUpdate threw: $($_.Exception.Message)" }
        }
        return @{ Status = [string]$Result.Return; Threw = $false; ExceptionMessage = '' }
    }
    #Legacy fallback (does not handle complex/special-character passwords reliably).
    try
    {
        $PasswordSet = Get-CimInstance -Namespace 'root\wmi' -ClassName 'Lenovo_SetBiosPassword' -ErrorAction Stop | Select-Object -First 1
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Get-CimInstance Lenovo_SetBiosPassword failed: $($_.Exception.Message)" }
    }
    if ($null -eq $PasswordSet)
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Lenovo_SetBiosPassword query returned no rows" }
    }
    try
    {
        $Result = Invoke-CimMethod -InputObject $PasswordSet -MethodName 'SetBiosPassword' -Arguments @{ parameter = "pap,$OldPassword,$NewPassword,ascii,us" } -ErrorAction Stop
    }
    catch
    {
        return @{ Status = ''; Threw = $true; ExceptionMessage = "Legacy SetBiosPassword threw: $($_.Exception.Message)" }
    }
    return @{ Status = [string]$Result.Return; Threw = $false; ExceptionMessage = '' }
}

function Invoke-RemediateRotate
{
    #Rotate to the target password: use the fallback list to unlock, set the new password, then advance the marker (SetUpdate commits; applies at next reboot).

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
            $Op = Invoke-LenovoSetSupervisorPassword -OldPassword $OldPw -NewPassword $NewPw
            if ($Op.Threw)
            {
                Write-LogEntry -Value "Lenovo WMI threw during rotate: $($Op.ExceptionMessage)" -Severity 3
                return @{ Stdout = "FAILED: pw rotate v=$CandidateVer->v=$TargetVersion (WMI threw)"; ExitCode = 1 }
            }
            if ($Op.Status -eq 'Success')
            {
                $Unlocked = $true
                $UsedVersion = $CandidateVer
                Write-LogEntry -Value "Unlock succeeded with v=$CandidateVer (Lenovo status=Success)" -Severity 1
                break
            }
            Write-LogEntry -Value "Unlock with v=$CandidateVer returned status=$($Op.Status) - trying next candidate" -Severity 2
        }
        if (-not $Unlocked)
        {
            Write-LogEntry -Value "All fallback candidates failed - drifted/recovery state" -Severity 3
            return @{ Stdout = "FAILED: pw drifted - no fallback password unlocked the BIOS (tried v=[$($Plan.FallbackOrder -join ',')])"; ExitCode = 1 }
        }
        #SetUpdate Success is authoritative; Lenovo applies the change at next reboot (no readback - see header).
        Write-LogEntry -Value "SetUpdate Success - rotate committed; will apply at next reboot" -Severity 1
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
    #Clear the password: use the fallback list to unlock, set an empty password, then write marker Version=0 (applies at next reboot).

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
            $Op = Invoke-LenovoSetSupervisorPassword -OldPassword $OldPw -NewPassword ''
            if ($Op.Threw)
            {
                Write-LogEntry -Value "Lenovo WMI threw during clear: $($Op.ExceptionMessage)" -Severity 3
                return @{ Stdout = "FAILED: pw clear (WMI threw)"; ExitCode = 1 }
            }
            if ($Op.Status -eq 'Success')
            {
                $Unlocked = $true
                $UsedVersion = $CandidateVer
                Write-LogEntry -Value "Clear succeeded with unlock v=$CandidateVer (Lenovo status=Success)" -Severity 1
                break
            }
            Write-LogEntry -Value "Clear-with-v=$CandidateVer returned status=$($Op.Status) - trying next candidate" -Severity 2
        }
        if (-not $Unlocked)
        {
            Write-LogEntry -Value "All clear candidates failed - drifted/recovery state" -Severity 3
            return @{ Stdout = "FAILED: pw clear - no fallback password unlocked the BIOS (tried v=[$($Plan.FallbackOrder -join ',')])"; ExitCode = 1 }
        }
        #SetUpdate Success is authoritative; Lenovo applies the clear at next reboot (no readback).
        Write-LogEntry -Value "SetUpdate Success - clear committed; will apply at next reboot" -Severity 1
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
    #Recover a drifted device (marker missing but BIOS pw set): use the fallback list to unlock, set the target, then write the marker (applies at next reboot).

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
            $Op = Invoke-LenovoSetSupervisorPassword -OldPassword $OldPw -NewPassword $NewPw
            if ($Op.Threw)
            {
                Write-LogEntry -Value "Lenovo WMI threw during drifted-rotate: $($Op.ExceptionMessage)" -Severity 3
                return @{ Stdout = "FAILED: pw drifted v=$CandidateVer->v=$TargetVersion (WMI threw)"; ExitCode = 1 }
            }
            if ($Op.Status -eq 'Success')
            {
                $Unlocked = $true
                $UsedVersion = $CandidateVer
                Write-LogEntry -Value "Drifted-unlock succeeded with v=$CandidateVer (Lenovo status=Success)" -Severity 1
                break
            }
            Write-LogEntry -Value "Drifted-unlock with v=$CandidateVer returned status=$($Op.Status) - trying next candidate" -Severity 2
        }
        if (-not $Unlocked)
        {
            Write-LogEntry -Value "All drifted-fallback candidates failed - recovery state" -Severity 3
            return @{ Stdout = "FAILED: pw drifted - no fallback password unlocked the BIOS (tried v=[$($Plan.FallbackOrder -join ',')])"; ExitCode = 1 }
        }
        #SetUpdate Success is authoritative; Lenovo applies the change at next reboot (no readback).
        Write-LogEntry -Value "SetUpdate Success - drifted-rotate committed; will apply at next reboot" -Severity 1
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

Write-LogEntry -Value "START - Lenovo BIOS password remediation (Intune) v$Version" -Severity 1
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
    if (-not [string]::IsNullOrWhiteSpace($Manufacturer) -and $Manufacturer -notlike 'LENOVO*')
    {
        Write-LogEntry -Value "Manufacturer is '$Manufacturer' (not Lenovo) - script does not apply" -Severity 1
        Write-LogEntry -Value "END - Lenovo BIOS password remediation (out of scope)" -Severity 1
        Write-Output "SKIPPED: pw out-of-scope (not Lenovo)"
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

#Cert-auth short-circuit: signed payloads required, password lifecycle does not apply. Exit 0 so Intune does not retry.
$PwState = Get-LenovoPasswordState
if ($null -ne $PwState -and $PwState.CertAuth)
{
    Write-LogEntry -Value "Lenovo BIOS certificate-based authentication is in use (PasswordState=128) - password lifecycle does not apply. Skipping remediation." -Severity 2
    Write-LogEntry -Value "END - Lenovo BIOS password remediation (cert-auth incompatible)" -Severity 1
    Write-Output "SKIPPED: pw cert-auth-incompatible (signed payloads required)"
    exit 0
}
if ($null -eq $PwState)
{
    Write-LogEntry -Value "Lenovo_BiosPasswordSettings query failed - proceeding (planner will return FailDegraded)" -Severity 2
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
$BiosSet = if ($null -ne $PwState) { $PwState.SupervisorSet } else { $null }
Write-LogEntry -Value "BIOS supervisor IsSet=$BiosSet" -Severity 1

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
        'FailInitNotSupportedV1'
        {
            $Stdout = "FAILED: pw init-not-supported-v1 (initial supervisor set requires SDBM)"
            $ExitCode = 1
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
Write-LogEntry -Value "END - Lenovo BIOS password remediation (exit $ExitCode)" -Severity 1
#Write STDOUT for Intune reporting.
Write-Output $Stdout
exit $ExitCode

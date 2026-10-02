# Dell BIOS management scripts

PowerShell scripts for managing BIOS passwords and settings on Dell systems. The posts that document them are listed in the [Dell section](https://www.configjon.com/bios-firmware-configuration/#dell) of the BIOS / Firmware Configuration page on my blog. For the other manufacturers and an overview of the repository, see the [main README](../README.md).

## Task sequence and interactive scripts

| Script | Purpose | Documentation |
| --- | --- | --- |
| [Manage-DellBiosPasswords-DellBIOSProvider.ps1](Manage-DellBiosPasswords-DellBIOSProvider.ps1) | Set, change, or clear Dell BIOS admin and system passwords using the DellBIOSProvider module. | [Dell BIOS Password Management - DellBIOSProvider](https://www.configjon.com/dell-bios-password-management/) |
| [Manage-DellBiosPasswords-WMI.ps1](Manage-DellBiosPasswords-WMI.ps1) | Set, change, or clear Dell BIOS admin and system passwords using WMI, with no module required. | [Dell BIOS Password Management - WMI](https://www.configjon.com/dell-bios-password-management-wmi/) |
| [Manage-DellBiosSettings-DellBIOSProvider.ps1](Manage-DellBiosSettings-DellBIOSProvider.ps1) | Get or set Dell BIOS settings, or set the boot order, using the DellBIOSProvider module. | [Dell BIOS Settings Management - DellBIOSProvider](https://www.configjon.com/dell-bios-settings-management/) |
| [Manage-DellBiosSettings-WMI.ps1](Manage-DellBiosSettings-WMI.ps1) | Get or set Dell BIOS settings, reset them to defaults, or set the boot order using WMI, with no module required. | [Dell BIOS Settings Management - WMI](https://www.configjon.com/dell-bios-settings-management-wmi/) |
| [Install-DellBiosProvider.ps1](Install-DellBiosProvider.ps1) | Install the Dell Command \| PowerShell Provider (DellBIOSProvider) module from the PowerShell Gallery or a local copy. | [Working with the Dell Command \| PowerShell Provider](https://www.configjon.com/working-with-the-dell-command-powershell-provider/) |

The password and settings scripts each come in two variants. The DellBIOSProvider variants use the DellBIOSProvider PowerShell module, which `Install-DellBiosProvider.ps1` installs. The WMI variants use WMI directly and need no module.

## Example settings files

Example files with commonly configured Dell BIOS settings. The CSV files can be passed to the settings scripts above with the `-CsvPath` parameter.

| File | Contents |
| --- | --- |
| [Settings_CSV_SecureBoot.csv](Settings_CSV_SecureBoot.csv) | Settings for enabling UEFI and Secure Boot |
| [Settings_CSV_TPM.csv](Settings_CSV_TPM.csv) | Settings for enabling and activating the TPM |
| [Settings_CSV_General.csv](Settings_CSV_General.csv) | Other common settings |
| [Settings_InScript_All.txt](Settings_InScript_All.txt) | Common settings formatted for use in the body of a settings script |

## Intune Remediations

The [Intune](Intune/) folder holds detection and remediation script pairs for Intune Remediations. Start with [Managing BIOS Passwords and Settings with Intune](https://www.configjon.com/bios-management-with-intune/). The whole series is listed under [Managing BIOS with Intune](https://www.configjon.com/bios-firmware-configuration/#managing-bios-with-intune) on the blog.

| Script | Purpose | Documentation |
| --- | --- | --- |
| [Manage-DellBiosPasswords-WMI-Detect.ps1](Intune/Manage-DellBiosPasswords-WMI-Detect.ps1) | Intune detection script that checks whether the Dell BIOS admin password is at the target version. | [BIOS password management with Intune Remediations](https://www.configjon.com/intune-bios-password-management/) |
| [Manage-DellBiosPasswords-WMI-Remediate.ps1](Intune/Manage-DellBiosPasswords-WMI-Remediate.ps1) | Intune remediation script that sets, reapplies, rotates, or clears the Dell BIOS admin password. | [BIOS password management with Intune Remediations](https://www.configjon.com/intune-bios-password-management/) |
| [Manage-DellBiosSettings-WMI-Detect.ps1](Intune/Manage-DellBiosSettings-WMI-Detect.ps1) | Intune detection script that checks whether Dell BIOS settings match the desired state. | [BIOS settings management with Intune Remediations](https://www.configjon.com/intune-bios-settings-management/) |
| [Manage-DellBiosSettings-WMI-Remediate.ps1](Intune/Manage-DellBiosSettings-WMI-Remediate.ps1) | Intune remediation script that applies Dell BIOS settings that have drifted from the desired state. | [BIOS settings management with Intune Remediations](https://www.configjon.com/intune-bios-settings-management/) |
| [Build-IntunePayload.ps1](../Tools/Build-IntunePayload.ps1) (in `Tools`) | Embed CMS-encrypted BIOS password files in an Intune remediation script. | [Setting up BIOS password certificates in Intune](https://www.configjon.com/intune-bios-password-certificates/) |

## Legacy scripts

The [Legacy Scripts](Legacy%20Scripts/) folder holds earlier versions of `Install-DellBiosProvider.ps1`, written for older releases of the DellBIOSProvider module. They are kept for reference and are not updated.

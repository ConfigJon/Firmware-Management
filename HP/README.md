# HP BIOS management scripts

PowerShell scripts for managing BIOS passwords and settings on HP systems. The posts that document them are listed in the [HP section](https://www.configjon.com/bios-firmware-configuration/#hp) of the BIOS / Firmware Configuration page on my blog. For the other manufacturers and an overview of the repository, see the [main README](../README.md).

## Task sequence and interactive scripts

| Script | Purpose | Documentation |
| --- | --- | --- |
| [Manage-HPBiosPasswords-WMI.ps1](Manage-HPBiosPasswords-WMI.ps1) | Set, change, or clear HP BIOS setup and power on passwords using WMI, with no module required. | [HP BIOS Password Management](https://www.configjon.com/hp-bios-password-management/) |
| [Manage-HPBiosPasswords-HPCMSL.ps1](Manage-HPBiosPasswords-HPCMSL.ps1) | Set, change, or clear HP BIOS setup and power on passwords using the HP Client Management Script Library (HPCMSL). | [HP BIOS Password Management (HPCMSL)](https://www.configjon.com/hp-bios-password-management-hpcmsl/) |
| [Manage-HPBiosSettings-WMI.ps1](Manage-HPBiosSettings-WMI.ps1) | Get or set HP BIOS settings using WMI, with no module required. | [HP BIOS Settings Management](https://www.configjon.com/hp-bios-settings-management/) |
| [Manage-HPBiosSettings-HPCMSL.ps1](Manage-HPBiosSettings-HPCMSL.ps1) | Get or set HP BIOS settings, or reset them to defaults, using the HP Client Management Script Library (HPCMSL). | [HP BIOS Settings Management (HPCMSL)](https://www.configjon.com/hp-bios-settings-management-hpcmsl/) |
| [Install-HPCMSL.ps1](Install-HPCMSL.ps1) | Install the HP Client Management Script Library (HPCMSL) modules from the PowerShell Gallery or a local copy. | [Installing the HP Client Management Script Library](https://www.configjon.com/installing-the-hp-client-management-script-library/) |

The password and settings scripts each come in two variants. The HPCMSL variants use the HP Client Management Script Library, which `Install-HPCMSL.ps1` installs. The WMI variants use WMI directly and need no module.

## Example settings files

Example files with commonly configured HP BIOS settings. The CSV files can be passed to the settings scripts above with the `-CsvPath` parameter.

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
| [Manage-HPBiosPasswords-WMI-Detect.ps1](Intune/Manage-HPBiosPasswords-WMI-Detect.ps1) | Intune detection script that checks whether the HP BIOS setup password is at the target version. | [BIOS password management with Intune Remediations](https://www.configjon.com/intune-bios-password-management/) |
| [Manage-HPBiosPasswords-WMI-Remediate.ps1](Intune/Manage-HPBiosPasswords-WMI-Remediate.ps1) | Intune remediation script that sets, reapplies, rotates, or clears the HP BIOS setup password. | [BIOS password management with Intune Remediations](https://www.configjon.com/intune-bios-password-management/) |
| [Manage-HPBiosSettings-WMI-Detect.ps1](Intune/Manage-HPBiosSettings-WMI-Detect.ps1) | Intune detection script that checks whether HP BIOS settings match the desired state. | [BIOS settings management with Intune Remediations](https://www.configjon.com/intune-bios-settings-management/) |
| [Manage-HPBiosSettings-WMI-Remediate.ps1](Intune/Manage-HPBiosSettings-WMI-Remediate.ps1) | Intune remediation script that applies HP BIOS settings that have drifted from the desired state. | [BIOS settings management with Intune Remediations](https://www.configjon.com/intune-bios-settings-management/) |
| [Build-IntunePayload.ps1](../Tools/Build-IntunePayload.ps1) (in `Tools`) | Embed CMS-encrypted BIOS password files in an Intune remediation script. | [Setting up BIOS password certificates in Intune](https://www.configjon.com/intune-bios-password-certificates/) |

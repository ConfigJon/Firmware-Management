# Lenovo BIOS management scripts

PowerShell scripts for managing BIOS passwords and settings on Lenovo systems. The posts that document them are listed in the [Lenovo section](https://www.configjon.com/bios-firmware-configuration/#lenovo) of the BIOS / Firmware Configuration page on my blog. For the other manufacturers and an overview of the repository, see the [main README](../README.md).

## Task sequence and interactive scripts

| Script | Purpose | Documentation |
| --- | --- | --- |
| [Manage-LenovoBiosPasswords.ps1](Manage-LenovoBiosPasswords.ps1) | Manage Lenovo BIOS supervisor, power on, system management, and hard drive passwords using WMI. | [Lenovo BIOS Password Management](https://www.configjon.com/lenovo-bios-password-management/) |
| [Manage-LenovoBiosSettings.ps1](Manage-LenovoBiosSettings.ps1) | Get or set Lenovo BIOS settings, or reset them to defaults, using WMI. | [Lenovo BIOS Settings Management](https://www.configjon.com/lenovo-bios-settings-management/) |

The Lenovo scripts use WMI directly and need no module.

## Example settings files

Example files with commonly configured Lenovo BIOS settings. The CSV files can be passed to the settings script above with the `-CsvPath` parameter.

| File | Contents |
| --- | --- |
| [Settings_CSV_SecureBoot.csv](Settings_CSV_SecureBoot.csv) | Settings for enabling UEFI and Secure Boot |
| [Settings_CSV_TPM.csv](Settings_CSV_TPM.csv) | Settings for enabling and activating the TPM |
| [Settings_CSV_General.csv](Settings_CSV_General.csv) | Other common settings |
| [Settings_InScript_All.txt](Settings_InScript_All.txt) | Common settings formatted for use in the body of the settings script |

## Intune Remediations

The [Intune](Intune/) folder holds detection and remediation script pairs for Intune Remediations. Start with [Managing BIOS Passwords and Settings with Intune](https://www.configjon.com/bios-management-with-intune/). The whole series is listed under [Managing BIOS with Intune](https://www.configjon.com/bios-firmware-configuration/#managing-bios-with-intune) on the blog.

| Script | Purpose | Documentation |
| --- | --- | --- |
| [Manage-LenovoBiosPasswords-WMI-Detect.ps1](Intune/Manage-LenovoBiosPasswords-WMI-Detect.ps1) | Intune detection script that checks whether the Lenovo BIOS supervisor password is at the target version. | [BIOS password management with Intune Remediations](https://www.configjon.com/intune-bios-password-management/) |
| [Manage-LenovoBiosPasswords-WMI-Remediate.ps1](Intune/Manage-LenovoBiosPasswords-WMI-Remediate.ps1) | Intune remediation script that rotates or clears the Lenovo BIOS supervisor password. | [BIOS password management with Intune Remediations](https://www.configjon.com/intune-bios-password-management/) |
| [Manage-LenovoBiosSettings-WMI-Detect.ps1](Intune/Manage-LenovoBiosSettings-WMI-Detect.ps1) | Intune detection script that checks whether Lenovo BIOS settings match the desired state. | [BIOS settings management with Intune Remediations](https://www.configjon.com/intune-bios-settings-management/) |
| [Manage-LenovoBiosSettings-WMI-Remediate.ps1](Intune/Manage-LenovoBiosSettings-WMI-Remediate.ps1) | Intune remediation script that applies Lenovo BIOS settings that have drifted from the desired state. | [BIOS settings management with Intune Remediations](https://www.configjon.com/intune-bios-settings-management/) |
| [Build-IntunePayload.ps1](../Tools/Build-IntunePayload.ps1) (in `Tools`) | Embed CMS-encrypted BIOS password files in an Intune remediation script. | [Setting up BIOS password certificates in Intune](https://www.configjon.com/intune-bios-password-certificates/) |

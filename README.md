# Firmware-Management

[![License](https://img.shields.io/github/license/ConfigJon/Firmware-Management)](LICENSE)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207-5391FE)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20WinPE-0078D4)
![Deploys with](https://img.shields.io/badge/deploys%20with-ConfigMgr%20%7C%20Intune-0078D4)
[![Docs](https://img.shields.io/badge/docs-configjon.com-2ea44f)](https://www.configjon.com/bios-firmware-configuration/)

PowerShell scripts for managing BIOS/firmware settings and passwords on Dell, HP, and Lenovo systems. Runs on Windows PowerShell 5.1 and PowerShell 7.

## Documentation

All documentation is on my blog. This repository holds the scripts, and the README in each manufacturer folder lists the files in that folder and links each script to the post that documents it.

- **[BIOS / Firmware Configuration](https://www.configjon.com/bios-firmware-configuration/)** - the index of all the posts
- **[BIOS Management Scripts v2 Released](https://www.configjon.com/bios-management-scripts-v2/)** - overview of the task sequence and interactive scripts
- **[Managing BIOS Passwords and Settings with Intune](https://www.configjon.com/bios-management-with-intune/)** - overview of the Intune scripts

The scripts also have built-in help. `Get-Help .\<script>.ps1 -Full` shows the full help for a script, and `Get-Help .\<script>.ps1 -Online` opens the post that documents it.

## What's in the repository

There are two sets of scripts:

- **Task sequence and interactive scripts** - for ConfigMgr or MDT task sequences, or for running by hand. Each manufacturer has a script for BIOS passwords and a script for BIOS settings. The Dell and HP scripts come in two variants: one uses the manufacturer's PowerShell module, and one uses WMI directly with no module. The Lenovo scripts use WMI.
- **Intune Remediations** - detection and remediation script pairs for BIOS passwords and for BIOS settings, in the `Intune` folder of each manufacturer.

| Folder | Contents | Posts |
| --- | --- | --- |
| [Dell](Dell/) | Password and settings scripts in DellBIOSProvider and WMI variants, the DellBIOSProvider module installer, example settings files, and the Intune scripts | [Dell posts](https://www.configjon.com/bios-firmware-configuration/#dell) |
| [HP](HP/) | Password and settings scripts in HPCMSL and WMI variants, the HPCMSL installer, example settings files, and the Intune scripts | [HP posts](https://www.configjon.com/bios-firmware-configuration/#hp) |
| [Lenovo](Lenovo/) | Password and settings scripts, example settings files, and the Intune scripts | [Lenovo posts](https://www.configjon.com/bios-firmware-configuration/#lenovo) |
| [Tools](Tools/) | [Build-IntunePayload.ps1](Tools/Build-IntunePayload.ps1): Embed CMS-encrypted BIOS password files in an Intune remediation script. | [Setting up BIOS password certificates in Intune](https://www.configjon.com/intune-bios-password-certificates/) |

## License

Released under the [MIT License](LICENSE).

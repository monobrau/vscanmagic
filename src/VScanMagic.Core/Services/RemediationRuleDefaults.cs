using VScanMagic.Core.Models;

namespace VScanMagic.Core.Services;

public static class RemediationRuleDefaults
{
    public static List<RemediationRule> GetAll() =>
    [
        new RemediationRule
        {
            Pattern = "*Windows Server 2012*",
            WordText = "This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. Consider planning a migration to a supported operating system version.",
            TicketText = "- This end-of-support operating system represents an infrastructure project\r\n  - Consider planning a migration to a supported operating system version",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*end-of-life*",
            WordText = "This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. Consider planning a migration to a supported operating system version.",
            TicketText = "- This end-of-support operating system represents an infrastructure project\r\n  - Consider planning a migration to a supported operating system version",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*out of support*",
            WordText = "This end-of-support operating system represents an infrastructure project beyond the scope of quarterly vulnerability remediation. Consider planning a migration to a supported operating system version.",
            TicketText = "- This end-of-support operating system represents an infrastructure project\r\n  - Consider planning a migration to a supported operating system version",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Windows 10*",
            WordText = "Windows 10 reached End of Life on October 14, 2025, and is no longer supported by Microsoft unless you have extended support licensing. If Windows Updates are functional and no extension licensing is in place, there is nothing further to be done other than considering an upgrade to Windows 11 or retiring the machine. For systems with extension licensing, continue to verify Windows Update status through ConnectWise Automate.",
            TicketText = "- Windows 10 reached End of Life on October 14, 2025\r\n  - No longer supported unless you have extended support licensing\r\n  - If Windows Updates are functional and no extension licensing in place:\r\n    * Nothing to be done other than considering upgrade to Windows 11 or retiring machine\r\n  - For systems with extension licensing:\r\n    * Continue to verify Windows Update status through ConnectWise Automate",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Windows*",
            WordText = "Windows patch inconsistencies should be investigated via ConnectWise Automate. Systems with lower vulnerability counts may indicate that patching is working correctly and awaiting the latest patch cycles. For systems with high vulnerability counts, verify Windows Update status and investigate any potential issues preventing patch installation.",
            TicketText = "- Investigate via ConnectWise Automate\r\n  - Verify Windows Update status on affected systems\r\n  - Check for any issues preventing patch installation",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*printer*",
            WordText = "Network printers and IoT devices require manual firmware updates via manufacturer-provided tools and interfaces. Consult the manufacturer's documentation for firmware update procedures.",
            TicketText = "- Requires manual firmware updates via manufacturer tools\r\n  - Consult manufacturer documentation for update procedures",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*ripple20*",
            WordText = "Ripple20 vulnerabilities affect the Treck TCP/IP stack in millions of IoT and embedded devices. Remediation requires firmware updates from the device manufacturer. Identify affected devices via network scanning; check vendor security advisories (e.g., Intel, HP, Cisco, Schneider Electric). Where patching is not immediately possible, implement network segmentation and firewall rules to restrict access to vulnerable devices.",
            TicketText = "- Ripple20 affects Treck TCP/IP stack in IoT/embedded devices\r\n  - Update firmware from device manufacturer\r\n  - Check vendor security advisories for patches\r\n  - Where patching not possible: network segmentation and firewall restrictions",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*rippl20*",
            WordText = "Ripple20 vulnerabilities affect the Treck TCP/IP stack in millions of IoT and embedded devices. Remediation requires firmware updates from the device manufacturer. Identify affected devices via network scanning; check vendor security advisories (e.g., Intel, HP, Cisco, Schneider Electric). Where patching is not immediately possible, implement network segmentation and firewall rules to restrict access to vulnerable devices.",
            TicketText = "- Ripple20 affects Treck TCP/IP stack in IoT/embedded devices\r\n  - Update firmware from device manufacturer\r\n  - Check vendor security advisories for patches\r\n  - Where patching not possible: network segmentation and firewall restrictions",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Ripple20*",
            WordText = "Ripple20 vulnerabilities affect the Treck TCP/IP stack in millions of IoT and embedded devices. Remediation requires firmware updates from the device manufacturer. Identify affected devices via network scanning; check vendor security advisories (e.g., Intel, HP, Cisco, Schneider Electric). Where patching is not immediately possible, implement network segmentation and firewall rules to restrict access to vulnerable devices.",
            TicketText = "- Ripple20 affects Treck TCP/IP stack in IoT/embedded devices\r\n  - Update firmware from device manufacturer\r\n  - Check vendor security advisories for patches\r\n  - Where patching not possible: network segmentation and firewall restrictions",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft Teams*",
            WordText = "Microsoft Teams can be updated via RMM script deployed through ConnectWise Automate. This can be remediated by cleaning up unused user profile installed versions using: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM.",
            TicketText = "- Update via RMM script deployed through ConnectWise Automate\r\n  - Can be remediated by cleaning up unused user profile installed versions\r\n  - Script path: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*VMware*",
            WordText = "VMware ESXi and vSphere updates are released by Broadcom. Remediation requires downloading patches from the Broadcom Support Portal and applying via vSphere Lifecycle Manager (recommended) or by manually installing the offline bundle or ISO. Host reboot is required; migrate or shut down VMs before patching.",
            TicketText = "- Download patches from Broadcom Support Portal\r\n  - Apply via vSphere Lifecycle Manager (recommended) or manual offline bundle/ISO\r\n  - Host reboot required; migrate or shut down VMs before patching",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*vSphere*",
            WordText = "VMware ESXi and vSphere updates are released by Broadcom. Remediation requires downloading patches from the Broadcom Support Portal and applying via vSphere Lifecycle Manager (recommended) or by manually installing the offline bundle or ISO. Host reboot is required; migrate or shut down VMs before patching.",
            TicketText = "- Download patches from Broadcom Support Portal\r\n  - Apply via vSphere Lifecycle Manager (recommended) or manual offline bundle/ISO\r\n  - Host reboot required; migrate or shut down VMs before patching",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*ESXi*",
            WordText = "VMware ESXi and vSphere updates are released by Broadcom. Remediation requires downloading patches from the Broadcom Support Portal and applying via vSphere Lifecycle Manager (recommended) or by manually installing the offline bundle or ISO. Host reboot is required; migrate or shut down VMs before patching.",
            TicketText = "- Download patches from Broadcom Support Portal\r\n  - Apply via vSphere Lifecycle Manager (recommended) or manual offline bundle/ISO\r\n  - Host reboot required; migrate or shut down VMs before patching",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*vCenter*",
            WordText = "VMware vCenter Server updates are released by Broadcom. Remediation requires downloading the patch ISO from the Broadcom Support Portal and applying via the vCenter Lifecycle Manager plug-in, GUI installer, or Virtual Appliance Management Interface (VAMI). Plan for maintenance; vCenter services restart during patching.",
            TicketText = "- Download patch ISO from Broadcom Support Portal\r\n  - Apply via vCenter Lifecycle Manager, GUI installer, or VAMI\r\n  - Plan for maintenance; vCenter services restart during patching",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*FortiGate*",
            WordText = "FortiGate firewall firmware updates are available from the Fortinet Support Portal. For managed devices, use FortiCloud Fabric Manager (Management > Firmware) to download, backup, and install firmware. If automatic upgrade fails, apply manually via the web interface (System > Firmware). Backup the configuration before updating. Plan a maintenance window; the device reboots during the update.",
            TicketText = "- Download firmware from Fortinet Support Portal\r\n  - Prefer FortiCloud Fabric Manager (Management > Firmware) for managed devices\r\n  - If automatic upgrade fails, apply manually via web interface (System > Firmware)\r\n  - Backup configuration before updating\r\n  - Plan maintenance window; device reboots during update",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Fortinet*",
            WordText = "FortiGate firewall firmware updates are available from the Fortinet Support Portal. For managed devices, use FortiCloud Fabric Manager (Management > Firmware) to download, backup, and install firmware. If automatic upgrade fails, apply manually via the web interface (System > Firmware). Backup the configuration before updating. Plan a maintenance window; the device reboots during the update.",
            TicketText = "- Download firmware from Fortinet Support Portal\r\n  - Prefer FortiCloud Fabric Manager (Management > Firmware) for managed devices\r\n  - If automatic upgrade fails, apply manually via web interface (System > Firmware)\r\n  - Backup configuration before updating\r\n  - Plan maintenance window; device reboots during update",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*SonicWall*",
            WordText = "SonicWall firewall firmware updates are available from the MySonicWall portal. Download the firmware for your appliance model, then apply via the management interface (System > Settings > Firmware). Backup the configuration before updating. Plan a maintenance window; the device reboots during the update.",
            TicketText = "- Download firmware from MySonicWall portal\r\n  - Apply via management interface (System > Settings > Firmware)\r\n  - Backup configuration before updating\r\n  - Plan maintenance window; device reboots during update",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET (all versions)*",
            WordText = "Legacy and modern .NET runtimes are consolidated in this finding. .NET Framework cannot be upgraded to modern .NET without a migration project; modern .NET versions should be kept current or retargeted to .NET 8 (LTS). Out-of-support versions stop receiving security patches.",
            TicketText = "- Consolidated .NET finding (Framework, Core, Runtime)\r\n  - .NET Framework: migration project required to modern .NET; no in-place upgrade\r\n  - Modern .NET: retarget to .NET 8 (LTS) with minimal code changes where possible\r\n  - Main driver: security patches; out-of-support versions receive none",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Framework*",
            WordText = "Legacy .NET: .NET Framework (1.0 through 4.8) cannot be upgraded in-place to modern .NET (5, 6, 7, 8, 9). There is no direct upgrade path; attempting to switch runtimes without migration will likely break the application. Modern .NET is a different runtime with different APIs - some Framework APIs do not exist or behave differently. The older the Framework version (e.g. 3.5, 4.0), the more painful the migration. Migration is a real project: retarget the application to .NET 8 (LTS) or later, update deprecated or incompatible APIs, and test thoroughly. For MSP/client conversations: the main reason to migrate is that Framework versions go out of support and stop receiving security patches - not because end users will notice any difference.",
            TicketText = "- .NET Framework cannot be upgraded to modern .NET (5, 6, 7, 8, 9) without migration\r\n  - No direct upgrade path; switching runtimes without migration will likely break the app\r\n  - Modern .NET is a different runtime; many APIs differ or are missing\r\n  - Older Framework versions (3.5, 4.0) are more painful to migrate than 4.x\r\n  - Migration is a real project: retarget to .NET 8 (LTS), update APIs, test thoroughly\r\n  - Main driver: security patches (Framework goes out of support) - not user experience",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Core 2.*",
            WordText = "Legacy .NET: .NET Core 2.x is out of support. Migration to modern .NET (5/6/7/8) is required. .NET Framework apps need a full migration project; apps already on .NET Core can usually retarget with minimal code changes. The main driver is security: out-of-support versions stop receiving patches.",
            TicketText = "- Legacy .NET Core 2.x: out of support\r\n  - Retarget to .NET 8 (LTS) or later\r\n  - .NET Core apps: usually minimal code changes\r\n  - Main driver: security patches; out-of-support versions receive none",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Core 3.*",
            WordText = "Legacy .NET: .NET Core 3.x (including 3.1 LTS) has reached or is nearing end of support. Migration to modern .NET 8 (LTS) is recommended. Within the modern .NET lineage, retargeting is usually a project file edit and minimal code changes. The main driver is security: out-of-support versions stop receiving patches.",
            TicketText = "- Legacy .NET Core 3.x: end of support or nearing\r\n  - Retarget to .NET 8 (LTS) with minimal code changes\r\n  - Main driver: security patches; out-of-support versions receive none\r\n  - LTS versions (6, 8, 10) get 3 years support",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Core*",
            WordText = "Legacy .NET: .NET Core (1.x, 2.x, 3.x) is the predecessor to modern .NET (5+). Core 2.x and 3.x are out of support. Migration to .NET 8 (LTS) is recommended. .NET Framework requires a full migration project; apps on Core can usually retarget with minimal changes. The main driver is security patches.",
            TicketText = "- Legacy .NET Core: out of support or nearing\r\n  - Retarget to .NET 8 (LTS) or later\r\n  - Main driver: security patches; out-of-support versions receive none\r\n  - .NET Framework migration is a real project; Core retarget is usually low friction",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Runtime 5.*",
            WordText = "Modern .NET: .NET 5 is out of support (18-month lifecycle). Retarget to .NET 8 (LTS) with minimal code changes. Within modern .NET, upgrades are largely drop-in. The main driver is support: LTS versions get 3 years; non-LTS get 18 months. End users typically notice nothing; value is security patches and lower infrastructure costs.",
            TicketText = "- Modern .NET 5: out of support\r\n  - Retarget to .NET 8 (LTS) with minimal code changes\r\n  - Main driver: security patches\r\n  - LTS versions get 3 years support; non-LTS get 18 months",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Runtime 6.*",
            WordText = "Modern .NET: .NET 6 (LTS) maintains strong backwards compatibility. Retargeting to .NET 8 (LTS) is usually a project file edit and minimal code changes. The main reason to upgrade is support lifecycle: LTS versions get 3 years; staying current ensures security patches. End users typically notice nothing; value is security patches and lower infrastructure costs.",
            TicketText = "- Modern .NET 6: largely drop-in upgradeable to .NET 8\r\n  - Retarget to .NET 8 (LTS) with minimal code changes\r\n  - Main driver: support lifecycle and security patches\r\n  - LTS versions get 3 years support",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Runtime 7.*",
            WordText = "Modern .NET: .NET 7 (non-LTS) has an 18-month support window. Consider retargeting to .NET 8 (LTS) for 3-year support. Upgrades within modern .NET are largely drop-in. The main driver is support lifecycle and security patches. End users typically notice nothing.",
            TicketText = "- Modern .NET 7: non-LTS (18 months support)\r\n  - Consider retarget to .NET 8 (LTS) for 3-year support\r\n  - Largely drop-in upgradeable\r\n  - Main driver: support lifecycle and security patches",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Runtime 8.*",
            WordText = "Modern .NET: .NET 8 (LTS) is the current long-term support release with 3 years of support. Keep current with patch updates. If a newer LTS (e.g. .NET 10) is available, retargeting is usually minimal. End users typically notice nothing; value is security patches and lower infrastructure costs.",
            TicketText = "- Modern .NET 8 (LTS): keep current with patch updates\r\n  - 3 years support\r\n  - Retarget to newer LTS when available with minimal changes\r\n  - Main driver: security patches",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Runtime 9.*",
            WordText = "Modern .NET: .NET 9 (non-LTS) has an 18-month support window. For long-term support, consider .NET 8 (LTS) or plan for .NET 10 (LTS). Upgrades within modern .NET are largely drop-in. The main driver is support lifecycle and security patches.",
            TicketText = "- Modern .NET 9: non-LTS (18 months support)\r\n  - Consider .NET 8 (LTS) or .NET 10 (LTS) for longer support\r\n  - Largely drop-in upgradeable\r\n  - Main driver: support lifecycle and security patches",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Runtime*",
            WordText = "Modern .NET: .NET 5+ (Runtime) maintains strong backwards compatibility. Retargeting between versions is usually a project file edit and minimal code changes. The main reason to upgrade is support: LTS versions (6, 8, 10) get 3 years; non-LTS get 18 months. End users typically notice nothing; for MSP/client conversations, the value is security patches and lower infrastructure costs - not user-facing improvements.",
            TicketText = "- Modern .NET: largely drop-in upgradeable between versions\r\n  - Retarget to .NET 8 (LTS) with minimal code changes\r\n  - Main driver: security patches (older versions go out of support) - not user experience\r\n  - LTS versions get 3 years support; non-LTS get 18 months",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET 5.*",
            WordText = "Modern .NET: .NET 5 is out of support (18-month lifecycle). Retarget to .NET 8 (LTS) with minimal code changes. Within modern .NET, upgrades are largely drop-in. The main driver is support and security patches.",
            TicketText = "- Modern .NET 5: out of support\r\n  - Retarget to .NET 8 (LTS) with minimal code changes\r\n  - Main driver: security patches",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET 6.*",
            WordText = "Modern .NET: .NET 6 (LTS) is largely drop-in upgradeable to .NET 8 (LTS). Retargeting is usually a project file edit and minimal code changes. The main driver is support lifecycle and security patches. End users typically notice nothing.",
            TicketText = "- Modern .NET 6: largely drop-in upgradeable to .NET 8\r\n  - Retarget to .NET 8 (LTS) with minimal code changes\r\n  - Main driver: support lifecycle and security patches",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET 7.*",
            WordText = "Modern .NET: .NET 7 (non-LTS) has an 18-month support window. Consider retargeting to .NET 8 (LTS) for 3-year support. Upgrades within modern .NET are largely drop-in.",
            TicketText = "- Modern .NET 7: non-LTS\r\n  - Consider retarget to .NET 8 (LTS) for 3-year support\r\n  - Largely drop-in upgradeable",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET 8.*",
            WordText = "Modern .NET: .NET 8 (LTS) is the current long-term support release with 3 years of support. Keep current with patch updates. End users typically notice nothing; value is security patches and lower infrastructure costs.",
            TicketText = "- Modern .NET 8 (LTS): keep current with patch updates\r\n  - 3 years support\r\n  - Main driver: security patches",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET 9.*",
            WordText = "Modern .NET: .NET 9 (non-LTS) has an 18-month support window. For long-term support, consider .NET 8 (LTS) or plan for .NET 10 (LTS). Upgrades within modern .NET are largely drop-in.",
            TicketText = "- Modern .NET 9: non-LTS (18 months support)\r\n  - Consider .NET 8 or .NET 10 (LTS) for longer support\r\n  - Largely drop-in upgradeable",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*",
            WordText = "Determine what the device or software is (use Product/OS and affected hosts). Review the manufacturer's security advisories and vulnerability data for patches or firmware updates. Consider configuration mitigations (e.g. network segmentation, hardening) where patching is not immediately possible. If available via ConnectWise Automate/RMM, deploy updates via patch management or scripts; otherwise, manual updates may be required.",
            TicketText = "- Determine device/software identity (Product/OS, affected hosts)\r\n  - Review manufacturer security advisories and vulnerability data\r\n  - Check for firmware updates or patches\r\n  - Consider configuration mitigations where patching not possible\r\n  - Deploy via ConnectWise Automate/RMM if available; otherwise manual updates",
            IsDefault = true
        },
    ];
}
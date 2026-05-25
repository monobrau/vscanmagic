using VScanMagic.Core.Models;

namespace VScanMagic.Core.Services;

public static class RemediationRuleDefaults
{
    /// <summary>
    /// Increment when built-in rule text or metadata changes so persisted rules resync on load.
    /// </summary>
    public const int Revision = 3;

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
            WordText = "Windows 10 reached End of Life on October 14, 2025, and is no longer supported by Microsoft unless you have extended support licensing. If Windows Updates are functional and no extension licensing is in place, there is nothing further to be done other than considering an upgrade to Windows 11 or retiring the machine. For systems with extension licensing, continue to verify Windows Update status through RMM if the client has patch management deployed.",
            TicketText = "- Windows 10 reached End of Life on October 14, 2025\r\n  - No longer supported unless you have extended support licensing\r\n  - If Windows Updates are functional and no extension licensing in place:\r\n    * Nothing to be done other than considering upgrade to Windows 11 or retiring machine\r\n  - For systems with extension licensing:\r\n    * Verify Windows Update status via RMM if patch management is deployed for the client",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Windows*",
            WordText = "Windows patch inconsistencies should be investigated via RMM when the client has patch management deployed. Systems with lower vulnerability counts may indicate that patching is working correctly and awaiting the latest patch cycles. For systems with high vulnerability counts, verify Windows Update status and investigate any potential issues preventing patch installation.",
            TicketText = "- Investigate via RMM if patch management is deployed for the client\r\n  - Verify Windows Update status on affected systems\r\n  - Check for any issues preventing patch installation",
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
            WordText = "Microsoft Teams can be updated via RMM script when the client has patch management deployed. Unused user profile-installed versions can be cleaned up in RMM (example path in ConnectWise Automate: Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation).",
            TicketText = "- Update via RMM script if patch management is deployed for the client\r\n  - Can remediate by cleaning up unused user profile-installed versions\r\n  - Example script path (ConnectWise Automate): RR - Custom > R-Security - Teams Classic Cleanup Remediation",
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
            Pattern = "*Microsoft Visual C++ (all versions)*",
            GuidanceStyle = RemediationGuidanceStyle.SideBySideRuntime,
            WordText = "Microsoft Visual C++ Redistributables are side-by-side runtime libraries. Different major release lines (2008, 2010, 2012, 2013, and 2015–2022/v14) normally coexist on the same machine because applications depend on the line they were built against. Do not remove an older major line unless you have verified no installed application still requires it. Within the Visual Studio 2015–2022 (v14) line, Microsoft documents binary compatibility — install the latest v14 redistributable build for that architecture; it will not downgrade a newer v14 already present. Remediation is to patch within the same major line (latest vc_redist for that line, or Windows Update for centrally deployed runtimes), not to delete unrelated year packages. Updating the parent application may install a newer redistributable, but other software may still depend on older major lines.",
            TicketText = "- Visual C++ redistributables are side-by-side across major release lines (2008, 2010, 2012, 2013, 2015–2022/v14)\r\n  - Do NOT remove an older major line without verifying dependent applications\r\n  - Major lines are not interchangeable (2013 ≠ 2015–2022)\r\n  - Within 2015–2022/v14: install latest v14 build for the architecture; binary compatible per Microsoft\r\n  - Remediate by patching within the same major line via vc_redist or Windows Update\r\n  - Multiple Visual C++ versions on one host is normal",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Visual C++*",
            GuidanceStyle = RemediationGuidanceStyle.SideBySideRuntime,
            WordText = "Microsoft Visual C++ Redistributables are side-by-side runtime libraries. Different major release lines (2008, 2010, 2012, 2013, and 2015–2022/v14) normally coexist on the same machine because applications depend on the line they were built against. Do not remove an older major line unless you have verified no installed application still requires it. Within the Visual Studio 2015–2022 (v14) line, Microsoft documents binary compatibility — install the latest v14 redistributable build for that architecture; it will not downgrade a newer v14 already present. Remediation is to patch within the same major line (latest vc_redist for that line, or Windows Update for centrally deployed runtimes), not to delete unrelated year packages. Updating the parent application may install a newer redistributable, but other software may still depend on older major lines.",
            TicketText = "- Visual C++ redistributables are side-by-side across major release lines (2008, 2010, 2012, 2013, 2015–2022/v14)\r\n  - Do NOT remove an older major line without verifying dependent applications\r\n  - Major lines are not interchangeable (2013 ≠ 2015–2022)\r\n  - Within 2015–2022/v14: install latest v14 build for the architecture; binary compatible per Microsoft\r\n  - Remediate by patching within the same major line via vc_redist or Windows Update\r\n  - Multiple Visual C++ versions on one host is normal",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Visual C#*",
            WordText = "ConnectSecure sometimes groups atypical product names with Visual C++ redistributable findings. Verify the actual installed product on affected hosts before remediating. If it is a Visual C++ Redistributable, apply the same side-by-side guidance: patch within the same major release line and do not remove older major lines without verifying dependent applications.",
            TicketText = "- Verify actual product on affected hosts (may be Visual C++ redistributable data)\r\n  - If Visual C++ redistributable: patch within same major release line\r\n  - Do NOT remove older major redistributable lines without dependency check",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET (all versions)*",
            WordText = "Microsoft .NET has two incompatible families — the break is between .NET Framework and modern .NET, not at a single version number. Legacy .NET Framework (through 4.8) is not backward compatible with modern .NET (.NET Core 3.x and .NET 5+). .NET 5 was the rebranding of .NET Core; it continues the modern lineage with strong backward compatibility from Core 3.1 onward and between modern releases (5→8→10). Framework applications cannot move to modern .NET without a migration project. Modern .NET apps can usually retarget to a supported LTS release (.NET 8 or .NET 10) with minimal changes. For runtime-only findings, keep the matching branch current via Windows Update or vendor installers — installing a modern runtime does not satisfy a .NET Framework-only application. .NET 6 reached end of support on November 12, 2024.",
            TicketText = "- Two incompatible .NET families: .NET Framework vs modern .NET (Core 3.x / .NET 5+)\r\n  - The break is Framework vs modern — NOT \"before .NET 5 vs after .NET 5\" within Framework\r\n  - .NET Framework (through 4.8): migration project required; no in-place upgrade to modern .NET\r\n  - Modern .NET (.NET 5+): continues Core lineage; largely backward compatible within modern versions only\r\n  - .NET 6 out of support since November 12, 2024 — retarget to .NET 8 or .NET 10 LTS\r\n  - Patch supported runtimes via Windows Update; do not install modern .NET expecting it to replace Framework",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Framework*",
            WordText = "Legacy .NET Framework (1.0 through 4.8) is a separate runtime from modern .NET (.NET Core 3.x and .NET 5+). They are not backward compatible — Framework cannot be upgraded in-place to modern .NET. There is no direct upgrade path; attempting to switch runtimes without migration will likely break the application. Modern .NET is a different runtime with different APIs; some Framework APIs do not exist or behave differently. The older the Framework version (e.g. 3.5, 4.0), the more painful the migration. Migration is a real project: retarget the application to .NET 8 LTS or .NET 10 LTS, update deprecated or incompatible APIs, and test thoroughly. Installing modern .NET on a server does not remediate a .NET Framework 4.x dependency. Note: .NET Framework 4.8 on a supported Windows version still receives security updates via Windows Update — migration to modern .NET is recommended for long-term support but is not the only way to stay patched on supported OS versions.",
            TicketText = "- .NET Framework is NOT backward compatible with modern .NET (5+)\r\n  - No in-place upgrade path; switching runtimes without migration will likely break the app\r\n  - Installing modern .NET does not replace or satisfy .NET Framework dependencies\r\n  - Modern .NET is a different runtime; many APIs differ or are missing\r\n  - Migration is a real project: retarget to .NET 8/10 LTS, update APIs, test thoroughly\r\n  - .NET Framework 4.8 on supported Windows still gets WU security updates\r\n  - Main driver for migration: long-term support lifecycle, not immediate patch gap on 4.8",
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
            WordText = "Legacy .NET: .NET Core 3.x (including 3.1 LTS) reached end of support on December 13, 2022. Retarget to .NET 8 or .NET 10 (LTS) with minimal code changes within the modern .NET lineage. The main driver is security: out-of-support versions no longer receive patches.",
            TicketText = "- Legacy .NET Core 3.x: end of support since December 13, 2022\r\n  - Retarget to .NET 8 or .NET 10 (LTS) with minimal code changes\r\n  - Main driver: security patches; out-of-support versions receive none\r\n  - LTS versions get 3 years support",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft .NET Core*",
            WordText = "Modern .NET lineage: .NET Core (1.x–3.x) is the predecessor to .NET 5+ — same family, not .NET Framework. Core 2.x and 3.x are out of support. Retargeting to .NET 8 (LTS) is usually a project file edit and minimal code changes because modern .NET maintains backward compatibility within this lineage. .NET Framework is a separate, incompatible family requiring a full migration project. The main driver is security patches.",
            TicketText = "- .NET Core is the modern .NET lineage (continues as .NET 5+), NOT .NET Framework\r\n  - Retarget to .NET 8 (LTS) with minimal code changes within the modern lineage\r\n  - .NET Framework migration is a separate, incompatible project\r\n  - Main driver: security patches; out-of-support versions receive none",
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
            WordText = "Modern .NET: .NET 6 reached end of support on November 12, 2024 and no longer receives security updates. Retarget applications to .NET 8 LTS (through November 2026) or .NET 10 LTS. Remaining on .NET 6 is an unsupported-state risk, not a patchable finding.",
            TicketText = "- .NET 6 end of support: November 12, 2024 (no further security patches)\r\n  - Retarget apps to .NET 8 LTS or .NET 10 LTS\r\n  - Do not rely on Windows Update for ongoing .NET 6 security fixes\r\n  - Not interchangeable with .NET Framework",
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
            WordText = "Modern .NET: .NET 8 LTS is supported through November 2026. .NET 10 LTS (November 2025) is the newer long-term option. Keep the installed runtime patched via Windows Update or vendor installers. Retargeting between supported LTS releases is usually minimal. End users typically notice nothing; value is security patches.",
            TicketText = "- Modern .NET 8 LTS: supported through November 2026\r\n  - .NET 10 LTS is the newer long-term option\r\n  - Keep current with patch updates via WU or vendor installer\r\n  - Main driver: security patches",
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
            WordText = "Modern .NET (.NET 5 and later, continuing the .NET Core lineage) maintains strong backward compatibility within this family only — not with .NET Framework. Retargeting between modern versions (5→8→10) is usually a project file edit and minimal code changes. The main reason to upgrade is support lifecycle: LTS versions (8, 10) get 3 years; non-LTS get 18 months. End users typically notice nothing; for MSP/client conversations, the value is security patches — not user-facing improvements. Do not install a modern runtime expecting it to satisfy .NET Framework dependencies.",
            TicketText = "- Modern .NET (.NET 5+): backward compatible within modern lineage only — NOT with .NET Framework\r\n  - Retarget to .NET 8 LTS or .NET 10 LTS with minimal code changes\r\n  - Does not replace .NET Framework; Framework apps need a migration project\r\n  - Main driver: security patches (out-of-support versions receive none)\r\n  - LTS versions get 3 years support; non-LTS get 18 months",
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
            WordText = "Modern .NET: .NET 6 reached end of support on November 12, 2024. Retarget to .NET 8 LTS or .NET 10 LTS with minimal code changes. End users typically notice nothing; the driver is security support.",
            TicketText = "- Modern .NET 6: out of support since November 12, 2024\r\n  - Retarget to .NET 8 LTS or .NET 10 LTS with minimal code changes\r\n  - Main driver: security patches no longer available on .NET 6",
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
            WordText = "Modern .NET: .NET 8 LTS is supported through November 2026. .NET 10 LTS (November 2025) is the newer long-term option. Keep current with patch updates. End users typically notice nothing; value is security patches.",
            TicketText = "- Modern .NET 8 LTS: supported through November 2026\r\n  - .NET 10 LTS is the newer long-term option\r\n  - Keep current with patch updates\r\n  - Main driver: security patches",
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
            Pattern = "*Dot Net 6*",
            WordText = "Modern .NET: .NET 6 reached end of support on November 12, 2024 and no longer receives security updates. Retarget applications to .NET 8 LTS or .NET 10 LTS. This is the modern .NET lineage — not .NET Framework.",
            TicketText = "- .NET 6 out of support since November 12, 2024 (no security patches)\r\n  - Retarget apps to .NET 8 LTS or .NET 10 LTS\r\n  - Not interchangeable with .NET Framework",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*ASP.NET Core*",
            WordText = "ASP.NET Core is part of the modern .NET lineage (not .NET Framework). Keep the installed ASP.NET Core hosting/runtime bundle current via Windows Update or the vendor installer matching the app's target band (e.g. 8.0.x). Retarget applications to a supported LTS release (.NET 8 or .NET 10) when the reported runtime is out of support.",
            TicketText = "- ASP.NET Core is modern .NET, not .NET Framework\r\n  - Patch hosting/runtime bundle to latest patch in the same band\r\n  - Retarget out-of-support apps to .NET 8 or .NET 10 LTS\r\n  - Verify version after update on affected hosts",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Google Chrome*",
            GuidanceStyle = RemediationGuidanceStyle.AutoUpdate,
            WordText = "Google Chrome updates via the Google Update service (background scheduled tasks on Windows), not only when a user opens the browser. For this finding, confirm affected hosts reach the ConnectSecure target version during the normal update cycle rather than treating this as a manual patch project. Users may need to relaunch Chrome for an already-downloaded update to apply. If versions remain stale, check update policies and Google Update task health, then consider RMM (if deployed) or ConnectSecure Patch Now on managed hosts.",
            TicketText = "- Auto-updating software: verify first, patch only if stale\r\n  - Confirm installed version against ConnectSecure target on affected hosts\r\n  - Chrome updates via Google Update (background); relaunch may be needed to apply\r\n  - If still behind: check Chrome update policy / Google Update scheduled tasks\r\n  - Last resort on managed hosts: ConnectSecure Patch Now or RMM script (if deployed)\r\n  - ConnectSecure Solution below lists the target version",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Mozilla Firefox*",
            GuidanceStyle = RemediationGuidanceStyle.AutoUpdate,
            WordText = "Mozilla Firefox updates via Mozilla Maintenance Service and background update mechanisms, not only when a user opens the browser. Confirm affected hosts reach the ConnectSecure target version during the normal update cycle. Users may need to relaunch Firefox for an already-downloaded update to apply. If versions stay stale, check update settings and maintenance service status, then consider RMM (if deployed) or ConnectSecure Patch Now on managed hosts.",
            TicketText = "- Auto-updating software: verify first, patch only if stale\r\n  - Confirm installed version against ConnectSecure target on affected hosts\r\n  - Firefox updates in background; relaunch may be needed to apply\r\n  - If still behind: check Firefox update settings / Maintenance Service\r\n  - Last resort on managed hosts: ConnectSecure Patch Now or RMM script (if deployed)\r\n  - ConnectSecure Solution below lists the target version",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft Edge*",
            GuidanceStyle = RemediationGuidanceStyle.AutoUpdate,
            WordText = "Microsoft Edge updates primarily through the Microsoft Edge Update service (scheduled background tasks). Windows Update may also deliver Edge on some systems. Confirm affected hosts reach the ConnectSecure target version during the normal update cycle rather than treating this as a manual patch project. Users may need to relaunch Edge for an already-downloaded update to apply. If versions remain stale, check Edge Update task status and update channel, then consider managed deployment tools.",
            TicketText = "- Auto-updating software: verify first, patch only if stale\r\n  - Confirm installed version against ConnectSecure target on affected hosts\r\n  - Edge updates via Microsoft Edge Update (primary); WU may also apply updates\r\n  - Relaunch may be needed for a downloaded update to take effect\r\n  - If still behind on managed hosts: ConnectSecure Patch Now or RMM script\r\n  - ConnectSecure Solution below lists the target version",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*7-Zip*",
            WordText = "7-Zip can be remediated via ConnectSecure Patch Now (application patch) when agents are online. Verify installed version on affected hosts after the patch job completes; offline agents remain pending until they check in.",
            TicketText = "- Update 7-Zip to the ConnectSecure fix target\r\n  - Deploy via ConnectSecure Patch Now when agent is online\r\n  - Verify version on affected hosts after patch completes\r\n  - Offline agents: pending until next check-in",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Notepad++*",
            WordText = "Notepad++ can be remediated via ConnectSecure Patch Now (application patch) when agents are online. Verify installed version on affected hosts after patching; offline agents remain pending until check-in.",
            TicketText = "- Update Notepad++ to the ConnectSecure fix target\r\n  - Deploy via ConnectSecure Patch Now when agent is online\r\n  - Verify version on affected hosts\r\n  - Offline agents: pending until next check-in",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Adobe Acrobat*",
            WordText = "Adobe Acrobat updates are often per-user or delivered through Adobe's own updater rather than a simple enterprise patch. Check whether ConnectSecure Patch Now applies; otherwise use Adobe's managed update tooling or manual install on affected hosts. Verify version after update.",
            TicketText = "- Update Adobe Acrobat to the ConnectSecure fix target\r\n  - Check ConnectSecure Patch Now applicability first\r\n  - Otherwise use Adobe updater or manual install on affected hosts\r\n  - Verify version after update",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Adobe Reader*",
            WordText = "Adobe Reader updates are often per-user or delivered through Adobe's own updater. Check whether ConnectSecure Patch Now applies; otherwise use Adobe's managed update tooling or manual install on affected hosts. Verify version after update.",
            TicketText = "- Update Adobe Reader to the ConnectSecure fix target\r\n  - Check ConnectSecure Patch Now applicability first\r\n  - Otherwise use Adobe updater or manual install\r\n  - Verify version after update",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*TeamViewer*",
            WordText = "TeamViewer should be updated to the ConnectSecure fix target. When patchable, deploy via ConnectSecure Patch Now on online agents and verify version afterward. Otherwise use TeamViewer's built-in updater, RMM (if deployed), or vendor installer.",
            TicketText = "- Update TeamViewer to the ConnectSecure fix target\r\n  - Use ConnectSecure Patch Now when patchable and agent is online\r\n  - Verify version on affected hosts\r\n  - Fallback: vendor updater, RMM (if deployed), or manual install",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Zoom*",
            WordText = "Zoom client updates should match the ConnectSecure fix target. Deploy via ConnectSecure Patch Now when patchable, or use Zoom's installer/updater or RMM (if deployed). Verify version on affected hosts after update.",
            TicketText = "- Update Zoom to the ConnectSecure fix target\r\n  - Use ConnectSecure Patch Now when patchable\r\n  - Otherwise vendor updater, RMM (if deployed), or manual install\r\n  - Verify version on affected hosts",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Cisco Webex*",
            WordText = "Cisco Webex should be updated to the ConnectSecure fix target. When patchable, use ConnectSecure Patch Now on online agents; otherwise deploy via vendor installer or RMM. Verify version after update.",
            TicketText = "- Update Cisco Webex to the ConnectSecure fix target\r\n  - Deploy via ConnectSecure Patch Now when patchable\r\n  - Verify version on affected hosts\r\n  - Fallback: vendor installer or RMM",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft Defender*",
            WordText = "Microsoft Defender platform and signature updates are normally delivered through Windows Update or Microsoft Defender management policies. Verify the reported engine/platform version against Windows Update history or RMM patch status when patch management is deployed for the client.",
            TicketText = "- Update Microsoft Defender platform/signatures\r\n  - Verify via Windows Update or Defender management policy\r\n  - Check RMM patch status on affected hosts if patch management is deployed\r\n  - Confirm engine/platform version after update",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*Microsoft 365*",
            WordText = "Microsoft 365 Apps updates are typically delivered through Office Click-to-Run / Microsoft 365 update channels rather than generic application patching. Verify update channel configuration and deployment in Microsoft 365 admin or Intune; use RMM or Intune Office update workflows where configured.",
            TicketText = "- Update Microsoft 365 Apps to latest supported build\r\n  - Verify Office update channel (Current Channel, etc.)\r\n  - Use Intune or RMM Office update workflow if configured\r\n  - Confirm version on affected hosts",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*OS-OUT-OF-SUPPORT*",
            GuidanceStyle = RemediationGuidanceStyle.Infrastructure,
            WordText = "The operating system on affected hosts is out of vendor support. This is an infrastructure upgrade project — migrate to a supported Windows release or retire the system. Patching individual applications will not remediate OS-level out-of-support findings.",
            TicketText = "- Operating system is out of vendor support\r\n  - Plan upgrade to supported Windows version or retire system\r\n  - Not resolved by application patching alone\r\n  - Treat as infrastructure project",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*TLSv1.0*",
            GuidanceStyle = RemediationGuidanceStyle.Configuration,
            WordText = "TLS 1.0 is deprecated and should be disabled on servers, applications, and devices where it is explicitly enabled. On many current Windows 10/11 clients, Schannel may still allow TLS 1.0 by default — confirm whether the finding is on an application, server role, appliance, or OS hardening target before changing defaults. Disable TLS 1.0 in favor of TLS 1.2 or 1.3. May require vendor guidance or registry/Schannel hardening on Windows.",
            TicketText = "- Disable TLS 1.0 on affected services where it is enabled\r\n  - Prefer TLS 1.2 or 1.3 only\r\n  - Confirm scope: app, server, appliance, or OS Schannel setting\r\n  - Windows: legacy TLS may still be enabled by default on some in-market releases\r\n  - Test dependent applications after change",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*TLSv1.1*",
            GuidanceStyle = RemediationGuidanceStyle.Configuration,
            WordText = "TLS 1.1 is deprecated and should be disabled on servers, applications, and devices where it is explicitly enabled. On many current Windows 10/11 clients, Schannel may still allow TLS 1.1 by default — confirm whether the finding is on an application, server role, appliance, or OS hardening target. Disable TLS 1.1 in favor of TLS 1.2 or 1.3. May require vendor guidance or Schannel hardening on Windows.",
            TicketText = "- Disable TLS 1.1 on affected services where it is enabled\r\n  - Prefer TLS 1.2 or 1.3 only\r\n  - Confirm scope: app, server, appliance, or OS Schannel setting\r\n  - Windows: legacy TLS may still be enabled by default on some in-market releases\r\n  - Test dependent applications after change",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*SSL_CA_Expired*",
            WordText = "An SSL/TLS certificate on the affected service has expired. Renew the certificate with a trusted public CA or appropriate internal PKI, install the new cert chain, and verify clients no longer receive certificate errors. Include intermediate certificates as needed.",
            TicketText = "- Renew expired SSL/TLS certificate\r\n  - Use trusted public CA or internal PKI as appropriate\r\n  - Install full cert chain on the service\r\n  - Verify no client certificate warnings remain",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*SSL_SelfSigned_CA*",
            WordText = "The affected service is using a self-signed or untrusted certificate authority. Replace with a certificate issued by a trusted CA (public or internal PKI) unless this is a deliberate lab/dev exception documented and segmented.",
            TicketText = "- Replace self-signed/untrusted certificate\r\n  - Issue cert from trusted public CA or internal PKI\r\n  - Document and segment deliberate dev/lab exceptions only\r\n  - Verify trust chain on clients",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*SSL_CA_Validity*",
            WordText = "Review SSL/TLS certificate validity on the affected service — expiration horizon, chain completeness, and trusted issuer. Renew or replace certificates before expiry and ensure intermediate CAs are installed correctly.",
            TicketText = "- Review SSL/TLS certificate validity\r\n  - Check expiration date and full chain\r\n  - Renew or replace before expiry\r\n  - Ensure intermediate CAs are installed",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*smb-protocols*",
            WordText = "Legacy or insecure SMB protocol versions may be enabled on the affected host or appliance. Disable SMBv1 where possible and prefer SMB signing/encryption per Microsoft hardening guidance. Verify dependent file/print shares after changes.",
            TicketText = "- Harden SMB configuration\r\n  - Disable SMBv1 where possible\r\n  - Enable SMB signing/encryption per vendor guidance\r\n  - Test file/print shares after changes",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*http-server-header*",
            WordText = "HTTP server banner/header disclosure is an informational finding. Where practical, configure the web server or appliance to suppress or minimize version banners (IIS, Apache, nginx, embedded devices). This is configuration hardening, not a version upgrade.",
            TicketText = "- Reduce HTTP server version banner disclosure\r\n  - Configure web server/appliance to minimize headers\r\n  - Configuration hardening — not necessarily a version upgrade\r\n  - Verify service still functions after change",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*whois-ip*",
            WordText = "WHOIS/IP information disclosure is typically an external scan informational finding. Review whether the exposed asset should be reachable from the internet; restrict exposure, use CDN/WAF masking where appropriate, and document accepted risk if the service must remain public.",
            TicketText = "- Review internet-exposed asset and WHOIS/IP disclosure\r\n  - Restrict exposure or segment if not required publicly\r\n  - Consider CDN/WAF where appropriate\r\n  - Document accepted risk if exposure is required",
            IsDefault = false
        },
        new RemediationRule
        {
            Pattern = "*",
            WordText = "Determine what the device or software is (use Product/OS and affected hosts). Review the manufacturer's security advisories and vulnerability data for patches or firmware updates. Consider configuration mitigations (e.g. network segmentation, hardening) where patching is not immediately possible. If the client has RMM or scripting available, deploy updates via patch management or scripts; otherwise, manual updates may be required.",
            TicketText = "- Determine device/software identity (Product/OS, affected hosts)\r\n  - Review manufacturer security advisories and vulnerability data\r\n  - Check for firmware updates or patches\r\n  - Consider configuration mitigations where patching not possible\r\n  - Deploy via RMM or scripting if available for the client; otherwise manual updates",
            IsDefault = true
        },
    ];
}
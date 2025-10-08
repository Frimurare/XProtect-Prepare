==========================================================
Enhanced Milestone XProtect Configuration Script v1.9 - User Guide
==========================================================

## HOW TO RUN THIS SCRIPT (Quick Start)

**SCRIPT VERSION NOTICE:** This is the PowerShell script version (.ps1) - completely transparent and "open like a book" so you can see exactly what it does. This version won't trigger overly sensitive email filters or antivirus programs that sometimes get nervous about compiled executables. You can read every line of code and verify it's safe before running.

**To run the script:**

**Option 1 - Using the batch file (Recommended):**
1. Extract all files from the ZIP to a folder
2. Right-click "launch.bat"
3. Select "Run as administrator"
4. The script will start automatically

**Option 2 - Direct PowerShell method:**
1. Extract all files from the ZIP to a folder
2. Right-click "Xprotect-Pre-Check-2025-10-06-v1.9.ps1"
3. Select "Run with PowerShell"
4. If prompted about execution policy, type "Y" and press Enter
5. Follow the menu options
6. If you have trouble just fix execution policies temporary ( in powershell, Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass )

**Important:** The script MUST run as administrator to configure system settings.

**After running:** Check the detailed logs in C:\Milestonecheck-ulfh\ to see exactly what was configured.

==========================================================

## Hello there, and welcome to the Enhanced Milestone XProtect Configuration Script v1.9!

This comprehensive tool was created by me, Ulf Holmström, Happy Problem Solver at Manvarg AB (Ex Solution Engineer at Milestone Systems) in Sweden. I built this script to make deploying and optimizing Milestone XProtect systems as smooth as possible. After years of watching technicians spend hours on repetitive configuration tasks, I thought, "There's got to be a better way!" So here we are – your new best friend for XProtect deployments.

## What does this enhanced script do?

Think of this script as your personal XProtect configuration assistant that handles all the tedious, time-consuming tasks that nobody enjoys doing manually:

1. **Secure Antivirus Protection:** Automatically sets up recommended Windows Defender exclusions for XProtect – BUT ONLY ON STORAGE DRIVES. The script now includes enhanced security by NEVER adding antivirus exceptions to your OS drive (C:), keeping your system fully protected while optimizing storage drives for video recording performance.

2. **Storage Drive Optimization:** Checks your storage drives for proper 64KB block size allocation and automatically disables indexing on storage drives (but keeps it on your OS drive, naturally). This is crucial for smooth 24/7 video recording.

3. **Windows Bloatware Removal:** Gets rid of all that unnecessary stuff Windows comes with – BUT NOW SMARTER! The script removes:
   - Standard bloatware (Cortana, Solitaire, Weather, etc.)
   - **Microsoft Teams (BOTH new app AND classic versions)**
   - **Outlook App (new Windows Outlook)**
   - **Office Hub/365 installer (the trial installer icon)**
   - OneDrive
   
   **IMPORTANT:** Xbox apps are NO LONGER removed by default because they can cause popup issues about missing screen recording components. If you really need to remove them, there's a separate option with clear warnings.

4. **System Performance Tuning:** Sets your system to Never Sleep mode (critical for 24/7 recording!), disables unnecessary Windows services, optimizes Remote Desktop for management, and keeps the essential tools you actually need.

5. **Windows Update Management:** Gives you control over automatic updates – enable them for security or disable them for stability during critical recording periods. Your choice. Windows Update is disabled LAST to prevent installation failures.

6. **NTP Time Server for Cameras (CRITICAL!):** **NEW IN V1.9:** Configures Windows Time Service as a reliable NTP (Network Time Protocol) server for your CCTV cameras. This is ABSOLUTELY CRITICAL for surveillance systems! The script:
   - Configures Windows Time Service as a reliable NTP server
   - Syncs the server time from Swedish NTP pool (se.pool.ntp.org)
   - Opens firewall (UDP port 123) for camera connections
   - Enables forensically accurate timestamps on all recordings
   
   **Why this matters:** Inaccurate timestamps can make your recordings legally inadmissible! Courts require precise, synchronized timestamps. This configuration ensures all your cameras have the exact same time, synchronized to atomic time sources.

7. **SNMP Monitoring Setup (OPTIONAL):** SNMP installation is completely optional in the Complete Setup. Most users don't need it! SNMP is only useful if you have network monitoring systems like PRTG, Nagios, or Zabbix. The script now explains what SNMP is and when you need it, so you can make an informed decision.

8. **Milestone PowerShell Tools Installation:** Automatically downloads and installs the official Milestone PowerShell Tools from PowerShell Gallery. These powerful cmdlets let you manage XProtect systems programmatically.

9. **Complete Automated Setup:** The crown jewel – a one-click option that runs everything automatically in the CORRECT ORDER. **INCLUDES NTP SERVER CONFIGURATION** for camera time sync! You'll still be asked whether you want optional SNMP installation.

10. **Comprehensive Logging:** Detailed logs with timestamps, plus automatic registry backups before any changes.

## What's NEW in Version 1.9?

### NTP Time Server - CRITICAL for CCTV Systems!
This is arguably the most important feature addition since the script was created. Version 1.9 adds comprehensive NTP (Network Time Protocol) server configuration, which is ABSOLUTELY CRITICAL for surveillance systems.

**The Problem:**
Many CCTV installations have cameras with drifting clocks. Camera A might say 14:32:15, while Camera B says 14:31:48, and Camera C says 14:33:02. When an incident occurs, you're trying to correlate events across multiple cameras, but the timestamps don't match. Even worse - inaccurate timestamps can make your recordings legally inadmissible in court!

**The Solution in v1.9:**
The script now configures your XProtect server as a reliable NTP time server that your cameras can synchronize to. This ensures:
- **Forensic Accuracy:** All cameras have synchronized timestamps to the exact millisecond
- **Legal Compliance:** Timestamps are traceable to atomic time sources (se.pool.ntp.org)
- **Event Correlation:** You can accurately correlate events across multiple cameras
- **Audit Trail:** Investigators can trust your timeline of events

**What it does:**
- Configures Windows Time Service (w32time) as a reliable NTP server
- Points the server to Swedish atomic time sources (0.se.pool.ntp.org, 1.se.pool.ntp.org, etc.)
- Opens Windows Firewall for NTP (UDP port 123) to allow camera connections
- Sets the server as "reliable" so cameras trust it as an authoritative time source
- Verifies the configuration and displays status information

**How to use it:**
1. Run the script and choose option 7 (Configure NTP Time Server)
2. After configuration, point your cameras to use this server's IP address as their NTP server
3. All cameras will now maintain forensically accurate time

**Included in Complete Setup:**
The Complete System Setup (option 8) now includes NTP configuration automatically. It's too important to be optional - every CCTV system needs accurate time synchronization!

**Technical Details:**
- Uses standard Windows Time Service (no third-party software needed)
- More reliable than NetTime or other third-party solutions for serving time
- Compatible with all IP cameras that support NTP (which is basically all of them)
- Works on Windows 11, Windows Server 2022, and Windows Server 2025

**Why this matters more than you think:**
I've seen cases where critical evidence was thrown out of court because the timestamps were inaccurate and couldn't be proven to be synchronized. Insurance claims denied because investigators couldn't establish a reliable timeline. Don't let this happen to you!

## What's NEW in Version 1.8?

### SNMP Made Optional (Big Improvement!)
Many users contacted me saying they were running the complete setup but didn't know what SNMP was or whether they needed it. After explaining it, most realized they didn't need it at all! Version 1.8 fixes this:

**Changes to SNMP handling:**
- **SNMP is now OPTIONAL in Complete System Setup** – you'll be asked if you want it
- **Better explanation of what SNMP is** – the script now explains it's for network monitoring
- **Clear guidance on when you need it** – only install if you use tools like PRTG, Nagios, Zabbix
- **Time warning added** – "This may take a while" message during installation
- **Still available as standalone option** – Menu option 5 for manual installation

**What is SNMP?**
SNMP (Simple Network Management Protocol) is a networking protocol that allows monitoring systems to collect performance data from your servers. Think of it as a way for network monitoring tools to "talk" to your XProtect server and get health information.

**Do you need SNMP?**
You ONLY need SNMP if you have:
- Network monitoring tools like PRTG, Nagios, Zabbix, SolarWinds, etc.
- Enterprise management platforms that use SNMP for server monitoring
- A requirement from your IT department to enable SNMP for monitoring

**If you're unsure, you probably don't need it!** XProtect works perfectly fine without SNMP. It's only needed for external monitoring systems, not for XProtect itself.

### Why This Matters
The old Complete Setup would install SNMP automatically, which meant:
- Longer setup time (SNMP installation can take several minutes)
- Confusion about what was being installed
- Unnecessary features on most systems

Now in v1.8, you get a clean, fast setup by default, with the option to add SNMP only if you actually need it!

## Why all these optimizations matter

Let me paint you a picture. You've just deployed a beautiful new XProtect system with 64 cameras recording 24/7. Everything looks great during testing. But three weeks later, you get a call: "The system is sluggish, and we're missing recordings!"

What happened? Windows Search is frantically indexing millions of video files. The antivirus is scanning every single frame that gets written to disk. The system went to sleep at 3 AM and missed six hours of footage. OneDrive is trying to sync 2TB of surveillance video to the cloud. Teams is running in the background eating resources.

This script prevents all of that nonsense before it starts. It's like baby-proofing your house, but for enterprise video surveillance.

## The "Never Sleep" mode – why it's critical

Here's something that keeps me up at night (ironically): surveillance servers that go to sleep. Windows, bless its heart, thinks it's being helpful by putting computers to sleep to save energy. That's great for laptops, but catastrophic for 24/7 recording systems.

Picture this: It's 2:47 AM, nothing exciting is happening, so Windows thinks, "I'll just take a little nap." Suddenly, 32 cameras stop recording. The break-in happens at 3:15 AM, but you've got nothing because your "always-on" surveillance system was having a snooze.

This script sets your power management to High Performance mode and configures it to NEVER sleep. Your cameras will thank you, your clients will thank you, and you'll sleep better knowing your systems are actually awake when they need to be.

## About the "All-in-One" complete setup

The Complete System Setup option is my personal favorite. I designed it for those times when you're deploying multiple identical systems and you want consistency. Hit option 8, answer a few questions, and come back to a fully optimized XProtect-ready system.

**Version 1.9 Execution Order:**
1. Interactive drive selection for antivirus exceptions
2. Automatic protection of OS drive (C:) - NEVER excluded from antivirus scanning
3. Storage optimization with automatic indexing disable (storage drives only)
4. Complete bloatware cleanup (Teams, Outlook, Office Hub)
5. Xbox apps preservation (no longer removed - prevents popup issues)
6. Full performance optimization (power, services, RDP)
7. **NTP Time Server configuration (AUTOMATIC - too critical to skip!)**
8. Milestone PowerShell Tools installation
9. **SNMP installation - OPTIONAL** (you'll be asked if you want it)
10. Windows Update management (LAST STEP to prevent installation failures)
11. Comprehensive logging

It even times itself and gives you a nice summary at the end. Very satisfying.

## Enhanced Security

**CRITICAL SECURITY FEATURE:** The script includes enhanced protection for your operating system:

- **NEVER** adds antivirus exceptions to the C: drive - your OS stays fully protected
- **Interactive drive selection** - even in complete setup, you choose which storage drives to optimize
- **Automatic filtering** - attempts to include C: drive are automatically blocked with warnings
- **Clear messaging** - always shows which drives are being excluded and why
- **Security-first approach** - performance optimization without compromising system security

This means your Windows installation, system files, and installed programs remain fully protected by antivirus scanning, while your video storage drives get the performance optimizations they need.

## Is this safe to use?

Absolutely. I've put my name on this tool, and my reputation matters to me. The script creates registry backups before making any changes, logs everything it does, and includes restore options for the optimizations.

If you're the cautious type (and you should be in IT!), you can run each section individually to see exactly what it does. Or drop me an email at ulf@manvarg.se and I'll happily walk you through any part of it.

The script version is completely transparent PowerShell code – no tricks, no hidden functionality, just straightforward system optimization with enhanced security.

## A bit of personal background

This enhanced version came together during a particularly rainy late summer in 2025, as autumn's first storms lashed against my office windows here in Sweden. Fueled by way too much coffee and a determination to solve the repetitive deployment headaches I kept hearing about from technicians in the field.

Version 1.8 represents direct feedback from users who said "We don't know what SNMP is" or "The complete setup takes forever" - listening to users makes better software!

## How to use this tool

**Getting Started:**
1. Right-click the script file and select "Run as administrator" (this is essential – the script needs admin privileges)
2. The script will show you a friendly menu with 9 clear options

**For Individual Tasks:**
- Choose options 1-7 to run specific configurations
- Each section guides you through the process step-by-step
- **Option 5 (SNMP)** now includes a clear explanation of what it is and whether you need it
- **Option 7 (NTP Time Server)** - NEW! Configures camera time synchronization
- Check the logs afterward to see exactly what was done

**For Complete Deployment:**
- Choose option 8 for the full automated setup
- Select which storage drives to optimize (C: drive automatically protected)
- **NTP Time Server will be configured automatically** (too critical to skip!)
- **You'll be asked if you want SNMP installation** (most users can say No)
- Windows Update will be kept enabled during installations automatically
- Answer questions about final Windows Update configuration (enable/disable/skip)
- Watch as it systematically optimizes your system in the correct order
- Review the comprehensive summary when it's done

**Bloatware Removal Options:**
- Option 1: Remove all standard bloatware (Xbox apps excluded)
- Option 2: Select specific apps manually
- Option 3: OneDrive only
- Option 4: Teams + Outlook + Office Hub (the cloud services)
- Option 5: Full cleanup (everything except Xbox)
- Option 6: Xbox apps removal (REQUIRES typing "YES" + shows warnings)

**After Running:**
- Always restart the system to ensure all changes take effect
- Check the detailed logs in C:\Milestonecheck-ulfh\
- Install your XProtect software on the optimized system
- **IMPORTANT: Configure your cameras to use this server's IP as their NTP server**
- Configure camera connections and storage
- Verify time synchronization is working on all cameras

**Pro Tips:**
- Run this BEFORE installing XProtect for best results
- Use the complete setup for standardized deployments
- **ALWAYS configure NTP** - accurate timestamps are legally critical for CCTV
- **Skip SNMP unless you have monitoring tools** - most users don't need it
- Keep the log files for documentation and troubleshooting
- The registry backups are in C:\Milestonecheck-ulfh\RegistryBackups if you ever need to restore something
- Don't worry about Xbox apps unless you specifically need to remove them
- **After setup, verify NTP with:** `w32tm /query /status`

## Support and feedback

Questions? Suggestions? Found a bug? I'd love to hear from you! Reach out to me at ulf@manvarg.se. I'm always interested in ways to make this tool even more useful for the XProtect community.

This tool is provided as a free resource to make everyone's life easier. Use it, share it with colleagues, and may all your XProtect deployments be smooth and trouble-free!

Thanks for using the Enhanced Milestone XProtect Configuration Script v1.9!

Have a fantastic day, and happy deploying!

**/Ulf Holmström**  
**Happy Problem Solver, Manvarg AB**  
**Ex Solution Engineer at Milestone Systems**  
**October 2025**

P.S. – If this tool saves you even half the time it took me to write it, then it was worth every line of code. Version 1.9's NTP server configuration ensures your recordings have forensically accurate timestamps - which could save you in court!

==========================================================

## VERSION HISTORY AND CHANGELOG

### Version 1.9 (October 2025) - NTP TIME SERVER FOR CAMERAS (CRITICAL!)

**MAJOR FEATURE: Network Time Protocol Server Configuration**
- **ADDED:** Complete NTP Time Server configuration for camera synchronization
- **ADDED:** Windows Time Service (w32time) configured as reliable NTP server
- **ADDED:** Automatic sync to Swedish NTP pool (0.se.pool.ntp.org, 1.se.pool.ntp.org, etc.)
- **ADDED:** Firewall configuration for NTP (UDP port 123 inbound)
- **ADDED:** Verification and status display of NTP configuration
- **ADDED:** Menu option 7 for standalone NTP configuration
- **IMPROVED:** Complete System Setup now includes automatic NTP configuration
- **IMPROVED:** Better documentation about why accurate time is critical for CCTV
- **UPDATED:** Menu structure (now 10 options instead of 9)

**Why This Update Is Critical:**
Time synchronization is one of the most overlooked aspects of CCTV systems, yet it's absolutely critical for:
- **Legal Compliance:** Courts require precise, synchronized timestamps for evidence admissibility
- **Forensic Analysis:** Investigators need to correlate events across multiple cameras
- **Event Timeline:** Accurate reconstruction of incidents requires synchronized time
- **Insurance Claims:** Many claims are denied due to inaccurate or unsynchronized timestamps

Previous solutions like NetTime were unreliable for serving time to cameras. Windows Time Service (w32time) is built into Windows, highly reliable, and industry-standard. Version 1.9 properly configures it as an NTP server that cameras can trust.

**Technical Implementation:**
- Sets AnnounceFlags to 5 (reliable time source)
- Enables NtpServer in Windows Time Service
- Configures manual peer list with Swedish NTP pool servers
- Opens firewall using netsh advfirewall for maximum compatibility
- Works on Windows 11, Server 2022, and Server 2025
- Includes comprehensive error handling and verification

**User Experience:**
- Standalone menu option with full explanation of importance
- Automatic configuration in Complete Setup (not optional - too critical!)
- Clear post-configuration instructions for camera setup
- Status verification and troubleshooting commands provided

**Installation Time:**
NTP configuration is very fast (5-10 seconds) compared to SNMP installation (2-5 minutes), making it perfect for inclusion in the Complete Setup without adding significant time.

### Version 1.8 (October 2025) - SNMP MADE OPTIONAL

**MAJOR IMPROVEMENT: User-Driven SNMP Configuration**
- **CHANGED:** SNMP is now OPTIONAL in Complete System Setup (option 7)
- **ADDED:** Clear explanation of what SNMP is and when you need it
- **ADDED:** "This may take a while" warning during SNMP installation
- **IMPROVED:** Complete Setup now asks if you want SNMP before installing
- **IMPROVED:** Menu option 5 updated with better SNMP explanation
- **IMPROVED:** More user-friendly guidance about SNMP requirements
- **UPDATED:** Documentation significantly expanded with SNMP information

**Why This Update Matters:**
User feedback showed that many people were running the complete setup without knowing what SNMP was or whether they needed it. After explaining, most realized they didn't need it! SNMP is only necessary if you use network monitoring systems like PRTG, Nagios, or Zabbix. Making it optional:
- Speeds up the complete setup process for most users
- Reduces confusion about what's being installed
- Gives users informed choice about optional features
- Still available for those who actually need network monitoring

**Technical Details:**
- SNMP installation moved to optional Phase 5 in Complete Setup
- User prompted with Y/N choice for SNMP installation
- Clear explanation displayed before prompting
- Installation skipped if user chooses No
- Summary report shows whether SNMP was installed or skipped
- Standalone SNMP option (menu #5) enhanced with detailed explanation

**Installation Time Note:**
SNMP installation can take 2-5 minutes depending on Windows Update connectivity. The script now warns users about this to set proper expectations.

### Version 1.7 (October 2025) - CLOUD SERVICES CLEANUP & XBOX PROTECTION

**MAJOR IMPROVEMENTS: Better Bloatware Removal**
- **IMPROVED:** Microsoft Teams removal now handles BOTH new Teams app (AppX) AND classic Teams installation
- **ADDED:** Outlook App removal (new "Outlook for Windows" that Microsoft is pushing)
- **ADDED:** Office Hub/365 installer removal (removes the trial Office icon)
- **CHANGED:** Xbox apps are NO LONGER removed by default (prevents popup issues)
- **ADDED:** Separate Xbox removal option (#6) with clear warnings about consequences
- **ADDED:** Confirmation requirement (type "YES") for Xbox removal to prevent accidents
- **IMPROVED:** Enhanced bloatware removal menu with better organization
- **IMPROVED:** More detailed removal reporting and logging
- **UPDATED:** Complete System Setup now explicitly mentions Xbox preservation

**Why This Update Mattered:**
Field feedback showed that Xbox Game Bar removal was causing annoying popup messages about missing screen recording components on some systems. Additionally, Microsoft's push toward cloud services meant more Teams, Outlook, and Office Hub installers on fresh Windows installations.

### Version 1.6 (October 2025) - CRITICAL BUG FIX

**MAJOR FIX: Installation Order Correction**
- **FIXED:** Windows Update is now disabled LAST in the execution order
- **FIXED:** SNMP installation now happens BEFORE Windows Update can be disabled
- **FIXED:** Milestone PowerShell Tools installation now happens BEFORE Windows Update can be disabled
- **IMPROVED:** Complete System Setup follows correct dependency order
- **ADDED:** Automatic Windows Update enablement during installation phases
- **ADDED:** Warning messages about Windows Update requirements
- **ADDED:** Clear phase indicators showing when installations happen

### Version 1.5 (September 2025) - Security Enhancement

**MAJOR IMPROVEMENT: OS Drive Protection**
- **ADDED:** Automatic exclusion of C: drive from antivirus exceptions for enhanced security
- **ADDED:** Interactive drive selection in Complete System Setup
- **ADDED:** Clear security warnings about OS drive protection
- **IMPROVED:** Better user guidance for storage drive selection

### Version 1.4 and Earlier

Previous versions included:
- Basic antivirus exception configuration
- Storage optimization tools
- Bloatware removal functionality (including Xbox apps - later changed in v1.7)
- System performance optimization
- Windows Update management
- SNMP installation capabilities
- Milestone PowerShell Tools installation
- Complete automated setup option
- Registry backup and logging features

**Note:** Versions prior to 1.5 did not include OS drive protection. Versions prior to 1.6 had incorrect installation ordering. Versions prior to 1.7 removed Xbox apps by default. Versions prior to 1.8 made SNMP mandatory in complete setup. Upgrading to v1.8 is highly recommended for all users.

==========================================================

## FREQUENTLY ASKED QUESTIONS

**Q: Why is NTP Time Server configuration so important?**
A: **Accurate time synchronization is CRITICAL for CCTV systems!** Here's why:
- **Legal Admissibility:** Courts can reject footage with inaccurate or unsynchronized timestamps
- **Forensic Analysis:** Investigators need precise timeline correlation across multiple cameras
- **Event Reconstruction:** You can't accurately reconstruct incidents if cameras have different times
- **Regulatory Compliance:** Many regulations require accurate, traceable time sources
- **Insurance Claims:** Claims are often denied if timestamps can't be verified as accurate

Without NTP, cameras drift - sometimes by minutes or even hours. This makes your expensive surveillance system nearly useless for its primary purpose: providing legally admissible evidence.

**Q: How do I configure my cameras to use the NTP server?**
A: After running the script:
1. Find the "Time" or "Date & Time" settings in each camera's web interface
2. Set NTP Server to your XProtect server's IP address
3. Enable NTP synchronization
4. Verify that time is syncing correctly

Most IP cameras sync time every few hours automatically once configured.

**Q: Can I use an external NTP server instead?**
A: You could, but there are advantages to using your XProtect server:
- **No internet required:** Cameras don't need internet access
- **Network isolation:** Cameras stay on your isolated CCTV network
- **Reliability:** No dependency on external services
- **Centralized control:** One time source for your entire system

The XProtect server itself syncs to Swedish atomic time sources, so it's just as accurate as using external NTP directly.

**Q: Will this work with all camera brands?**
A: Yes! NTP is an industry standard protocol supported by virtually all IP cameras from all manufacturers (Axis, Hikvision, Dahua, Hanwha, etc.). If your camera has network connectivity, it almost certainly supports NTP.

**Q: Do I need SNMP?**
A: **Most users DON'T need SNMP!** You only need it if you have network monitoring tools like:
- PRTG Network Monitor
- Nagios
- Zabbix
- SolarWinds
- Other enterprise monitoring systems

If you don't use any of these tools, you can safely skip SNMP installation. XProtect works perfectly without it.

**Q: What exactly is SNMP?**
A: SNMP (Simple Network Management Protocol) is a networking protocol that allows monitoring systems to collect performance data from your servers. Think of it as a way for external monitoring tools to check your server's health (CPU usage, disk space, etc.). It's NOT required for XProtect to function.

**Q: How long does SNMP installation take?**
A: SNMP installation typically takes 2-5 minutes, depending on your Windows Update connectivity. The script will warn you about this before starting. If you skip SNMP, the complete setup finishes much faster.

**Q: Why don't you remove Xbox apps by default anymore?**
A: After extensive field testing, we discovered that removing Xbox Game Bar causes Windows to display popup messages about "missing components needed for screen recording" on some systems. While most recording servers don't use screen recording, the popups were more annoying than the apps themselves. If you specifically need to remove them, use option 6 in the bloatware menu.

**Q: What's the difference between the new Teams and classic Teams?**
A: Microsoft is transitioning from the classic Teams (installed as a regular application) to the new Teams (installed as an AppX package). Some Windows installations have one, some have both. Version 1.7+ removes both types automatically.

**Q: Should I remove the Outlook App?**
A: Yes, for recording servers. The "Outlook for Windows" app is Microsoft's new email client that they're pushing. It's unnecessary on a dedicated recording server and just consumes resources.

**Q: Will this script work on Windows Server editions?**
A: Yes! The script works on both Windows desktop and server editions. Some bloatware apps may not be present on Server editions, and the script will simply skip them.

**Q: Can I run this script on a system that already has XProtect installed?**
A: Yes, but it's better to run it BEFORE installing XProtect. If XProtect is already installed, the script will still work, but you should test thoroughly afterward and may need to restart the XProtect services.

**Q: What if I accidentally said Yes to SNMP but don't need it?**
A: SNMP doesn't hurt anything if installed - it just sits there unused. If you really want to remove it, you can use Windows Features to uninstall the SNMP Client capability.

**Q: How do I know which drives are storage drives vs. the OS drive?**
A: The C: drive is always your OS drive (where Windows is installed). Any other drive letters (D:, E:, F:, etc.) are typically storage drives. The script automatically protects C: from antivirus exceptions.

**Q: Can I install SNMP later if I skip it now?**
A: Absolutely! Just run the script again and choose option 5 (Enable SNMP Capabilities). You don't need to re-run the complete setup.

**Q: What happens if my XProtect server's time is wrong?**
A: The script configures the server to sync from Swedish atomic time sources (se.pool.ntp.org), which are accurate to within milliseconds of UTC. After running the script, the server will have accurate time. You can verify with: `w32tm /query /status`

**Q: My cameras are still not syncing time. What should I check?**
A: Troubleshooting steps:
1. Verify Windows Time service is running: `Get-Service w32time`
2. Check NTP status: `w32tm /query /status`
3. Verify firewall allows UDP port 123: Check Windows Firewall rules
4. Ensure cameras can reach the server's IP address (ping test)
5. Check camera's NTP configuration (correct IP address entered?)
6. Some cameras need a reboot after changing NTP settings

**Q: Can I use this NTP server for other devices too?**
A: Yes! Any device that supports NTP can sync from this server - other computers, network equipment, access control systems, etc. That's one of the benefits of having a local NTP server.

**Q: Does this replace the need for a GPS time source?**
A: For most installations, yes. The server syncs to atomic time sources via internet, which is sufficient for legal requirements in most jurisdictions. GPS time sources are only needed for extremely critical applications or sites without reliable internet access.

==========================================================
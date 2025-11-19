# VScanMagic v3.1.2 Release Notes

## 🎉 Release Highlights

VScanMagic v3.1.2 introduces **enhanced system tracking** with username and IP address integration, improved **Microsoft Teams remediation guidance**, and **Excel color key improvements** for better visual clarity.

## ✨ New Features

### 👤 Enhanced System Information in Reports
- **Username tracking** - Reports now display usernames alongside hostnames when available
- **Top Ten Vulnerabilities Report** displays affected systems as "hostname (username)"
- **Automatic detection** of username columns in Excel files (supports multiple column name variations)
- **Graceful fallback** - Works with or without username data in source files

### 🌐 IP Address Integration in Ticket Instructions
- **IP addresses included** for each affected system in Ticket Instructions
- **Clear format**: Each system listed as "hostname - IP address"
- **Easier system identification** and troubleshooting for technicians
- **Bulleted list format** for improved readability

### 💾 Improved Data Structures
- **Complete system objects** stored with hostname, IP address, and username
- **Better system tracking** across consolidated vulnerability entries
- **Enhanced data consolidation** for accurate reporting
- **Backward compatible** - All changes work with existing Excel files

## 🔧 Improvements

### 🎯 Microsoft Teams Remediation Guidance
- **Detailed cleanup instructions** for Teams Classic vulnerabilities
- **Script path included** in both Word reports and Ticket Instructions:
  - `Select Scripts > RR - Custom > RR - Custom - R-Security Remediation > R-Security - Teams Classic Cleanup Remediation in RMM`
- **Clear guidance** on cleaning up unused user profile installed versions
- **Step-by-step instructions** for technicians

### 🎨 Excel Color Key Enhancement
- **Changed "Do not touch" color** from red to orange in Excel pivot table
- **Prevents confusion** with red conditional formatting for EPSS scores > 0.01
- **Improved visual clarity** in the color-coded key
- **Better color distinction** between key and data highlighting

## 📋 Technical Changes

### Modified Components
- **Column mapping system** - Added Username field detection with multiple column name variations
- **`Read-SheetData` function** - Enhanced to capture username field from Excel data
- **`Get-Top10Vulnerabilities` function** - Stores complete system objects (hostname, IP, username)
- **Word report formatter** - Updated to display "hostname (username)" format
- **Ticket Instructions formatter** - Updated to display "hostname - IP address" format
- **Excel color key** - Changed ColorIndex from 3 (red) to 46 (orange) for "Do not touch" entry

### Username Column Detection
Automatically detects username columns with these variations:
- Username
- User Name
- User
- Account
- Login
- Login Name

### Backward Compatibility
- ✅ All changes are backward compatible
- ✅ Reports function correctly with or without username/IP data
- ✅ Existing Excel files work without modification
- ✅ No configuration changes required

## 🔄 Migration from v3.1.1

- **No breaking changes**
- **No action required** for existing workflows
- **New system information appears automatically** when available in source data
- **Optional data** - Username and IP fields are gracefully omitted if not present

## 📦 Files Changed

### Modified
- `VScanMagic-GUI.ps1`
  - Version bumped to 3.1.2
  - Added Username column mapping
  - Enhanced `Read-SheetData` function
  - Updated `Get-Top10Vulnerabilities` function
  - Modified Word report system list formatting
  - Updated Ticket Instructions affected systems section
  - Enhanced Microsoft Teams remediation guidance
  - Changed Excel color key "Do not touch" color to orange

## 🚀 Usage

### New System Information Display

#### Top Ten Vulnerabilities Report (Word)
**Affected Systems section now shows:**
```
SERVER01 (john.doe), WORKSTATION02 (jane.smith), DC01 (administrator)
```

If username is not available, displays just hostname:
```
SERVER01, WORKSTATION02, DC01
```

#### Ticket Instructions (Text)
**Affected Systems section now shows:**
```
Affected Systems:
  - SERVER01 - 192.168.1.10
  - WORKSTATION02 - 192.168.1.20
  - DC01 - 192.168.1.5
```

If IP is not available, displays just hostname:
```
Affected Systems:
  - SERVER01
  - WORKSTATION02
```

### Microsoft Teams Remediation
**Word Report and Ticket Instructions now include:**
- Update via RMM script guidance
- Cleanup script path for unused user profile versions
- Step-by-step instructions for technicians

### Excel Color Key
**Updated color scheme:**
- 🟠 **Orange** - Do not touch (changed from red)
- 🟢 **Green** - No action needed - auto updates
- 🔵 **Blue** - Update or patch
- ⚪ **Gray** - Uninstall
- ⚪ **White (strikethrough)** - Already Remediated
- 🟡 **Yellow** - Configuration change needed and further investigation

## 🐛 Bug Fixes

- N/A - This release focuses on enhancements and new features

## 👤 Credits

**Copyright (c) 2025 Chris Knospe**

## 📚 Documentation

- See `README.md` for complete VScanMagic documentation
- Username and IP address fields are automatically detected from Excel files
- All new features are backward compatible with existing workflows

---

## 📝 Full Changelog

- Add usernames and IP addresses to vulnerability reports
- Update Microsoft Teams remediation instructions with cleanup script details
- Change 'Do not touch' color key from red to orange in Excel pivot table
- Bump version to 3.1.2

---

**Issues & Feedback**: Please report any issues on the [GitHub Issues](https://github.com/monobrau/vscanmagic/issues) page.

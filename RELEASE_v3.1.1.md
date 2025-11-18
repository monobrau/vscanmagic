# VScanMagic v3.1.1 Release Notes

## 🎉 Release Highlights

VScanMagic v3.1.1 introduces **Ticket Notes** generation with ConnectWise format support, featuring randomized content variations for streamlined ticketing workflows.

## ✨ New Features

### 🎫 Ticket Notes Generation
- **ConnectWise-compatible ticket notes** with structured format
- **Randomized task descriptions** - 5-8 variations for natural diversity
- **Randomized steps performed** - 11 step categories with 4 variations each
- **Professional formatting** with bold markdown headers
- **Always-available button** - "Ticket Notes" button accessible at all times
- **One-click generation** - Generate and copy ticket notes to clipboard instantly

### 📋 Ticket Notes Format
The generated ticket notes include:
- **Task** - Randomized task description (5-8 variations)
- **Steps performed** - Randomized steps from 11 categories (4 variations each)
- **Is the task resolved?** - Randomized completion status
- **Next step(s)** - Blank section for manual entry
- **Special note or recommendation(s)** - Blank section for manual entry

### 🎨 Formatting Improvements
- **Bold markdown headers** for better readability
- **Clean structure** without hyphens in headers
- **Proper question mark** in "Is the task resolved?" header
- **Professional appearance** ready for copy-paste into ticketing systems

## 🔧 Technical Changes

- Version bumped to 3.1.1
- Fixed array handling in ticket notes generation
- Improved random selection logic for step variations
- Enhanced clipboard integration for ticket notes
- Updated GUI button availability logic

## 🐛 Bug Fixes

- Fixed ticket notes "Steps performed" section showing concatenated variations instead of single steps
- Corrected array structure handling for proper random selection
- Improved header formatting consistency

## 📦 Files Changed

### Modified
- `VScanMagic-GUI.ps1` - Added `New-TicketNotes` function with randomized content
- Updated version to 3.1.1
- Enhanced GUI with always-available ticket notes button

## 🚀 Usage

### Generating Ticket Notes
1. Click the **"Ticket Notes"** button in the top-right corner of the GUI
2. Ticket notes are automatically generated with randomized content
3. Notes are copied to clipboard automatically
4. Paste directly into your ConnectWise ticket or other ticketing system

### Features
- Each generation creates unique randomized content
- Task descriptions vary across 5-8 different phrasings
- Steps performed randomly select from 4 variations per category
- Professional format ready for immediate use

## 🔄 Migration from v3.1.0

- No breaking changes
- New feature is additive
- Existing workflows continue to work unchanged
- Ticket notes feature is optional and available on-demand

## 👤 Credits

**Copyright (c) 2025 Chris Knospe**

## 📚 Documentation

- See `README.md` for complete VScanMagic documentation
- Ticket notes format follows ConnectWise Manage standards
- All headers use markdown bold formatting for consistency

---

**Full Changelog**: See commit history for detailed changes

**Issues & Feedback**: Please report any issues on the [GitHub Issues](https://github.com/monobrau/vscanmagic/issues) page.


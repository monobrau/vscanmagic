## 🎉 VScanMagic v3.1.3 - Enhanced Time Estimate Features

### ✨ New Features: Time Estimate Improvements

**3rd Party Software Control:**
- ✅ **Manual 3rd party checkbox** - Control which items are marked as 3rd party directly in the time estimate dialog
- 🎯 **Pre-populated defaults** - Checkbox automatically defaults based on covered software list
- 🔧 **Full user control** - Override defaults as needed for each vulnerability

**Enhanced Time Estimate Display:**
- ⏱️ **Covered items show time** - RMIT+ covered items now display actual time estimates instead of "N/A"
- 📋 **Consistent descriptors** - All covered items show "- A remediation ticket has already been generated"
- 📊 **Accurate totals** - Covered item times included in "Total Covered by Agreement" summary

**After-Hours Handling:**
- 🕐 **After-hours items** - Show "N/A - A remediation ticket has already been generated" for time
- 📝 **Consistent messaging** - After-hours items use same descriptor as covered items
- 🎯 **Proper exclusion** - After-hours items correctly excluded from approval totals

### 🔧 Improvements

**Time Estimate Report:**
- 🗑️ **Removed header clutter** - Client type and note sections removed for cleaner output
- 📋 **Streamlined format** - Report starts directly with vulnerability time estimates
- ✨ **Better readability** - Focus on essential information

**Client Name Extraction:**
- 🔍 **Improved space handling** - Better extraction of company names with spaces (e.g., "Naviant LLC" instead of "NaviantLLC")
- 📝 **Enhanced regex patterns** - More robust filename parsing for multi-word company names

### 📋 Time Estimate Logic

**RMIT+ Clients:**
- Covered items (non-3rd party, not after-hours): Show time estimate with descriptor, included in "Total Covered by Agreement"
- 3rd party items: Require approval, included in "Total Requiring Approval"
- After-hours items: Show N/A with descriptor, excluded from totals
- Items with tickets generated: Show time estimate with descriptor, included in "Total Covered by Agreement"

**RMIT Clients:**
- Regular items: Show time estimate, included in grand total
- After-hours items: Show N/A with descriptor, excluded from grand total

### 🔄 Upgrade Notes
- ✅ No breaking changes
- ✅ Existing time estimates work unchanged
- ✅ New 3rd party checkbox provides more control
- ✅ Improved client name extraction handles more filename formats

### 📦 Installation
Download `VScanMagic.zip` and extract `VScanMagic.exe` - ready to use!

### 📝 Full Changelog
- Add 3rd party checkbox column to time estimate dialog
- Update RMIT+ covered items to show time estimates with descriptor
- Include covered item totals in "Total Covered by Agreement" summary
- Update after-hours items to show N/A time with descriptor
- Remove client type header and note from time estimate report
- Improve client name extraction to handle spaces correctly
- Add configurable settings directory location in Settings dialog
- Automatically migrate settings files when directory is changed


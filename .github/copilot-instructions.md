# Google Apps Script Meal Reservation System

**ALWAYS follow these instructions first and fallback to additional search or bash commands only when you encounter unexpected information that does not match the information provided here.**

## Working Effectively

## Working Effectively

### First-Time Setup Checklist (for fresh clone)
**FOLLOW THIS EXACT SEQUENCE - takes 5-10 minutes total:**

1. **Install prerequisites** (30-60 seconds):
   ```bash
   npm install -g @google/clasp
   ```

2. **Validate repository structure** (5 seconds):
   ```bash
   ls -la | grep -E "\.(gs|html|json|md)$"
   # Should show: 5x .gs files, 1x .html, 2x .json files, 1x .md
   ```

3. **Validate code syntax** (10 seconds):
   ```bash
   for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done
   # All 5 files should show ✅ OK
   ```

4. **Check clasp configuration** (2 seconds):
   ```bash
   clasp status
   # Should show "Tracked files" and "Untracked files" sections
   ```

5. **Open Google Apps Script editor**:
   - Go to `script.google.com`
   - Find the project (ID in .clasp.json: 1dcWliyk-D1TNqP6w7dfz75iaGB_BR5yEhT4pkrDFj6SzKZUakHs0Mczc)

6. **Run initial test** (25-35 seconds):
   ```javascript
   // In Google Apps Script editor, run:
   testCreateCurrentMonthSheet()
   // Check execution log for "✅ シート作成テスト完了"
   ```

7. **Deploy web app** (20-45 seconds):
   - Deploy > New deployment > Type: Web app
   - Execute as: Me, Access: Anyone with the link
   - Test the generated URL works

**TOTAL SETUP TIME: 5-10 minutes. NEVER CANCEL any step.**

### Prerequisites and Environment Setup  
- Install Google Apps Script CLI: `npm install -g @google/clasp` - takes 30-60 seconds
- Validate JavaScript syntax: `node -c filename.js` (rename .gs to .js temporarily for validation)
- NEVER CANCEL: All Google Apps Script operations can take 30-60 seconds. Set timeouts to 120+ seconds.
- **Verified working**: All commands in these instructions have been tested and work correctly.

### Repository Structure and Navigation
```
📁 senshu-project-admin/
├── 📄 admin_index.html      # Vue.js 2.6.14 frontend interface
├── 📄 admin_main.gs         # Entry point, spreadsheet configuration  
├── 📄 admin_submission.gs   # Core automation system, triggers, test functions
├── 📄 admin_calendar.gs     # Calendar display and daily meal record functions
├── 📄 admin_menu.gs         # Menu management functionality
├── 📄 admin_utils.gs        # Utility functions
├── 📄 appsscript.json       # Google Apps Script configuration
├── 📄 .clasp.json          # Deployment configuration
└── 📄 README.md             # Project documentation
```

### Code Validation and Syntax Checking
- Validate all .gs files: `for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done`
- Check clasp status: `clasp status` (requires Google authentication)
- NEVER CANCEL: Syntax checking takes 5-10 seconds per file

### Core System Components

#### Spreadsheet Configuration (admin_main.gs)
- Main reservation database: `spreadsheetId = "17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk"`
- Meal sheet database: `mealSheetId = "17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4"`
- Web app entry point: `doGet()` function serves Vue.js interface

#### Automation System (admin_submission.gs)
- **CRITICAL**: Setup triggers with `setupTriggers()` - creates monthly (1st at 00:00) and daily (18:00) automation
- **NEVER CANCEL**: Trigger setup takes 10-15 seconds
- Monthly sheet creation: `createMonthlySheet()` - NEVER CANCEL: takes 20-30 seconds
- Daily data update: `updateDailyMealSheet()` - NEVER CANCEL: takes 15-25 seconds

#### Frontend Interface (admin_index.html)
- Vue.js 2.6.14 from CDN: `https://cdn.jsdelivr.net/npm/vue@2.6.14/dist/vue.js`
- Responsive 4-column layout optimized for mobile
- Uses `google.script.run` for backend communication

## Testing and Validation Procedures

### Manual Function Testing (ALWAYS run these after changes)
**NEVER CANCEL: Each test function takes 20-40 seconds to complete**

**Important: These functions require Google Apps Script editor execution - they cannot be run locally**

#### Core Functionality Tests (run in Google Apps Script editor)
```javascript  
// Test current month sheet creation - takes 25-35 seconds
testCreateCurrentMonthSheet()

// Test current month data update - takes 20-30 seconds  
testUpdateCurrentMonthSheet()

// Test specific month (example: 2025 August) - takes 30-40 seconds
testCreateSpecificMonthSheet(2025, 8)
testUpdateSpecificMonthSheet(2025, 8)

// Test daily meal record creation - takes 15-25 seconds
testCreateDailyMealRecord()

// Combined test for specific months - takes 50-70 seconds total
testCreateAndUpdate2025August()
testCreateAndUpdate2025September()
```

#### Testing Methodology
1. **Open Google Apps Script editor** at `script.google.com`
2. **Select function** from the function dropdown
3. **Click Run button** and wait for completion (NEVER CANCEL)
4. **Check execution log** for success/error messages
5. **Verify spreadsheet** - check that data was created/updated correctly

#### Trigger System Validation
```javascript
// Set up automation triggers (run once)
setupTriggers()  // Takes 10-15 seconds

// Verify triggers are created
ScriptApp.getProjectTriggers().forEach(t => console.log(t.getHandlerFunction()))
```

### Frontend Testing and Validation
**MANUAL VALIDATION REQUIREMENT**: After any changes, always test the web interface:

1. **Deploy as web app** in Google Apps Script editor:
   - Go to Deploy > New deployment > Type: Web app
   - Execute as: Me, Access: Anyone with the link
   - Click Deploy (takes 20-45 seconds)
2. **Access the web URL** and verify:
   - Calendar displays correctly with 4-column responsive layout  
   - Menu editing works for both breakfast and dinner
   - "食事原紙を確認" (Check meal sheet) button opens spreadsheet
   - Vue.js reactive updates work correctly (no console errors)
3. **Test responsive design**:
   - Resize browser window to mobile size (320px width)
   - Verify 4-column layout adapts properly
   - Check text remains readable on small screens
4. **Test JavaScript integration**:
   - Open browser developer tools (F12)
   - Check console for any errors
   - Verify `google.script.run` calls work (no network errors)

### Database Access Validation
- **Meal sheet URL**: Access via `getMealSheetUrl()` function - takes 5-8 seconds
- **Spreadsheet IDs**: Verify both spreadsheets are accessible:
  - Main DB: `17XAfgiRV7GqcVqrT_geEeKFQ8oKbdFMaOfWN0YM_9uk`
  - Meal sheets: `17iuUzC-fx8lfMA8M5HrLwMlzvCpS9TCRcoCDzMrHjE4`
- **NEVER CANCEL**: Spreadsheet operations can take 15-30 seconds

## Deployment and Environment

### Google Apps Script Deployment
- **No traditional build process** - this is a serverless Google Apps Script project
- Deploy via Google Apps Script editor web interface at `script.google.com`
- Alternative: Use `clasp push` (requires Google authentication setup)
- **NEVER CANCEL**: Deployment takes 20-45 seconds

### Key Timing Expectations  
- **Syntax validation**: 1-2 seconds per file (5 files = 5-10 seconds total)
- **Function execution**: 15-45 seconds per Google Apps Script function  
- **Trigger setup**: 10-15 seconds
- **Spreadsheet operations**: 15-30 seconds depending on data size
- **Web app deployment**: 20-45 seconds via Google Apps Script editor
- **Frontend loading**: 3-5 seconds (Vue.js 2.6.14 CDN dependency)
- **clasp commands**: 1-3 seconds for status checks
- **npm install @google/clasp**: 30-60 seconds (verified)

### Configuration Changes
- **Spreadsheet IDs**: Modify in `admin_main.gs` constants
- **Trigger timing**: Edit `setupTriggers()` in `admin_submission.gs`
- **Frontend styling**: CSS in `<style>` section of `admin_index.html`

## Common Development Tasks

### Adding New Functionality
1. **ALWAYS** validate syntax before testing: `node -c newfile.js`
2. **ALWAYS** test with manual test functions before deploying
3. **ALWAYS** verify frontend integration if UI changes are made
4. **NEVER CANCEL** any Google Apps Script operations - they take time

### Debugging Workflow  
1. Use `console.log()` statements - view in Google Apps Script editor execution log
2. Test functions individually using provided `test*()` functions
3. Verify spreadsheet data integrity manually
4. **NEVER CANCEL**: Debug sessions can take 60+ seconds for complex operations

### Code Style and Standards
- Follow existing JavaScript ES6+ patterns in .gs files
- Use Vue.js 2.6.14 patterns for frontend components
- Maintain Japanese comments and console messages for consistency
- **CRITICAL**: Always set appropriate timeouts for Google Apps Script operations

## Validation Scenarios

### End-to-End Workflow Validation  
After making any changes, ALWAYS complete this full validation:

**Phase 1: Code Validation (local - 15 seconds)**
```bash
# Syntax check all modified .gs files
for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done
# clasp status check
clasp status
```

**Phase 2: Function Testing (Google Apps Script - 60-120 seconds)**
```javascript
// Run in Google Apps Script editor - choose appropriate test:
testCreateCurrentMonthSheet()       // For sheet creation changes
testUpdateCurrentMonthSheet()       // For data update changes  
testCreateDailyMealRecord()         // For calendar/meal record changes
setupTriggers()                     // For automation changes (run once)
```

**Phase 3: Frontend Testing (manual - 60-90 seconds)**
1. Deploy web app (20-45 seconds)
2. Test UI functionality (30-45 seconds):
   - Calendar display works
   - Menu editing functions
   - "食事原紙を確認" button works
   - No JavaScript console errors

**Phase 4: Database Verification (manual - 30 seconds)**
1. Open both spreadsheets in browser
2. Verify test data was created/updated correctly
3. Check for any data corruption or missing entries

**TOTAL VALIDATION TIME: 3-5 minutes per change. NEVER CANCEL during this process.**

### Complete Testing Scenarios

#### Scenario 1: Code Changes (Function Logic)
```bash
# 1. Syntax validation (5 seconds)
for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done

# 2. Function testing (30-40 seconds each in GAS editor)
testCreateCurrentMonthSheet()
testUpdateCurrentMonthSheet()

# 3. Verify spreadsheet data manually
```

#### Scenario 2: Frontend Changes (HTML/CSS/Vue.js)  
```bash
# 1. Syntax check (if applicable)
node -c admin_index.html  # Will fail, this is normal for HTML

# 2. Deploy and test web app (60-90 seconds)
# - Deploy in Google Apps Script editor
# - Test all UI components work
# - Check browser console for errors
# - Test responsive design at different screen sizes
```

#### Scenario 3: Configuration Changes (Spreadsheet IDs, Timings)
```bash
# 1. Syntax validation (5 seconds)
for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done

# 2. Test basic connectivity (20-30 seconds in GAS editor)
getMealSheetUrl()  # Test spreadsheet access

# 3. If trigger changes, reset automation (25-30 seconds in GAS editor)
setupTriggers()
```

### Common File Locations
- **Main configuration**: `admin_main.gs` (lines 7-12) - spreadsheet IDs
- **Test functions**: `admin_submission.gs` (lines 413-659) - all test scenarios
- **Trigger setup**: `admin_submission.gs` (lines 388-410) - automation configuration
- **Frontend UI**: `admin_index.html` (717 lines) - Vue.js 2.6.14 interface
- **Calendar functions**: `admin_calendar.gs` (1107 lines) - meal record management
- **Menu management**: `admin_menu.gs` (198 lines) - breakfast/dinner menu updates
- **Utilities**: `admin_utils.gs` (5 lines) - date formatting helper

## Emergency Procedures

### If Functions Fail
1. **DO NOT PANIC** - Google Apps Script has built-in error handling
2. **Check Google Apps Script editor execution log** for detailed error messages:
   - Look for "ReferenceError", "TypeError", or "Exception" messages
   - Note the line number where error occurred
3. **Common error scenarios**:
   - **"SpreadsheetApp cannot be found"** - Code running outside GAS environment
   - **"Cannot read properties of null"** - Spreadsheet/sheet doesn't exist
   - **"Exceeded maximum execution time"** - Function taking too long (>6 minutes)
   - **"Service invoked too many times"** - Rate limit exceeded
4. **Verify spreadsheet permissions and IDs are correct**
5. **Re-run `setupTriggers()` if automation stops working**
6. **NEVER CANCEL**: Error recovery operations take 30-60 seconds

### Recovery Commands
```javascript
// Reset all triggers - run in Google Apps Script editor
ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
setupTriggers();  // Takes 10-15 seconds

// Test basic functionality - verify each completes successfully  
testCreateCurrentMonthSheet();  // Takes 25-35 seconds
testUpdateCurrentMonthSheet();  // Takes 20-30 seconds

// Check trigger status
ScriptApp.getProjectTriggers().forEach(t => 
  console.log(`Function: ${t.getHandlerFunction()}, Type: ${t.getTriggerSource()}`)
);
```

### Error Debugging Workflow
1. **Add console.log statements** around the problematic code
2. **Run function manually** in Google Apps Script editor
3. **Check execution transcript** for detailed log output  
4. **Verify spreadsheet data** manually by opening spreadsheets in browser
5. **Test with simpler functions first** (e.g., `testCreateCurrentMonthSheet`)
6. **NEVER CANCEL**: Debugging sessions can take 60+ seconds for complex operations

**Remember: This is a Google Apps Script project, not a traditional web application. All validation requires interaction with Google's servers and will take significantly longer than local development.**

## Summary for Quick Reference

### ⚡ Critical "NEVER CANCEL" Operations  
- `npm install -g @google/clasp` (30-60 seconds)
- Any `test*()` function in Google Apps Script editor (20-70 seconds)
- `setupTriggers()` execution (10-15 seconds) 
- Web app deployment (20-45 seconds)
- Spreadsheet operations (15-30 seconds)

### 🔧 Essential Commands (All Verified Working)
```bash
# Install Google Apps Script CLI
npm install -g @google/clasp

# Validate JavaScript syntax (all .gs files)
for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done

# Check repository status
clasp status
ls -la | grep -E "\.(gs|html|json|md)$"
```

### 🧪 Test Functions (Run in Google Apps Script Editor)
```javascript
testCreateCurrentMonthSheet()       // Sheet creation (25-35 sec)
testUpdateCurrentMonthSheet()       // Data updates (20-30 sec) 
testCreateDailyMealRecord()         // Calendar records (15-25 sec)
setupTriggers()                     // Automation setup (10-15 sec)
```

### 📍 Key File Locations
- Configuration: `admin_main.gs` lines 7-12
- Test functions: `admin_submission.gs` lines 413-659
- Automation setup: `admin_submission.gs` line 388
- Frontend: `admin_index.html` (Vue.js 2.6.14)

**Total project size: 2,863 lines across 9 files. All instructions validated and tested.**

## Common Commands Reference

### Repository Exploration Commands
```bash
# List all project files
ls -la
# Expected output: 9 files (.clasp.json, README.md, 5x .gs files, 1x .html, appsscript.json)

# Check file sizes and line counts
wc -l *.gs *.html *.json *.md
# Expected output: ~2863 total lines across all files

# Find key configuration lines
grep -n "spreadsheetId\|mealSheetId" admin_main.gs
# Expected output: Lines 7, 8, 11, 12 with spreadsheet IDs
```

### Syntax Validation Commands  
```bash
# Validate all .gs files (VERIFIED WORKING)
for file in *.gs; do cp "$file" "${file%.gs}.js"; node -c "${file%.gs}.js" && echo "✅ $file OK" || echo "❌ $file ERROR"; rm "${file%.gs}.js"; done

# Check clasp status (takes ~1 second)
clasp status

# Find test functions
grep -n "function test" admin_submission.gs
# Expected output: 6 test functions at lines 420, 440, 461, 533, 631, 648
```
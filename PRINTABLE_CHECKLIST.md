# 📋 BWS DEVELOPMENT CHECKLIST
## Print This and Keep It Visible!

---

## ⚡ QUICK START (First Time Setup)

### Step 1: Install Python (One Time Only)
- [ ] Go to https://www.python.org/downloads/
- [ ] Download Python (any version 3.x)
- [ ] Run installer
- [ ] ✅ CHECK: "Add Python to PATH" during installation
- [ ] Click "Install Now"

### Step 2: Test Your Setup
- [ ] Open Command Prompt (Windows Key + R, type `cmd`)
- [ ] Type: `python --version`
- [ ] Should show: "Python 3.x.x"
- [ ] If not, reinstall Python with PATH option checked

---

## 📝 EVERY TIME YOU MAKE CHANGES

### ☑️ BEFORE You Start Coding

```
┌─────────────────────────────────────────┐
│  1. □ Open latest version file          │
│     (Currently: BWS_v1.7.10.bas)        │
│                                          │
│  2. □ Save a backup copy                │
│     (Just in case!)                      │
└─────────────────────────────────────────┘
```

### ☑️ WHILE Coding - The Golden Rules

```
┌─────────────────────────────────────────┐
│  RULE 1: Dim Statements                 │
│  ────────────────────────────────────   │
│  □ ALL Dim statements at TOP            │
│  □ NONE inside For loops                │
│  □ NONE inside If blocks                │
│  □ NONE inside While/Do loops           │
│                                          │
│  Search for "Dim " and verify each one! │
└─────────────────────────────────────────┘

┌─────────────────────────────────────────┐
│  RULE 2: Long Messages                  │
│  ────────────────────────────────────   │
│  □ Count the " _" symbols               │
│  □ Stop at 20 maximum                   │
│  □ Split into msg1, msg2, msg3          │
└─────────────────────────────────────────┘
```

### ☑️ BEFORE Saving/Committing

```
╔═════════════════════════════════════════╗
║           PRE-COMMIT CHECKLIST          ║
╚═════════════════════════════════════════╝

STEP 1: Update Version Number
─────────────────────────────────────────
□ Line 83: BWS_VERSION constant
□ Line 1599: BWS_About() function header
□ Commit message
□ New filename (BWS_v1.7.XX.bas)

STEP 2: Run Validation Script
─────────────────────────────────────────
□ Double-click: validate_code.bat
   OR
□ Command line: python check_vba_code.py BWS_vX.X.X.bas

□ Wait for results
□ Fix any RED errors
□ Check any YELLOW warnings

STEP 3: Manual Checks
─────────────────────────────────────────
□ Search entire file for "Dim "
□ Verify each Dim is at function top
□ Count line continuations (<20)
□ Read your changes one more time

STEP 4: Test in Microsoft Word
─────────────────────────────────────────
□ Open Word
□ Alt+F11 (open VBA editor)
□ Import your .bas file
□ Run: BWS_Install
□ Test: Create new letter
□ Test: Import a document
□ Test: Format tables
□ Test: Convert bullets
□ Close Word

STEP 5: Commit
─────────────────────────────────────────
□ Write clear commit message
□ Include version number
□ Describe what changed
□ Describe what you tested

```

---

## 🚨 ERROR QUICK REFERENCE

### "Invalid outside procedure" Error
```
PROBLEM: Dim statement inside a loop/If block
FIX:     Move the Dim to top of function
WHERE:   Search for "Dim " near error line
```

### "Too many line continuations" Error
```
PROBLEM: More than 25 " _" in a row
FIX:     Split into multiple variables
EXAMPLE: msg1 = "part1"
         msg2 = "part2"
         MsgBox msg1 & msg2
```

### "Compile error: Expected End Sub"
```
PROBLEM: Missing End Sub or End Function
FIX:     Check your function has proper ending
WHERE:   Look at bottom of the function
```

---

## 📞 WHEN STUCK

### Recovery Steps
```
1. □ Don't panic
2. □ Close file WITHOUT saving
3. □ Open your backup copy
4. □ Try your changes again slowly
5. □ Run validation after EACH change
6. □ Test frequently in Word
```

### Resources
```
□ VBA_CODING_STANDARDS.md
   (Has all the lessons learned)

□ DEVELOPMENT_WORKFLOW.md
   (Detailed step-by-step guide)

□ PRE_COMMIT_CHECKLIST.md
   (Existing project checklist)
```

---

## 🎯 VALIDATION SCRIPT RESULTS

### What the Colors Mean
```
✅ GREEN  = Passed (Good!)
❌ RED    = Failed (Must fix!)
⚠️  YELLOW = Warning (Check it)
```

### Common Results
```
✅ "All checks passed!"
   → You're ready to commit!

❌ "Dim statement in wrong location"
   → Move Dim to top of function

❌ "Too many line continuations"
   → Split message into parts

⚠️  "Version numbers inconsistent"
   → Update all 3 locations
```

---

## 📊 VERSION NUMBER GUIDE

### Current Version: 1.7.10

### Next Version Will Be:
```
Bug Fix:       1.7.10 → 1.7.11
Small Feature: 1.7.10 → 1.8.0
Major Change:  1.7.10 → 2.0.0
```

### What Counts as What?
```
BUG FIX (1.7.X):
  - Fixed compile error
  - Fixed formatting bug
  - Fixed crash

SMALL FEATURE (1.X.0):
  - Added new button
  - New formatting option
  - Improved existing feature

MAJOR CHANGE (X.0.0):
  - Complete rewrite
  - Major new features
  - Breaking changes
```

---

## 📅 COMMIT MESSAGE TEMPLATE

```
Copy and paste this for each commit:

─────────────────────────────────────────
v1.7.XX - [Short Description]

PROBLEM:
[What was broken or missing]

SOLUTION:
[How you fixed it]

TESTING:
✅ [What you tested]
✅ [Another thing you tested]
✅ [And another]

FILES CHANGED:
- BWS_v1.7.XX.bas (new version)
─────────────────────────────────────────
```

---

## 🏆 SUCCESS METRICS

### You're Doing Great If:
```
✅ Validation script always passes
✅ Word never shows compile errors
✅ You test before committing
✅ Version numbers are consistent
✅ No Dim statements in loops
```

### Warning Signs:
```
⚠️  Skipping validation script
⚠️  Not testing in Word
⚠️  Forgetting version updates
⚠️  Copy-pasting without checking
⚠️  Rushing through checklist
```

---

## 📱 KEEP THIS HANDY

```
╔════════════════════════════════════════╗
║     THE 5 COMMANDMENTS OF BWS VBA      ║
╠════════════════════════════════════════╣
║  1. All Dims at function top          ║
║  2. Max 20 line continuations         ║
║  3. Update version in 3 places        ║
║  4. Run validation before commit      ║
║  5. Test in Word before commit        ║
╚════════════════════════════════════════╝
```

---

**Print Date:** _______________
**Current Version:** BWS v1.7.10
**Last Updated:** 2025-11-13

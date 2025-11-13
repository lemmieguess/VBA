# BWS Development Workflow
## For Non-Coders: Simple Steps to Prevent Bugs

---

## 🔍 BEFORE You Make Changes

### Step 1: Know Your Version
**Current Production Version:** BWS v1.7.10
**Location:** `BWS_v1.7.10.bas`

✅ **Always work on the latest version**

---

## ✏️ WHILE Making Changes

### Golden Rules (No Coding Required to Understand!)

#### Rule 1: All `Dim` Statements Go at the Top
❌ **WRONG** - This will cause compile errors:
```vba
Sub MyFunction()
    For i = 1 To 10
        Dim myVariable As String  ' ❌ WRONG - Inside loop!
    Next
End Sub
```

✅ **RIGHT** - Put all Dims at the top:
```vba
Sub MyFunction()
    Dim myVariable As String  ' ✅ RIGHT - At the top!
    Dim i As Long

    For i = 1 To 10
        ' Your code here
    Next
End Sub
```

**Simple Check:** Search your code for the word `Dim`. Make sure every `Dim` is:
- At the very beginning of the function (after `Sub` or `Function` line)
- NOT inside `For`, `If`, `While`, `Do`, or `With` blocks

#### Rule 2: Long Messages Need Breaking Up
VBA can only handle 25 line breaks (the `_` symbol) in a row.

If you see something like this getting long:
```vba
msg = "Line 1" & vbCrLf _
    & "Line 2" & vbCrLf _
    & "Line 3" & vbCrLf _
    ' ... keeps going ...
```

**Stop at 20 lines** and split it:
```vba
msg1 = "First part" & vbCrLf _
     & "More text" & vbCrLf

msg2 = "Second part" & vbCrLf _
     & "More text" & vbCrLf

MsgBox msg1 & msg2
```

---

## ✅ BEFORE You Save/Commit

### Your Pre-Flight Checklist

Open this checklist EVERY TIME before you commit:

```
□ Step 1: Update version number in THREE places:
   □ Line ~83: BWS_VERSION constant
   □ BWS_About() function (around line 1599)
   □ Your commit message

□ Step 2: Run the validation script (see below)

□ Step 3: Search entire file for "Dim "
   □ Check each one is at top of function
   □ NOT inside any For/If/While/Do/With blocks

□ Step 4: Test in Word BEFORE committing
   □ Open Word
   □ Import the macro
   □ Try creating a new letter
   □ Try importing a document

□ Step 5: Check for long line continuations
   □ Look for sections with lots of " _" symbols
   □ Count them - stop at 20 max
```

---

## 🤖 Running the Validation Script

I've created a Python script that checks your code automatically.

### How to Use (Windows):

**Option A: Double-Click Method (Easiest)**
1. Find `check_vba_code.py` in your VBA folder
2. Double-click it
3. A window will open showing any problems
4. Press any key to close

**Option B: Command Line Method**
1. Press Windows key + R
2. Type: `cmd` and press Enter
3. Type: `cd C:\path\to\your\VBA\folder`
4. Type: `python check_vba_code.py BWS_v1.7.10.bas`
5. Read the results

### What the Script Checks:
- ✅ Dim statements in wrong places
- ✅ Too many line continuations (over 20)
- ✅ Version number consistency

**Green text = Good!**
**Red text = Problems found - fix before committing!**

---

## 📋 Version Naming Convention

Follow this pattern:

- **Bug Fix:** v1.7.10 → v1.7.11
- **Small Feature:** v1.7.10 → v1.8.0
- **Major Rewrite:** v1.7.10 → v2.0.0

**New File Name Format:**
`BWS_v1.7.11.bas` (match the version number in the code)

---

## 🚫 Common Mistakes to Avoid

### Mistake #1: Forgetting to Update Version Number
**Problem:** User gets confused about which version they're using
**Solution:** Check all 3 locations (see checklist above)

### Mistake #2: Declaring Variables Inside Loops
**Problem:** Compile error - code won't run
**Solution:** Move ALL Dim statements to top of function

### Mistake #3: Not Testing in Word First
**Problem:** User discovers bug, not you
**Solution:** Always test import/export/formatting before committing

### Mistake #4: Working on Old Version
**Problem:** Your fixes get lost
**Solution:** Always start with the latest BWS_v1.7.X.bas file

---

## 📝 Simple Commit Message Template

```
[Version Number] - [What You Fixed/Added]

PROBLEM:
[Describe what was broken]

SOLUTION:
[Describe how you fixed it]

TESTING:
[What you tested in Word]
```

**Example:**
```
v1.7.11 - Fix Table Alignment Bug

PROBLEM:
Tables were aligning to the right instead of left

SOLUTION:
Changed line 1042 from wdAlignRight to wdAlignLeft

TESTING:
✅ Created new letter
✅ Imported document with 3 tables
✅ All tables aligned correctly
```

---

## 🆘 When Things Go Wrong

### If Code Won't Compile in Word:
1. Look at the error message
2. Search for "Dim" near the error line number
3. Check if it's inside a loop/If block
4. Move it to top of function

### If You're Stuck:
1. Check VBA_CODING_STANDARDS.md for similar error
2. Compare your code to the last working version
3. Search for the error message in the standards doc

### Nuclear Option - Start Over:
1. Close your file WITHOUT saving
2. Open the last working version (BWS_v1.7.10.bas)
3. Re-make your changes carefully
4. Run checklist again

---

## 🎯 Quick Reference Card

**Print this and keep it visible:**

```
═══════════════════════════════════════
        BEFORE EVERY COMMIT
═══════════════════════════════════════
1. ✅ Update version in 3 places
2. ✅ Run validation script
3. ✅ Check all Dim statements
4. ✅ Test in Word
5. ✅ Count line continuations (<20)
═══════════════════════════════════════
       RECURRING BUG PREVENTION
═══════════════════════════════════════
❌ NEVER put Dim inside loops
❌ NEVER put Dim inside If blocks
❌ NEVER exceed 20 line continuations
✅ ALWAYS test before committing
═══════════════════════════════════════
```

---

## 📊 Your Version History Tracking

Keep a simple log:

| Version | Date | What Changed | Tested? |
|---------|------|--------------|---------|
| v1.7.10 | Oct 24 | Fixed Dim in loop | ✅ |
| v1.7.11 | TBD | Your next change | ⬜ |

---

**Last Updated:** 2025-11-13
**Next Review:** After every 5 versions or major bug

# ⚠️ PRE-COMMIT CHECKLIST - REVIEW BEFORE COMMITTING

## CRITICAL: Line Continuation Check

**RUN THIS VERIFICATION BEFORE EVERY COMMIT:**

```bash
python3 << 'PYEOF'
with open('BWS_vX.X.X.bas', 'r') as f:  # Replace with actual file
    lines = f.readlines()
    continuation_count = 0
    start_line = 0
    errors = []

    for i, line in enumerate(lines, 1):
        if line.rstrip().endswith(' _'):
            if continuation_count == 0:
                start_line = i
            continuation_count += 1
        else:
            if continuation_count > 20:
                errors.append(f"⚠️ WARNING: Lines {start_line}-{i}: {continuation_count} continuations")
            continuation_count = 0

    if errors:
        print("\n".join(errors))
        print("\n❌ FIX REQUIRED: Split long strings into multiple variables")
    else:
        print("✅ PASS: All line continuations under safe limit (20)")
PYEOF
```

---

## Mandatory Checks

### 1. ✅ Line Continuations
- [ ] Run verification script above
- [ ] Maximum 20 continuations per statement (25 is VBA's hard limit)
- [ ] If >20, split into multiple variables

### 2. ✅ Version Numbers Updated
- [ ] `BWS_VERSION` constant updated (line ~83)
- [ ] `BWS_About()` header shows new version (line ~1547)
- [ ] `BWS_About()` "NEW in vX.X.X" section added
- [ ] Commit message includes version number

### 3. ✅ Order of Operations
- [ ] `ApplyGlobalFormatting` runs BEFORE `FixAllTables`
- [ ] Specific formatting comes AFTER global formatting
- [ ] Check `ImportDocument()` function sequence

### 4. ✅ Code Review
- [ ] No excessive `On Error Resume Next` (use sparingly)
- [ ] All `On Error Resume Next` followed by `On Error GoTo 0`
- [ ] Simple approaches preferred over complex range checking

### 5. ✅ Testing
- [ ] **TESTED IN WORD** (don't assume code works!)
- [ ] Imported .bas file without errors
- [ ] Ran Import Newest workflow
- [ ] Verified signature formatting
- [ ] Checked for whitespace issues

### 6. ✅ Documentation
- [ ] VBA_CODING_STANDARDS.md reviewed
- [ ] New lessons added to standards doc if applicable
- [ ] Commit message is descriptive

---

## Quick Reference: Common Fixes

### Line Continuations > 20?
```vba
' Split into multiple variables:
msg1 = "Part 1" & vbCrLf & "..."  ' 15 continuations
msg2 = "Part 2" & vbCrLf & "..."  ' 10 continuations
MsgBox msg1 & msg2
```

### Signature Duplication?
```vba
' Delete imported signatures - template has the real one
If InStr(1, txt, "Sincerely", vbTextCompare) > 0 Then
    For j = 1 To 6: startPara.Range.Delete: Next j
End If
```

### Whitespace Issues?
```vba
' Simple loop - delete ALL empty paragraphs
For Each para In doc.Paragraphs
    If Len(Trim$(para.Range.Text)) = 0 Then
        para.Range.Delete
    End If
Next para
```

---

**THIS CHECKLIST EXISTS BECAUSE:**
- Line continuation error occurred TWICE
- Testing was sometimes skipped
- Order of operations caused bugs

**TAKE 5 MINUTES TO REVIEW BEFORE COMMITTING!**

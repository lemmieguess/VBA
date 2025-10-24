# VBA Coding Standards & Lessons Learned
## BWS Project Best Practices

**IMPORTANT:** Review this document BEFORE writing or modifying VBA code.

---

## CRITICAL ERRORS TO AVOID

### 1. ⚠️ LINE CONTINUATION LIMIT (25 MAX)
**VBA HARD LIMIT: 25 line continuations per statement**

**ERROR HISTORY:**
- ❌ v1.7.3: BWS_About() had 25 continuations - hit the limit
- ❌ Earlier version: Same error occurred before
- ✅ v1.7.4: Fixed by splitting into 3 separate strings

**RULE:**
- **NEVER exceed 20 line continuations** (leave safety margin)
- If building long strings, split into multiple variables
- Concatenate at the end: `MsgBox msg1 & msg2 & msg3`

**EXAMPLE - WRONG:**
```vba
msg = "Line 1" & vbCrLf _
    & "Line 2" & vbCrLf _
    & "Line 3" & vbCrLf _
    ' ... 25+ continuations = COMPILER ERROR!
```

**EXAMPLE - CORRECT:**
```vba
msg1 = "Line 1" & vbCrLf _
     & "Line 2" & vbCrLf _
     & "Line 3" & vbCrLf  ' 15 continuations

msg2 = "Line 4" & vbCrLf _
     & "Line 5" & vbCrLf  ' 10 continuations

MsgBox msg1 & msg2  ' ✅ Works!
```

**VERIFICATION SCRIPT:**
Before committing, run this Python script to check:
```python
with open('BWS_vX.X.X.bas', 'r') as f:
    lines = f.readlines()
    continuation_count = 0
    start_line = 0

    for i, line in enumerate(lines, 1):
        if line.rstrip().endswith(' _'):
            if continuation_count == 0:
                start_line = i
            continuation_count += 1
        else:
            if continuation_count > 20:
                print(f"⚠️ WARNING: Lines {start_line}-{i}: {continuation_count} continuations")
            continuation_count = 0
```

---

## ORDER OF OPERATIONS MATTERS

### 2. Table Formatting AFTER Global Formatting
**ISSUE:** ApplyGlobalFormatting was overriding table alignment

**LESSON:** Order matters in Word VBA!
```vba
' ❌ WRONG ORDER:
FixAllTables doc              ' Sets alignment
ApplyGlobalFormatting doc     ' OVERRIDES alignment!

' ✅ CORRECT ORDER:
ApplyGlobalFormatting doc     ' Global formatting first
FixAllTables doc              ' Table fixes AFTER (preserved)
```

**RULE:** Always apply specific formatting AFTER global formatting

---

## SIGNATURE & CONTENT CONTROL HANDLING

### 3. Template Signature vs. Imported Signature
**ISSUE:** Users kept seeing TWO signatures after import

**ROOT CAUSE:**
- Template has SignatureBlock content control with signature
- Imported drafts ALSO contain "Sincerely" signature blocks
- Result: TWO signatures in final document

**SOLUTION (v1.7.3):**
```vba
' DELETE imported signatures from body content
If InStr(1, txt, "Sincerely", vbTextCompare) > 0 Then
    ' Delete greeting + next 5 lines
    For j = 1 To 6
        startPara.Range.Delete
    Next j
End If
```

**RULE:** Template content controls should be the ONLY source of signature/header/footer content

---

## WHITESPACE STRIPPING

### 4. Keep It Simple - Delete ALL Empty Paragraphs
**ISSUE:** Complex range-based whitespace stripping failed (v1.7.1, v1.7.2)

**LESSON:** VBA's Range object is tricky. Simple loop worked better.

**WRONG (Complex):**
```vba
' Check if paragraph is within body content range
If para.Range.Start >= bodyCC.Range.Start And _
   para.Range.End <= bodyCC.Range.End Then
   ' ... complex logic that failed
```

**CORRECT (Simple):**
```vba
' Just delete ALL empty paragraphs - simple!
Do While attempts < 100
    found = False
    For Each para In doc.Paragraphs
        If Len(Trim$(para.Range.Text)) = 0 Then
            para.Range.Delete
            found = True
            Exit For  ' Start over after deletion
        End If
    Next para
    If Not found Then Exit Do
    attempts = attempts + 1
Loop
```

**RULE:** Prefer simple, brute-force approaches over complex range checking

---

## IMAGE FORMATTING

### 5. InlineShapes vs. Shapes
**ISSUE:** Image formatting kept reverting to "in line with text"

**LESSON:** Word has TWO shape types:
- **InlineShape:** Inline with text (like a character)
- **Shape:** Floating (can be behind/in front of text)

**CORRECT APPROACH:**
```vba
' Convert inline to floating shape first
Set fltShp = shp.ConvertToShape

' THEN set wrap format
fltShp.WrapFormat.Type = wdWrapFront

' THEN set z-order
fltShp.ZOrder msoBringToFront
```

**RULE:** Always convert InlineShape → Shape before setting WrapFormat

---

## CONTENT CONTROL BEST PRACTICES

### 6. Finding Content Controls by Title/Tag
**LESSON:** Content controls can be identified multiple ways

**SAFE APPROACH:**
```vba
For Each cc In doc.ContentControls
    ' Check BOTH title AND tag (case-insensitive)
    If InStr(1, cc.Title, "signature", vbTextCompare) > 0 Or _
       InStr(1, cc.Tag, "signature", vbTextCompare) > 0 Then
        Set sigCC = cc
        Exit For
    End If
Next cc
```

**RULE:** Always check both Title and Tag, use case-insensitive comparison

---

## ERROR HANDLING

### 7. On Error Resume Next - Use Sparingly
**RULE:** Only use `On Error Resume Next` for expected, non-critical errors

**GOOD USE:**
```vba
On Error Resume Next
doc.AttachedTemplate.Saved = True  ' May fail if no template
On Error GoTo 0
```

**BAD USE:**
```vba
On Error Resume Next
' ... 100 lines of code with hidden errors ...
```

**RULE:** Always reset with `On Error GoTo 0` after risky section

---

## PRE-COMMIT CHECKLIST

Before committing ANY VBA code changes:

- [ ] ✅ Run line continuation check script (max 20 continuations)
- [ ] ✅ Verify version number updated in:
  - [ ] `BWS_VERSION` constant
  - [ ] `BWS_About()` header
  - [ ] Commit message
- [ ] ✅ Check order of operations in `ImportDocument()`
- [ ] ✅ Test in Word before committing (don't just assume it works!)
- [ ] ✅ Review this document for applicable lessons

---

## VERSION HISTORY

### Errors Made & Fixed:
1. **v1.7.3 → v1.7.4:** Line continuation limit exceeded (25 max)
2. **v1.7.1 → v1.7.2:** Complex whitespace stripping failed, simplified in v1.7.3
3. **v1.7.0 → v1.7.2:** Duplicate signatures (3 iterations to fix)
4. **v1.6.6:** Table formatting order (FixAllTables before vs. after global)
5. **Earlier:** Line continuation limit exceeded (first occurrence)

---

## REFERENCE: VBA LIMITS

| Limit | Value | Notes |
|-------|-------|-------|
| Line continuations | 25 max | **HARD LIMIT** - use 20 max for safety |
| Line length | 1,024 characters | Rare issue |
| Module size | 64KB | Split into multiple modules if needed |
| Variable name | 255 characters | Non-issue |
| Procedure lines | ~65,000 | Non-issue |

---

**Last Updated:** 2025-10-24 (v1.7.4)
**Maintained By:** Claude Code
**Project:** BWS (Bridgewater Studio) Macro Development

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

### 2. ⚠️ VARIABLE DECLARATIONS MUST BE AT FUNCTION LEVEL
**VBA HARD RULE: All Dim statements must be at function/subroutine level**

**ERROR HISTORY:**
- ❌ v1.7.5 → v1.7.6: Dim statements inside For loop - compile error
- Error: "Invalid outside procedure" or "Expected end of statement"

**RULE:**
- **ALL Dim statements at function level** (at the top, before any code)
- **NEVER declare variables inside loops, If blocks, or other structures**
- Declare all variables at the top, even if only used in a specific block

**EXAMPLE - WRONG:**
```vba
Sub MyFunction()
    For i = 1 To 10
        Dim tempVar As String  ' ❌ COMPILER ERROR!
        tempVar = "test"
    Next i
End Sub
```

**EXAMPLE - CORRECT:**
```vba
Sub MyFunction()
    Dim tempVar As String  ' ✅ Declared at function level
    Dim i As Long

    For i = 1 To 10
        tempVar = "test"   ' ✅ Only assignment/Set in loop
    Next i
End Sub
```

**VERIFICATION CHECKLIST:**
- [ ] All Dim statements are at the top of the function
- [ ] No Dim statements inside For/While/Do loops
- [ ] No Dim statements inside If/Select blocks
- [ ] No Dim statements inside With blocks

---

## ORDER OF OPERATIONS MATTERS

### 3. Table Formatting AFTER Global Formatting
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

### 4. Template Signature vs. Imported Signature
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

### 5. Empty Paragraph Detection - Beware the Paragraph Mark!
**CRITICAL ISSUE (v1.7.5):** Whitespace stripping was CREATING whitespace instead of removing it!

**THREE BUGS DISCOVERED:**

#### Bug #1: Wrong Empty Paragraph Check
**ISSUE:** `Len(Trim$(para.Range.Text)) = 0` NEVER matches empty paragraphs!

**WHY?** Every paragraph contains a paragraph mark (Chr(13)) with length = 1

**WRONG:**
```vba
If Len(Trim$(para.Range.Text)) = 0 Then  ' ❌ NEVER TRUE for empty paragraphs!
    para.Range.Delete
End If
```

**CORRECT:**
```vba
txt = para.Range.Text
If Len(Trim$(txt)) <= 1 Then  ' ✅ Catches paragraph mark (length 1)
    para.Range.Delete
End If
```

#### Bug #2: Deleting Multiple Paragraphs
**ISSUE:** Deleting the SAME range multiple times doesn't work

**WRONG:**
```vba
For j = 1 To 6
    startPara.Range.Delete  ' ❌ Deleting same range 6 times!
Next j
```

**CORRECT:**
```vba
' Create range spanning 6 paragraphs, delete once
Set deleteRange = startPara.Range
endParaIndex = startPara.Range.Paragraphs(1).Index + 5
If endParaIndex <= doc.Paragraphs.Count Then
    deleteRange.End = doc.Paragraphs(endParaIndex).Range.End
End If
deleteRange.Delete  ' ✅ Delete entire range at once
```

#### Bug #3: Order of Operations
**ISSUE:** Whitespace stripping ran BEFORE operations that create empty paragraphs

**WRONG ORDER:**
```vba
StripAllWhitespace doc        ' ❌ Runs too early
FormatSignatureBlock doc      ' Creates empty paragraphs
' Result: Empty paragraphs remain!
```

**CORRECT ORDER:**
```vba
FormatSignatureBlock doc      ' May create empty paragraphs
' ... other formatting ...
StripAllWhitespace doc        ' ✅ Runs LAST - cleans up everything
```

**RULES:**
1. Empty paragraph check: Use `<= 1` not `= 0` (paragraph mark has length 1)
2. Multi-paragraph deletion: Create range spanning N paragraphs, delete once
3. Whitespace stripping: Run LAST, after all formatting that might create empty paragraphs
4. Prefer simple, brute-force approaches over complex range checking

---

## RANGE & PARAGRAPH MANIPULATION

### 6. Paragraph Objects Don't Have .Index Property
**CRITICAL ISSUE (v1.7.7):** "Method or data member not found" error

**ROOT CAUSE:**
Paragraph objects in Word VBA do NOT have an .Index property!

**ERROR:**
```vba
paraIndex = para.Range.Paragraphs(1).Index  ' ❌ .Index doesn't exist!
```

**VBA FACT:**
Paragraph objects have properties like:
- Range, Style, Format, Alignment, OutlineLevel
- They do NOT have .Index to get their position

**WRONG APPROACH - Using .Index:**
```vba
' ❌ This will cause compile error!
paraIndex = startPara.Range.Paragraphs(1).Index
endParaIndex = paraIndex + 5
deleteRange.End = doc.Paragraphs(endParaIndex).Range.End
```

**CORRECT APPROACH - Using Range Methods:**
```vba
' ✅ Use Range.MoveEnd to extend range by N paragraphs
Set deleteRange = startPara.Range
deleteRange.MoveEnd Unit:=wdParagraph, Count:=5  ' Extend by 5 paragraphs
deleteRange.Delete  ' Delete 6 paragraphs total (start + 5)
```

**USEFUL RANGE METHODS:**
```vba
' Extending ranges:
rng.MoveEnd Unit:=wdParagraph, Count:=5    ' Extend end by 5 paragraphs
rng.MoveStart Unit:=wdParagraph, Count:=2  ' Move start forward 2 paragraphs

' Expanding ranges:
rng.Expand Unit:=wdParagraph  ' Expand to include entire paragraph

' Collapsing ranges:
rng.Collapse Direction:=wdCollapseEnd  ' Collapse to end point
```

**RULES:**
1. Paragraph objects don't have .Index property
2. Use Range methods (MoveStart, MoveEnd, Expand) for range manipulation
3. Range methods handle edge cases automatically (won't go past document boundaries)
4. Range methods are simpler and more reliable than index-based approaches

**BENEFITS:**
- Cleaner code (11 lines → 1 line in v1.7.7)
- No reliance on non-existent properties
- Automatic boundary handling
- More maintainable

---

## IMAGE FORMATTING

### 7. InlineShapes vs. Shapes
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

### 8. Finding Content Controls by Title/Tag
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

### 9. On Error Resume Next - Use Sparingly
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
1. **v1.7.6 → v1.7.7:** Paragraph.Index property doesn't exist - compile error (use Range.MoveEnd instead)
2. **v1.7.5 → v1.7.6:** Dim statements inside For loop - compile error (VBA requires all Dim at function level)
3. **v1.7.4 → v1.7.5:** THREE critical bugs - wrong empty paragraph check (= 0 instead of <= 1), wrong deletion loop (same range 6x), wrong order (whitespace before signature deletion)
4. **v1.7.3 → v1.7.4:** Line continuation limit exceeded (25 max)
5. **v1.7.1 → v1.7.2:** Complex whitespace stripping failed, simplified in v1.7.3
6. **v1.7.0 → v1.7.2:** Duplicate signatures (3 iterations to fix)
7. **v1.6.6:** Table formatting order (FixAllTables before vs. after global)
8. **Earlier:** Line continuation limit exceeded (first occurrence)

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

**Last Updated:** 2025-10-24 (v1.7.7)
**Maintained By:** Claude Code
**Project:** BWS (Bridgewater Studio) Macro Development

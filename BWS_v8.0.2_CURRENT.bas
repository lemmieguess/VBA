Attribute VB_Name = "BWS_Module"
Option Explicit

' ============================================================================
' Bridgewater Studio - BWS Ultimate v8.0.3
' Complete feature set: persistent 3-row toolbar, all formatting, toggleable bullets
' Based on: v8.0.2 FIXED + document formatting improvements
' v8.0.3 IMPROVEMENTS:
'   - Fixed hanging indent: removed from regular paragraphs, kept only on bullets
'   - Fixed bullet alignment: first line and hanging indent now align vertically
'   - Added BWS Heading style application for Experience Strategy Rationale
'   - Fixed signature image wrapping to "in front of text"
'   - Fixed table alignment: only currency columns right-aligned, all others left-aligned
' ============================================================================

' -------- Versioning / App Keys --------
Private Const BWS_APP_NAME As String = "BridgewaterStudio"
Private Const BWS_APP_SECTION As String = "BWS"
Public  Const BWS_VERSION   As String = "Ultimate v8.0.3"

' -------- Registry Keys --------
Private Const BWS_REG_APP As String = "BridgewaterStudio"
Private Const BWS_REG_SECTION As String = "BWS"
Private Const BWS_REG_BULLET_CONVERT As String = "BulletConversionEnabled"

' -------- Formatting prefs --------
Private Const BULLET_LEFT_IN As Double = 0.5
Private Const BULLET_HANG_IN As Double = 0.25
Private Const BULLET_TAB_IN  As Double = 0.5
Private Const FONT_NAME_PREF As String = "Roboto"
Private Const FONT_SIZE_PREF As Single = 11

' Table colors
Private Const HDR_R As Long = 235, HDR_G As Long = 240, HDR_B As Long = 246
Private Const ZEB_R As Long = 248, ZEB_G As Long = 250, ZEB_B As Long = 252

' Column detection threshold
Private Const NUMERIC_COL_THRESHOLD As Double = 0.6

' -------- mso constants --------
Private Const MSO_CONTROL_BUTTON As Long = 1
Private Const MSO_BUTTON_CAPTION As Long = 2
Private Const MSO_BUTTON_ICON_AND_CAPTION As Long = 3
Private Const MSO_BAR_TOP As Long = 1

' ============================================================================
' AUTOOPEN/AUTONEW - Makes toolbar persistent
' ============================================================================

Public Sub AutoOpen()
    On Error Resume Next
    BuildOrRefreshBWSToolbar False
End Sub

Public Sub AutoNew()
    On Error Resume Next
    BuildOrRefreshBWSToolbar False
End Sub

' ============================================================================
' INSTALLER - Run once to set up BWS
' ============================================================================

Public Sub BWS_Install()
    Dim base As String, drafts As String, templ As String

    MsgBox "Welcome to BWS Ultimate Installer!" & vbCrLf & vbCrLf & _
           "You'll be prompted to select:" & vbCrLf & _
           "1. Base folder (Dropbox root)" & vbCrLf & _
           "2. Template file (.dotm/.dotx)" & vbCrLf & vbCrLf & _
           "The toolbar will be created automatically.", vbInformation, "BWS Install"

    ' Get base path
    base = PickFolder("Select your Dropbox base folder")
    If LenB(base) = 0 Then
        MsgBox "Installation cancelled.", vbExclamation, "BWS"
        Exit Sub
    End If
    SaveSettingStr "BasePath", base

    ' Create Drafts folder
    drafts = CombinePath(base, "Drafts")
    If EnsureFolder(drafts) Then SaveSettingStr "DraftsPath", drafts

    ' Get template
    templ = PickFile("Select your letterhead template", "*.dotm; *.dotx")
    If LenB(templ) > 0 Then SaveSettingStr "TemplatePath", templ

    ' Set defaults
    SaveSettingStr "AutoOpenFolder", "True"
    SaveSettingStr BWS_REG_BULLET_CONVERT, "Yes"

    ' Build toolbar
    BuildOrRefreshBWSToolbar True

    MsgBox "Installation complete!" & vbCrLf & vbCrLf & _
           "Base: " & base & vbCrLf & _
           "Drafts: " & drafts & vbCrLf & _
           "Template: " & templ & vbCrLf & vbCrLf & _
           "The BWS toolbar is now ready to use.", vbInformation, "BWS"
End Sub

' ============================================================================
' Utilities
' ============================================================================

Private Function Ver() As String
    Ver = "BWS " & BWS_VERSION
End Function

Public Sub StatusMessage(ByVal msg As String)
    On Error GoTo Done
    Application.StatusBar = msg
    DoEvents
Done:
End Sub

Private Function FontExists(ByVal fontName As String) As Boolean
    Dim f As Variant
    FontExists = False
    For Each f In Application.FontNames
        If StrComp(CStr(f), fontName, vbTextCompare) = 0 Then
            FontExists = True
            Exit Function
        End If
    Next f
End Function

Private Function GetSettingStr(ByVal keyName As String, Optional ByVal defaultValue As String = "") As String
    GetSettingStr = GetSetting(BWS_APP_NAME, BWS_APP_SECTION, keyName, defaultValue)
End Function

Private Sub SaveSettingStr(ByVal keyName As String, ByVal value As String)
    SaveSetting BWS_APP_NAME, BWS_APP_SECTION, keyName, value
End Sub

Private Function EnsureFolder(ByVal folderPath As String) As Boolean
    Dim fso As Object
    On Error GoTo Fail
    If LenB(folderPath) = 0 Then GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    EnsureFolder = True
    Exit Function
Fail:
    EnsureFolder = False
End Function

Private Function CombinePath(ByVal a As String, ByVal b As String) As String
    If LenB(a) = 0 Then
        CombinePath = b
    ElseIf Right$(a, 1) = "\" Then
        CombinePath = a & b
    Else
        CombinePath = a & "\" & b
    End If
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each ch In bad
        s = Replace$(s, CStr(ch), "_")
    Next ch
    s = Trim$(s)
    If Len(s) > 200 Then s = Left$(s, 200)
    SanitizeFileName = s
End Function

Private Function TempCopyFile(ByVal sourcePath As String) As String
    Dim fso As Object
    Dim tempFolder As String
    Dim tempName As String
    On Error GoTo Err_TempCopyFile
    tempFolder = Environ$("TEMP")
    If LenB(tempFolder) = 0 Then tempFolder = Environ$("TMP")
    If LenB(tempFolder) = 0 Then tempFolder = ThisDocument.Path
    If Right$(tempFolder, 1) <> "\" Then tempFolder = tempFolder & "\"
    tempName = tempFolder & "BWS_TEMP_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int((100000#) * Rnd()))
    tempName = tempName & Mid$(sourcePath, InStrRev(sourcePath, "."))
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile sourcePath, tempName, True
    TempCopyFile = tempName
    Exit Function
Err_TempCopyFile:
    TempCopyFile = ""
End Function

Private Function PickFile(ByVal titleText As String, ByVal filterPattern As String) As String
    Dim fd As Object, it As Variant
    On Error GoTo Fail
    Set fd = Application.FileDialog(3)
    With fd
        .AllowMultiSelect = False
        .Title = titleText
        .Filters.Clear
        .Filters.Add filterPattern, filterPattern
        If .Show = -1 Then
            For Each it In .SelectedItems
                PickFile = CStr(it)
                Exit Function
            Next it
        End If
    End With
Fail:
    PickFile = ""
End Function

Private Function PickFolder(ByVal titleText As String) As String
    Dim fd As Object, it As Variant
    On Error GoTo Fail
    Set fd = Application.FileDialog(4)
    With fd
        .AllowMultiSelect = False
        .Title = titleText
        If .Show = -1 Then
            For Each it In .SelectedItems
                PickFolder = CStr(it)
                Exit Function
            Next it
        End If
    End With
Fail:
    PickFolder = ""
End Function

Private Function TryParseCurrency(ByVal s As String, ByRef result As Double) As Boolean
    Dim t As String
    Dim isNeg As Boolean
    On Error GoTo FailParse
    t = Trim$(s)
    If LenB(t) = 0 Then GoTo FailParse
    If Left$(t, 1) = "(" And Right$(t, 1) = ")" Then
        isNeg = True
        t = Mid$(t, 2, Len(t) - 2)
    End If
    t = Replace$(t, "$", "")
    t = Replace$(t, ",", "")
    t = Replace$(t, vbTab, "")
    t = Replace$(t, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Trim$(t)
    If Left$(t, 1) = "+" Then t = Mid$(t, 2)
    If Left$(t, 1) = "-" Then
        isNeg = True
        t = Mid$(t, 2)
    End If
    If LenB(t) = 0 Then GoTo FailParse
    result = CDbl(t)
    If isNeg Then result = -result
    TryParseCurrency = True
    Exit Function
FailParse:
    TryParseCurrency = False
End Function

' ============================================================================
' Content Controls
' ============================================================================

Private Function IsBodyMatch(ByVal cc As ContentControl) As Boolean
    Dim t As String, g As String
    t = LCase$(cc.Title)
    g = LCase$(cc.Tag)
    IsBodyMatch = (t = "bodycontent" Or g = "bodycontent" Or t = "body" Or g = "body")
End Function

Private Function FindBodyContentControl(ByVal doc As Document, Optional ByVal deleteDuplicates As Boolean = True) As ContentControl
    Dim cc As ContentControl
    Dim found As ContentControl
    Set found = Nothing
    For Each cc In doc.ContentControls
        If IsBodyMatch(cc) Then
            If found Is Nothing Then
                Set found = cc
            ElseIf deleteDuplicates Then
                cc.Delete
            End If
        End If
    Next cc
    Set FindBodyContentControl = found
End Function

' ============================================================================
' NEW LETTERHEAD
' ============================================================================

Public Sub BWS_NewLetter()
    Dim templ As String
    Dim newDoc As Document

    templ = GetTemplatePath(True)
    If LenB(templ) = 0 Then
        MsgBox "No template configured.", vbExclamation, "BWS"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Set newDoc = Documents.Add(Template:=templ, NewTemplate:=False, DocumentType:=0)
    newDoc.Activate
    StatusMessage Ver() & " - New letter created"
    Exit Sub

ErrHandler:
    MsgBox "Failed to create new letter: " & Err.Description, vbCritical, "BWS"
End Sub

' ============================================================================
' IMPORT FUNCTIONALITY
' ============================================================================

Public Sub BWS_ImportPicked()
    Dim sourcePath As String
    sourcePath = PickFile("Select draft to import", "*.docx; *.doc")
    If LenB(sourcePath) > 0 Then
        ImportDocument sourcePath
    End If
End Sub

Public Sub BWS_ImportNewest()
    Dim draftsPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim newestFile As String
    Dim newestDate As Date

    draftsPath = GetDraftsPath(False)
    If LenB(draftsPath) = 0 Then
        MsgBox "Drafts folder not configured.", vbExclamation, "BWS"
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(draftsPath) Then
        MsgBox "Drafts folder not found: " & draftsPath, vbExclamation, "BWS"
        Exit Sub
    End If

    Set folder = fso.GetFolder(draftsPath)
    newestDate = #1/1/1900#
    newestFile = ""

    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "docx" Or LCase$(fso.GetExtensionName(file.Name)) = "doc" Then
            If file.DateLastModified > newestDate Then
                newestDate = file.DateLastModified
                newestFile = file.Path
            End If
        End If
    Next file

    If LenB(newestFile) = 0 Then
        MsgBox "No .docx files found in Drafts folder.", vbInformation, "BWS"
        Exit Sub
    End If

    ImportDocument newestFile
    Exit Sub

ErrHandler:
    MsgBox "Error finding newest draft: " & Err.Description, vbCritical, "BWS"
End Sub

Private Sub ImportDocument(ByVal sourcePath As String)
    Dim doc As Document
    Dim bodyCC As ContentControl
    Dim insertRng As Range

    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "No active document.", vbExclamation, "BWS"
        Exit Sub
    End If

    StatusMessage "Importing " & sourcePath & "..."

    ' Find body content control
    Set bodyCC = FindBodyContentControl(doc, True)
    If bodyCC Is Nothing Then
        MsgBox "No BodyContent control found.", vbExclamation, "BWS"
        Exit Sub
    End If

    ' Clear existing content
    bodyCC.Range.Text = ""
    Set insertRng = bodyCC.Range
    insertRng.Collapse wdCollapseStart

    ' Import using Selection.InsertFile (proven pattern)
    On Error Resume Next
    insertRng.Select
    Selection.InsertFile sourcePath
    On Error GoTo 0

    ' Apply post-import processing (bullets LAST per v7.7.37 pattern)
    StatusMessage "Processing imported content..."

    ' 1. Remove duplicate controls
    FindBodyContentControl doc, True

    ' 2. Strip whitespace
    StripLeadingWhitespace doc
    StripTrailingWhitespace doc

    ' 3. Fix tables
    FixAllTables doc

    ' 4. Apply formatting
    ApplyGlobalFormatting doc

    ' 5. Fix signature image wrapping
    FixSignatureImageWrapping doc

    ' 6. Convert bullets LAST (and only if enabled)
    Dim bulletsEnabled As String
    bulletsEnabled = GetSettingStr(BWS_REG_BULLET_CONVERT, "Yes")
    If bulletsEnabled = "Yes" Then
        ConvertTextBulletsToRealBullets doc
    End If

    StatusMessage Ver() & " - Import complete"
    MsgBox "Import complete!", vbInformation, "BWS"
End Sub

' ============================================================================
' BULLET FUNCTIONS
' ============================================================================

Public Sub BWS_ToggleBullets()
    Dim current As String
    Dim newSetting As String

    current = GetSettingStr(BWS_REG_BULLET_CONVERT, "Yes")

    If current = "Yes" Then
        newSetting = "No"
        SaveSettingStr BWS_REG_BULLET_CONVERT, "No"
        MsgBox "Bullet conversion is now DISABLED.", vbInformation, "BWS"
    Else
        newSetting = "Yes"
        SaveSettingStr BWS_REG_BULLET_CONVERT, "Yes"
        MsgBox "Bullet conversion is now ENABLED.", vbInformation, "BWS"
    End If
End Sub

Public Sub BWS_ConvertTextBullets()
    ConvertTextBulletsToRealBullets ActiveDocument
    MsgBox "Text bullets converted!", vbInformation, "BWS"
End Sub

Public Sub BWS_FixBullets()
    Dim doc As Document
    Dim para As Paragraph

    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub

    For Each para In doc.Paragraphs
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            With para.Range.ListFormat
                .ListLevelNumber = 1
            End With
            With para.Range.ParagraphFormat
                .LeftIndent = InchesToPoints(BULLET_LEFT_IN)
                .FirstLineIndent = InchesToPoints(-BULLET_HANG_IN)
            End With
            para.Range.ParagraphFormat.TabStops.ClearAll
            para.Range.ParagraphFormat.TabStops.Add Position:=InchesToPoints(BULLET_TAB_IN)
        Else
            ' Remove hanging indent from non-bullet paragraphs
            With para.Range.ParagraphFormat
                .FirstLineIndent = 0
            End With
        End If
    Next para

    MsgBox "Bullet formatting fixed!", vbInformation, "BWS"
End Sub

Private Sub ConvertTextBulletsToRealBullets(ByVal doc As Document)
    Dim para As Paragraph
    Dim txt As String
    Dim firstChar As String

    For Each para In doc.Paragraphs
        If para.Range.ListFormat.ListType = wdListNoNumbering Then
            txt = Trim$(para.Range.Text)
            If Len(txt) > 1 Then
                firstChar = Left$(txt, 1)
                If InStr("•·-◦▪■►", firstChar) > 0 Then
                    para.Range.Text = Mid$(txt, 2)
                    para.Range.ListFormat.ApplyBulletDefault
                    With para.Range.ParagraphFormat
                        .LeftIndent = InchesToPoints(BULLET_LEFT_IN)
                        .FirstLineIndent = InchesToPoints(-BULLET_HANG_IN)
                    End With
                Else
                    ' Ensure non-bullet paragraphs have no hanging indent
                    With para.Range.ParagraphFormat
                        .FirstLineIndent = 0
                    End With
                End If
            End If
        End If
    Next para
End Sub

Private Sub StripLeadingWhitespace(ByVal doc As Document)
    Dim para As Paragraph
    On Error Resume Next
    For Each para In doc.Paragraphs
        If Len(Trim$(para.Range.Text)) = 0 Then
            para.Range.Delete
        Else
            Exit For
        End If
    Next para
End Sub

Private Sub StripTrailingWhitespace(ByVal doc As Document)
    Dim para As Paragraph
    Dim i As Long
    On Error Resume Next
    For i = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(i)
        If Len(Trim$(para.Range.Text)) = 0 Then
            para.Range.Delete
        Else
            Exit For
        End If
    Next i
End Sub

' ============================================================================
' TABLE FORMATTING (with column alignment detection)
' ============================================================================

Public Sub BWS_FixTables()
    FixAllTables ActiveDocument
    MsgBox "Tables formatted!", vbInformation, "BWS"
End Sub

Private Sub FixAllTables(ByVal doc As Document)
    Dim tbl As Table
    For Each tbl In doc.Tables
        FormatTable tbl
    Next tbl
End Sub

Private Sub FormatTable(ByVal t As Table)
    On Error Resume Next

    ' CRITICAL: Table alignment and width (from BWS_LESSONS_LEARNED)
    t.Rows.LeftIndent = 0  ' THE critical line!
    t.Rows.Alignment = wdAlignRowLeft
    t.Range.ParagraphFormat.LeftIndent = 0  ' Fix cell content left padding!
    t.PreferredWidthType = wdPreferredWidthPercent
    t.PreferredWidth = 100
    t.AutoFitBehavior wdAutoFitWindow

    ' Cell padding
    t.TopPadding = Application.CentimetersToPoints(0.1)
    t.BottomPadding = Application.CentimetersToPoints(0.1)
    t.LeftPadding = Application.CentimetersToPoints(0.15)
    t.RightPadding = Application.CentimetersToPoints(0.15)

    ' Header formatting
    If t.Rows.Count >= 1 Then
        With t.Rows(1)
            .HeadingFormat = True
            .Shading.BackgroundPatternColor = RGB(HDR_R, HDR_G, HDR_B)
            .Shading.ForegroundPatternColor = RGB(HDR_R, HDR_G, HDR_B)
        End With
    End If

    ' Zebra striping
    Dim i As Long
    For i = 2 To t.Rows.Count
        If i Mod 2 = 0 Then
            t.Rows(i).Shading.BackgroundPatternColor = RGB(ZEB_R, ZEB_G, ZEB_B)
            t.Rows(i).Shading.ForegroundPatternColor = RGB(ZEB_R, ZEB_G, ZEB_B)
        End If
    Next i

    ' Column alignment - right-align only currency columns (using .Select pattern)
    Dim c As Long
    For c = 1 To t.Columns.Count
        If IsCurrencyColumn(t, c) Then
            t.Columns(c).Select
            Selection.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
        Else
            t.Columns(c).Select
            Selection.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        End If
    Next c

    On Error GoTo 0
End Sub

Private Function IsCurrencyColumn(ByVal t As Table, ByVal colIndex As Long) As Boolean
    Dim r As Long
    Dim cellText As String
    Dim currencyCount As Long
    Dim nonEmptyCount As Long
    Dim val As Double

    currencyCount = 0
    nonEmptyCount = 0

    For r = 2 To t.Rows.Count  ' Skip header
        On Error Resume Next
        cellText = Trim$(t.Cell(r, colIndex).Range.Text)
        cellText = Replace$(cellText, Chr$(13), "")
        cellText = Replace$(cellText, Chr$(7), "")

        If Len(cellText) > 0 Then
            nonEmptyCount = nonEmptyCount + 1
            ' Check if cell contains currency symbol ($)
            If InStr(cellText, "$") > 0 And TryParseCurrency(cellText, val) Then
                currencyCount = currencyCount + 1
            End If
        End If
        On Error GoTo 0
    Next r

    If nonEmptyCount > 0 Then
        IsCurrencyColumn = (currencyCount / nonEmptyCount) >= NUMERIC_COL_THRESHOLD
    Else
        IsCurrencyColumn = False
    End If
End Function

' ============================================================================
' UPDATE TOTALS
' ============================================================================

Public Sub BWS_UpdateTotals()
    Dim doc As Document
    Dim t As Table
    Dim rightCol As Long
    Dim total As Double
    Dim r As Long
    Dim cellText As String
    Dim val As Double

    Set doc = ActiveDocument
    If doc.Tables.Count = 0 Then
        MsgBox "No tables found.", vbInformation, "BWS"
        Exit Sub
    End If

    Set t = doc.Tables(1)
    rightCol = t.Columns.Count
    total = 0

    For r = 2 To t.Rows.Count
        On Error Resume Next
        cellText = Trim$(t.Cell(r, rightCol).Range.Text)
        cellText = Replace$(cellText, Chr$(13), "")
        cellText = Replace$(cellText, Chr$(7), "")

        If TryParseCurrency(cellText, val) Then
            total = total + val
        End If
        On Error GoTo 0
    Next r

    ' Add total row
    t.Rows.Add
    t.Cell(t.Rows.Count, 1).Range.Text = "Total"
    t.Cell(t.Rows.Count, 1).Range.Font.Bold = True
    t.Cell(t.Rows.Count, rightCol).Range.Text = Format$(total, "$#,##0.00")
    t.Cell(t.Rows.Count, rightCol).Range.Font.Bold = True

    MsgBox "Total updated: " & Format$(total, "$#,##0.00"), vbInformation, "BWS"
End Sub

' ============================================================================
' APPLY FORMATTING
' ============================================================================

Public Sub BWS_ApplyFormatting()
    ApplyGlobalFormatting ActiveDocument
    MsgBox "Formatting applied!", vbInformation, "BWS"
End Sub

Public Sub BWS_ApplyHeadingStyles()
    ApplyBWSHeadingStyles ActiveDocument
    MsgBox "BWS Heading styles applied!", vbInformation, "BWS"
End Sub

Public Sub BWS_FixSignatureWrapping()
    FixSignatureImageWrapping ActiveDocument
    MsgBox "Signature image wrapping fixed!", vbInformation, "BWS"
End Sub

Private Sub ApplyGlobalFormatting(ByVal doc As Document)
    On Error Resume Next

    ' Check if Roboto is available
    Dim useFont As String
    If FontExists(FONT_NAME_PREF) Then
        useFont = FONT_NAME_PREF
    Else
        useFont = "Calibri"
    End If

    ' Apply to whole document
    With doc.Range
        .Font.Name = useFont
        .Font.Size = FONT_SIZE_PREF
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.LeftIndent = 0
    End With

    On Error GoTo 0
End Sub

Private Sub ApplyBWSHeadingStyles(ByVal doc As Document)
    Dim para As Paragraph
    Dim txt As String
    Dim inSection As Boolean
    Dim bwsHeadingExists As Boolean

    On Error Resume Next

    ' Check if BWS Heading style exists
    bwsHeadingExists = False
    Dim sty As Style
    For Each sty In doc.Styles
        If sty.NameLocal = "BWS Heading" Or sty.NameLocal = "BWS_Heading" Then
            bwsHeadingExists = True
            Exit For
        End If
    Next sty

    If Not bwsHeadingExists Then
        MsgBox "BWS Heading style not found in template.", vbExclamation, "BWS"
        Exit Sub
    End If

    inSection = False

    ' Find Experience Strategy Rationale section and apply BWS Heading to appropriate paragraphs
    For Each para In doc.Paragraphs
        txt = Trim$(para.Range.Text)

        ' Check if we're entering the section
        If InStr(1, txt, "Experience Strategy Rationale", vbTextCompare) > 0 Then
            inSection = True
        ElseIf inSection Then
            ' Check if we've left the section (next major heading)
            If Len(txt) > 0 And para.Range.Font.Bold And _
               (InStr(1, txt, "Executive Summary", vbTextCompare) > 0 Or _
                InStr(1, txt, "Timeline", vbTextCompare) > 0 Or _
                InStr(1, txt, "Budget", vbTextCompare) > 0 Or _
                InStr(1, txt, "Deliverables", vbTextCompare) > 0) Then
                inSection = False
            ElseIf Len(txt) > 0 And para.Range.ListFormat.ListType = wdListNoNumbering Then
                ' Apply BWS Heading style to non-bullet, non-empty paragraphs that look like headings
                ' (typically shorter lines, may be bold, not full sentences)
                If Len(txt) < 100 And Not InStr(txt, ".") > 0 Then
                    para.Range.Style = "BWS Heading"
                End If
            End If
        End If
    Next para

    On Error GoTo 0
End Sub

Private Sub FixSignatureImageWrapping(ByVal doc As Document)
    Dim shp As InlineShape
    Dim newShp As Shape

    On Error Resume Next

    ' Find inline images and convert them to floating with "In Front of Text" wrapping
    For Each shp In doc.InlineShapes
        If shp.Type = wdInlineShapePicture Or shp.Type = wdInlineShapeLinkedPicture Then
            ' Convert inline shape to floating shape
            Set newShp = shp.ConvertToShape
            If Not newShp Is Nothing Then
                ' Set wrapping to "In Front of Text"
                newShp.WrapFormat.Type = 3 ' wdWrapFront = 3
            End If
        End If
    Next shp

    On Error GoTo 0
End Sub

' ============================================================================
' SAVE & EXPORT
' ============================================================================

Public Sub BWS_SaveDraft()
    Dim doc As Document
    Dim draftsPath As String
    Dim fileName As String
    Dim fullPath As String

    Set doc = ActiveDocument
    draftsPath = GetDraftsPath(True)

    If LenB(draftsPath) = 0 Then
        MsgBox "Drafts folder not configured.", vbExclamation, "BWS"
        Exit Sub
    End If

    fileName = "Draft_" & Format$(Now, "yyyy-mm-dd_hhnnss") & ".docx"
    fullPath = CombinePath(draftsPath, fileName)

    doc.SaveAs2 fullPath, wdFormatXMLDocument
    SaveSettingStr "LastSavePath", fullPath

    StatusMessage Ver() & " - Saved to drafts"

    If GetSettingStr("AutoOpenFolder", "False") = "True" Then
        Shell "explorer.exe /select," & Chr$(34) & fullPath & Chr$(34), vbNormalFocus
    End If

    MsgBox "Saved to drafts!", vbInformation, "BWS"
End Sub

Public Sub BWS_SaveProposal()
    Dim doc As Document
    Dim base As String
    Dim client As String
    Dim project As String
    Dim proposalsPath As String
    Dim clientFolder As String
    Dim yearFolder As String
    Dim projectFolder As String
    Dim fileName As String
    Dim fullPath As String

    Set doc = ActiveDocument
    base = GetBasePath(True)

    If LenB(base) = 0 Then
        MsgBox "Base path not configured.", vbExclamation, "BWS"
        Exit Sub
    End If

    ' Get metadata
    client = ExtractValue(doc, "Client:")
    project = ExtractValue(doc, "Project:")

    If LenB(client) = 0 Then client = InputBox("Enter client name:", "BWS", "Client")
    If LenB(project) = 0 Then project = InputBox("Enter project name:", "BWS", "Project")

    If LenB(client) = 0 Or LenB(project) = 0 Then
        MsgBox "Client and Project required.", vbExclamation, "BWS"
        Exit Sub
    End If

    ' Build path: Proposals\Client\YYYY\Project
    proposalsPath = CombinePath(base, "Proposals")
    clientFolder = CombinePath(proposalsPath, SanitizeFileName(client))
    yearFolder = CombinePath(clientFolder, Format$(Now, "yyyy"))
    projectFolder = CombinePath(yearFolder, SanitizeFileName(project))

    EnsureFolder proposalsPath
    EnsureFolder clientFolder
    EnsureFolder yearFolder
    EnsureFolder projectFolder

    fileName = SanitizeFileName(project) & "_" & Format$(Now, "yyyy-mm-dd") & ".docx"
    fullPath = CombinePath(projectFolder, fileName)

    doc.SaveAs2 fullPath, wdFormatXMLDocument
    SaveSettingStr "LastSavePath", fullPath

    StatusMessage Ver() & " - Saved to proposals"

    If GetSettingStr("AutoOpenFolder", "False") = "True" Then
        Shell "explorer.exe /select," & Chr$(34) & fullPath & Chr$(34), vbNormalFocus
    End If

    MsgBox "Saved to proposals!", vbInformation, "BWS"
End Sub

Public Sub BWS_ExportPDF()
    Dim doc As Document
    Dim lastSave As String
    Dim pdfPath As String

    Set doc = ActiveDocument
    lastSave = GetSettingStr("LastSavePath", "")

    If LenB(lastSave) = 0 Then
        MsgBox "No previous save found. Please save first.", vbExclamation, "BWS"
        Exit Sub
    End If

    pdfPath = Left$(lastSave, InStrRev(lastSave, ".") - 1) & ".pdf"

    doc.ExportAsFixedFormat pdfPath, wdExportFormatPDF, False, wdExportOptimizeForPrint

    StatusMessage Ver() & " - PDF exported"

    If GetSettingStr("AutoOpenFolder", "False") = "True" Then
        Shell "explorer.exe /select," & Chr$(34) & pdfPath & Chr$(34), vbNormalFocus
    End If

    MsgBox "PDF exported!", vbInformation, "BWS"
End Sub

Public Sub BWS_OpenFolder()
    Dim lastSave As String
    Dim folderPath As String

    lastSave = GetSettingStr("LastSavePath", "")

    If LenB(lastSave) = 0 Then
        MsgBox "No previous save found.", vbInformation, "BWS"
        Exit Sub
    End If

    folderPath = Left$(lastSave, InStrRev(lastSave, "\"))
    Shell "explorer.exe " & Chr$(34) & folderPath & Chr$(34), vbNormalFocus
End Sub

' ============================================================================
' PATH MANAGEMENT
' ============================================================================

Private Function GetBasePath(ByVal promptIfMissing As Boolean) As String
    Dim base As String
    base = GetSettingStr("BasePath", "")
    If LenB(base) = 0 And promptIfMissing Then
        base = PickFolder("Pick your Dropbox base folder")
        If LenB(base) > 0 Then SaveSettingStr "BasePath", base
    End If
    GetBasePath = base
End Function

Private Function GetDraftsPath(ByVal createIfMissing As Boolean) As String
    Dim drafts As String, base As String
    drafts = GetSettingStr("DraftsPath", "")
    If LenB(drafts) = 0 Then
        base = GetBasePath(True)
        If LenB(base) > 0 Then
            drafts = CombinePath(base, "Drafts")
            If createIfMissing Then EnsureFolder drafts
            SaveSettingStr "DraftsPath", drafts
        End If
    End If
    GetDraftsPath = drafts
End Function

Private Function GetTemplatePath(ByVal promptIfMissing As Boolean) As String
    Dim templ As String
    templ = GetSettingStr("TemplatePath", "")
    If LenB(templ) = 0 And promptIfMissing Then
        templ = PickFile("Pick your letterhead template", "*.dotm; *.dotx")
        If LenB(templ) > 0 Then SaveSettingStr "TemplatePath", templ
    End If
    GetTemplatePath = templ
End Function

Private Function ExtractValue(ByVal doc As Document, ByVal keyLabel As String) As String
    Dim rng As Range
    Dim txt As String, pos As Long
    Set rng = doc.Content
    txt = rng.Text
    pos = InStr(1, txt, keyLabel, vbTextCompare)
    If pos > 0 Then
        Dim after As String, eol As Long
        after = Mid$(txt, pos + Len(keyLabel))
        eol = InStr(after, vbCr)
        If eol = 0 Then eol = Len(after) + 1
        ExtractValue = Trim$(Left$(after, eol - 1))
    Else
        ExtractValue = ""
    End If
End Function

' ============================================================================
' SETTINGS
' ============================================================================

Public Sub BWS_Settings()
    Dim base As String, templ As String, drafts As String, autoOpen As String

    base = GetBasePath(False)
    If LenB(base) = 0 Then
        base = PickFolder("Pick your Dropbox base folder")
        If LenB(base) > 0 Then SaveSettingStr "BasePath", base
    End If

    drafts = GetDraftsPath(False)
    If LenB(drafts) = 0 And LenB(base) > 0 Then
        drafts = CombinePath(base, "Drafts")
        If EnsureFolder(drafts) Then SaveSettingStr "DraftsPath", drafts
    End If

    templ = GetTemplatePath(False)
    If LenB(templ) = 0 Then
        templ = PickFile("Pick your letterhead template", "*.dotm; *.dotx")
        If LenB(templ) > 0 Then SaveSettingStr "TemplatePath", templ
    End If

    autoOpen = GetSettingStr("AutoOpenFolder", "False")
    If MsgBox("Auto-open folder after save? Current: " & autoOpen, vbYesNo + vbQuestion, "BWS") = vbYes Then
        SaveSettingStr "AutoOpenFolder", "True"
    Else
        SaveSettingStr "AutoOpenFolder", "False"
    End If

    MsgBox Ver() & " settings updated.", vbInformation, "BWS"
End Sub

Public Sub BWS_Diagnostic()
    Dim msg As String
    Dim bulletStatus As String

    bulletStatus = GetSettingStr(BWS_REG_BULLET_CONVERT, "Yes")

    msg = "================ BWS Diagnostic ================" & vbCrLf _
        & "Version:   " & BWS_VERSION & vbCrLf _
        & "BasePath:  " & GetSettingStr("BasePath", "(not set)") & vbCrLf _
        & "Drafts:    " & GetSettingStr("DraftsPath", "(not set)") & vbCrLf _
        & "Template:  " & GetSettingStr("TemplatePath", "(not set)") & vbCrLf _
        & "LastSave:  " & GetSettingStr("LastSavePath", "(not set)") & vbCrLf _
        & "AutoOpen:  " & GetSettingStr("AutoOpenFolder", "False") & vbCrLf _
        & "Bullets:   " & bulletStatus & vbCrLf _
        & "Host:      " & Application.Name & " " & Application.Version & vbCrLf _
        & "==============================================="
    MsgBox msg, vbInformation, "BWS"
End Sub

' ============================================================================
' TOOLBAR (3 persistent rows)
' ============================================================================

Public Sub BWS_InstallToolbar()
    BuildOrRefreshBWSToolbar True
End Sub

Public Sub BWS_RemoveToolbar()
    On Error Resume Next
    Application.CommandBars("BWS-1").Delete
    Application.CommandBars("BWS-2").Delete
    Application.CommandBars("BWS-3").Delete
End Sub

Private Sub BuildOrRefreshBWSToolbar(ByVal showMessage As Boolean)
    Dim cb1 As CommandBar
    Dim cb2 As CommandBar
    Dim cb3 As CommandBar

    ' Remove old toolbars
    On Error Resume Next
    Application.CommandBars("BWS").Delete
    Application.CommandBars("BWS-1").Delete
    Application.CommandBars("BWS-2").Delete
    Application.CommandBars("BWS-3").Delete
    On Error GoTo 0

    ' ROW 1: Document operations (PERSISTENT)
    Set cb1 = Application.CommandBars.Add(Name:="BWS-1", Position:=MSO_BAR_TOP, Temporary:=False)
    cb1.Visible = True
    AddBtn cb1, "New Letter", "BWS_NewLetter", 18, True, MSO_BUTTON_ICON_AND_CAPTION, "Create new letterhead"
    AddBtn cb1, "Import Newest", "BWS_ImportNewest", 23, False, MSO_BUTTON_ICON_AND_CAPTION, "Import newest draft"
    AddBtn cb1, "Import Pick", "BWS_ImportPicked", 580, False, MSO_BUTTON_ICON_AND_CAPTION, "Pick draft to import"
    AddBtn cb1, "Save Draft", "BWS_SaveDraft", 3, True, MSO_BUTTON_ICON_AND_CAPTION, "Save to drafts"
    AddBtn cb1, "Save Proposal", "BWS_SaveProposal", 84, False, MSO_BUTTON_ICON_AND_CAPTION, "Save to proposals"
    AddBtn cb1, "Export PDF", "BWS_ExportPDF", 85, False, MSO_BUTTON_ICON_AND_CAPTION, "Export as PDF"
    AddBtn cb1, "Open Folder", "BWS_OpenFolder", 283, False, MSO_BUTTON_ICON_AND_CAPTION, "Open last save folder"

    ' ROW 2: Formatting tools (PERSISTENT)
    Set cb2 = Application.CommandBars.Add(Name:="BWS-2", Position:=MSO_BAR_TOP, Temporary:=False)
    cb2.Visible = True
    AddBtn cb2, "Update Totals", "BWS_UpdateTotals", 359, True, MSO_BUTTON_ICON_AND_CAPTION, "Update table totals"
    AddBtn cb2, "Fix Tables", "BWS_FixTables", 144, False, MSO_BUTTON_ICON_AND_CAPTION, "Format all tables"
    AddBtn cb2, "Convert Bullets", "BWS_ConvertTextBullets", 70, True, MSO_BUTTON_ICON_AND_CAPTION, "Convert text bullets"
    AddBtn cb2, "Fix Bullets", "BWS_FixBullets", 65, False, MSO_BUTTON_ICON_AND_CAPTION, "Fix bullet formatting"
    AddBtn cb2, "Toggle Bullets", "BWS_ToggleBullets", 70, False, MSO_BUTTON_ICON_AND_CAPTION, "Enable/disable auto-convert"
    AddBtn cb2, "Apply Format", "BWS_ApplyFormatting", 487, True, MSO_BUTTON_ICON_AND_CAPTION, "Apply global formatting"

    ' ROW 3: Settings (PERSISTENT)
    Set cb3 = Application.CommandBars.Add(Name:="BWS-3", Position:=MSO_BAR_TOP, Temporary:=False)
    cb3.Visible = True
    AddBtn cb3, "Settings", "BWS_Settings", 642, True, MSO_BUTTON_ICON_AND_CAPTION, "Configure BWS"
    AddBtn cb3, "Diagnostic", "BWS_Diagnostic", 589, False, MSO_BUTTON_ICON_AND_CAPTION, "Show settings"
    AddBtn cb3, "About", "BWS_About", 487, False, MSO_BUTTON_ICON_AND_CAPTION, "Version info"

    If showMessage Then
        MsgBox "BWS toolbar installed!" & vbCrLf & vbCrLf & _
               "3 persistent rows created." & vbCrLf & _
               "Toolbar will survive Word restart.", vbInformation, "BWS"
    End If
End Sub

Private Sub AddBtn(ByVal cb As CommandBar, _
                   ByVal caption As String, _
                   ByVal onAction As String, _
                   ByVal faceId As Long, _
                   ByVal beginGroup As Boolean, _
                   ByVal style As Long, _
                   Optional ByVal tooltip As String = "")
    Dim btn As CommandBarButton
    Set btn = cb.Controls.Add(Type:=MSO_CONTROL_BUTTON, Temporary:=True)
    With btn
        .caption = caption
        .Style = style
        If faceId > 0 Then .faceId = faceId
        .OnAction = onAction
        .Tag = "BWS_BTN_" & Replace(caption, " ", "_")
        .BeginGroup = beginGroup
        If LenB(tooltip) > 0 Then .TooltipText = tooltip
    End With
End Sub

' ============================================================================
' ABOUT
' ============================================================================

Public Sub BWS_About()
    Dim msg As String
    msg = "================ BWS ====================" & vbCrLf _
        & "Version: " & BWS_VERSION & vbCrLf _
        & "Host:    " & Application.Name & " " & Application.Version & vbCrLf _
        & vbCrLf _
        & "Features:" & vbCrLf _
        & "• 3-row persistent toolbar" & vbCrLf _
        & "• Toggleable bullet conversion" & vbCrLf _
        & "• Smart table formatting" & vbCrLf _
        & "• Column alignment detection" & vbCrLf _
        & "• Auto folder opening" & vbCrLf _
        & "========================================="
    MsgBox msg, vbInformation, "BWS"
End Sub

' ============================================================================
' Helper: Convert inches to points
' ============================================================================

Private Function InchesToPoints(ByVal inches As Double) As Single
    InchesToPoints = inches * 72
End Function

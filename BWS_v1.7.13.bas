Attribute VB_Name = "BWS_Module"
Option Explicit

' ============================================================================
' Bridgewater Studio - BWS v1.7.13 - Table Indent & Image Preservation Fix
' Based on: v1.7.12 with table formatting and signature image fixes
'
' v1.7.13 CRITICAL FIXES:
' - Fixed table indentation: Added RightIndent = 0 for table rows and cells
'   Problem: Tables had negative right indent (-0.25") after import
'   Solution: Explicitly set RightIndent = 0 for both table.Rows and each cell
' - Fixed table cell spacing: Added SpaceBefore = 0 and SpaceAfter = 0 for all cells
'   Ensures consistent zero spacing in all table cells
' - Fixed signature image deletion in StripAllWhitespace
'   Problem: Signature images were deleted during whitespace cleanup
'   Solution: Check para.Range.InlineShapes.Count > 0 before deleting paragraphs
'   Paragraphs containing images are now preserved
'
' v1.7.12 CRITICAL FIX:
' - Fixed toolbar duplication when installed in Normal.dotm
'   Problem: AutoOpen runs every time a document opens, recreating toolbars
'   Solution: Added BWS_ToolbarsExist() check to skip recreation if toolbars exist
'   BuildOrRefreshBWSToolbar now only recreates when:
'   * Toolbars don't exist yet, OR
'   * Explicitly requested via BWS_Install/BWS_InstallToolbar (showMessage=True)
'   This prevents multiple toolbar instances in Normal.dotm scenario
'
' v1.7.11 CRITICAL FIX:
' - Fixed Dim statement inside If block in ExtractValue function (line 1472)
'   Moved "Dim after As String, eol As Long" to function-level declaration
'   This is the THIRD time we've had Dim placement errors:
'   * v1.7.5 → v1.7.6: Dim inside For loop
'   * v1.7.9 → v1.7.10: Dim inside If block (FormatSignatureBlock)
'   * v1.7.10 → v1.7.11: Dim inside If block (ExtractValue) - THIS FIX
' - Added automated validation workflow (check_vba_code.py) to prevent future occurrences
'
' v1.7.3 CRITICAL FIXES:
' - DELETE imported signatures: Remove ALL "Sincerely" signatures from body content
'   Template already has signature - imported signatures cause duplication
'   Solution: Find and DELETE "Sincerely" + next 5 lines from body content
' - Simplified whitespace: Delete ALL empty paragraphs from entire document
'   Previous approach was too complex and didn't work
'   New approach: Simple loop deleting every empty paragraph (up to 100 attempts)
' - Template signature wrapper: Fix image formatting (msoBringToFront)
'
' v1.7.2 BUGFIXES:
' - One-click import: BWS_ImportNewest now creates new document + imports
'   Combined "New Letter" + "Import Newest" into single button click
'   Workflow: Click "Import Newest" → New doc from template → Auto-import
'   Eliminates manual two-step process
'
' v1.6.6.3 FIX:
' - Signature block spacing: 2 blank lines after "Sincerely"
' - Image positioning: Behind text (msoSendToBack) to avoid obscuring name
'
' v1.6.6.2 FIX:
' - Template save prompt: DisplayAlerts = wdAlertsNone during import
'
' v1.6.6.1 FIX:
' - Suppress template save prompt: Set AttachedTemplate.Saved = True
'   Prevents "save changes to BWS letterhead template?" dialog
'   Caused by ApplyHeaderStyles applying "BWS Header" style from template
'
' v1.6.6 FIXES:
' - Fixed signature block: Reverted to v1.6.4 approach, removed formatting ops
'   Keeps all signature content (Sincerely + 3 lines), only fixes image wrapping
' - Fixed hanging indent on paragraphs: ApplyGlobalFormatting now clears indents
'   All regular paragraphs get LeftIndent=0, FirstLineIndent=0 explicitly
' - Fixed table alignment: Apply cell-by-cell instead of column-wide Selection
'   Each cell formatted individually to prevent table-wide overrides
'
' v1.6.5 FIXES:
' - Fixed bullet hanging indent: Override ListTemplate.ListLevels properties
'   Applied to both ConvertTextBulletsToRealBullets and BWS_FixBullets
' - Fixed table alignment order: Move FixAllTables AFTER ApplyGlobalFormatting
'   Prevents global formatting from overriding currency column alignment
'
' v1.6.4 FIXES & FEATURES:
' - Fixed bullet indent: Removed hanging indent for vertical alignment
'   Bullets and text now both align at 0.25" (no overhang)
' - NEW: BWS_Settings() function to change configuration without reinstall
'   Change base folder, drafts folder, template, auto-open, bullet conversion
'
' v1.6.3 FIXES:
' - Fixed bullet hanging indent: Now aligns vertically (0.25" both)
' - Fixed table alignment: Left-align all columns except currency
' - Fixed signature image wrapping: Now "In Front of Text"
' - Added BWS Header style to Executive Summary sections
'
' v1.6.2 ENHANCEMENT:
' - Installer now remembers existing configuration
' - User can choose to keep existing settings or reconfigure
' - Eliminates need to re-select folders on every reinstall
'
' v1.6.1 BUGFIX:
' - Fixed line spacing bug: Changed from wdLineSpaceMultiple to
'   wdLineSpaceExactly with 13.8 points (276 twips = 1.15x spacing)
' - This resolves vertically compressed text issue
'
' v1.6 FEATURES:
' - Signature block detection with 0.125" indent (Story 1)
' - Bullet font fixed to Calibri 11pt (Story 2)
' - Automatic page margin setting including -0.062" top (Story 3)
' - Table headers with #D9D9D9 gray background (Story 4)
' - Metadata line formatting with 0" indent (Story 5)
' - BWS Header style with single line spacing (Story 6)
' ============================================================================

' -------- Versioning / App Keys --------
Private Const BWS_APP_NAME As String = "BridgewaterStudio"
Private Const BWS_APP_SECTION As String = "BWS"
Public  Const BWS_VERSION   As String = "v1.7.13"

' -------- Registry Keys --------
Private Const BWS_REG_APP As String = "BridgewaterStudio"
Private Const BWS_REG_SECTION As String = "BWS"
Private Const BWS_REG_BULLET_CONVERT As String = "BulletConversionEnabled"

' -------- Formatting prefs --------
Private Const BULLET_LEFT_IN As Double = 0.25
Private Const BULLET_HANG_IN As Double = 0.25
Private Const BULLET_TAB_IN  As Double = 0.5
Private Const FONT_NAME_PREF As String = "Roboto"
Private Const BULLET_FONT As String = "Calibri"  ' NEW v1.6: Story 2
Private Const FONT_SIZE_PREF As Single = 11

' NEW v1.6: Story 1 - Signature block formatting
Private Const SIGNATURE_INDENT As Double = 0.125  ' inches

' NEW v1.6: Story 3 - Page margin specs
Private Const MARGIN_TOP As Double = 0.5          ' inches
Private Const MARGIN_BOTTOM As Double = 0.625    ' inches
Private Const MARGIN_LEFT As Double = 0.375      ' inches
Private Const MARGIN_RIGHT As Double = 0.5       ' inches

' NEW v1.6: Story 4 - Table colors (changed from v8.0.2)
Private Const HDR_R As Long = 217, HDR_G As Long = 217, HDR_B As Long = 217  ' #D9D9D9
Private Const ZEB_R As Long = 248, ZEB_G As Long = 250, ZEB_B As Long = 252

' NEW v1.6: Story 6 - Header formatting
Private Const HEADER_FONT As String = "Roboto"
Private Const HEADER_SIZE As Single = 16
Private Const HEADER_COLOR_R As Long = 15
Private Const HEADER_COLOR_G As Long = 71
Private Const HEADER_COLOR_B As Long = 97

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
    Dim existingBase As String, existingTemplate As String
    Dim response As VbMsgBoxResult
    Dim useExisting As Boolean

    ' Check for existing configuration
    existingBase = GetSettingStr("BasePath", "")
    existingTemplate = GetSettingStr("TemplatePath", "")

    ' If settings exist, ask user if they want to keep them
    If LenB(existingBase) > 0 Then
        response = MsgBox("Found existing configuration:" & vbCrLf & vbCrLf & _
                         "Base: " & existingBase & vbCrLf & _
                         "Template: " & existingTemplate & vbCrLf & vbCrLf & _
                         "Keep this configuration?" & vbCrLf & vbCrLf & _
                         "Yes = Keep existing" & vbCrLf & _
                         "No = Choose new folders" & vbCrLf & _
                         "Cancel = Exit installer", _
                         vbYesNoCancel + vbQuestion, "BWS Install")

        If response = vbCancel Then
            MsgBox "Installation cancelled.", vbExclamation, "BWS"
            Exit Sub
        ElseIf response = vbYes Then
            ' Use existing settings
            useExisting = True
            base = existingBase
            templ = existingTemplate
            drafts = GetSettingStr("DraftsPath", CombinePath(base, "Drafts"))
        End If
    End If

    ' If no existing settings or user chose to reconfigure
    If Not useExisting Then
        MsgBox "Welcome to BWS v1.7.13 Installer!" & vbCrLf & vbCrLf & _
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
    End If

    ' Build toolbar (always refresh)
    BuildOrRefreshBWSToolbar True

    MsgBox "Installation complete!" & vbCrLf & vbCrLf & _
           "Base: " & base & vbCrLf & _
           "Drafts: " & drafts & vbCrLf & _
           "Template: " & templ & vbCrLf & vbCrLf & _
           "The BWS toolbar is now ready to use.", vbInformation, "BWS"
End Sub

' ============================================================================
' SETTINGS - Manage BWS configuration without reinstalling
' ============================================================================

Public Sub BWS_Settings()
    Dim choice As VbMsgBoxResult
    Dim newValue As String
    Dim currentBase As String, currentDrafts As String, currentTemplate As String
    Dim currentAutoOpen As String, currentBullets As String

    ' Get current settings
    currentBase = GetSettingStr("BasePath", "(not set)")
    currentDrafts = GetSettingStr("DraftsPath", "(not set)")
    currentTemplate = GetSettingStr("TemplatePath", "(not set)")
    currentAutoOpen = GetSettingStr("AutoOpenFolder", "False")
    currentBullets = GetSettingStr(BWS_REG_BULLET_CONVERT, "Yes")

    ' Show settings menu
    Dim msg As String
    msg = "BWS Settings" & vbCrLf & vbCrLf & _
          "Current configuration:" & vbCrLf & _
          "1. Base folder: " & currentBase & vbCrLf & _
          "2. Drafts folder: " & currentDrafts & vbCrLf & _
          "3. Template file: " & currentTemplate & vbCrLf & _
          "4. Auto-open folder when saving: " & currentAutoOpen & vbCrLf & _
          "5. Bullet conversion: " & currentBullets & vbCrLf & vbCrLf & _
          "What would you like to change?" & vbCrLf & vbCrLf & _
          "Enter 1-5 to change a setting, or click Cancel to exit."

    newValue = InputBox(msg, "BWS Settings", "")
    If LenB(newValue) = 0 Then Exit Sub

    Select Case newValue
        Case "1"  ' Base folder
            newValue = PickFolder("Select new base folder (Dropbox root)")
            If LenB(newValue) > 0 Then
                SaveSettingStr "BasePath", newValue
                MsgBox "Base folder updated to:" & vbCrLf & newValue, vbInformation, "BWS Settings"
            End If

        Case "2"  ' Drafts folder
            newValue = PickFolder("Select new Drafts folder")
            If LenB(newValue) > 0 Then
                SaveSettingStr "DraftsPath", newValue
                MsgBox "Drafts folder updated to:" & vbCrLf & newValue, vbInformation, "BWS Settings"
            End If

        Case "3"  ' Template
            newValue = PickFile("Select new letterhead template", "*.dotm; *.dotx")
            If LenB(newValue) > 0 Then
                SaveSettingStr "TemplatePath", newValue
                MsgBox "Template updated to:" & vbCrLf & newValue, vbInformation, "BWS Settings"
            End If

        Case "4"  ' Auto-open folder
            choice = MsgBox("Auto-open folder when saving?" & vbCrLf & vbCrLf & _
                           "Currently: " & currentAutoOpen & vbCrLf & vbCrLf & _
                           "Click Yes to enable, No to disable.", _
                           vbYesNoCancel + vbQuestion, "BWS Settings")
            If choice = vbYes Then
                SaveSettingStr "AutoOpenFolder", "True"
                MsgBox "Auto-open folder enabled", vbInformation, "BWS Settings"
            ElseIf choice = vbNo Then
                SaveSettingStr "AutoOpenFolder", "False"
                MsgBox "Auto-open folder disabled", vbInformation, "BWS Settings"
            End If

        Case "5"  ' Bullet conversion
            choice = MsgBox("Enable automatic bullet conversion?" & vbCrLf & vbCrLf & _
                           "Currently: " & currentBullets & vbCrLf & vbCrLf & _
                           "Click Yes to enable, No to disable.", _
                           vbYesNoCancel + vbQuestion, "BWS Settings")
            If choice = vbYes Then
                SaveSettingStr BWS_REG_BULLET_CONVERT, "Yes"
                MsgBox "Bullet conversion enabled", vbInformation, "BWS Settings"
            ElseIf choice = vbNo Then
                SaveSettingStr BWS_REG_BULLET_CONVERT, "No"
                MsgBox "Bullet conversion disabled", vbInformation, "BWS Settings"
            End If

        Case Else
            MsgBox "Invalid choice. Please enter 1-5.", vbExclamation, "BWS Settings"
    End Select
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

    ' NEW v1.6: Story 3 - Set page margins
    SetPageMargins newDoc

    StatusMessage Ver() & " - New letter created"
    Exit Sub

ErrHandler:
    MsgBox "Failed to create new letter: " & Err.Description, vbCritical, "BWS"
End Sub

' ============================================================================
' NEW v1.6: Story 3 - PAGE MARGINS
' ============================================================================

Private Sub SetPageMargins(ByVal doc As Document)
    ' Story 3: Set precise page margins including negative top margin
    On Error Resume Next
    With doc.PageSetup
        .TopMargin = Application.InchesToPoints(MARGIN_TOP)       ' 0.5"
        .BottomMargin = Application.InchesToPoints(MARGIN_BOTTOM) ' 0.625"
        .LeftMargin = Application.InchesToPoints(MARGIN_LEFT)     ' 0.375"
        .RightMargin = Application.InchesToPoints(MARGIN_RIGHT)   ' 0.5"
        .DifferentFirstPageHeaderFooter = True
    End With
    On Error GoTo 0
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
    Dim templ As String
    Dim newDoc As Document

    ' NEW v1.7: Create new document from template first (one-click workflow)
    templ = GetTemplatePath(True)
    If LenB(templ) = 0 Then
        MsgBox "No template configured.", vbExclamation, "BWS"
        Exit Sub
    End If

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

    ' NEW v1.7: Create new document from template
    StatusMessage "Creating new letter from template..."
    Set newDoc = Documents.Add(Template:=templ, NewTemplate:=False, DocumentType:=0)
    newDoc.Activate
    SetPageMargins newDoc

    ' Import into the newly created document
    ImportDocument newestFile
    Exit Sub

ErrHandler:
    MsgBox "Error finding newest draft: " & Err.Description, vbCritical, "BWS"
End Sub

Private Sub ImportDocument(ByVal sourcePath As String)
    Dim doc As Document
    Dim bodyCC As ContentControl
    Dim insertRng As Range
    Dim originalAlerts As Long

    Set doc = ActiveDocument
    If doc Is Nothing Then
        MsgBox "No active document.", vbExclamation, "BWS"
        Exit Sub
    End If

    ' Disable alerts to suppress template save prompts
    originalAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = wdAlertsNone

    On Error GoTo ErrorHandler

    StatusMessage "Importing " & sourcePath & "..."

    ' NEW v1.6: Story 3 - Set page margins first
    SetPageMargins doc

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

    ' Apply post-import processing
    StatusMessage "Processing imported content..."

    ' 1. Remove duplicate controls
    FindBodyContentControl doc, True

    ' 2. Apply global formatting FIRST
    ApplyGlobalFormatting doc

    ' 3. Fix tables AFTER global formatting (so table alignment isn't overridden)
    FixAllTables doc

    ' NEW v1.6: Story 5 - Format metadata lines
    FormatMetadataLines doc

    ' NEW v1.6: Story 1 - Format signature block (this may create empty paragraphs)
    FormatSignatureBlock doc

    ' 4. Convert bullets (and only if enabled)
    Dim bulletsEnabled As String
    bulletsEnabled = GetSettingStr(BWS_REG_BULLET_CONVERT, "Yes")
    If bulletsEnabled = "Yes" Then
        ConvertTextBulletsToRealBullets doc
    End If

    ' 5. Strip whitespace LAST - after all formatting that might create empty paragraphs
    StripAllWhitespace doc

    ' Mark template as unmodified to prevent "save changes to template?" prompt
    On Error Resume Next
    If Not doc.AttachedTemplate Is Nothing Then
        doc.AttachedTemplate.Saved = True
    End If
    On Error GoTo 0

    ' Restore alerts
    Application.DisplayAlerts = originalAlerts

    StatusMessage Ver() & " - Import complete"
    MsgBox "Import complete!", vbInformation, "BWS"
    Exit Sub

ErrorHandler:
    ' Restore alerts even on error
    Application.DisplayAlerts = originalAlerts
    MsgBox "Import error: " & Err.Description, vbCritical, "BWS"
End Sub

' ============================================================================
' NEW v1.6: Story 1 - SIGNATURE BLOCK FORMATTING
' ============================================================================

Private Sub FormatSignatureBlock(ByVal doc As Document)
    ' Two-step process:
    ' 1. REMOVE any "Sincerely" signatures from body content (imported doc)
    ' 2. Fix image formatting in template signature wrapper only

    Dim cc As ContentControl
    Dim shp As InlineShape
    Dim fltShp As Shape
    Dim sigCC As ContentControl
    Dim bodyCC As ContentControl
    Dim para As Paragraph
    Dim txt As String
    Dim parasToDelete As Collection
    Dim i As Long
    Dim startPara As Paragraph
    Dim deleteRange As Range
    Dim isInSigWrapper As Boolean
    Dim tooCloseToSig As Boolean

    On Error Resume Next

    ' STEP 0: Find signature wrapper FIRST to know where NOT to delete
    Set sigCC = Nothing
    For Each cc In doc.ContentControls
        If InStr(1, cc.Title, "signature", vbTextCompare) > 0 Or _
           InStr(1, cc.Tag, "signature", vbTextCompare) > 0 Or _
           InStr(1, cc.Title, "SignatureBlock", vbTextCompare) > 0 Or _
           InStr(1, cc.Tag, "SignatureBlock", vbTextCompare) > 0 Then
            Set sigCC = cc
            Exit For
        End If
    Next cc

    ' STEP 1: Remove imported signatures from body content
    Set bodyCC = FindBodyContentControl(doc, False)
    If Not bodyCC Is Nothing Then
        Set parasToDelete = New Collection

        ' Find all "Sincerely" paragraphs in body content
        For Each para In doc.Paragraphs
            If para.Range.Start >= bodyCC.Range.Start And _
               para.Range.End <= bodyCC.Range.End Then
                txt = Trim$(para.Range.Text)
                If LenB(txt) > 0 Then
                    ' Check for signature keywords
                    If InStr(1, txt, "Sincerely", vbTextCompare) > 0 Or _
                       InStr(1, txt, "Best regards", vbTextCompare) > 0 Or _
                       InStr(1, txt, "Regards", vbTextCompare) > 0 Then
                        ' Only add to delete list if NOT too close to signature wrapper
                        ' This prevents deletion from accidentally removing template signature
                        tooCloseToSig = False

                        If Not sigCC Is Nothing Then
                            ' If paragraph is within 10 positions of signature wrapper, skip it
                            If para.Range.Start >= (sigCC.Range.Start - 10) Then
                                tooCloseToSig = True
                            End If
                        End If

                        If Not tooCloseToSig Then
                            parasToDelete.Add para
                        End If
                    End If
                End If
            End If
        Next para

        ' Delete signature paragraphs from body content
        ' Work backwards to avoid index issues when deleting
        For i = parasToDelete.Count To 1 Step -1
            On Error Resume Next
            Set startPara = parasToDelete(i)

            ' Create range starting at this paragraph
            Set deleteRange = startPara.Range

            ' Extend range to cover next 5 paragraphs (6 total)
            deleteRange.MoveEnd Unit:=wdParagraph, Count:=5

            ' CRITICAL: Ensure we don't delete into signature wrapper area
            If Not sigCC Is Nothing Then
                If deleteRange.End > sigCC.Range.Start Then
                    ' Trim back to just before signature wrapper
                    deleteRange.End = sigCC.Range.Start - 1
                End If
            End If

            ' Also ensure we don't delete beyond body content control
            If deleteRange.End > bodyCC.Range.End Then
                deleteRange.End = bodyCC.Range.End
            End If

            ' Delete the range (if there's anything left to delete)
            If deleteRange.End > deleteRange.Start Then
                deleteRange.Delete
            End If
            On Error GoTo 0
        Next i
    End If

    ' STEP 2: Fix image formatting in template signature wrapper
    ' (sigCC is already set from STEP 0)
    If Not sigCC Is Nothing Then
        ' Fix all images in signature wrapper to be "in front of text"
        For Each shp In sigCC.Range.InlineShapes
            If shp.Type = wdInlineShapePicture Or shp.Type = wdInlineShapeLinkedPicture Then
                Set fltShp = shp.ConvertToShape
                fltShp.WrapFormat.Type = wdWrapFront
                fltShp.ZOrder msoBringToFront  ' In front of text
            End If
        Next shp
    End If

    On Error GoTo 0
End Sub

' ============================================================================
' NEW v1.6: Story 5 - METADATA LINE FORMATTING
' ============================================================================

Private Sub FormatMetadataLines(ByVal doc As Document)
    ' Story 5: Format metadata lines like "Prepared by:", "Client:", "Project:", "Date:"
    ' Apply: 0" indent, 0pt spacing, Single line spacing

    Dim para As Paragraph
    Dim txt As String

    On Error Resume Next

    For Each para In doc.Paragraphs
        txt = para.Range.Text
        If LenB(txt) > 0 Then
            ' Check if line starts with metadata keywords (case-insensitive)
            If InStr(1, txt, "Prepared by:", vbTextCompare) > 0 Or _
               InStr(1, txt, "Client:", vbTextCompare) > 0 Or _
               InStr(1, txt, "Project:", vbTextCompare) > 0 Or _
               InStr(1, txt, "Date:", vbTextCompare) > 0 Or _
               InStr(1, txt, "Subject:", vbTextCompare) > 0 Or _
               InStr(1, txt, "To:", vbTextCompare) > 0 Or _
               InStr(1, txt, "From:", vbTextCompare) > 0 Then

                With para.Range.ParagraphFormat
                    .LeftIndent = 0                         ' 0" indent
                    .FirstLineIndent = 0
                    .SpaceAfter = 0                         ' 0pt spacing
                    .SpaceBefore = 0
                    .LineSpacingRule = wdLineSpaceSingle   ' Single spacing
                End With
            End If
        End If
    Next para

    On Error GoTo 0
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

            ' Override Word's default list formatting to remove hanging indent
            ' Bullet at left margin, text aligns vertically at BULLET_LEFT_IN
            With para.Range.ListFormat.ListTemplate.ListLevels(1)
                .NumberPosition = 0  ' Bullet at left margin
                .TextPosition = InchesToPoints(BULLET_LEFT_IN)  ' Text starts here
                .TabPosition = InchesToPoints(BULLET_LEFT_IN)
                .Alignment = wdListLevelAlignLeft
            End With

            With para.Range.ParagraphFormat
                .LeftIndent = InchesToPoints(BULLET_LEFT_IN)  ' All lines at this position
                .FirstLineIndent = 0  ' No hanging indent - text aligns vertically
            End With
            ' NEW v1.6: Story 2 - Force Calibri font for bullets
            para.Range.Font.Name = BULLET_FONT
            para.Range.ParagraphFormat.TabStops.ClearAll
            para.Range.ParagraphFormat.TabStops.Add Position:=InchesToPoints(BULLET_TAB_IN)
        End If
    Next para

    MsgBox "Bullet formatting fixed!", vbInformation, "BWS"
End Sub

Private Sub ConvertTextBulletsToRealBullets(ByVal doc As Document)
    Dim para As Paragraph
    Dim txt As String
    Dim firstChar As String

    ' STEP 1: Convert text bullets (•, -, etc.) to real bullets
    For Each para In doc.Paragraphs
        If para.Range.ListFormat.ListType = wdListNoNumbering Then
            txt = Trim$(para.Range.Text)
            If Len(txt) > 1 Then
                firstChar = Left$(txt, 1)
                If InStr("•·-◦▪■►", firstChar) > 0 Then
                    para.Range.Text = Mid$(txt, 2)
                    para.Range.ListFormat.ApplyBulletDefault
                End If
            End If
        End If
    Next para

    ' STEP 2: Fix formatting for ALL bullets (both newly converted and existing)
    For Each para In doc.Paragraphs
        ' Check if this paragraph is a bullet (any type of list)
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            On Error Resume Next

            ' Override Word's default list formatting to remove hanging indent
            ' Bullet at left margin, text aligns vertically at BULLET_LEFT_IN
            With para.Range.ListFormat.ListTemplate.ListLevels(1)
                .NumberPosition = 0  ' Bullet at left margin
                .TextPosition = InchesToPoints(BULLET_LEFT_IN)  ' Text starts here
                .TabPosition = InchesToPoints(BULLET_LEFT_IN)
                .Alignment = wdListLevelAlignLeft
            End With

            With para.Range.ParagraphFormat
                .LeftIndent = InchesToPoints(BULLET_LEFT_IN)  ' All lines at this position
                .FirstLineIndent = 0  ' No hanging indent - text aligns vertically
            End With

            ' NEW v1.6: Story 2 - Force Calibri font for bullets
            para.Range.Font.Name = BULLET_FONT

            On Error GoTo 0
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

Private Sub StripAllWhitespace(ByVal doc As Document)
    ' Aggressively strip ALL empty paragraphs from entire document
    ' Simple approach: keep looping and deleting until no more empty paragraphs
    ' NEW v1.7.13: Don't delete paragraphs that contain images (InlineShapes)
    Dim para As Paragraph
    Dim found As Boolean
    Dim attempts As Long
    Dim txt As String
    Dim hasImages As Boolean

    On Error Resume Next

    ' Keep removing empty paragraphs until none left (up to 100 attempts)
    attempts = 0
    Do While attempts < 100
        found = False
        For Each para In doc.Paragraphs
            txt = para.Range.Text
            ' Empty paragraph = just paragraph mark (Chr(13)) or whitespace + paragraph mark
            ' So check if trimmed length <= 1 (paragraph mark only)
            If Len(Trim$(txt)) <= 1 Then
                ' NEW v1.7.13: Check if paragraph contains images before deleting
                hasImages = False
                If para.Range.InlineShapes.Count > 0 Then
                    hasImages = True
                End If

                ' Only delete if paragraph doesn't contain images
                If Not hasImages Then
                    para.Range.Delete
                    found = True
                    Exit For  ' Start over after deletion
                End If
            End If
        Next para
        If Not found Then Exit Do  ' No more empty paragraphs
        attempts = attempts + 1
    Loop

    On Error GoTo 0
End Sub

Private Sub StripBodyContentWhitespace(ByVal doc As Document)
    ' Legacy function - now calls StripAllWhitespace
    StripAllWhitespace doc
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
    t.Rows.RightIndent = 0  ' NEW v1.7.13: Also zero right indent
    t.Rows.Alignment = wdAlignRowLeft
    t.PreferredWidthType = wdPreferredWidthPercent
    t.PreferredWidth = 100
    t.AutoFitBehavior wdAutoFitWindow

    ' Cell padding
    t.TopPadding = Application.CentimetersToPoints(0.1)
    t.BottomPadding = Application.CentimetersToPoints(0.1)
    t.LeftPadding = Application.CentimetersToPoints(0.15)
    t.RightPadding = Application.CentimetersToPoints(0.15)

    ' NEW v1.6: Story 4 - Header formatting with #D9D9D9 gray and bold text
    If t.Rows.Count >= 1 Then
        With t.Rows(1)
            .HeadingFormat = True
            .Shading.BackgroundPatternColor = RGB(HDR_R, HDR_G, HDR_B)  ' #D9D9D9
            .Shading.ForegroundPatternColor = RGB(HDR_R, HDR_G, HDR_B)
        End With
        ' Make header text bold
        t.Rows(1).Range.Font.Bold = True
    End If

    ' Zebra striping
    Dim i As Long
    For i = 2 To t.Rows.Count
        If i Mod 2 = 0 Then
            t.Rows(i).Shading.BackgroundPatternColor = RGB(ZEB_R, ZEB_G, ZEB_B)
            t.Rows(i).Shading.ForegroundPatternColor = RGB(ZEB_R, ZEB_G, ZEB_B)
        End If
    Next i

    ' Column alignment - apply cell-by-cell to avoid table-wide overrides
    Dim c As Long
    Dim r As Long
    Dim isCurrency As Boolean

    For c = 1 To t.Columns.Count
        isCurrency = IsCurrencyColumn(t, c)

        ' Apply alignment to each cell individually
        For r = 1 To t.Rows.Count
            On Error Resume Next
            With t.Cell(r, c).Range.ParagraphFormat
                .LeftIndent = 0
                .RightIndent = 0  ' NEW v1.7.13: Zero right indent
                .FirstLineIndent = 0
                .SpaceBefore = 0  ' NEW v1.7.13: Zero spacing before
                .SpaceAfter = 0   ' NEW v1.7.13: Zero spacing after
                If isCurrency Then
                    .Alignment = wdAlignParagraphRight
                Else
                    .Alignment = wdAlignParagraphLeft
                End If
            End With
            On Error GoTo 0
        Next r
    Next c

    On Error GoTo 0
End Sub

Private Function IsCurrencyColumn(ByVal t As Table, ByVal colIndex As Long) As Boolean
    ' Only returns True if column contains currency symbols ($, €, £, etc.)
    Dim r As Long
    Dim cellText As String
    Dim currencyCount As Long
    Dim nonEmptyCount As Long

    currencyCount = 0
    nonEmptyCount = 0

    For r = 2 To t.Rows.Count  ' Skip header
        On Error Resume Next
        cellText = Trim$(t.Cell(r, colIndex).Range.Text)
        cellText = Replace$(cellText, Chr$(13), "")
        cellText = Replace$(cellText, Chr$(7), "")

        If Len(cellText) > 0 Then
            nonEmptyCount = nonEmptyCount + 1
            ' Check if cell contains currency symbols
            If InStr(cellText, "$") > 0 Or _
               InStr(cellText, "€") > 0 Or _
               InStr(cellText, "£") > 0 Or _
               InStr(cellText, "¥") > 0 Then
                currencyCount = currencyCount + 1
            End If
        End If
        On Error GoTo 0
    Next r

    ' Column is currency if >= 50% of non-empty cells have currency symbols
    If nonEmptyCount > 0 Then
        IsCurrencyColumn = (currencyCount / nonEmptyCount) >= 0.5
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
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.FirstLineIndent = 0  ' Clear any hanging indent
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
        .ParagraphFormat.LineSpacing = 13.8  ' 276 twips = 1.15x spacing (matches reference)
    End With

    ' NEW v1.6: Story 6 - Apply header styles
    ApplyHeaderStyles doc

    On Error GoTo 0
End Sub

' ============================================================================
' NEW v1.6: Story 6 - HEADER STYLE FORMATTING
' ============================================================================

Private Sub ApplyHeaderStyles(ByVal doc As Document)
    ' Story 6: Apply BWS Header style with single line spacing
    ' Detect likely headers (larger font, standalone lines, etc.)
    ' Also apply "BWS Header" style to Executive Summary section headings

    Dim para As Paragraph
    Dim txt As String
    Dim inExecSummary As Boolean
    Dim bwsHeaderStyleExists As Boolean

    On Error Resume Next

    ' Check if BWS Header style exists in template
    bwsHeaderStyleExists = False
    Dim testStyle As Style
    Set testStyle = doc.Styles("BWS Header")
    If Not testStyle Is Nothing Then
        bwsHeaderStyleExists = True
    End If
    On Error GoTo 0

    inExecSummary = False

    For Each para In doc.Paragraphs
        txt = Trim$(para.Range.Text)

        On Error Resume Next

        ' Track if we're in Executive Summary or Experience Strategy Rationale sections
        If InStr(1, txt, "Executive Summary", vbTextCompare) > 0 Or _
           InStr(1, txt, "Experience Strategy", vbTextCompare) > 0 Or _
           InStr(1, txt, "Rationale", vbTextCompare) > 0 Then
            inExecSummary = True
        ElseIf InStr(1, txt, "Proposal", vbTextCompare) > 0 Or _
               InStr(1, txt, "Timeline", vbTextCompare) > 0 Or _
               InStr(1, txt, "Budget", vbTextCompare) > 0 Then
            inExecSummary = False
        End If

        ' Apply BWS Header style to headings in Executive Summary sections
        If inExecSummary And bwsHeaderStyleExists Then
            ' Detect headings: short lines, ending with ":", or all caps
            If Len(txt) > 0 And Len(txt) < 80 And _
               (Right$(txt, 1) = ":" Or txt = UCase$(txt) Or para.Range.Font.Bold = True) Then
                para.Style = doc.Styles("BWS Header")
            End If
        End If

        ' Apply general header formatting for large/caps/bold text
        If para.Range.Font.Size >= 14 Or _
           (Len(txt) > 0 And Len(txt) < 100 And txt = UCase$(txt) And para.Range.Font.Bold = True) Then

            With para.Range
                .Font.Name = HEADER_FONT
                .Font.Size = HEADER_SIZE
                .Font.Color = RGB(HEADER_COLOR_R, HEADER_COLOR_G, HEADER_COLOR_B)
                .Font.Bold = True
            End With

            With para.Range.ParagraphFormat
                .LineSpacingRule = wdLineSpaceSingle  ' Single line spacing
                .SpaceAfter = 0                       ' 0pt after
                .SpaceBefore = 12                     ' 12pt before for separation
            End With
        End If

        On Error GoTo 0
    Next para
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
    Dim after As String, eol As Long
    Set rng = doc.Content
    txt = rng.Text
    pos = InStr(1, txt, keyLabel, vbTextCompare)
    If pos > 0 Then
        after = Mid$(txt, pos + Len(keyLabel))
        eol = InStr(after, vbCr)
        If eol = 0 Then eol = Len(after) + 1
        ExtractValue = Trim$(Left$(after, eol - 1))
    Else
        ExtractValue = ""
    End If
End Function

' ============================================================================
' DIAGNOSTIC
' ============================================================================

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

Private Function BWS_ToolbarsExist() As Boolean
    ' Check if all 3 BWS toolbars exist and are valid
    Dim exists As Boolean
    On Error Resume Next
    exists = False

    ' Check if all 3 toolbars exist
    If Not Application.CommandBars("BWS-1") Is Nothing Then
        If Not Application.CommandBars("BWS-2") Is Nothing Then
            If Not Application.CommandBars("BWS-3") Is Nothing Then
                exists = True
            End If
        End If
    End If

    On Error GoTo 0
    BWS_ToolbarsExist = exists
End Function

Private Sub BuildOrRefreshBWSToolbar(ByVal showMessage As Boolean)
    Dim cb1 As CommandBar
    Dim cb2 As CommandBar
    Dim cb3 As CommandBar

    ' NEW v1.7.12: Check if toolbars already exist
    ' Only recreate if they don't exist OR if explicitly requested (showMessage=True)
    ' This prevents duplicate toolbars when macros are in Normal.dotm
    If BWS_ToolbarsExist() And Not showMessage Then
        ' Toolbars already exist and this is an AutoOpen call, skip recreation
        Exit Sub
    End If

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
        MsgBox "BWS v1.7.13 toolbar installed!" & vbCrLf & vbCrLf & _
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
    Dim msg2 As String
    Dim msg3 As String

    ' Build message in parts to avoid VBA's 25-line-continuation limit
    ' Part 1: Header and recent versions (v1.7.13, v1.7.12, v1.7.11)
    msg = "================ BWS v1.7.13 =================" & vbCrLf _
        & "Version: " & BWS_VERSION & vbCrLf _
        & "Host:    " & Application.Name & " " & Application.Version & vbCrLf _
        & vbCrLf _
        & "NEW in v1.7.13:" & vbCrLf _
        & "• Fixed table indent: RightIndent = 0" & vbCrLf _
        & "• Fixed cell spacing: Before/After = 0" & vbCrLf _
        & "• Fixed signature image preservation" & vbCrLf _
        & vbCrLf _
        & "From v1.7.12:" & vbCrLf _
        & "• Fixed toolbar duplication in Normal.dotm" & vbCrLf _
        & "• AutoOpen now checks if toolbars exist" & vbCrLf _
        & vbCrLf _
        & "From v1.7.11:" & vbCrLf _
        & "• Fixed Dim in ExtractValue (inside If block)" & vbCrLf

    ' Part 2: Older versions (v1.6.6, v1.6.5)
    msg2 = vbCrLf _
        & "From v1.6.6:" & vbCrLf _
        & "• Signature: Fixed content preservation" & vbCrLf _
        & "• Signature spacing: 2 blank lines" & vbCrLf _
        & "• Image behind text (not obscuring)" & vbCrLf _
        & "• Paragraphs: Clear hanging indent" & vbCrLf _
        & "• Tables: Cell-by-cell alignment" & vbCrLf _
        & vbCrLf _
        & "From v1.6.5:" & vbCrLf _
        & "• Bullet indent: ListTemplate override" & vbCrLf _
        & "• Table order: Fixed alignment override" & vbCrLf _
        & "• Config memory in installer" & vbCrLf

    ' Part 3: Core features and footer
    msg3 = vbCrLf _
        & "Core Features:" & vbCrLf _
        & "• Line spacing: 276 twips (v1.6.1 fix)" & vbCrLf _
        & "• Signature block: 0.125"" indent" & vbCrLf _
        & "• Calibri font for bullets" & vbCrLf _
        & "• Auto page margins (-0.062"" top)" & vbCrLf _
        & "• Table headers: #D9D9D9 gray + bold" & vbCrLf _
        & "• Metadata formatting (0"" indent)" & vbCrLf _
        & "• 3-row persistent toolbar" & vbCrLf _
        & "• Smart table formatting" & vbCrLf _
        & "=============================================="

    MsgBox msg & msg2 & msg3, vbInformation, "BWS"
End Sub

' ============================================================================
' Helper: Convert inches to points
' ============================================================================

Private Function InchesToPoints(ByVal inches As Double) As Single
    InchesToPoints = inches * 72
End Function

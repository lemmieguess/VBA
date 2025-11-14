# BWS UserForm Settings Dialog - Complete Setup Guide

## What You Need to Know

**This is a ONE-TIME manual setup** that takes about 10-15 minutes. After setup, the professional dialog will be permanent in your Normal.dotm.

**Important:** You CANNOT auto-generate a UserForm in VBA. You must:
1. Manually create the form
2. Manually drag controls onto it
3. Manually set properties
4. Paste the code I provide

**If you skip this:** The Settings button will still work - it'll just use the InputBox interface you saw in the screenshot. No errors, just less pretty.

---

## Part 1: Create the UserForm (5 minutes)

### Step 1.1: Open VBA Editor
1. Open Microsoft Word
2. Press **Alt+F11** (this opens the VBA Editor)
3. You should see a window titled "Microsoft Visual Basic for Applications"

### Step 1.2: Create New UserForm
1. In the **Project Explorer** pane (usually left side), find **"Normal"** or **"Normal (Normal.dotm)"**
2. Right-click on **"Normal"** (or any folder under it)
3. Select: **Insert > UserForm**
4. A blank gray form appears in the center (called "UserForm1")

### Step 1.3: Set Form Properties
1. Click once on the blank form to select it
2. Press **F4** to open the **Properties Window** (usually bottom-left)
3. In the Properties Window, find these properties and change them:

| Property | Change To |
|----------|-----------|
| **(Name)** | `BWS_SettingsForm` |
| **Caption** | `BWS Settings` |
| **Width** | `480` |
| **Height** | `360` |
| **StartUpPosition** | `2 - CenterScreen` |

**CRITICAL:** The **(Name)** must be exactly `BWS_SettingsForm` or it won't work!

---

## Part 2: Add Controls to the Form (10 minutes)

### Step 2.1: Open the Toolbox
- Go to **View > Toolbox** (or it may already be open)
- You should see a floating window with icons (Label, TextBox, CommandButton, etc.)

### Step 2.2: Add Labels (3 labels)

**For each label below:**
1. Click the **"A"** icon (Label) in the Toolbox
2. Drag a rectangle on the form where you want it
3. Press **F4** to open Properties
4. Set the properties shown below

**Label 1 - Base Folder:**
- **(Name)**: `lblBasePath`
- **Caption**: `Base Folder:`
- **Left**: `12`
- **Top**: `12`
- **Width**: `90`
- **Height**: `18`

**Label 2 - Drafts Folder:**
- **(Name)**: `lblDraftsPath`
- **Caption**: `Drafts Folder:`
- **Left**: `12`
- **Top**: `42`
- **Width**: `90`
- **Height**: `18`

**Label 3 - Template File:**
- **(Name)**: `lblTemplatePath`
- **Caption**: `Template File:`
- **Left**: `12`
- **Top**: `72`
- **Width**: `90`
- **Height**: `18`

### Step 2.3: Add TextBoxes (3 textboxes)

**For each textbox:**
1. Click the **"ab|"** icon (TextBox) in the Toolbox
2. Drag a rectangle on the form
3. Press **F4** and set properties

**TextBox 1 - Base Path:**
- **(Name)**: `txtBasePath`
- **Left**: `108`
- **Top**: `12`
- **Width**: `270`
- **Height**: `20`

**TextBox 2 - Drafts Path:**
- **(Name)**: `txtDraftsPath`
- **Left**: `108`
- **Top**: `42`
- **Width**: `270`
- **Height**: `20`

**TextBox 3 - Template Path:**
- **(Name)**: `txtTemplatePath`
- **Left**: `108`
- **Top**: `72`
- **Width**: `270`
- **Height**: `20`

### Step 2.4: Add Browse Buttons (3 buttons)

**For each button:**
1. Click the **button icon** (CommandButton) in the Toolbox
2. Drag a small rectangle on the form
3. Press **F4** and set properties

**Button 1 - Browse Base:**
- **(Name)**: `btnBrowseBase`
- **Caption**: `...`
- **Left**: `384`
- **Top**: `12`
- **Width**: `30`
- **Height**: `20`

**Button 2 - Browse Drafts:**
- **(Name)**: `btnBrowseDrafts`
- **Caption**: `...`
- **Left**: `384`
- **Top**: `42`
- **Width**: `30`
- **Height**: `20`

**Button 3 - Browse Template:**
- **(Name)**: `btnBrowseTemplate`
- **Caption**: `...`
- **Left**: `384`
- **Top**: `72`
- **Width**: `30`
- **Height**: `20`

### Step 2.5: Add Checkboxes (2 checkboxes)

**For each checkbox:**
1. Click the **checkbox icon** (CheckBox) in the Toolbox
2. Drag a rectangle on the form
3. Press **F4** and set properties

**CheckBox 1 - Auto-Open:**
- **(Name)**: `chkAutoOpen`
- **Caption**: `Auto-open folder when saving`
- **Left**: `12`
- **Top**: `108`
- **Width**: `300`
- **Height**: `20`

**CheckBox 2 - Bullet Conversion:**
- **(Name)**: `chkBulletConvert`
- **Caption**: `Enable automatic bullet conversion`
- **Left**: `12`
- **Top**: `138`
- **Width**: `300`
- **Height**: `20`

### Step 2.6: Add Save/Cancel Buttons (2 buttons)

**Button 4 - Save:**
- **(Name)**: `btnSave`
- **Caption**: `Save`
- **Left**: `240`
- **Top**: `280`
- **Width**: `80`
- **Height**: `30`
- **Default**: `True` (scroll down in properties to find this)

**Button 5 - Cancel:**
- **(Name)**: `btnCancel`
- **Caption**: `Cancel`
- **Left**: `330`
- **Top**: `280`
- **Width**: `80`
- **Height**: `30`
- **Cancel**: `True` (scroll down in properties to find this)

---

## Part 3: Add the Code (2 minutes)

### Step 3.1: Open Form Code Window
1. In **Project Explorer**, find your form: **"BWS_SettingsForm"**
2. **Double-click** on **"BWS_SettingsForm"** (NOT the form itself, but the name in Project Explorer)
3. A code window opens (should be mostly empty)

### Step 3.2: Paste This Code

**DELETE** any existing code in that window, then paste this entire block:

```vba
Option Explicit

' Load current settings when form initializes
Private Sub UserForm_Initialize()
    ' Load current settings from registry
    txtBasePath.Text = GetSettingStr("BasePath", "")
    txtDraftsPath.Text = GetSettingStr("DraftsPath", "")
    txtTemplatePath.Text = GetSettingStr("TemplatePath", "")
    chkAutoOpen.Value = (GetSettingStr("AutoOpenFolder", "False") = "True")
    chkBulletConvert.Value = (GetSettingStr(BWS_REG_BULLET_CONVERT, "Yes") = "Yes")
End Sub

' Browse for base folder
Private Sub btnBrowseBase_Click()
    Dim newPath As String
    newPath = PickFolder("Select your Dropbox base folder")
    If LenB(newPath) > 0 Then
        txtBasePath.Text = newPath
    End If
End Sub

' Browse for drafts folder
Private Sub btnBrowseDrafts_Click()
    Dim newPath As String
    newPath = PickFolder("Select your Drafts folder")
    If LenB(newPath) > 0 Then
        txtDraftsPath.Text = newPath
    End If
End Sub

' Browse for template file
Private Sub btnBrowseTemplate_Click()
    Dim newPath As String
    newPath = PickFile("Select your letterhead template", "*.dotm; *.dotx", "")
    If LenB(newPath) > 0 Then
        txtTemplatePath.Text = newPath
    End If
End Sub

' Save button - save all settings and close
Private Sub btnSave_Click()
    ' Save all settings to registry
    SaveSettingStr "BasePath", txtBasePath.Text
    SaveSettingStr "DraftsPath", txtDraftsPath.Text
    SaveSettingStr "TemplatePath", txtTemplatePath.Text
    SaveSettingStr "AutoOpenFolder", IIf(chkAutoOpen.Value, "True", "False")
    SaveSettingStr BWS_REG_BULLET_CONVERT, IIf(chkBulletConvert.Value, "Yes", "No")

    ' Show confirmation
    MsgBox "Settings saved successfully!", vbInformation, "BWS Settings"

    ' Close the form
    Unload Me
End Sub

' Cancel button - close without saving
Private Sub btnCancel_Click()
    Unload Me
End Sub
```

---

## Part 4: Save and Test (1 minute)

### Step 4.1: Save Everything
1. In VBA Editor, press **Ctrl+S** (or File > Save)
2. This saves the UserForm into Normal.dotm
3. **Close the VBA Editor** (or leave it open, doesn't matter)

### Step 4.2: Test It!
1. Back in Word, find the **BWS toolbar**
2. Click the **"Settings"** button
3. You should now see the **professional dialog** with all your controls!
4. Test the browse buttons (...) - they should open folder/file pickers
5. Test Save/Cancel

---

## Troubleshooting

### Problem: Settings button still shows InputBox
**Solution:** The form name must be exactly `BWS_SettingsForm`
- Check the **(Name)** property in Step 1.3
- It's case-sensitive and must match exactly

### Problem: "Object not found" error
**Solution:** The code must be in the **form's code module**, not BWS_Module
- Double-click **"BWS_SettingsForm"** in Project Explorer
- Paste code there, not in BWS_Module

### Problem: Browse buttons don't work
**Solution:** Make sure BWS_Module is loaded in Normal.dotm
- The PickFolder and PickFile functions must be available
- Run BWS_Install first to ensure everything is set up

### Problem: Want to delete and start over
**Solution:**
1. In Project Explorer, right-click **"BWS_SettingsForm"**
2. Select **Remove BWS_SettingsForm**
3. Click **No** when asked to export
4. Start over from Part 1

### Problem: Don't want to do this, prefer InputBox
**Solution:**
- Just skip this entire guide!
- The Settings button will use InputBox (still fully functional)
- No errors, no problems

---

## Visual Layout Reference

Your form should look something like this:

```
┌─────────────────────────────────────────────────────────┐
│ BWS Settings                                      [X]   │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  Base Folder:     [________________________]  [...]    │
│                                                         │
│  Drafts Folder:   [________________________]  [...]    │
│                                                         │
│  Template File:   [________________________]  [...]    │
│                                                         │
│  ☐ Auto-open folder when saving                        │
│                                                         │
│  ☐ Enable automatic bullet conversion                  │
│                                                         │
│                                                         │
│                                                         │
│                                                         │
│                                          [Save] [Cancel]│
└─────────────────────────────────────────────────────────┘
```

---

## Summary

**What you're creating:**
- 1 UserForm (the container)
- 3 Labels (text descriptions)
- 3 TextBoxes (editable paths)
- 5 Buttons (3 browse + Save + Cancel)
- 2 Checkboxes (True/False settings)

**Total controls:** 14 (but they're simple to add)

**Time investment:** 10-15 minutes

**Benefit:** Professional, modern settings dialog forever!

**Alternative:** Skip this entirely and keep using InputBox (still works fine)

---

## Need Help?

If you get stuck on any step, just ask! I can clarify any part of this process.

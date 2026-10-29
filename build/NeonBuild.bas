Attribute VB_Name = "NeonBuild"
' =====================================================================
'  NeonCenterline - plugin BUILDER (for the author; not part of the plugin).
'  Import into GlobalMacros (File > Import) and run the macro NeonBuildPlugin:
'    1) creates (or clears) the project NeonCenterline.gms in the user GMS folder;
'    2) adds modules from the 'src' folder (UTF-8 sources are read and added with
'       AddFromString, so Russian text is converted to the Windows code page);
'    3) creates the settings window frmNeonSettings and inserts its code;
'    4) inserts the code of ThisMacroStorage (CorelDRAW start: toolbar update after a plugin
'       update);
'    5) then SAVE the project in the VBA editor (select NeonCenterline, Ctrl+S).
'  Requires access to the VBA editor object model (Application.VBE).
'  This file is ASCII-only on purpose: it is imported with File > Import.
' =====================================================================
Option Explicit

Private Const PRJ As String = "NeonCenterline"

Public Sub NeonBuildPlugin()
    Dim src As String, vbe_ As Object, gp As Object, vp As Object, i As Long, mods As Variant, m As Variant
    Dim comp As Object, rep As String

    src = InputBox("Folder with plugin 'src' files (NeonUI.bas etc.):", _
                   "NeonBuild", GetSetting("NeonCenterline", "Build", "Src", Environ$("USERPROFILE") & "\Downloads\NeonCenterline\src"))
    If Len(src) = 0 Then
        Exit Sub
    End If
    If Right$(src, 1) = "\" Then
        src = Left$(src, Len(src) - 1)
    End If
    If Len(Dir$(src & "\NeonUI.bas")) = 0 Then
        MsgBox "NeonUI.bas not found in:" & vbCr & src, vbExclamation, "NeonBuild"
        Exit Sub
    End If
    SaveSetting "NeonCenterline", "Build", "Src", src

    On Error Resume Next
    Set vbe_ = Application.VBE
    On Error GoTo 0
    If vbe_ Is Nothing Then
        MsgBox "No access to the VBA editor object model (Application.VBE)." & vbCr & _
               "Build manually: see BUILD.md.", vbExclamation, "NeonBuild"
        Exit Sub
    End If

    Set gp = FindGms()
    If gp Is Nothing Then
        Set gp = GMSManager.Projects.Create(PRJ)
    End If
    If gp Is Nothing Then
        MsgBox "Could not create " & PRJ & ".gms", vbExclamation, "NeonBuild"
        Exit Sub
    End If
    Set vp = FindVbProject(vbe_, gp)
    If vp Is Nothing Then
        MsgBox "VBA project for " & gp.FullFileName & " not found in the editor.", vbExclamation, "NeonBuild"
        Exit Sub
    End If
    For i = vp.VBComponents.Count To 1 Step -1
        Set comp = vp.VBComponents(i)
        ' forms named frmNeon* are kept and refilled (a removed form keeps its name until the
        ' project is saved, so re-adding it fails with error 75); everything else is removed
        If comp.Type = 1 Or comp.Type = 2 Or (comp.Type = 3 And LCase$(Left$(comp.Name, 7)) <> "frmneon") Then
            On Error Resume Next
            comp.Name = comp.Name & "_old" & i
            On Error GoTo 0
            vp.VBComponents.Remove comp
        End If
    Next i

    mods = Array("NeonLang.bas", "NeonUI.bas", "NeonUpdate.bas", "NeonAddon.bas", "NeonTools.bas", "NeonCut.bas", "NeonMono.bas", "NeonScript.bas")
    For Each m In mods
        If Len(Dir$(src & "\" & m)) = 0 Then
            MsgBox "Missing file: " & src & "\" & m, vbExclamation, "NeonBuild"
            Exit Sub
        End If
        ImportUtf8Module vp, src & "\" & m
        rep = rep & "  + " & m & vbCr
    Next m

    If Len(Dir$(src & "\frmNeonSettings.code.txt")) > 0 Then
        BuildSettingsForm vp, src & "\frmNeonSettings.code.txt"
        rep = rep & "  + frmNeonSettings" & vbCr
    End If
    If Len(Dir$(src & "\frmNeonRun.code.txt")) > 0 Then
        BuildRunForm vp, src & "\frmNeonRun.code.txt"
        rep = rep & "  + frmNeonRun" & vbCr
    End If
    If Len(Dir$(src & "\ThisMacroStorage.code.txt")) > 0 Then
        If SetStorageCode(vp, src & "\ThisMacroStorage.code.txt") Then
            rep = rep & "  + ThisMacroStorage (start: toolbar update)" & vbCr
        Else
            rep = rep & "  ! ThisMacroStorage not found: no toolbar update at start" & vbCr
        End If
    End If

    On Error Resume Next
    gp.Dirty = True
    On Error GoTo 0
    MsgBox "Built: " & gp.FullFileName & vbCr & rep & vbCr & _
           "Now SAVE the project: VBA editor (Alt+F11) -> select '" & PRJ & "' -> Ctrl+S.", vbInformation, "NeonBuild"
End Sub

Private Function FindGms() As Object
    Dim i As Long, p As Object
    For i = 1 To GMSManager.Projects.Count
        Set p = GMSManager.Projects(i)
        If LCase$(p.FileName) = LCase$(PRJ & ".gms") Then
            Set FindGms = p
            Exit Function
        End If
    Next i
End Function

Private Function FindVbProject(vbe_ As Object, gp As Object) As Object
    Dim vp As Object, fn As String
    For Each vp In vbe_.VBProjects
        fn = ""
        On Error Resume Next
        fn = vp.FileName
        On Error GoTo 0
        If LCase$(fn) = LCase$(gp.FullFileName) Then
            Set FindVbProject = vp
            Exit Function
        End If
    Next vp
    For Each vp In vbe_.VBProjects
        If LCase$(vp.Name) = LCase$(PRJ) Then
            Set FindVbProject = vp
            Exit Function
        End If
    Next vp
End Function

Private Sub BuildSettingsForm(vp As Object, ByVal codeFile As String)
    Dim fc As Object, d As Object, code As String
    Set fc = GetForm(vp, "frmNeonSettings")
    fc.Properties("Width") = 360
    fc.Properties("Height") = 450
    fc.Properties("Caption") = "NeonCenterline"
    Set d = fc.Designer
    AddCtl d, "Forms.Label.1", "lblTool", 12, 14, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboTool", 190, 10, 70, 20
    AddCtl d, "Forms.Label.1", "lblHint", 12, 36, 326, 26
    AddCtl d, "Forms.Label.1", "lblEnd", 12, 66, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboEnd", 190, 62, 148, 20
    AddCtl d, "Forms.CheckBox.1", "chkAsk", 12, 90, 326, 18
    AddCtl d, "Forms.Label.1", "lblType", 12, 120, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboType", 12, 138, 326, 20
    AddCtl d, "Forms.Label.1", "lblCut", 12, 168, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboCut", 190, 164, 70, 20
    AddCtl d, "Forms.CheckBox.1", "chkFit", 12, 190, 326, 18
    AddCtl d, "Forms.Label.1", "lblTrim", 30, 214, 100, 18
    AddCtl d, "Forms.ComboBox.1", "cboTrim", 130, 210, 208, 20
    AddCtl d, "Forms.Label.1", "lblRef", 12, 236, 326, 18
    AddCtl d, "Forms.ComboBox.1", "cboRef", 12, 254, 326, 20
    AddCtl d, "Forms.Label.1", "lblLang", 12, 288, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboLang", 190, 284, 148, 20
    AddCtl d, "Forms.CheckBox.1", "chkUpd", 12, 312, 326, 18
    AddCtl d, "Forms.Label.1", "lblVer", 12, 338, 326, 18
    AddCtl d, "Forms.CommandButton.1", "btnUpdate", 12, 360, 160, 24
    AddCtl d, "Forms.CommandButton.1", "btnAbout", 178, 360, 160, 24
    AddCtl d, "Forms.CommandButton.1", "btnOK", 180, 394, 76, 24
    AddCtl d, "Forms.CommandButton.1", "btnCancel", 262, 394, 76, 24
    d.Controls("cboLang").Style = 2
    d.Controls("cboType").Style = 2
    d.Controls("cboTrim").Style = 2
    d.Controls("cboRef").Style = 2
    d.Controls("lblHint").WordWrap = True
    d.Controls("btnOK").Default = True
    d.Controls("btnCancel").Cancel = True
    code = ReadUtf8(codeFile)
    With fc.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        .AddFromString code
    End With
End Sub

' window shown before a run: cutter, neon type, cut step, fitting
Private Sub BuildRunForm(vp As Object, ByVal codeFile As String)
    Dim fc As Object, d As Object, code As String
    Set fc = GetForm(vp, "frmNeonRun")
    fc.Properties("Width") = 360
    fc.Properties("Height") = 360
    fc.Properties("Caption") = "NeonCenterline"
    Set d = fc.Designer
    AddCtl d, "Forms.Label.1", "lblTool", 12, 14, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboTool", 190, 10, 70, 20
    AddCtl d, "Forms.Label.1", "lblHint", 12, 36, 326, 18
    AddCtl d, "Forms.Label.1", "lblEnd", 12, 58, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboEnd", 190, 54, 148, 20
    AddCtl d, "Forms.Label.1", "lblType", 12, 86, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboType", 12, 104, 326, 20
    AddCtl d, "Forms.Label.1", "lblCut", 12, 134, 170, 18
    AddCtl d, "Forms.ComboBox.1", "cboCut", 190, 130, 70, 20
    AddCtl d, "Forms.CheckBox.1", "chkFit", 12, 156, 326, 18
    AddCtl d, "Forms.Label.1", "lblTrim", 30, 180, 100, 18
    AddCtl d, "Forms.ComboBox.1", "cboTrim", 130, 176, 208, 20
    AddCtl d, "Forms.Label.1", "lblRef", 12, 202, 326, 18
    AddCtl d, "Forms.ComboBox.1", "cboRef", 12, 220, 326, 20
    AddCtl d, "Forms.CheckBox.1", "chkAsk", 12, 254, 326, 18
    AddCtl d, "Forms.CommandButton.1", "btnOK", 160, 284, 96, 26
    AddCtl d, "Forms.CommandButton.1", "btnCancel", 262, 284, 76, 26
    d.Controls("cboType").Style = 2
    d.Controls("cboTrim").Style = 2
    d.Controls("cboRef").Style = 2
    d.Controls("btnOK").Default = True
    d.Controls("btnCancel").Cancel = True
    code = ReadUtf8(codeFile)
    With fc.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        .AddFromString code
    End With
End Sub

' code of the project's ThisMacroStorage (document module: GlobalMacroStorage events)
Private Function SetStorageCode(vp As Object, ByVal codeFile As String) As Boolean
    Dim i As Long, comp As Object, code As String
    For i = 1 To vp.VBComponents.Count
        If vp.VBComponents(i).Type = 100 Then
            Set comp = vp.VBComponents(i)
            Exit For
        End If
    Next i
    If comp Is Nothing Then
        Exit Function
    End If
    code = ReadUtf8(codeFile)
    With comp.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        .AddFromString code
    End With
    SetStorageCode = True
End Function

' existing form: clear its controls and code; otherwise add a new one
Private Function GetForm(vp As Object, ByVal nm As String) As Object
    Dim i As Long, fc As Object, d As Object
    For i = 1 To vp.VBComponents.Count
        If LCase$(vp.VBComponents(i).Name) = LCase$(nm) Then
            Set fc = vp.VBComponents(i)
            Exit For
        End If
    Next i
    If fc Is Nothing Then
        Set fc = vp.VBComponents.Add(3)
        fc.Name = nm
    Else
        Set d = fc.Designer
        For i = d.Controls.Count - 1 To 0 Step -1
            d.Controls.Remove d.Controls(i).Name
        Next i
        With fc.CodeModule
            If .CountOfLines > 0 Then
                .DeleteLines 1, .CountOfLines
            End If
        End With
    End If
    Set GetForm = fc
End Function

Private Sub AddCtl(d As Object, ByVal progId As String, ByVal nm As String, ByVal x As Single, ByVal y As Single, _
                   ByVal w As Single, ByVal h As Single)
    Dim c As Object
    Set c = d.Controls.Add(progId, nm, True)
    c.Left = x: c.Top = y: c.Width = w: c.Height = h
End Sub

Private Sub ImportUtf8Module(vp As Object, ByVal path As String)
    Dim txt As String, lines As Variant, i As Long, nm As String, code As String, comp As Object, p As Long
    txt = ReadUtf8(path)
    txt = Replace(txt, vbCrLf, vbLf)
    lines = Split(txt, vbLf)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 10) = "Attribute " Then
            If Left$(lines(i), 17) = "Attribute VB_Name" Then
                p = InStr(lines(i), """")
                nm = Mid$(lines(i), p + 1)
                nm = Left$(nm, InStr(nm, """") - 1)
            End If
        Else
            code = code & lines(i) & vbCrLf
        End If
    Next i
    If Len(nm) = 0 Then
        Err.Raise vbObjectError + 1, , "No VB_Name in " & path
    End If
    Set comp = vp.VBComponents.Add(1)
    comp.Name = nm
    With comp.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
        .AddFromString code
    End With
End Sub

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

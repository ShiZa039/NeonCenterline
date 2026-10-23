Attribute VB_Name = "NeonBuild"
' =====================================================================
'  NeonCenterline - plugin BUILDER (for the author; not part of the plugin).
'  Import into GlobalMacros (File > Import) and run the macro NeonBuildPlugin:
'    1) creates (or clears) the project NeonCenterline.gms in the user GMS folder;
'    2) adds modules from the 'src' folder (UTF-8 sources are read and added with
'       AddFromString, so Russian text is converted to the Windows code page);
'    3) creates the settings window frmNeonSettings and inserts its code;
'    4) then SAVE the project in the VBA editor (select NeonCenterline, Ctrl+S).
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
        If comp.Type = 1 Or comp.Type = 2 Or comp.Type = 3 Then
            On Error Resume Next
            comp.Name = comp.Name & "_old" & i
            On Error GoTo 0
            vp.VBComponents.Remove comp
        End If
    Next i

    mods = Array("NeonLang.bas", "NeonUI.bas", "NeonUpdate.bas", "NeonAddon.bas", "NeonMono.bas", "NeonScript.bas")
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
    Set fc = vp.VBComponents.Add(3)
    fc.Name = "frmNeonSettings"
    fc.Properties("Width") = 330
    fc.Properties("Height") = 262
    fc.Properties("Caption") = "NeonCenterline"
    Set d = fc.Designer
    AddCtl d, "Forms.Label.1", "lblTool", 12, 14, 150, 18
    AddCtl d, "Forms.ComboBox.1", "cboTool", 170, 10, 70, 20
    AddCtl d, "Forms.Label.1", "lblHint", 12, 36, 296, 26
    AddCtl d, "Forms.CheckBox.1", "chkAsk", 12, 64, 296, 18
    AddCtl d, "Forms.Label.1", "lblLang", 12, 94, 150, 18
    AddCtl d, "Forms.ComboBox.1", "cboLang", 170, 90, 138, 20
    d.Controls("cboLang").Style = 2
    AddCtl d, "Forms.CheckBox.1", "chkUpd", 12, 120, 296, 18
    AddCtl d, "Forms.Label.1", "lblVer", 12, 148, 296, 18
    AddCtl d, "Forms.CommandButton.1", "btnUpdate", 12, 172, 146, 24
    AddCtl d, "Forms.CommandButton.1", "btnAbout", 162, 172, 146, 24
    AddCtl d, "Forms.CommandButton.1", "btnOK", 150, 206, 76, 24
    AddCtl d, "Forms.CommandButton.1", "btnCancel", 232, 206, 76, 24
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

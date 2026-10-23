Attribute VB_Name = "NeonUpdate"
' =====================================================================
'  NeonCenterline — обновление с GitHub.
'
'  1) Читается update.json (адрес — NEON_UPDATE_URL в NeonUI):
'       { "version": "1.1.0",
'         "gms_url": "https://github.com/.../releases/download/v1.1.0/NeonCenterline.gms",
'         "sha256":  "64 hex-символа",
'         "notes_ru": "Что нового...", "notes_en": "What's new..." }
'  2) Если версия новее — спрашиваем, скачиваем .gms во временную папку, проверяем
'     контрольную сумму (certutil) и что это файл проекта VBA (сигнатура OLE).
'  3) Работающий проект нельзя заменить, пока он загружен в CorelDRAW, поэтому
'     запускается маленький скрипт (wscript, скрыто): ждёт закрытия CorelDRAW, сохраняет
'     старый файл как .bak и кладёт новый на его место. Новая версия — со следующего запуска.
' =====================================================================
Option Explicit

Private Const REG_APP As String = "NeonCenterline"

' раз в сутки, если включено (по умолчанию — да); молча, если обновлений нет или нет сети
Public Sub NeonDailyUpdateCheck()
    Dim last As String, today As String
    If GetSetting(REG_APP, "Update", "Daily", "1") <> "1" Then
        Exit Sub
    End If
    today = Format$(Date, "yyyy-mm-dd")
    last = GetSetting(REG_APP, "Update", "LastCheck", "")
    If last = today Then
        Exit Sub
    End If
    SaveSetting REG_APP, "Update", "LastCheck", today
    NeonUpdateCheck True
End Sub

' quiet = True: не сообщать «обновлений нет» и ошибки сети
Public Sub NeonUpdateCheck(ByVal quiet As Boolean)
    Dim js As String, ver As String, url As String, sha As String, notes As String, errText As String
    Dim pending As String, target As String, tmpF As String

    If InStr(NEON_UPDATE_URL, "/OWNER/") > 0 Or Len(NEON_UPDATE_URL) = 0 Then
        If Not quiet Then
            MsgBox NeonL("upd_nourl"), vbInformation, NeonL("app_title")
        End If
        Exit Sub
    End If

    pending = GetSetting(REG_APP, "Update", "Pending", "")
    If Len(pending) > 0 Then
        tmpF = GetSetting(REG_APP, "Update", "PendingFile", "")
        If VerCmp(pending, NEON_VERSION) > 0 And Len(tmpF) > 0 And FileExists(tmpF) Then
            ' скачано, но ещё не установлено (подменщик мог не дожить: выход из Windows,
            ' перезагрузка) — запустить его снова и не скачивать заново
            target = PluginGmsPath()
            If Len(target) > 0 Then
                StartSwapScript tmpF, target, errText
            End If
            If Not quiet Then
                MsgBox NeonLF("upd_pending", pending), vbInformation, NeonL("app_title")
            End If
            Exit Sub
        End If
        ' уже установлено или файла обновления нет — сбросить и проверить заново
        SaveSetting REG_APP, "Update", "Pending", ""
        SaveSetting REG_APP, "Update", "PendingFile", ""
    End If

    js = HttpGetText(NEON_UPDATE_URL & "?t=" & Format$(Now, "yyyymmddhhnnss"), errText)
    If Len(js) = 0 Then
        If Not quiet Then
            MsgBox NeonLF("upd_err", errText), vbExclamation, NeonL("app_title")
        End If
        Exit Sub
    End If
    ver = JsonStr(js, "version")
    url = JsonStr(js, "gms_url")
    sha = LCase$(JsonStr(js, "sha256"))
    If NeonLangCode() = "ru" Then
        notes = JsonStr(js, "notes_ru")
    Else
        notes = JsonStr(js, "notes_en")
    End If
    If Len(ver) = 0 Or Len(url) = 0 Then
        If Not quiet Then
            MsgBox NeonLF("upd_err", "update.json"), vbExclamation, NeonL("app_title")
        End If
        Exit Sub
    End If
    If VerCmp(ver, NEON_VERSION) <= 0 Then
        If Not quiet Then
            MsgBox NeonLF("upd_none", NEON_VERSION), vbInformation, NeonL("app_title")
        End If
        Exit Sub
    End If
    If MsgBox(NeonLF("upd_found", ver, NEON_VERSION, notes), vbYesNo + vbQuestion, NeonL("app_title")) <> vbYes Then
        Exit Sub
    End If

    target = PluginGmsPath()
    If Len(target) = 0 Then
        MsgBox NeonL("upd_nogms"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    tmpF = TempDir() & "\NeonCenterline_" & ver & ".gms"
    If Not HttpDownload(url, tmpF, errText) Then
        MsgBox NeonLF("upd_err", errText), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    If Not IsOleFile(tmpF) Then
        KillQuiet tmpF
        MsgBox NeonL("upd_bad"), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    If Len(sha) = 64 Then
        If FileSha256(tmpF) <> sha Then
            KillQuiet tmpF
            MsgBox NeonL("upd_bad"), vbExclamation, NeonL("app_title")
            Exit Sub
        End If
    End If
    If Not StartSwapScript(tmpF, target, errText) Then
        MsgBox NeonLF("upd_err", errText), vbExclamation, NeonL("app_title")
        Exit Sub
    End If
    SaveSetting REG_APP, "Update", "Pending", ver
    SaveSetting REG_APP, "Update", "PendingFile", tmpF
    MsgBox NeonLF("upd_ready", ver), vbInformation, NeonL("app_title")
End Sub

' ---------------------------------------------------------------------
'  Сеть
' ---------------------------------------------------------------------
Private Function HttpGetText(ByVal url As String, errText As String) As String
    Dim h As Object
    On Error GoTo Fail
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 3000, 3000, 5000, 5000
    h.Open "GET", url, False
    h.SetRequestHeader "User-Agent", "NeonCenterline/" & NEON_VERSION
    h.Send
    If h.Status <> 200 Then
        errText = "HTTP " & h.Status
        Exit Function
    End If
    HttpGetText = Utf8ToStr(h.ResponseBody)
    Exit Function
Fail:
    errText = Err.Description
End Function

Private Function HttpDownload(ByVal url As String, ByVal path As String, errText As String) As Boolean
    Dim h As Object, st As Object
    On Error GoTo Fail
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.SetTimeouts 5000, 5000, 30000, 60000
    h.Open "GET", url, False
    h.SetRequestHeader "User-Agent", "NeonCenterline/" & NEON_VERSION
    h.Send
    If h.Status <> 200 Then
        errText = "HTTP " & h.Status
        Exit Function
    End If
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write h.ResponseBody
    st.SaveToFile path, 2
    st.Close
    HttpDownload = True
    Exit Function
Fail:
    errText = Err.Description
End Function

Private Function Utf8ToStr(body As Variant) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write body
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    Utf8ToStr = st.ReadText
    st.Close
End Function

' ---------------------------------------------------------------------
'  JSON: строковое значение по ключу (простой разбор; экранирование \n \" \\ \uXXXX)
' ---------------------------------------------------------------------
Private Function JsonStr(ByVal js As String, ByVal key As String) As String
    Dim p As Long, i As Long, c As String, r As String
    p = InStr(js, """" & key & """")
    If p = 0 Then
        Exit Function
    End If
    p = InStr(p + Len(key) + 2, js, ":")
    If p = 0 Then
        Exit Function
    End If
    p = InStr(p, js, """")
    If p = 0 Then
        Exit Function
    End If
    i = p + 1
    Do While i <= Len(js)
        c = Mid$(js, i, 1)
        If c = "\" Then
            i = i + 1
            c = Mid$(js, i, 1)
            Select Case c
            Case "n"
                r = r & vbCrLf
            Case "t"
                r = r & vbTab
            Case "u"
                r = r & ChrW(CLng("&H" & Mid$(js, i + 1, 4)))
                i = i + 4
            Case Else
                r = r & c
            End Select
        ElseIf c = """" Then
            Exit Do
        Else
            r = r & c
        End If
        i = i + 1
    Loop
    JsonStr = r
End Function

' сравнение версий «1.2.10» (по числам); >0 — a новее b
Public Function VerCmp(ByVal a As String, ByVal b As String) As Long
    Dim pa As Variant, pb As Variant, i As Long, x As Long, y As Long, n As Long
    pa = Split(a, "."): pb = Split(b, ".")
    n = UBound(pa)
    If UBound(pb) > n Then
        n = UBound(pb)
    End If
    For i = 0 To n
        x = 0: y = 0
        If i <= UBound(pa) Then
            x = Val(pa(i))
        End If
        If i <= UBound(pb) Then
            y = Val(pb(i))
        End If
        If x <> y Then
            VerCmp = IIf(x > y, 1, -1)
            Exit Function
        End If
    Next i
    VerCmp = 0
End Function

' ---------------------------------------------------------------------
'  Файлы
' ---------------------------------------------------------------------
' полный путь к .gms этого плагина (проект, в котором лежит модуль NeonUpdate)
Private Function PluginGmsPath() As String
    Dim p As Object, i As Long
    On Error Resume Next
    For i = 1 To GMSManager.Projects.Count
        Set p = GMSManager.Projects(i)
        If LCase$(p.FileName) = "neoncenterline.gms" Then
            PluginGmsPath = p.FullFileName
            Exit Function
        End If
    Next i
End Function

Private Function IsOleFile(ByVal path As String) As Boolean
    Dim f As Integer, b(0 To 7) As Byte
    On Error GoTo Fail
    If FileLen(path) < 1024 Then
        Exit Function
    End If
    f = FreeFile
    Open path For Binary Access Read As #f
    Get #f, 1, b
    Close #f
    IsOleFile = (b(0) = &HD0 And b(1) = &HCF And b(2) = &H11 And b(3) = &HE0)
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

' SHA-256 файла через certutil (есть во всех Windows); пусто — не удалось
Private Function FileSha256(ByVal path As String) As String
    Dim sh As Object, outF As String, f As Integer, ln As String, s As String
    On Error GoTo Fail
    outF = TempDir() & "\NeonCenterline_hash.txt"
    KillQuiet outF
    Set sh = CreateObject("WScript.Shell")
    sh.Run "cmd /c certutil -hashfile """ & path & """ SHA256 > """ & outF & """", 0, True
    f = FreeFile
    Open outF For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = LCase$(Replace(Trim$(ln), " ", ""))
        If IsHex64(ln) Then
            s = ln
        End If
    Loop
    Close #f
    KillQuiet outF
    FileSha256 = s
    Exit Function
Fail:
    On Error Resume Next
    Close #f
End Function

Private Function IsHex64(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) <> 64 Then
        Exit Function
    End If
    For i = 1 To 64
        c = Mid$(s, i, 1)
        If InStr("0123456789abcdef", c) = 0 Then
            Exit Function
        End If
    Next i
    IsHex64 = True
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' временная папка: через WScript.Shell (Environ$ отдаёт ANSI и портит путь с русским именем
' пользователя на нерусской Windows)
Private Function TempDir() As String
    Dim t As String
    On Error Resume Next
    t = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%")
    If Len(t) = 0 Or t = "%TEMP%" Then
        t = Environ$("TEMP")
    End If
    TempDir = t
End Function

Private Sub KillQuiet(ByVal path As String)
    On Error Resume Next
    Kill path
End Sub

' скрипт-«подменщик»: ждёт закрытия CorelDRAW и кладёт новый .gms на место старого
' (файл скрипта — в UTF-16, чтобы пути с русскими буквами не портились)
Private Function StartSwapScript(ByVal newF As String, ByVal target As String, errText As String) As Boolean
    Dim vbs As String, code As String, st As Object
    On Error GoTo Fail
    vbs = TempDir() & "\NeonCenterline_update.vbs"
    code = "On Error Resume Next" & vbCrLf & _
           "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
           "Set wmi = GetObject(""winmgmts:\\.\root\cimv2"")" & vbCrLf & _
           "If wmi Is Nothing Then WScript.Quit" & vbCrLf & _
           "For i = 1 To 43200" & vbCrLf & _
           "  Set ps = wmi.ExecQuery(""SELECT ProcessId FROM Win32_Process WHERE Name='CorelDRW.exe'"")" & vbCrLf & _
           "  Err.Clear" & vbCrLf & _
           "  n = -1" & vbCrLf & _
           "  n = ps.Count" & vbCrLf & _
           "  If Err.Number <> 0 Then WScript.Quit" & vbCrLf & _
           "  If n = 0 Then Exit For" & vbCrLf & _
           "  WScript.Sleep 2000" & vbCrLf & _
           "Next" & vbCrLf & _
           "WScript.Sleep 1500" & vbCrLf & _
           "src = """ & newF & """" & vbCrLf & _
           "dst = """ & target & """" & vbCrLf & _
           "If fso.FileExists(src) Then" & vbCrLf & _
           "  If fso.FileExists(dst) Then fso.CopyFile dst, dst & "".bak"", True" & vbCrLf & _
           "  Err.Clear" & vbCrLf & _
           "  fso.CopyFile src, dst, True" & vbCrLf & _
           "  If Err.Number = 0 Then fso.DeleteFile src, True" & vbCrLf & _
           "End If" & vbCrLf & _
           "fso.DeleteFile WScript.ScriptFullName, True" & vbCrLf
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "unicode"
    st.Open
    st.WriteText code
    st.SaveToFile vbs, 2
    st.Close
    CreateObject("WScript.Shell").Run "wscript.exe //nologo """ & vbs & """", 0, False
    StartSwapScript = True
    Exit Function
Fail:
    errText = Err.Description
End Function

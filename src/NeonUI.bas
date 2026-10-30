Attribute VB_Name = "NeonUI"
' =====================================================================
'  NeonCenterline — интерфейс плагина: кнопки панели, настройки, «О программе»,
'  проверка обновлений. Расчёт — в модулях NeonMono (моноширинные шрифты) и
'  NeonScript (курсив); здесь только вызов и сообщения.
'
'  Команды для кнопок панели (AppUI.xslt):
'    NeonCenterline.NeonUI.RunMono      — центральная линия, «Моно»
'    NeonCenterline.NeonUI.RunScript    — центральная линия, «Курсив»
'    NeonCenterline.NeonUI.ShowSettings — настройки
'    NeonCenterline.NeonUI.CheckUpdates — проверить обновления
'    NeonCenterline.NeonUI.ShowAbout    — о программе
'  Инструменты — в модулях NeonTools (длина, выделить линии, упростить контур) и NeonCut.
'                                                        SAB:CORE, 2026
' =====================================================================
Option Explicit

Public Const NEON_VERSION As String = "1.0.2"
' адрес файла с описанием последней версии (см. update.json в репозитории)
Public Const NEON_UPDATE_URL As String = "https://raw.githubusercontent.com/ShiZa039/NeonCenterline/main/update.json"
' запасные адреса того же файла (raw.githubusercontent.com бывает недоступен у части провайдеров)
Public Const NEON_UPDATE_URL2 As String = "https://cdn.jsdelivr.net/gh/ShiZa039/NeonCenterline@main/update.json"
Public Const NEON_UPDATE_URL3 As String = "https://github.com/ShiZa039/NeonCenterline/raw/main/update.json"
' страница поддержки автора: НЕ сам сервис донатов, а DONATE.md в репозитории — ссылки там можно менять без выпуска новой версии (пусто — пункт не показывается)
Public Const NEON_DONATE_URL As String = "https://github.com/ShiZa039/NeonCenterline/blob/main/DONATE.md"
Private Const REG_APP As String = "NeonCenterline"
' результат окна перед расчётом (frmNeonRun): True — «Построить» / «Упростить»
Public NeonRunConfirmed As Boolean
' режим окна frmNeonRun: 0 — центральные линии, 1 — упрощение контура, 2 — люверсы
Public NeonRunMode As Long
' версия набора кнопок панели «Неон». Поднимать вместе с меткой [NeonToolbar N] в
' addon/NeonCenterline/AppUI.xslt, когда меняется состав панели (tools/gen_addon.py сверяет)
Public Const NEON_TOOLBAR_VER As String = "7"
' файлы панели обновлены в этом сеансе (InstallToolbar): в памяти CorelDRAW ещё старый состав
' панели, сбрасывать её можно только после перезапуска
Public NeonTbInstalled As Boolean
' событиям окон (ThisMacroStorage) в этом сеансе панель больше проверять не нужно
Private NeonTbQuiet As Boolean

' ---------------------------------------------------------------------
'  Кнопки
' ---------------------------------------------------------------------
Public Sub RunMono()
    RunMode 0
End Sub

Public Sub RunScript()
    RunMode 1
End Sub

Public Sub ShowSettings()
    Dim f As Object
    NeonToolbarCheck
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmNeonSettings")
    On Error GoTo 0
    If f Is Nothing Then
        SettingsFallback
    Else
        f.Show
        Set f = Nothing
    End If
End Sub

Public Sub CheckUpdates()
    NeonUpdateCheck False
End Sub

Public Sub ShowAbout()
    Dim msg As String
    msg = NeonLF("about", NEON_VERSION)
    If Len(NEON_DONATE_URL) > 0 Then
        If MsgBox(msg & NeonL("about_donate"), vbYesNo + vbInformation, NeonL("app_title")) = vbYes Then
            OpenUrl NEON_DONATE_URL
        End If
    Else
        MsgBox msg, vbInformation, NeonL("app_title")
    End If
End Sub

' ---------------------------------------------------------------------
'  Запуск расчёта
' ---------------------------------------------------------------------
Private Sub RunMode(ByVal mode As Long)
    Dim tool As Double, rc As Long, nIn As Long, nOut As Long
    Dim tAll As Double, tCore As Double, errText As String, mname As String
    Dim before As Collection, cutRep As String, newCol As Collection, eo As String

    If GetSetting(REG_APP, "UI", "AskTool", "1") = "1" Then
        If Not AskRunParams() Then
            Exit Sub
        End If
    End If
    tool = NeonGetTool()

    Set before = Nothing
    If NeonCutFitOn() And NeonCutStep() > 0 And Not ActiveDocument Is Nothing Then
        Set before = NeonLineIds()
    End If
    ' отступ свободного конца: пусто — D/2; в режиме «образец» у курсива концы задаёт сам режим
    eo = GetSetting(REG_APP, "Params", "EndOff", "")
    If Len(eo) > 0 Then
        NeonMonoSetEndOff ParseNum(eo)
        NeonScriptSetEndOff ParseNum(eo)
    Else
        NeonMonoSetEndOff -1
        NeonScriptSetEndOff -1
    End If
    If mode = 1 And Val(GetSetting(REG_APP, "Script", "RefEnd", "1")) > 0 Then
        NeonScriptSetEndOff -1
    End If
    If mode = 0 Then
        mname = NeonL("mode_mono")
        rc = NeonMonoRun(tool, nIn, nOut, tAll, tCore, errText)
    Else
        mname = NeonL("mode_script")
        NeonScriptSetRef CLng(Val(GetSetting(REG_APP, "Script", "RefEnd", "1")))
        rc = NeonScriptRun(tool, nIn, nOut, tAll, tCore, errText)
    End If

    ' подгонка новых линий под кратность реза
    If rc = 0 And Not before Is Nothing Then
        On Error Resume Next
        Set newCol = NeonLinesNew(before)
        If Not newCol Is Nothing Then
            If newCol.Count > 0 Then
                ActiveDocument.BeginCommandGroup NeonL("cut_title")
                cutRep = vbCrLf & vbCrLf & NeonCutFitShapes(newCol)
                ActiveDocument.EndCommandGroup
            End If
        End If
        If Err.Number <> 0 Then
            cutRep = vbCrLf & vbCrLf & NeonLF("err_fail", Err.Description)
        End If
        On Error GoTo 0
    End If

    Select Case rc
    Case 0
        If mode = 1 Then
            MsgBox NeonLF("done", mname, FmtNum(tool), nIn, nOut, Format$(tAll, "0.0"), Format$(tCore, "0.0")) & cutRep & _
                   NeonL("script_beta"), vbExclamation, NeonL("app_title")
        Else
            MsgBox NeonLF("done", mname, FmtNum(tool), nIn, nOut, Format$(tAll, "0.0"), Format$(tCore, "0.0")) & cutRep, _
                   vbInformation, NeonL("app_title")
        End If
    Case 1
        MsgBox NeonL("err_nodoc"), vbExclamation, NeonL("app_title")
    Case 2
        MsgBox NeonL("err_nosel"), vbExclamation, NeonL("app_title")
    Case 3
        MsgBox NeonL("err_noclosed"), vbExclamation, NeonL("app_title")
    Case Else
        MsgBox NeonLF("err_fail", errText), vbCritical, NeonL("app_title")
    End Select

    ' раз в день — тихая проверка обновлений (после расчёта, чтобы не задерживать работу)
    If rc = 0 Then
        NeonToolbarCheck
        NeonDailyUpdateCheck
    End If
End Sub

' окно параметров перед расчётом; без собранной формы — прежний вопрос о диаметре
Private Function AskRunParams() As Boolean
    Dim f As Object, s As String, tool As Double
    NeonRunMode = 0
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmNeonRun")
    On Error GoTo 0
    If Not f Is Nothing Then
        NeonRunConfirmed = False
        f.Show
        AskRunParams = NeonRunConfirmed
        Unload f
        Exit Function
    End If
    s = InputBox(NeonL("prompt_tool"), NeonL("app_title"), FmtNum(NeonGetTool()))
    If Len(Trim$(s)) = 0 Then
        Exit Function
    End If
    tool = ParseNum(s)
    If tool <= 0 Or tool > 100 Then
        MsgBox NeonL("bad_tool"), vbExclamation, NeonL("app_title")
        Exit Function
    End If
    SaveSetting REG_APP, "Params", "Tool", FmtNum(tool)
    AskRunParams = True
End Function

' ---------------------------------------------------------------------
'  Обновление панели «Неон» после обновления плагина.
'  CorelDRAW вписывает кнопки дополнения в рабочее пространство один раз и сам новые кнопки
'  в стоящую панель не добавляет. Поэтому, если в Addons лежат файлы панели этой версии, а панель
'  ещё не обновлялась, — сброс только панели «Неон» к составу из AppUI.xslt (как «Сброс» в
'  настройке панелей; остальное рабочее пространство не трогается):
'    - при запуске CorelDRAW (ThisMacroStorage, src = 1) — сразу, но не засчитывая: рабочее
'      пространство может догрузиться позже и вернуть панели прежний состав;
'    - при открытии окна документа (src = 2) и при командах плагина (src = 0): если на панели нет
'      кнопок нового набора (NeonBarProbe) — сброс; засчитывается, когда кнопки на месте (или
'      CorelDRAW не даёт их проверить); три неудачные попытки — больше не пытаться.
'  Не в том сеансе, где файлы только что обновлены (NeonTbInstalled).
'  Если файлы панели в Addons старые (или нет значков этого набора) — при команде плагина один раз
'  предложить обновить их (InstallToolbar, права администратора) и перезапустить CorelDRAW.
'  Вручную: макрос ResetToolbar (сброс и отчёт, что найдено).
' ---------------------------------------------------------------------
Public Sub NeonToolbarCheck(Optional ByVal src As Long = 0)
    Dim st As Long, rep As String, pr As Long, n As Long, t As String
    If src = 2 And NeonTbQuiet Then
        Exit Sub
    End If
    If GetSetting(REG_APP, "UI", "ToolbarVer", "") = NEON_TOOLBAR_VER Then
        NeonTbQuiet = True
        Exit Sub
    End If
    On Error GoTo Quit
    st = NeonToolbarFilesState()
    If st = 0 Then
        ' панель не ставилась (кнопки добавлены вручную или не нужны) — нечего обновлять
        SaveSetting REG_APP, "UI", "ToolbarVer", NEON_TOOLBAR_VER
        NeonTbQuiet = True
    ElseIf st = 1 Then
        ' предлагать обновить файлы — только при командах плагина
        If src = 0 And GetSetting(REG_APP, "UI", "ToolbarAsked", "") <> NEON_TOOLBAR_VER Then
            SaveSetting REG_APP, "UI", "ToolbarAsked", NEON_TOOLBAR_VER
            If MsgBox(NeonL("tb_outdated"), vbYesNo + vbQuestion, NeonL("app_title")) = vbYes Then
                InstallToolbar
            End If
        End If
        If src <> 0 Then
            NeonTbQuiet = True
        End If
    ElseIf NeonTbInstalled Then
        NeonTbQuiet = True
    ElseIf src = 1 Then
        ResetNeonBar rep
    Else
        pr = NeonBarProbe()
        If pr <> 1 Then
            If Not ResetNeonBar(rep) Then
                ' панели «Неон» нет (закрыта или удалена) — событиям окон до перезапуска не искать
                NeonTbQuiet = True
                Exit Sub
            End If
            pr = NeonBarProbe()
        End If
        If pr = 0 Then
            ' после сброса кнопок нового набора на панели нет — попробовать позже (не чаще раза
            ' за сеанс от событий окон), после трёх попыток оставить как есть
            NeonTbQuiet = True
            t = GetSetting(REG_APP, "UI", "ToolbarTries", "")
            If Left$(t, Len(NEON_TOOLBAR_VER) + 1) = NEON_TOOLBAR_VER & ":" Then
                n = Val(Mid$(t, Len(NEON_TOOLBAR_VER) + 2))
            End If
            n = n + 1
            SaveSetting REG_APP, "UI", "ToolbarTries", NEON_TOOLBAR_VER & ":" & n
            If n < 3 Then
                Exit Sub
            End If
        End If
        SaveSetting REG_APP, "UI", "ToolbarVer", NEON_TOOLBAR_VER
        NeonTbQuiet = True
    End If
Quit:
End Sub

' есть ли на панели «Неон» все кнопки этого набора (NeonTbProbeItems: GUID кнопки, GUID списка
' кнопки-группы, подпись): 1 — все, 0 — каких-то нет, -1 — не понять (панели нет или CorelDRAW
' не отдаёт ни кодов, ни подписей кнопок)
Private Function NeonBarProbe() As Long
    Dim cb As Object, ctl As Object, nm As String, nl As String, have As String, i As Long, n As Long
    Dim items As Variant, it As Variant, f As Variant, j As Long, s As String, c As String, known As Boolean, hit As Boolean
    On Error Resume Next
    NeonBarProbe = -1
    For Each cb In Application.CommandBars
        nm = "": nl = ""
        nm = cb.Name
        nl = cb.NameLocal
        If InStr(1, nm, "NeonCenterline", vbTextCompare) > 0 Or InStr(1, nl, "NeonCenterline", vbTextCompare) > 0 Then
            n = 0
            n = cb.Controls.Count
            For i = 1 To n
                Set ctl = Nothing
                Set ctl = cb.Controls(i)
                If Not ctl Is Nothing Then
                    s = "": c = ""
                    s = ctl.ID
                    c = ctl.Caption
                    If Len(s) > 0 Or Len(c) > 0 Then
                        known = True
                    End If
                    have = have & "|" & LCase$(s) & "|" & LCase$(c) & "|"
                End If
            Next i
        End If
    Next cb
    If Not known Then
        Exit Function
    End If
    ' коды — частью строки (могут прийти в скобках), подпись — целиком
    items = NeonTbProbeItems()
    For Each it In items
        hit = False
        f = Split(it, vbTab)
        For j = 0 To UBound(f)
            c = LCase$(f(j))
            If Len(c) > 0 Then
                If j = 2 Then
                    c = "|" & c & "|"
                End If
                If InStr(1, have, c, vbBinaryCompare) > 0 Then
                    hit = True
                End If
            End If
        Next j
        If Not hit Then
            NeonBarProbe = 0
            Exit Function
        End If
    Next it
    NeonBarProbe = 1
End Function

' сбросить панель «Неон» вручную: Инструменты > Макросы > NeonCenterline.NeonUI.ResetToolbar.
' Показывает, какие панели найдены и что на них после сброса (для проверки)
Public Sub ResetToolbar()
    Dim rep As String, ok As Boolean
    If NeonToolbarFilesState() <> 2 Then
        ' в Addons файлы панели старой версии (или их нет): сброс вернёт старый набор кнопок
        If MsgBox(NeonLF("tb_files_old", ToolbarFilesInfo(), NEON_TOOLBAR_VER), vbYesNo + vbQuestion, NeonL("app_title")) = vbYes Then
            InstallToolbar
        End If
        Exit Sub
    End If
    ok = ResetNeonBar(rep, True)
    rep = rep & "probe: " & NeonBarProbe() & vbLf & vbLf & ToolbarFilesInfo() & vbLf & _
          "NeonCenterline.gms: [NeonToolbar " & NEON_TOOLBAR_VER & "], " & NeonAddonDll()
    If ok Then
        SaveSetting REG_APP, "UI", "ToolbarVer", NEON_TOOLBAR_VER
        NeonTbQuiet = True
        MsgBox NeonLF("tb_reset_done", rep), vbInformation, NeonL("app_title")
    Else
        MsgBox NeonLF("tb_reset_none", rep), vbExclamation, NeonL("app_title")
    End If
End Sub

' файлы панели в Addons текущего CorelDRAW: 0 — нет, 1 — старые, 2 — этой версии (метка в AppUI.xslt
' и файл значков этого набора, NeonAddonDll, целиком)
Public Function NeonToolbarFilesState() As Long
    Dim p As String, f As Integer, txt As String, n As Long
    On Error GoTo Fail
    p = Application.AddonPath
    If Right$(p, 1) <> "\" Then
        p = p & "\"
    End If
    p = p & "NeonCenterline\"
    If Len(Dir$(p & "AppUI.xslt")) = 0 Then
        Exit Function
    End If
    f = FreeFile
    Open p & "AppUI.xslt" For Binary Access Read As #f
    txt = Space$(LOF(f))
    Get #f, , txt
    Close #f
    NeonToolbarFilesState = 1
    If InStr(1, txt, "[NeonToolbar " & NEON_TOOLBAR_VER & "]", vbBinaryCompare) > 0 Then
        On Error Resume Next
        n = -1
        n = FileLen(p & NeonAddonDll())
        If n = NeonAddonDllSize() Then
            NeonToolbarFilesState = 2
        End If
    End If
    Exit Function
Fail:
    NeonToolbarFilesState = 0
End Function

' установка файлов панели не удалась — при следующей команде плагина спросить снова
Public Sub NeonTbAskAgain()
    On Error Resume Next
    SaveSetting REG_APP, "UI", "ToolbarAsked", ""
End Sub

' файлы панели в Addons текущего CorelDRAW: путь, метка набора кнопок, дата, файлы значков (для отчёта ResetToolbar)
Private Function ToolbarFilesInfo() As String
    Dim p As String, f As Integer, txt As String, k As Long, e As Long, r As String
    On Error GoTo Fail
    p = Application.AddonPath
    If Right$(p, 1) <> "\" Then
        p = p & "\"
    End If
    p = p & "NeonCenterline\AppUI.xslt"
    r = p & ": "
    If Len(Dir$(p)) = 0 Then
        ToolbarFilesInfo = r & "-"
        Exit Function
    End If
    f = FreeFile
    Open p For Binary Access Read As #f
    txt = Space$(LOF(f))
    Get #f, , txt
    Close #f
    k = InStr(1, txt, "[NeonToolbar ", vbBinaryCompare)
    If k > 0 Then
        e = InStr(k, txt, "]", vbBinaryCompare)
    End If
    If k > 0 And e > k Then
        r = r & Mid$(txt, k, e - k + 1)
    Else
        r = r & "[NeonToolbar 1]"
    End If
    r = r & ", " & Format$(FileDateTime(p), "yyyy-mm-dd hh:nn")
    p = Left$(p, Len(p) - Len("AppUI.xslt"))
    txt = Dir$(p & "Resources*.dll")
    Do While Len(txt) > 0
        r = r & vbLf & txt & ": " & FileLen(p & txt) & ", " & Format$(FileDateTime(p & txt), "yyyy-mm-dd hh:nn")
        txt = Dir$()
    Loop
    ToolbarFilesInfo = r
    Exit Function
Fail:
    ToolbarFilesInfo = r & Err.Description
End Function

' сбросить панель «Неон» к составу из AppUI.xslt; True — панель найдена и сброшена.
' rep — отчёт: имя панели, число кнопок до и после сброса (detail — и подписи и коды кнопок)
Private Function ResetNeonBar(ByRef rep As String, Optional ByVal detail As Boolean = False) As Boolean
    Dim cb As Object, nm As String, nl As String, n0 As Long, n1 As Long, i As Long, cap As String, idt As String
    On Error Resume Next
    rep = ""
    For Each cb In Application.CommandBars
        nm = "": nl = ""
        nm = cb.Name
        nl = cb.NameLocal
        If InStr(1, nm, "NeonCenterline", vbTextCompare) > 0 Or InStr(1, nl, "NeonCenterline", vbTextCompare) > 0 Then
            n0 = -1: n1 = -1
            n0 = cb.Controls.Count
            Err.Clear
            cb.Reset
            If Err.Number = 0 Then
                ResetNeonBar = True
                n1 = cb.Controls.Count
                If n1 < 0 Then
                    ' после сброса прежний объект панели может стать недействительным
                    Err.Clear
                    n1 = Application.CommandBars.Item(nm).Controls.Count
                End If
                rep = rep & nm & " / " & nl & ": " & n0 & " -> " & n1 & vbLf
                If detail Then
                    For i = 1 To n1
                        cap = "?": idt = ""
                        cap = Application.CommandBars.Item(nm).Controls(i).Caption
                        idt = Application.CommandBars.Item(nm).Controls(i).ID
                        rep = rep & "    " & i & ". " & cap & "  [" & idt & "]" & vbLf
                    Next i
                End If
            Else
                rep = rep & nm & " / " & nl & ": " & Err.Description & vbLf
                Err.Clear
            End If
        End If
    Next cb
End Function

' окно перед расстановкой люверсов (frmNeonRun в режиме 2): отступ, шаг, диаметр и цвет точки,
' усиление углов. Без собранной формы — отступ и шаг стандартными окнами. False — отмена
Public Function AskGrommetParams() As Boolean
    Dim f As Object, s As String, v As Double
    NeonRunMode = 2
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmNeonRun")
    On Error GoTo 0
    If Not f Is Nothing Then
        NeonRunConfirmed = False
        f.Show
        AskGrommetParams = NeonRunConfirmed
        Unload f
        NeonRunMode = 0
        Exit Function
    End If
    NeonRunMode = 0
    s = InputBox(NeonL("lbl_grom_edge"), NeonL("grom_title"), FmtNum(NeonGromEdge()))
    If Len(Trim$(s)) = 0 Then
        Exit Function
    End If
    v = ParseNum(s)
    If v <= 0 Or v > 500 Then
        MsgBox NeonL("bad_grom"), vbExclamation, NeonL("grom_title")
        Exit Function
    End If
    SaveSetting REG_APP, "Grommet", "Edge", FmtNum(v)
    s = InputBox(NeonL("lbl_grom_step"), NeonL("grom_title"), FmtNum(NeonGromStep()))
    If Len(Trim$(s)) = 0 Then
        Exit Function
    End If
    v = ParseNum(s)
    If v < 20 Or v > 10000 Then
        MsgBox NeonL("bad_grom"), vbExclamation, NeonL("grom_title")
        Exit Function
    End If
    SaveSetting REG_APP, "Grommet", "Step", FmtNum(v)
    AskGrommetParams = True
End Function

' окно перед упрощением контура (то же окно frmNeonRun в режиме 1): допуск, показывать ли окно;
' без собранной формы — вопрос о допуске стандартным окном. False — отмена
Public Function AskSimplifyParams() As Boolean
    Dim f As Object, s As String, t As Double
    If GetSetting(REG_APP, "Simplify", "Ask", "1") <> "1" Then
        AskSimplifyParams = True
        Exit Function
    End If
    NeonRunMode = 1
    On Error Resume Next
    Set f = VBA.UserForms.Add("frmNeonRun")
    On Error GoTo 0
    If Not f Is Nothing Then
        NeonRunConfirmed = False
        f.Show
        AskSimplifyParams = NeonRunConfirmed
        Unload f
        NeonRunMode = 0
        Exit Function
    End If
    NeonRunMode = 0
    s = InputBox(NeonL("prompt_simp"), NeonL("simp_title"), FmtNum(NeonSimplifyTol()))
    If Len(Trim$(s)) = 0 Then
        Exit Function
    End If
    t = ParseNum(s)
    If t < 0.01 Or t > 5 Then
        MsgBox NeonL("bad_simp"), vbExclamation, NeonL("simp_title")
        Exit Function
    End If
    SaveSetting REG_APP, "Simplify", "Tol", FmtNum(t)
    AskSimplifyParams = True
End Function

' ---------------------------------------------------------------------
'  Настройки (общие для обеих кнопок; диаметр — тот же ключ, что у прежнего макроса)
' ---------------------------------------------------------------------
Public Function NeonGetTool() As Double
    Dim t As Double
    t = ParseNum(GetSetting(REG_APP, "Params", "Tool", "8"))
    If t <= 0 Then
        t = 8
    End If
    NeonGetTool = t
End Function

' без формы настроек (не собрана): по шагам через стандартные окна
Private Sub SettingsFallback()
    Dim s As String, t As Double, r As Long
    s = InputBox(NeonL("prompt_tool"), NeonL("settings_title"), FmtNum(NeonGetTool()))
    If Len(Trim$(s)) > 0 Then
        t = ParseNum(s)
        If t > 0 And t <= 100 Then
            SaveSetting REG_APP, "Params", "Tool", FmtNum(t)
        Else
            MsgBox NeonL("bad_tool"), vbExclamation, NeonL("settings_title")
        End If
    End If
    s = InputBox(NeonL("prompt_simp"), NeonL("settings_title"), FmtNum(NeonSimplifyTol()))
    If Len(Trim$(s)) > 0 Then
        t = ParseNum(s)
        If t >= 0.01 And t <= 5 Then
            SaveSetting REG_APP, "Simplify", "Tol", FmtNum(t)
        Else
            MsgBox NeonL("bad_simp"), vbExclamation, NeonL("settings_title")
        End If
    End If
    r = MsgBox(NeonL("fallback_ask"), vbYesNo + vbQuestion, NeonL("settings_title"))
    SaveSetting REG_APP, "UI", "AskTool", IIf(r = vbYes, "1", "0")
    r = MsgBox(NeonL("fallback_lang"), vbYesNoCancel + vbQuestion, NeonL("settings_title"))
    If r = vbYes Then
        SaveSetting REG_APP, "UI", "Lang", "ru"
    ElseIf r = vbNo Then
        SaveSetting REG_APP, "UI", "Lang", "en"
    Else
        SaveSetting REG_APP, "UI", "Lang", ""
    End If
    r = MsgBox(NeonL("fallback_upd"), vbYesNo + vbQuestion, NeonL("settings_title"))
    SaveSetting REG_APP, "Update", "Daily", IIf(r = vbYes, "1", "0")
End Sub

' число из строки: и запятая, и точка — десятичный разделитель
Public Function ParseNum(ByVal s As String) As Double
    s = Replace(Trim$(s), ",", ".")
    If Len(s) = 0 Then
        ParseNum = 0
    Else
        ParseNum = Val(s)
    End If
End Function

Public Function FmtNum(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then
        s = "0" & s
    End If
    FmtNum = s
End Function

Public Sub OpenUrl(ByVal url As String)
    On Error Resume Next
    CreateObject("Shell.Application").ShellExecute url
End Sub

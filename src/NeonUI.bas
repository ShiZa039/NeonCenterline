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
'                                                        SAB:CORE, 2026
' =====================================================================
Option Explicit

Public Const NEON_VERSION As String = "1.0.0"
' адрес файла с описанием последней версии (см. update.json в репозитории)
Public Const NEON_UPDATE_URL As String = "https://raw.githubusercontent.com/OWNER/NeonCenterline/main/update.json"
' страница поддержки автора (пусто — пункт не показывается)
Public Const NEON_DONATE_URL As String = ""
Private Const REG_APP As String = "NeonCenterline"

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
    Dim tool As Double, s As String, rc As Long, nIn As Long, nOut As Long
    Dim tAll As Double, tCore As Double, errText As String, mname As String

    tool = NeonGetTool()
    If GetSetting(REG_APP, "UI", "AskTool", "1") = "1" Then
        s = InputBox(NeonL("prompt_tool"), NeonL("app_title"), FmtNum(tool))
        If Len(Trim$(s)) = 0 Then
            Exit Sub
        End If
        tool = ParseNum(s)
        If tool <= 0 Or tool > 100 Then
            MsgBox NeonL("bad_tool"), vbExclamation, NeonL("app_title")
            Exit Sub
        End If
        SaveSetting REG_APP, "Params", "Tool", FmtNum(tool)
    End If

    If mode = 0 Then
        mname = NeonL("mode_mono")
        rc = NeonMonoRun(tool, nIn, nOut, tAll, tCore, errText)
    Else
        mname = NeonL("mode_script")
        rc = NeonScriptRun(tool, nIn, nOut, tAll, tCore, errText)
    End If

    Select Case rc
    Case 0
        MsgBox NeonLF("done", mname, FmtNum(tool), nIn, nOut, Format$(tAll, "0.0"), Format$(tCore, "0.0")), _
               vbInformation, NeonL("app_title")
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
        NeonDailyUpdateCheck
    End If
End Sub

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

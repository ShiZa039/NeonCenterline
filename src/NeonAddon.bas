Attribute VB_Name = "NeonAddon"
' =====================================================================
'  NeonCenterline — установка панели инструментов в текущий CorelDRAW.
'  ФАЙЛ СГЕНЕРИРОВАН tools/gen_addon.py из addon/NeonCenterline — не править вручную.
'  Файлы панели хранятся здесь кодами Unicode; макрос пишет их во временную папку и
'  копирует (с правами администратора) в Application.AddonPath\NeonCenterline.
' =====================================================================
Option Explicit

Public Sub InstallToolbar()
    Dim tmp As String, dst As String, cmd As String, sh As Object
    On Error GoTo Fail
    tmp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\NeonCenterline_addon"
    MakeDir tmp
    WriteUtf8 tmp & "\AppUI.xslt", AddonText_AppUI_xslt()
    WriteUtf8 tmp & "\UserUI.xslt", AddonText_UserUI_xslt()
    WriteUtf8 tmp & "\Coreldrw.addon", AddonText_Coreldrw_addon()
    dst = Application.AddonPath
    If Right$(dst, 1) <> "\" Then
        dst = dst & "\"
    End If
    dst = dst & "NeonCenterline"
    cmd = "/c mkdir """ & dst & """ 2>nul & xcopy /Y /I """ & tmp & "\*"" """ & dst & """"
    Set sh = CreateObject("Shell.Application")
    sh.ShellExecute "cmd.exe", cmd, "", "runas", 0
    MsgBox NeonLF("addon_done", dst), vbInformation, NeonL("app_title")
    Exit Sub
Fail:
    MsgBox NeonLF("err_fail", Err.Description), vbExclamation, NeonL("app_title")
End Sub

Private Sub MakeDir(ByVal p As String)
    On Error Resume Next
    MkDir p
End Sub

Private Sub WriteUtf8(ByVal path As String, ByVal txt As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText txt
    st.SaveToFile path, 2
    st.Close
End Sub

Private Function AU(ByVal hx As String) As String
    Dim a As Variant, i As Long, r As String
    If Len(hx) = 0 Then
        Exit Function
    End If
    a = Split(hx, " ")
    For i = 0 To UBound(a)
        r = r & ChrW(CLng("&H" & a(i)))
    Next i
    AU = r
End Function

Private Function AddonText_AppUI_xslt() As String
    Dim s As String
    s = s & AU("3C 3F 78 6D 6C 20 76 65 72 73 69 6F 6E 3D 22 31 2E 30 22 3F 3E A 3C 21 2D 2D A 20 20 4E 65 6F 6E 43 65 6E 74 65 72 6C 69 6E 65 20 2014 20 43A 43E 43C 430 43D 434 44B 20 438 20 43F 430 43D 435 43B 44C 20 438 43D 441 442 440 443 43C")
    s = s & AU("435 43D 442 43E 432 20 28 43 6F 72 65 6C 44 52 41 57 20 41 64 64 6F 6E 73 29 2E A 20 20 41A 43E 43C 430 43D 434 44B 20 432 44B 437 44B 432 430 44E 442 20 43C 430 43A 440 43E 441 44B 20 43F 440 43E 435 43A 442 430 20 4E 65 6F 6E 43 65 6E 74 65")
    s = s & AU("72 6C 69 6E 65 2E 67 6D 73 20 28 43F 43E 43B 44C 437 43E 432 430 442 435 43B 44C 441 43A 430 44F 20 43F 430 43F 43A 430 20 47 4D 53 29 2E A 20 20 47 55 49 44 20 43D 435 20 43C 435 43D 44F 442 44C 3A 20 43F 43E 20 43D 438 43C 20 43 6F 72 65 6C")
    s = s & AU("44 52 41 57 20 443 437 43D 430 451 442 20 43A 43D 43E 43F 43A 438 20 432 20 43D 430 441 442 440 43E 435 43D 43D 44B 445 20 440 430 431 43E 447 438 445 20 43F 440 43E 441 442 440 430 43D 441 442 432 430 445 2E A 2D 2D 3E A 3C 78 73 6C 3A 73 74 79 6C 65")
    s = s & "sheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"" x"
    s = s & "mlns:frmwrk=""Corel Framework Data"">" & vbLf & "  <xsl:output method=""xml"" encodin"
    s = s & "g=""UTF-8"" indent=""yes""/>" & vbLf & "" & vbLf & "  <frmwrk:uiconfig>" & vbLf & "    <frmwrk:applicationI"
    s = s & "nfo userConfiguration=""true"" />" & vbLf & "  </frmwrk:uiconfig>" & vbLf & "" & vbLf & "  <xsl:template "
    s = s & "match=""node()|@*"">" & vbLf & "    <xsl:copy>" & vbLf & "      <xsl:apply-templates select=""n"
    s = s & "ode()|@*""/>" & vbLf & "    </xsl:copy>" & vbLf & "  </xsl:template>" & vbLf & "" & vbLf & "  <xsl:template match="""
    s = s & "uiConfig/items"">" & vbLf & "    <xsl:copy>" & vbLf & "      <xsl:apply-templates select=""nod"
    s = s & "e()|@*""/>" & vbLf & "      <itemData guid=""4e6d43bf-0b2d-500c-b154-b553b6087991""" & vbLf & ""
    s = s & "                dynamicCommand=""NeonCenterline.NeonUI.RunMono""" & vbLf & "       "
    s = s & "         dynamicCategory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "      "
    s = s & AU("20 20 20 20 20 20 20 20 20 20 75 73 65 72 43 61 70 74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 41C 43E 43D 43E 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20")
    s = s & "   <itemData guid=""c68c5d4f-d668-5fb4-937e-adb7430cae58""" & vbLf & "             "
    s = s & "   dynamicCommand=""NeonCenterline.NeonUI.RunScript""" & vbLf & "                dy"
    s = s & "namicCategory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "                u"
    s = s & AU("73 65 72 43 61 70 74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 41A 443 440 441 438 432 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20 3C 69 74 65 6D 44")
    s = s & "ata guid=""ce37ed0a-544f-575c-ac35-ef4d4aab66c9""" & vbLf & "                dynami"
    s = s & "cCommand=""NeonCenterline.NeonUI.ShowSettings""" & vbLf & "                dynamicC"
    s = s & "ategory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "                userCap"
    s = s & AU("74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 43D 430 441 442 440 43E 439 43A 438 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20 3C 69 74 65 6D 44 61 74 61")
    s = s & " guid=""1025a06f-bc40-5284-9292-2c3cf477d6a2""" & vbLf & "                dynamicCo"
    s = s & "mmand=""NeonCenterline.NeonUI.CheckUpdates""" & vbLf & "                dynamicCate"
    s = s & "gory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "                userCaptio"
    s = s & AU("6E 3D 22 41D 435 43E 43D 3A 20 43F 440 43E 432 435 440 438 442 44C 20 43E 431 43D 43E 432 43B 435 43D 438 44F 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20 3C")
    s = s & "itemData guid=""c5cc733d-ee55-5173-b5ef-17a954845da7""" & vbLf & "                d"
    s = s & "ynamicCommand=""NeonCenterline.NeonUI.ShowAbout""" & vbLf & "                dynami"
    s = s & "cCategory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "                userC"
    s = s & AU("61 70 74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 43E 20 43F 440 43E 433 440 430 43C 43C 435 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 3C 2F 78 73 6C 3A 63")
    s = s & "opy>" & vbLf & "  </xsl:template>" & vbLf & "" & vbLf & "  <xsl:template match=""uiConfig/commandBars"">" & vbLf & ""
    s = s & "    <xsl:copy>" & vbLf & "      <xsl:apply-templates select=""node()|@*""/>" & vbLf & "      <"
    s = s & "commandBarData guid=""772fbf90-5bc1-5c4d-93a0-d1fced7f98b5""" & vbLf & "           "
    s = s & "           nonLocalizableName=""NeonCenterline""" & vbLf & "                      u"
    s = s & "serCaption=""NeonCenterline""" & vbLf & "                      type=""toolbar"">" & vbLf & "    "
    s = s & "    <toolbar>" & vbLf & "          <item guidRef=""4e6d43bf-0b2d-500c-b154-b553b60"
    s = s & "87991""/>" & vbLf & "          <item guidRef=""c68c5d4f-d668-5fb4-937e-adb7430cae58"
    s = s & """/>" & vbLf & "          <item guidRef=""ce37ed0a-544f-575c-ac35-ef4d4aab66c9""/>" & vbLf & " "
    s = s & "       </toolbar>" & vbLf & "      </commandBarData>" & vbLf & "    </xsl:copy>" & vbLf & "  </xsl:temp"
    s = s & "late>" & vbLf & "</xsl:stylesheet>" & vbLf & ""
    AddonText_AppUI_xslt = s
End Function

Private Function AddonText_UserUI_xslt() As String
    Dim s As String
    s = s & AU("3C 3F 78 6D 6C 20 76 65 72 73 69 6F 6E 3D 22 31 2E 30 22 3F 3E A 3C 21 2D 2D A 20 20 4E 65 6F 6E 43 65 6E 74 65 72 6C 69 6E 65 20 2014 20 440 430 437 43C 435 449 435 43D 438 435 20 43F 430 43D 435 43B 438 20 432 20 440 430 431 43E")
    s = s & AU("447 435 43C 20 43F 440 43E 441 442 440 430 43D 441 442 432 435 2E 20 41F 440 438 43C 435 43D 44F 435 442 441 44F 20 43E 434 438 43D 20 440 430 437 20 43D 430 A 20 20 440 430 431 43E 447 435 435 20 43F 440 43E 441 442 440 430 43D 441 442 432 43E 3B 20 43F 43E 441 43B")
    s = s & AU("435 20 438 437 43C 435 43D 435 43D 438 44F 20 2014 20 437 430 43F 443 441 43A 20 43 6F 72 65 6C 44 52 41 57 20 441 20 437 430 436 430 442 43E 439 20 46 38 20 28 441 431 440 43E 441 20 440 430 431 43E 447 435 433 43E A 20 20 43F 440 43E 441 442 440 430 43D")
    s = s & AU("441 442 432 430 21 29 20 438 43B 438 20 432 43A 43B 44E 447 438 442 44C 20 43F 430 43D 435 43B 44C 20 432 440 443 447 43D 443 44E 3A 20 41E 43A 43D 43E 20 3E 20 41F 430 43D 435 43B 438 20 438 43D 441 442 440 443 43C 435 43D 442 43E 432 20 3E 20 4E 65 6F 6E 43")
    s = s & "enterline." & vbLf & "-->" & vbLf & "<xsl:stylesheet version=""1.0""" & vbLf & "                xmlns:xsl"
    s = s & "=""http://www.w3.org/1999/XSL/Transform""" & vbLf & "                xmlns:frmwrk="""
    s = s & "Corel Framework Data""" & vbLf & "                exclude-result-prefixes=""frmwrk"""
    s = s & ">" & vbLf & "  <xsl:output method=""xml"" encoding=""UTF-8"" indent=""yes""/>" & vbLf & "" & vbLf & "  <frmwr"
    s = s & "k:uiconfig>" & vbLf & "    <frmwrk:compositeNode xPath=""/uiConfig/commandBars/com"
    s = s & "mandBarData[@guid='772fbf90-5bc1-5c4d-93a0-d1fced7f98b5']""/>" & vbLf & "    <frmw"
    s = s & "rk:compositeNode xPath=""/uiConfig/frame""/>" & vbLf & "  </frmwrk:uiconfig>" & vbLf & "" & vbLf & "  <xs"
    s = s & "l:template match=""node()|@*"">" & vbLf & "    <xsl:copy>" & vbLf & "      <xsl:apply-template"
    s = s & AU("73 20 73 65 6C 65 63 74 3D 22 6E 6F 64 65 28 29 7C 40 2A 22 2F 3E A 20 20 20 20 3C 2F 78 73 6C 3A 63 6F 70 79 3E A 20 20 3C 2F 78 73 6C 3A 74 65 6D 70 6C 61 74 65 3E A A 20 20 3C 21 2D 2D 20 43F 430 43D 435 43B")
    s = s & AU("44C 20 2014 20 440 44F 434 43E 43C 20 441 43E 20 441 442 430 43D 434 430 440 442 43D 43E 439 20 43F 430 43D 435 43B 44C 44E 20 438 43D 441 442 440 443 43C 435 43D 442 43E 432 20 28 441 43E 441 435 434 43D 438 439 20 44D 43B 435 43C 435 43D 442 29 20 2D 2D 3E A 20")
    s = s & " <xsl:template match=""uiConfig/containers/container[@guid='bee85f91-3a"
    s = s & "d9-dc8d-48b5-d2a87c8b2109']/container[@guid='Framework_MainFrame-layou"
    s = s & "t']/dockHost[@guid='894bf987-2ec1-8f83-41d8-68f6797d0db4']/toolbar[@gu"
    s = s & "idRef='c2b44f69-6dec-444e-a37e-5dbf7ff43dae']"">" & vbLf & "    <xsl:copy-of selec"
    s = s & "t="".""/>" & vbLf & "    <xsl:if test=""not(../toolbar[@guidRef='772fbf90-5bc1-5c4d-"
    s = s & "93a0-d1fced7f98b5'])"">" & vbLf & "      <toolbar guidRef=""772fbf90-5bc1-5c4d-93a0"
    s = s & "-d1fced7f98b5""/>" & vbLf & "    </xsl:if>" & vbLf & "  </xsl:template>" & vbLf & "</xsl:stylesheet>" & vbLf & ""
    AddonText_UserUI_xslt = s
End Function

Private Function AddonText_Coreldrw_addon() As String
    Dim s As String
    AddonText_Coreldrw_addon = s
End Function

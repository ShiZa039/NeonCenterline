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
    WriteUtf8 tmp & "\config.xml", AddonText_config_xml()
    WriteBase64 tmp & "\Resources.dll", AddonBin_Resources_dll()
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

Private Sub WriteBase64(ByVal path As String, ByVal b64 As String)
    Dim doc As Object, el As Object, st As Object
    Set doc = CreateObject("MSXML2.DOMDocument")
    Set el = doc.createElement("b")
    el.DataType = "bin.base64"
    el.Text = b64
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write el.nodeTypedValue
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
    s = s & "                icon=""guid://4e6d43bf-0b2d-500c-b154-b553b6087991""" & vbLf & "   "
    s = s & "             dynamicCommand=""NeonCenterline.NeonUI.RunMono""" & vbLf & "          "
    s = s & "      dynamicCategory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "         "
    s = s & AU("20 20 20 20 20 20 20 75 73 65 72 43 61 70 74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 41C 43E 43D 43E 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20")
    s = s & "<itemData guid=""c68c5d4f-d668-5fb4-937e-adb7430cae58""" & vbLf & "                "
    s = s & "icon=""guid://c68c5d4f-d668-5fb4-937e-adb7430cae58""" & vbLf & "                dyn"
    s = s & "amicCommand=""NeonCenterline.NeonUI.RunScript""" & vbLf & "                dynamicC"
    s = s & "ategory=""2cc24a3e-fe24-4708-9a74-9c75406eebcd""" & vbLf & "                userCap"
    s = s & AU("74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 41A 443 440 441 438 432 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20 3C 69 74 65 6D 44 61 74 61 20 67 75")
    s = s & "id=""ce37ed0a-544f-575c-ac35-ef4d4aab66c9""" & vbLf & "                icon=""guid:/"
    s = s & "/ce37ed0a-544f-575c-ac35-ef4d4aab66c9""" & vbLf & "                dynamicCommand="
    s = s & """NeonCenterline.NeonUI.ShowSettings""" & vbLf & "                dynamicCategory="""
    s = s & AU("32 63 63 32 34 61 33 65 2D 66 65 32 34 2D 34 37 30 38 2D 39 61 37 34 2D 39 63 37 35 34 30 36 65 65 62 63 64 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 75 73 65 72 43 61 70 74 69 6F 6E 3D 22 41D 435 43E")
    s = s & AU("43D 3A 20 43D 430 441 442 440 43E 439 43A 438 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20 3C 69 74 65 6D 44 61 74 61 20 67 75 69 64 3D 22 31 30")
    s = s & "25a06f-bc40-5284-9292-2c3cf477d6a2""" & vbLf & "                icon=""guid://1025a"
    s = s & "06f-bc40-5284-9292-2c3cf477d6a2""" & vbLf & "                dynamicCommand=""NeonC"
    s = s & "enterline.NeonUI.CheckUpdates""" & vbLf & "                dynamicCategory=""2cc24a"
    s = s & AU("33 65 2D 66 65 32 34 2D 34 37 30 38 2D 39 61 37 34 2D 39 63 37 35 34 30 36 65 65 62 63 64 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 75 73 65 72 43 61 70 74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 43F 440 43E")
    s = s & AU("432 435 440 438 442 44C 20 43E 431 43D 43E 432 43B 435 43D 438 44F 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 20 20 3C 69 74 65 6D 44 61 74 61 20 67 75 69")
    s = s & "d=""c5cc733d-ee55-5173-b5ef-17a954845da7""" & vbLf & "                icon=""guid://"
    s = s & "c5cc733d-ee55-5173-b5ef-17a954845da7""" & vbLf & "                dynamicCommand="""
    s = s & "NeonCenterline.NeonUI.ShowAbout""" & vbLf & "                dynamicCategory=""2cc2"
    s = s & AU("34 61 33 65 2D 66 65 32 34 2D 34 37 30 38 2D 39 61 37 34 2D 39 63 37 35 34 30 36 65 65 62 63 64 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 75 73 65 72 43 61 70 74 69 6F 6E 3D 22 41D 435 43E 43D 3A 20 43E")
    s = s & AU("20 43F 440 43E 433 440 430 43C 43C 435 22 A 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 20 65 6E 61 62 6C 65 3D 22 74 72 75 65 22 2F 3E A 20 20 20 20 3C 2F 78 73 6C 3A 63 6F 70 79 3E A 20 20 3C 2F 78 73 6C 3A 74 65")
    s = s & "mplate>" & vbLf & "" & vbLf & "  <xsl:template match=""uiConfig/commandBars"">" & vbLf & "    <xsl:copy>" & vbLf & ""
    s = s & "      <xsl:apply-templates select=""node()|@*""/>" & vbLf & "      <commandBarData "
    s = s & "guid=""772fbf90-5bc1-5c4d-93a0-d1fced7f98b5""" & vbLf & "                      nonL"
    s = s & "ocalizableName=""NeonCenterline""" & vbLf & "                      userCaption=""Neo"
    s = s & "nCenterline""" & vbLf & "                      type=""toolbar"">" & vbLf & "        <toolbar>" & vbLf & " "
    s = s & "         <item guidRef=""4e6d43bf-0b2d-500c-b154-b553b6087991""/>" & vbLf & "      "
    s = s & "    <item guidRef=""c68c5d4f-d668-5fb4-937e-adb7430cae58""/>" & vbLf & "          <"
    s = s & "item guidRef=""ce37ed0a-544f-575c-ac35-ef4d4aab66c9""/>" & vbLf & "        </toolba"
    s = s & "r>" & vbLf & "      </commandBarData>" & vbLf & "    </xsl:copy>" & vbLf & "  </xsl:template>" & vbLf & "</xsl:sty"
    s = s & "lesheet>" & vbLf & ""
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

Private Function AddonText_config_xml() As String
    Dim s As String
    s = s & "<?xml version=""1.0""?>" & vbLf & "<config>" & vbLf & "  <resources>" & vbLf & "    <resource name=""e6794"
    s = s & "41a-cbd6-5d39-a310-03516f03525e"" path=""Resources.dll"">" & vbLf & "      <resource"
    s = s & "Map>" & vbLf & "        <resEntry id=""772fbf90-5bc1-5c4d-93a0-d1fced7f98b5"" strin"
    s = s & "g=""20000""/>" & vbLf & "        <resEntry id=""4e6d43bf-0b2d-500c-b154-b553b6087991"
    s = s & """ string=""20001"" icon=""101""/>" & vbLf & "        <resEntry id=""c68c5d4f-d668-5fb4"
    s = s & "-937e-adb7430cae58"" string=""20002"" icon=""102""/>" & vbLf & "        <resEntry id="""
    s = s & "ce37ed0a-544f-575c-ac35-ef4d4aab66c9"" string=""20003"" icon=""103""/>" & vbLf & "    "
    s = s & "    <resEntry id=""1025a06f-bc40-5284-9292-2c3cf477d6a2"" string=""20004"""
    s = s & " icon=""103""/>" & vbLf & "        <resEntry id=""c5cc733d-ee55-5173-b5ef-17a954845d"
    s = s & "a7"" string=""20005"" icon=""103""/>" & vbLf & "      </resourceMap>" & vbLf & "    </resource>" & vbLf & " "
    s = s & " </resources>" & vbLf & "</config>" & vbLf & ""
    AddonText_config_xml = s
End Function

Private Function AddonBin_Resources_dll() As String
    AddonBin_Resources_dll = AddonBinPart0() & AddonBinPart1() & AddonBinPart2()
End Function

Private Function AddonBinPart0() As String
    Dim s As String
    s = s & "TVqQAAMAAAAEAAAA//8AALgAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA4fug4AtAnNIbgB"
    s = s & "TM0hVGhpcyBwcm9ncmFtIGNhbm5vdCBiZSBydW4gaW4gRE9TIG1vZGUuDQ0KJAAAAAAAAABQRQAAZIYDANxX22oAAAAAAAAAAPAA"
    s = s & "LiILAgIpAAAAAABiAAAAAAAAAAAAAAAQAAAAAACAAQAAAAAQAAAAAgAABAAAAAAAAAAFAAIAAAAAAACQAAAABAAAQuoAAAIAYAEA"
    s = s & "ACAAAAAAAAAQAAAAAAAAAAAQAAAAAAAAEAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAIAAAGAAAAAAwAAAAYAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAudGV4dAAAACAAAAAAEAAAAAIAAAAEAAAAAAAAAAAAAAAAAAAgAABgLmlkYXRhAAAYAAAAACAAAAAC"
    s = s & "AAAABgAAAAAAAAAAAAAAAAAAQAAAwC5yc3JjAAAAAGAAAAAwAAAAYAAAAAgAAAAAAAAAAAAAAAAAAEAAAEAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////AAAAAAAAAAD//////////wAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMAAwAAACgAAIAGAAAAWAEAgA4AAACIAQCAAAAAAAAAAAAAAAAA"
    s = s & "AAAJAAEAAACAAACAAgAAAJgAAIADAAAAsAAAgAQAAADIAACABQAAAOAAAIAGAAAA+AAAgAcAAAAQAQCACAAAACgBAIAJAAAAQAEA"
    s = s & "gAAAAAAAAAAAAAAAAAAAAQAAAAAA+AEAAAAAAAAAAAAAAAAAAAAAAQAAAAAACAIAAAAAAAAAAAAAAAAAAAAAAQAAAAAAGAIAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAQAAAAAAKAIAAAAAAAAAAAAAAAAAAAAAAQAAAAAAOAIAAAAAAAAAAAAAAAAAAAAAAQAAAAAASAIAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAQAAAAAAWAIAAAAAAAAAAAAAAAAAAAAAAQAAAAAAaAIAAAAAAAAAAAAAAAAAAAAAAQAAAAAAeAIAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAQDjBAAAcAEAgAAAAAAAAAAAAAAAAAAAAQAAAAAAiAIAAAAAAAAAAAAAAAAAAAAAAwBlAAAAsAEAgGYAAADIAQCAZwAA"
    s = s & "AOABAIAAAAAAAAAAAAAAAAAAAAEAAAAAAJgCAAAAAAAAAAAAAAAAAAAAAAEAAAAAAKgCAAAAAAAAAAAAAAAAAAAAAAEAAAAAALgC"
    s = s & "AADIMgAAaAQAAAAAAAAAAAAAMDcAAIgJAAAAAAAAAAAAALhAAACoEAAAAAAAAAAAAABgUQAAaAQAAAAAAAAAAAAAyFUAAIgJAAAA"
    s = s & "AAAAAAAAAFBfAACoEAAAAAAAAAAAAAD4bwAAaAQAAAAAAAAAAAAAYHQAAIgJAAAAAAAAAAAAAOh9AACoEAAAAAAAAAAAAACQjgAA"
    s = s & "3AAAAAAAAAAAAAAAcI8AADAAAAAAAAAAAAAAAKCPAAAwAAAAAAAAAAAAAADQjwAAMAAAAAAAAAAAAAAAKAAAABAAAAAgAAAAAQAg"
    s = s & "AAAAAABABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVVVUGKysriEdFSOA8Oz3OJiYmTwAAAAB/f38Cf39/AgAAAAAfHx8gLCsunzk3"
    s = s & "O8YoKChyqqqqAwAAAAEAAAAALy8vh4+Qj//u8+r/2t7W/1paW/klJSU3AAAAAQEBAQAYGBgfSEhJ3rq9t//f5tn/h4eH/y0tLYAA"
    s = s & "AAAAAAAADk9PT+b////9uHT1/ejC//nM0sf/JyUpmwAAAAEiIiIPNzY5y9ja1P/54//6y5T9/P////1PT07mAAAADgAAABRcXFvq"
    s = s & "+/v8/38P4/7Nk//7y9TE/icjKo8AAAAALCotsry+vP///f/7hiHi/4UZ5v77+fz/XFxb6gAAABQAAAATW1pZ6Pz8/f+JIOX+2ar/"
    s = s & "/MDIuf8kIClvJSMle6Oio//////7m0Po/3QA4v+WN+v++ff8/1paWugAAAATAAAAE1paWen8/P3/hx3l/tWk//7Cy7v/Gxgd14SD"
    s = s & "hv3////9qF3r/4ge5v+9f/H/jSXo/vz6/f9aWlnpAAAAEwAAABNaWlnp/Pz9/4cd5f7ZqP//sbup/1lWXf////3/wIfy/3IA4f/q"
    s = s & "1/r/0ab1/4gd5/78+/3/WlpZ6QAAABMAAAATWlpZ6fz8/f+IHub+0aD6/9Xa0P/s8ef92LH5/m8A3//Wrfr//////8WT8v+KIOj+"
    s = s & "/Pv9/1paWekAAAATAAAAE1paWen8+/3/iiHo/saT8v//////5sz9/3EA4P/Ij/n/4+rd/szSx//Tovv/iB7m/vz8/f9aWlnpAAAA"
    s = s & "EwAAABNaWlnp/Pv9/4ge5/7OovT/9u/8/3kJ4v+waPD//f/4/k1KT/+zvKz/2Kf//4cd5f78/P3/WlpZ6QAAABMAAAATWlpZ6fz6"
    s = s & "/f+LIuj+xpHz/5Y56f+WOun//////XZ2dv8cGh/mwsq6/9al//+HHeX+/Pz9/1paWekAAAATAAAAE1paWen59/v/lTXq/nwN5P+N"
    s = s & "KOb/////+5aYlP8oJyqcKCQqeMHJuf/Xpv/8iiLm/vz8/P9XWVfoAAAAEQAAABNbW1rp/fz9/4kh5v55CuD//O3/+6+zq/8qKSzN"
    s = s & "kZGRByAeJH3Fzr3+05///HgC4P79+v3/aGln8QAAACYAAAADPz4/yevs5//lvf/69Nn/+8zRx/8zMTTkQUFBHwAAAAAcHB9zqKyk"
    s = s & "///o//nCgvn8/////V5eXfUAAAchAAAAACgoKFJiYmL9zNLG/7C1rf9BQELiLzMvNgAAAAAAAAACKCgeGURDRN66vbf/4OjZ/42N"
    s = s & "jf8vLi+QAQEAAAAAAAEAAAAAISEhRSooLaImJCiMLS0tHAAAAAAAAAADAAAAAQAAAAAqKiokLy0vojg2O8YoKCh4f38/BAAAAAHH"
    s = s & "5wAAg8EAAIGBAACBAQAAgwEAAIABAACAAQAAgAEAAIABAACAAQAAgAEAAIBBAACAwQAAgcEAAMPBAADn5wAAKAAAABgAAAAwAAAA"
    s = s & "AQAgAAAAAABgCQAAAAAAAAAAAAAAAAAAAAAAAAAAAAB/f38CAAAAAEZGRjYxMTGhKysr2iwsLME9PT1kAQEBAAAAAAAAAAABAAAA"
    s = s & "AAAAAAAAAAAAf39/AgAAAABVVVUhMTExhSsrK6kzMzOAUVFRHAAAAAAAAAACAAAAAAAAAAIAAAAAPz8/UDQ0NPFnZmf/m5qc/4GA"
    s = s & "gv9CQkL/Nzc3pV9fXwgAAAACAAAAAQAAAABVVVUDAAAAAD4+PloyMjLlWVhZ/3l4e/9VVVX/NTU130FBQUYAAAAAAAAAAgAAAABb"
    s = s & "W1sZMjIy7YqKiv/////6/////v////3Q0ND7R0dH/zo6OnEAAAAAVVVVAz8/PwQAAAAAPT09Rjk5Of+fn6D/+vz5+/////319vT7"
    s = s & "goKC/zExMeZdXV0TAAAAAAAAAAAzMzNnT05P//n6+Pv37/7/rWzr/9Oy8//////8l5aY/SsrK8y/v78EAAAAAwAAAABVVVUnLy8v"
    s = s & "65uam/7////89ez9/8qg8v/89v//9vf1+05OTv8zMzNnAAAAAAAAAAArKyuBZmRn/v////vUrvb/aADg/51K6P/9/f3+s7O0/yws"
    s = s & "LOSLi4sLAAAAAGNjYxctLi3Wf3+A//////37+fv/oEzr/2wA4f/Wsvb/////+2ZlZ/4rKyuBAAAAAAAAAAAqKip9ZGJl//////vV"
    s = s & "r/b/cATh/7Z17v/////8lJKV/ysrK8J/f38Ef39/DC0tLbtra2z/+ff6/P///v+zce//eADj/3sE4//Sqvb/////+2RhZf8qKip9"
    s = s & "AAAAAAAAAAAqKip+ZGJl//////vVr/b/bgDh/7Vz7v/////8kY+S/ysrK7YEBAQALy8vnFpaWv/s6u37/////8KM8v91AOP/fgbk"
    s = s & "/3kA4//Tq/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/bgDh/7Vz7v/////8kpGU/ywsLKs8PDxyR0dH/9zb"
    s = s & "3vv/////0631/3YF4/+KIOf/oEns/3MA4v/Urfb/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/bgDh/7Vz7v//"
    s = s & "///+kI6R/y4uLu46Ojr7wsHD/v/////jzPf/gBDl/3YE4v/auvf/wYjz/2wA4f/Vr/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+"
    s = s & "ZGJl//////vVr/b/bgDh/7Vz7v//////kY+S/yIiIv+wr7D//////+7h+f+MJ+j/cgDj/8GJ8///////tXLw/24A4f/Vr/b/////"
    s = s & "+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/bgDh/7Vz7f//////f32A/4WEhf7/////+PT8/5k96v9vAOH/tHDv////"
    s = s & "/v/+/f//uHbw/24A4f/Vr/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/bgDh/7d18P//////zMrO//Tz9f/+"
    s = s & "/f7/qFvs/24A4f+iUOz//fv////+////////uHbw/24A4f/Vr/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/"
    s = s & "bgDh/7h28P////////////////+6ffH/cADh/5Iy6P/06P3/7/Dt/7e0uf//////t3Xv/24A4f/Vr/b/////+2RiZf8qKip+AAAA"
    s = s & "AAAAAAAqKip+ZGJl//////vVr/b/bgDh/7h28P/+/f/////+/8ye9P90AOP/hx/m/+nX+f//////bm1w/4B+gf//////tXPu/24A"
    s = s & "4f/Vr/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/bgDh/7Zy8P//////2rj4/3oH5P98DOT/3cD2//////+Y"
    s = s & "l5n+Hh8e/5KRlP//////tXPu/24A4f/Vr/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl//////vVr/b/bQDh/7+C8v/u3/v/"
    s = s & "ghfl/3QA4//KnfH//////6yqrv4yMzL/MDAw+46NkP//////tXPu/24A4f/Vr/b/////+2RiZf8qKip+AAAAAAAAAAAqKip+ZGJl"
    s = s & "//////vUrvb/cQDi/65j7/+bQOr/bQDh/7yB7///////ysnM+zs7PP9ERESdKysrsJKQk//////9tXPu/24A4f/Vr/b/////+2Ri"
    s = s & "Zf8qKip/AAAAAAAAAAAqKip9ZGJk//////vUrPb/eQHj/4MP5v9zAOL/q2Lt///////d3d78SUhJ/zg6Osp/f38MKioqs5KQk///"
    s = s & "///8tXPu/24A4f/Wsfb/////+2JgY/8rKyt7AAAAAAAAAAArKyuAZmRn//////vQpfX/eADj/3kA5P+ZQOn//Pn+/+vq7P1YV1j/"
    s = s & "NDQ03lVVVSEAAAAALCwsuZKQk//////8tHLu/3AD4f/Sqfb//////Gxqbf8sLCyPAQEBAAAAAAAuLi5tWFdZ//z/+vzjy/j/ewfj"
    s = s & "/40p5//w5vr/+/v6/Wppa/4wMDDuUFBQNgAAAAD///8CKioquZORlP7////9t3nu/2QA3//Bi/D//////IiGif4qKiqsAgICAAAA"
    s = s & "AABFRUUsNDQ0/MfIx/z////+5M74//bt///////9hoWG/DAwMPxJSUlTAAAAAB8fHwgAAAAALy8vl2dnaP/////78+f9/8SV8P/3"
    s = s & "7v//////+l5dXv8vLy+FAAAAAAAAAAIBAQEANTU1pldXV//NzM39///8/O/x7f2Lioz/Nzc3/0JCQnYAAAAAMzMzBX9/fwIAAAAA"
    s = s & "RERENDU1NfyUlJT////9+/////37/fr7jIyM/zMzM/ZGRkYoAAAAAAAAAAEAAAACTU1NFzU1Naw7Ozv7WVha/0tLTP8xMTHcPz8/"
    s = s & "WAAAAAAAAAADAAAAAAAAAAAAAAACAAAAAD47Pl41NTXtWlpa/3h3ev9XVlf/NDQ06D4+PlEAAAAAAAAAAgAAAAAAAAABAAAAAAAA"
    s = s & "AABCQkJBLy8vdzY2NmJZWVkUAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAVVVVAwAAAABISEgqMTExhyoqKqkxMTGCTU1NJAAAAABV"
    s = s & "VVUDAAAAAPH/jwDgfwcAwH4DAMA8AwCAOAEAwDADAMAgAwDAIAMAwAADAMAAAwDAAAMAwAADAMAAAwDAAAMAwAADAMAAAwDAAAMA"
    s = s & "wAQDAIAMAQDAHAEAwDwBAMB+AwDg/wcA//+PACgAAAAgAAAAQAAAAAEAIAAAAAAAgBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AH9/fwIAAAAAVVVVA0hISFFBQUGjNzc3zT09PbdHR0dkPz8/FAAAAAAAAAADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA"
    s = s & "AAIAAAAAT09PHURERGhBQUGIQ0NDdUpKSjAAAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAABAAAAADc3NxdBQUGyNDQ0/zs7O/9L"
    s = s & "S0v/QEBA/zExMf9AQEDSPz8/OAAAAABVVVUDAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAEVFRWo7OzvaMjIy/zg4OP80NDT/"
    s = s & "Nzc37ENDQ4w8PDwRAAAAAAAAAAEAAAAAAAAAAQAAAAI2NjYOQEBAwTU1Nf96enr70dDS/PDv8f7d3N79lpaW+zk5Of89PT3oQkJC"
    s = s & "KgAAAAEAAAACAAAAAAAAAAAAAAABVVVVAwAAAAJCQkKZNTU1/09PT/+hoKH8ysjL/LCwsfthYWH/OTk5/0BAQLcfHx8IPz8/BAAA"
    s = s & "AAFVVVUDAAAAAEVFRW43Nzf/g4OD+f////7/////////////////////rq6v+Tk5Of9AQECrAAAAAAAAAAIAAAABAAAAADMzMwUA"
    s = s & "AAAAQ0NDcTU1Nf+AgID76ejp/f////////////////X19f54eHj6NDQ0/0RERGEAAAAAPz8/BAAAAAFVVVUGOTk5xkdHR/7x7/H8"
    s = s & "/////+7g+f+7ivD/4cj3//7//P/////9ampp/jMzM+pJSUkmAAAAAAAAAAI/Pz8EAAAAAExMTFM0NDT9ZmZm+/j4+P7/////+vr7"
    s = s & "/+DC9//z6vr//////+vr7PxHR0f/Ojo6x1VVVQYAAAABAAAAAGtraxMwMDDwbGxs///8//7+/f7/qV3t/2gA4P+NLOj/8OX5//7+"
    s = s & "/v+amZv8Ly8v/0REREcAAAAASEhIBwAAAABKSko6MzMz61JSUv7t7e7+/v3//////f+1cfH/cAPi/7Jx8P/+/v3//////mxsbP8w"
    s = s & "MDDwa2trEwAAAAAAAAAAcXFxEjAwMO1tbW3///3//v7+/v+eSev/dQTj/5U46v/07Pr//////46Oj/0wMDD8R0dHQAAAAAEAAAAA"
    s = s & "SkpKHzs7O+BGRkb/1NTU/f////////7/yZj0/3gA5P93AuP/nUXr//79/v///f/+bW1t/zAwMO1xcXESAAAAAAAAAABxcXESMDAw"
    s = s & "7W1tbf///f/+/v7+/6JO7P9wAOL/pVDs///////////9bm1u/zIyMuJNTU0kAAAAAEVFRRY9PT3GNDQ0/8fIx/z////////9/9q5"
    s = s & "+P95BeT/gw7m/3kB4/+gS+z//v7+///9//5tbW3/MDAw7XFxcRIAAAAAAAAAAHFxcRIwMDDtbW1t///9//7+/v7/oU3s/3AA4v+i"
    s = s & "S+z////+//////1wcHD/MjIy4E9PTyBVVVUDPz8/rTg4OP+vr6/7/////////f/jyvn/hBrm/30F5P+BC+b/eADj/59K7P/+/v7/"
    s = s & "//3//m1tbf8wMDDtcXFxEgAAAAAAAAAAcXFxEjAwMO1tbW3///3//v7+/v+hTez/cADi/6JM7P////7//////XBwcP8zMzPgTExM"
    s = s & "FEhISIUwMDD/kZGR+//////+//z/9On9/40p6P95AeP/gQ/m/4ga5/92AOP/n0rs//79/v///f/+bW1t/zAwMO1xcXESAAAAAAAA"
    s = s & "AABxcXESMDAw7W1tbf///f/+/v7+/6FN7P9wAOL/okzs/////v/////9cHBw/zAwMNdMTEx7LS0t/Hl5ef7//////fz9//r1/f+X"
    s = s & "O+r/ewLj/3UA4//Mn/X/q1vu/24A4f+hTez//v7+///9//5tbW3/MDAw7XFxcRIAAAAAAAAAAHFxcRIwMDDtbW1t///9//7+/v7/"
    s = s & "oU3s/3AA4v+iTOz////+//////9ubm7/NTU1+jg4OP9hYWH+8fHx/v38/f////7/rWTu/3cA4/95AeT/r2bu//////+iTez/cADi"
    s = s & "/6FN7P/+/v7///3//m1tbf8wMDDtcXFxEgAAAAAAAAAAcXFxEjAwMO1tbW3///3//v7+/v+hTez/cADi/6JM7P////7//////3Jy"
    s = s & "cf8qKir/TExM/evr6/7//////////7x/8f93AOP/eADk/59M7P/+/v////7//6FL7P9wAOL/oU3s//7+/v///f/+bW1t/zAwMO1x"
    s = s & "cXESAAAAAAAAAABxcXESMDAw7W1tbf///f/+/v7+/6FN7P9wAOL/okzs/////v//////ampr/zc3N//U1NT//////////v/Lm/T/"
    s = s & "fAnk/3oA5P+WOur/8+n9//78////////ok3s/3AA4v+hTez//v7+///9//5tbW3/MDAw7XFxcRIAAAAAAAAAAHFxcRIwMDDtbW1t"
    s = s & "///9//7+/v7/oU3s/3AA4v+iTOz////+//39/f9lZWX/r6+v//////////3/4MX5/30H5f97A+T/hhrm/+vZ+/////7//Pr/////"
    s = s & "//+iTez/cADi/6FN7P/+/v7///3//m1tbf8wMDDtcXFxEgAAAAAAAAAAcXFxEjAwMO1tbW3///3//v7+/v+hTez/cADi/6JN7P//"
    s = s & "////+vn6/9nZ2f/9/P/////+/+fQ+v+EGOb/ewPk/30I5f/gxPn////////////9+/7//////6JN7P9wAOL/oU3s//7+/v///f/+"
    s = s & "bW1t/zAwMO1xcXESAAAAAAAAAABxcXESMDAw7W1tbf///f/+/v7+/6FN7P9wAOL/ok3s///////+/P////////39/v/06f3/lDbp"
    s = s & "/3kA5P99CeT/ypr0///////18/f/uLi4//r5+///////okzs/3AA4v+hTez//v7+///9//5tbW3/MDAw7XFxcRIAAAAAAAAAAHFx"
    s = s & "cRIwMDDtbW1t///9//7+/v7/oU3s/3AA4v+iTez///////37///8+/3//v7//6BN7P94AOP/dwDj/7t/8f////z//////5uamv9e"
    s = s & "Xl3//////////v+iTOz/cADi/6FN7P/+/v7///3//m1tbf8wMDDtcXFxEgAAAAAAAAAAcXFxEjAwMO1tbW3///3//v7+/v+hTez/"
    s = s & "cADi/6JN7P///////Pr///79//+wae//egHk/3kA5P+wae///f39//////+2trb/Jycn/3BwcP/////////+/6JM7P9wAOL/oU3s"
    s = s & "//7+/v///f/+bW1t/zAwMO1xcXESAAAAAAAAAABxcXESMDAw7W1tbf///f/+/v7+/6FN7P9wAOL/ok3s//7+////////xI/z/3YA"
    s = s & "4/96AeP/mkHr//v5/f//////0dHS/0BAQP0tLS3+cHBx//////////7/okzs/3AA4v+hTez//v7+///9//5tbW3/MDAw7XFxcRIA"
    s = s & "AAAAAAAAAHFxcRIwMDDtbW1t///9//7+/v7/oU3s/3AA4v+gSOz//////8yf9P97BuT/egLj/44p6f/y5/r//////+Lh4v9DQ0P9"
    s = s & "Ozs7/zQ0NP9ubm7//////////v+iTOz/cADi/6FN7P/+/v7///3//m1tbf8wMDDtcXFxEgAAAAAAAAAAcXFxEjAwMO1tbW3///3/"
    s = s & "/v7+/v+hTuz/bgDh/6lZ7v/n0fr/gBDl/3sD4/+FGub/4sn4/////v/u7e//X19f/TY2Nv9GRkauMDAw3XBwcP/////+///+/6JM"
    s = s & "7P9wAOL/oU3s//7+/v///f/+bW1t/zAwMO1xcXESAAAAAAAAAABxcXESMDAw7W1tbf///f/+/v3+/6BM7P9zAOL/mDjq/5Ap6P96"
    s = s & "AOP/egfk/9m49/////z//fz+/3l5ePwuLi7/RUVFwEtLSyIyMjLZcHBx//////3///7/okzs/3AA4v+hTez//v7+///9//5tbW3/"
    s = s & "MDAw7WtraxMAAAAAAAAAAHFxcRIwMDDtbW1t///9//7+/v7/n0rs/3kA4/9/CeX/fQfl/3kD5P/Lm/T////+//////+KiYr9MjIy"
    s = s & "/0REROI1NTUYUVFRGTIyMuNwcHD//////f///v+iTOz/cADi/6FN7P/+/f7///3//mxsbP8wMDDsd3d3DwAAAAAAAAAAcXFxEjAw"
    s = s & "MO5tbW3///3//v7+/v+fSuz/eQLj/4YU5v91AOP/tnXw/////v//////q6us/jAwMP1BQUHrPDw8OwAAAABLS0siMjIy3nBwcP//"
    s = s & "///9///+/6NP7P9wAeL/o1Ls//7+/v/////+b29w/zAwMPJaWlofAAAAAAAAAABpaWkRMTEx6mxsbP///f/+//7+/55I7P90AOL/"
    s = s & "cwDi/6ha7v/8+fz//////729vv4wMC/8QUFB/z8/P1QAAAAAAAAAAU9PTyAyMjLfcnFy//////3///7/nELr/3QC4/+MIuj/9e77"
    s = s & "//////+amZr+Ly8v/0tLS0oAAAAAAAAAAVVVVQM6Ojq6SUlJ/uvq7P39/f3/27v3/4ks6P+mV+3/8ej5///////V1Nb/QUFB/Dw8"
    s = s & "PP9CQkJnAAAAACoqKgYAAAAAT09PHTMzM91paGn//Pv8/f7+/f+4ePH/bgDi/51D6//38vr///7//5SUlP4uLi7/S0tLRAAAAAA/"
    s = s & "Pz8EAAAAAENDQ3UzMzP/p6in+//////8//r/8+j6//3++///////7Ozt/1FRUfk3Nzf/Q0NDlwAAAABVVVUDAAAAAFVVVQNVVVUD"
    s = s & "Ojo6tj8/P/7c29z7//////f2+f/auPb/8un5/////////f3+V1dX/TQ0NOdVVVUSAAAAAH9/fwIAAAACRkZGHTo6OuFFRUX/xcXF"
    s = s & "+f////7//////////t/e4PxgYGD7ODg4/0JCQrMAAAADf39/AgAAAAEAAAAAVVVVAwAAAABCQkJVODg4/2BgYPvt7e39////////"
    s = s & "/////////Pz8/oeHhvk3Nzf/QkJCjQAAAAB/f38CAAAAAFVVVQMAAAAAPz8/WDs7O/k7Ozv/eHh5/5ubnPyGhof+RkZG/zY2Nv9E"
    s = s & "RESkNTU1GAAAAAEAAAACAAAAAAAAAAAAAAAAf39/AgAAAABBQUGZNzc3/1ZWVv+goKH8ycfK/LCvsfxqamr+NjY3/0BAQMZEREQP"
    s = s & "AAAAAgAAAAIAAAAAAAAAAFVVVQMAAAAAQ0NDOT4+PrgyMjLuLy8v/zAwMPg9PT3IQ0NDWwAAAAAAAAAAAAAAAQAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAABAAAAAQAAAAJERER6OTk54TIyMv84ODj/MzMz/zQ0NPBDQ0OePDw8FQAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAIAAAAAAAAAAExMTDJGRkZIS0tLPWZmZgoAAAAAf39/AgAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAABE"
    s = s & "REQlREREaEFBQYhERER0SEhINQAAAAAAAAAAf39/AgAAAAAAAAAA/H//f/Af/A/gD/AH4AfwB8AH4APAB8ADwAeAA8AHAAPABgAD"
    s = s & "wAQAA8AEAAPAAAADwAAAA8AAAAPAAAADwAAAA8AAAAPAAAADwAAAA8AAAAPAAAADwAAAA8AAIAPAAGADwADgA8AB4APAA+AD4APg"
    s = s & "A+AH8APwD/AH+D/8D////38oAAAAEAAAACAAAAABACAAAAAAAEAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVV"
    s = s & "VQMAAAAAKSkpPUhISeymqqL+tbyv/WBgYfosLCx/AAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABHh4eEUA/Qdnb"
    s = s & "3dj/893//eLC//75+/f+aWlp+hwcGD4AAAAAf39/AgAAAAAAAAABVVVVAz8/PwRISEgHAAAAARoYHWmlqaL//ez//HwJ4f9qAN3/"
    s = s & "16r9/Oru5P85NzvGAD8ABP///wEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAPUGu8ffq/bNp9P2dQ+z/wYry/4IV4/7//P/+fYB8"
    s = s & "5wAAACcBAQEAAAAAAn9/fwIlIiVgKiktoScnKWgSEhI2fH57yf/6/+qFGeb/wJfk/v///f+AEOT/3bP//M3WxvQfGyJwAAAAAFVV"
    s = s & "VQMwMDCnh4iG/9Tbzf+Ympf/KScr68HIuv/dtv7/fQrj/9rQ5P/u8un1lUHe/KlY8v70/OzoS0hMqQAAAAt/f38Cbm9t///7//zR"
    s = s & "of/7//n//IKDgv/U283pq1nz+Z1E6/zc4tjuusKz386m7/2DEOf+//f/2oCBfbEeHh4qAAAAAJCUjf/36P//bQDd/927+v/Q2cf+"
    s = s & "087W6IIP6PTUpPr9u8Oz5XJxcuz//P/+ghPl/tSn++u6wrTvIB0klDMzMwpZWFn0+f31/5k57f+xb+v/8vjo/smi6/iEEuj///b/"
    s = s & "/oqNiNcrKC3I6fLi/LZt9f6SMej////8/5ucmf80MzTCKycsmsLMuvbJl/T9iSHm//////2kT+z6plPw++/25/9HRUm8GhgadJOW"
    AddonBinPart0 = s
End Function

Private Function AddonBinPart1() As String
    Dim s As String
    s = s & "kP/97//8eALg/qlg7f//9f/9hYeE/woKCjNzdXL3/fj/9ogg5P6VO+j/gRDl/+LI+PejqZ7LJSMoXhwcHAlBQELj7vPq/9yu//qH"
    s = s & "KeT++OX//KGmnf9VVVUDMTAzuc/Uy//13P/5lDzq/dKd//r0+fD+SkhK2Q8PDxEAAAAAKCgoUmJhYv7b4df//f/6/9XX0/9KSkr2"
    s = s & "AAAAASgoKCZIR0jlys3G//j+8v/i6Nz/cnFy/ykpKWgBAQEAMzMzBQEBAQAeHh5EOjk8v1dXV+0/P0DFJysnQQAAAAIAAAAAKioq"
    s = s & "KjMyNaNLSkvUOjg8wSMjI1cBAQEAAAAAAgAAAAAAAAABAAAAAAAAAAAAAAATAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAABgAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAH9/fwL///8BAAAAAAEBAQAAAAACAAAAAAAAAAAAAAACVVVVAwAAAAD///8BVVVVAwAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAP4fAAD8DwAA/AcAAPgHAADYBwAAAAMAAAADAAAAAQAAAAAAAABAAACAwAAA"
    s = s & "gOAAAMHxAADj/wAA//8AAP//AAAoAAAAGAAAADAAAAABACAAAAAAAGAJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAABVVVUDAQEBADw8PIYyMjL3XFtc/pOSlf9vbXD+PT09/jQ0NMtNTU0uAAAAAH9/fwIAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVVVQMAAAAAPT09Xz09Pf+op6j+/P/5////////////1tbW/Wlpaf8w"
    s = s & "MDC/////BH9/fwIAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf39/AgAAAAFSUlIlMC8v7ZybnP7////+"
    s = s & "6tn6/7h67//RqPT//////+rp6vtPT1D/NTU1gQAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVV"
    s = s & "AwAAAAAuLi5+Y2Nk//v6+/zy5/r/jSbn/28A4f9xAOL/w5Hw//////67urz+MTEx9E9PTyAAAAAAf39/AgAAAAAAAAAAAAAAAAAA"
    s = s & "AAB/f38CPz8/BD8/PwQAAAADAAAAAHV1dQ0rKyvRrKut+/////21d+7/cADi/5xC6/+JHeb/gxPl//Hk+//////8Xl1e/i8vL4EA"
    s = s & "AAAAVVVVAwAAAAAAAAAAAAAAAQAAAAEAAAAAAAAAAAAAAAAAAAAAVVVVAzIyMkdMSkzk8PLv/fr0/v+HHOX/jijq//bs///Fj/T/"
    s = s & "bADh/7p/7//////+rauu6ycnJ5YDAwMAAAAAAQAAAAAAAAAAAAAAAAAAAABFRUU3Ly8vcTMzM2NQUFAjAAAAACsrK3qGg4m2////"
    s = s & "1tSv8/9zAOP/n1ff//j69v/h0fD/gx3g/40j6P/69f//8PHu8kdGR+QzMzNBAAAAAH9/fwIAAAAAW1tbDjc3N6U4ODj9WFdZ/1BQ"
    s = s & "Uf80NDTwQEBAfzExMde/vsD/////8Z9Q5v94AOf/xand/+jr5v//////plLw/3IA4v/ky/r+////+mxrbtwjIyNXAAAAAFVVVQMA"
    s = s & "AAAANzc3r05OTv/Gxsb+///8+/j69fyhoaD/NTU1/1NSU/Tt8Ov+7t78/4IQ5/+IGuf/2dHg/+np6ffk5uPysYTV+XUD5P+1de3+"
    s = s & "////27GwtMIxMTHFW1tbDgAAAAAAAAABNjY2/8LCw/3////+4sv4/+7g+//////+gYCB/WVjZv77/Pr4z6P2/3EA4v+tZ+v99ff1"
    s = s & "/6yqrue4u7TU5Nrp/YAK5/6LIub+/fz9+eXm5NA6OTqlQEBANwAAAAMAAAACTk1P/fn89v/p1fv/eQXj/5Y85//9/f3/wMDB/4yL"
    s = s & "jfnz9/DCnUro6XUA5f/KpO308Pfq125rcN7Gxcfy/////6NS6v56Bej/3Lz22f///516eHzDLS0tdwAAAAA/Pz8ETUxN//f69P/v"
    s = s & "3/3/ewTk/44n6P//////2drZ/7i7tv/w6vbzhhfm+YUW5f/68v//8vTx9zU0Neh0c3jw/////tGo9P9vAOH/tHHu/v3/+v+hoKP/"
    s = s & "KioqxEBAQFcAAAAANzc3/cbGx/7/////oEvr/3oH5v/bwvH/3eDb/9/i3P7FnufxcwDm/7Jv7f7/////r66x5C4uLr5ISEjS4+Ti"
    s = s & "+vv3/v+IG+X/fgzk/+vg9Pz6+vn9iomL/zc3N/89PT19Kioqrnl4e//////+yZ7v/24A4v+8g/D//v/9//////GjT+75dwbl/+DD"
    s = s & "+P/7/Pr8ampr9C4uLnkvLy/SlJOV/v////6/hvD/bwDi/7x98v///////////aKiov41NTX/MjIyYFFRU9fP0c726uHx/4gf5/+S"
    s = s & "Men/+vT///Tr+vx9C+P9jSjl//////zT09L/OTo6/0RERC02NjaITUxN//Pz8vz37/3/hhvm/3kA5P+hTuv/6tj6//7+/f9ZWFr+"
    s = s & "////AioqKrurqqz8////7bBx5/xwAOT/uHjw/69m7/92AOX/xJbt/uvy5NyQjZG/NTU1tQEBAQBQUFAmLy8v66Cfof7////+yZzw"
    s = s & "/38O5P9tAOH/z6b0//////9sa279AAAAADc3N4VQUFH/9vj09fHl+/+SMuj/bgDi/20A4f+SNOf+/v3+/+Hh4ew9Oz2vPDw8NwAA"
    s = s & "AAIAAAAANTU1e0hISP/h4OL6/////u/f/f/Xt/b//////+jo6P1FRUX/AAAAAVFRURwvMC/glJOV//////3s3Pn/rGbs/7Jx7f/p"
    s = s & "1/r//////Hx8ff8yMjLeZmZmDwAAAAB/f38Cf39/Bjc3N7NNTU7/r66v//j79fz////74eHg/GZmZv80NDTgAAAAAwAAAABDQ0NX"
    s = s & "MzQz/qCfoP7////6/////f////3////7rKys/zw8PP8/Pz9AAAAAAAAAAAIAAAABAAAAAFxcXAs1NTWBMjIy8lFQUv9kY2X/RURF"
    s = s & "/zQ0NM1KSkopAAAAAD8/PwQAAAAAPDw8YTo6OvxoaGn/mJeZ/5WUlv9mZWf/MzMz9T09PXwAAAAA////AQAAAAAAAAAAAAAAAQAA"
    s = s & "AAAAAAAATU1NITIyMmsqKiqDOzs7UlxcXAsAAAAAAAAAAAAAAAAAAAACAAAAAERERDwtLS2LLS0txiwsLL8wMDCXRkZGNgAAAAAA"
    s = s & "AAACAAAAAAAAAAAAAAAAAAAAAAAAAAF/f38CAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAA"
    s = s & "////A////wEAAAAAAAAAAFVVVQMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAT8/PwQ/Pz8EVVVVAwAAAAEAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAH9/fwJVVVUDAAAAAAAAAABVVVUDf39/AgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA/8B/AP/APwD/gB8A/4AfAP8ADwD/AA8A/wAPAIIADwAAAAcA"
    s = s & "AAAHAAAABwAAAAMAAAABAAAEAACABAAAgAYAAIAPAADADwAA4B+BAPA/9wD4f/8A////AP///wD///8AKAAAACAAAABAAAAAAQAg"
    s = s & "AAAAAACAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA"
    s = s & "AAA3NzcXQkJCuDAwMPozMzP/Z2dn/m5ubv9FRUX+Li4u/zc3N+5CQkJzMzMzCgAAAAMAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJ/f38CRERED0BAQL01NTX/dHR0/cjHyP/59/r/"
    s = s & "/Pv9/97d3/+np6f/Tk5O/j09Pf9CQkJoAAAAAD8/PwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVAwAAAABBQUF9NjY2/4eHhvr////////////////////////////////r6+v9T09P/DY2"
    s = s & "NuZQUFAmAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVVVQMA"
    s = s & "AAAARkZGQTk5OflWVlb7+vn6/v/////v5Pn/r2nu/6dX7f/ImvP/+vn6///////Jycn8QEA//z09Pb5EREQPAAAAAgAAAAEAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf39/AlVVVQM7OzuKPz8//9DQ0f3/////+PT8"
    s = s & "/5pA6/9yAOL/egDk/3EA4v/Ah/L////+//////+Xl5f6MjIy/0VFRV8AAAAAPz8/BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVVVQMAAAAASEhIIzAwMOVwcHD//Pz9/v///v/Ah/L/dQDj/4EM5f9/CeT/fAPk/4MT5v/r"
    s = s & "2fv//v7+/+zr7v5ISEj2Nzc3s39/fwYAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJVVVUDPz8/BAAAAAIAAAAA"
    s = s & "Pz8/BAAAAABERERpLy8v7bi3uPT/////8uf6/4wm6P96CeT/lDTp/9St9v+KIef/cwDi/69m7//+//3//////ouLi/8vLy/4S0tL"
    s = s & "PQAAAAB/f38CAAAAAAAAAAAAAAAAAAAAAAAAAAE/Pz8EAAAAAAAAAAAAAAAAAAAAAAAAAAN/f38Cf39/BDU1Na1XV1fv9fT3////"
    s = s & "///Rp/X/dwXk/4wj6P/XtPb//////7Rv8v91AeT/gxPm/+nW+f7+/v7/3Nzd5zs7O745OTlVAAAAAFVVVQMAAAAAAAAAAAAAAAB/"
    s = s & "f38CAAAAAAAAAABRUVEZSEhIQ0dHR0RSUlIlAAAAAAAAAABMTEwoMDAwmJ6enpf////V/P36+65j7/9yAOL/jSzn//vy///+/v//"
    s = s & "z7Lq/4kn4P96B+X/u33x//////7//f/xbm5u4i8vL7BmZmYPAAAAAAAAAAEAAAAAAAAAAAAAAAAvLy8QQ0NDizY2NtowMDD/MDAw"
    s = s & "/zQ0NOU/Pz+wS0tLIjo6OmRBQUH/0M/Q5////9bdzOz4ghLm/34I6P+jbdL/uby3//j2+v/4+vb/nkzo/3EB4/+bQOv/+fb7////"
    s = s & "//+ioqP/MzMz+kVFRUIAAAAAVVVVAwAAAAAAAAAANTU1E0BAQMM4ODj/WVlZ/5CQkPyYl5n8aWlo/zo6Ov8/Pz/gNDQ0zVpaWvPy"
    s = s & "8PP//////8iW9P93AeT/fgvl/+TM+v/3+fX//Pr////////HjPn/cgDj/4Yc5v/t3Pv/////9NbV2MU1NTWOQ0NDRAAAAAAAAAAC"
    s = s & "AAAAAAAAAABBQUGzNTU1/4aGhvry8vP9///////////9/P/+qamp+zs7O/8wMDD+kI+Q9Pn5+Pzz6/f9lzvr/3cC5P+LI+b/0MvV"
    s = s & "/sXFxP/y8PP109XS97iiyPKIIeD/dwDk/r6G7//7//Te////u2FhYds0NDTsUVFRFgAAAAAAAAABAAAAADMzM/9pamn9////////"
    s = s & "///28fr/8Ob5/////P//////mpqa/ikpKf7b2dv//////+zZ+/+GHOf/ewfj/76A9P//////2trc/42NjbOnp6nR5eni/KFK7P5x"
    s = s & "AOL/lzvp////////////oaKi2ioqKrJRUVEvAAAAAX9/fwIAAAAALi4u/be2t///////8OT5/5xH6/+JKef/4MX2//7+/v/b29z+"
    s = s & "R0dH/5STleH1+PHCv4Xy9ngA4/98BOT/yqrm8/H37uTQz9L5bm5u+v////3/////unzx/3cB4/6FFOn/5s776f///5T///9bT01P"
    s = s & "ijs7O5EAAAAAAAAAAwAAAAA0NDT/zMvN///////UsfT/dwHj/3IA4/+2du////7//9fW2f6bm5v/8PDy9fv9+8aVP+PgcgDf/ZUz"
    s = s & "6//y6vj38fTv0YiHisswMDDEu7u86/////3w4/v/gxLl/noB5v+/g/Hp////3vf2+PdoaGj/Li4umMzMzAUAAAAAVVVVAzAwMP2+"
    s = s & "vb///v7+/+ze+P+FFub/dQPk/72A8v//////8fDy/6Khov/DxsD89Oj9/Ywf7P91AeP/sGfv////////////fHt7+R8fH/OPkJD7"
    s = s & "//////37/f+iT+3/cwHi/5w/7v/v6vP/1NTT/5qZm/kuLi7cQUFBt0JCQioAAAAALi4u/5GRkf///////Pv8/6JO7P90AeL/mTzs"
    s = s & "//Lt9v/Hycb+5uPo/9Xaz/XBj+3legHm/oEO5f/lzvn+/v3+/9PS0+dERETRNzc32VBQUM3h3uL3/////9Wv9v94AOT/dwDj/sui"
    s = s & "7/76+vr/+vf7/3l5ef82Njb/Pz8/4kBAQDs4ODi/SkpK/dTS1f3/////2LP5/3cA5P9/C+b/2771/+/w7v/////+//3/96VR7/90"
    s = s & "AeL/ljbq//v4/P//////lJSU9SEhIaBRUVFqLS0t6qCfofz//v//7+P6/5Qw6f9zAeP/ok3u/////////////////qmpqfs4ODj/"
    s = s & "PT095UNDQ2IzMzP/trW3//Pz8//f0en/iiPk/3YB5P/Bi/L//////////fzkzffjgQ3l/HgA5f/Nn/b//////uPi5PxhYWH/NTU1"
    s = s & "zEZGRlA5OTn1UVJR/ubl5///////y57z/3gB5P9+BuT/xZDz//bu/f/+//z//////5eXl/4uLi7/Q0NDOTk5ObtxcXHRysnL/fb3"
    s = s & "8f+3dvH/dwDk/5U26v/m0fn//////8WN9f94AuT/gRHk/+PO9P3////+vLu9/zMzM/9DQ0NuAAAAAUBAQKYyMjL/tLS0+//////4"
    s = s & "9Pv/kS7o/3kA4/+AC+X/kDDo/97E9v/+/v7/3Nrd/z09Pf0AAAAANjY2cDw8PPfd3N78////8NK75/eFF+f/dwDk/7Zz8f/dv/j/"
    s = s & "jCbn/nkA5P+vZu7/4+fe7dva3dKLi4zJOjo630FBQUoAAAAARUVFSTc3N/RaWlr8+fj6//7+/f/TrPb/eAnk/3sC5P9qAOH/t3fw"
    s = s & "//7+/v/l4+f/TU1N/wAAAABFRUVCNDQ0/nl4eP37+/zu+Pj3+bFo8P91AOP/gQ7l/4IQ5v96AeT/hRfm//Hk/P/z9fH+x8bHuDAw"
    s = s & "MI0+Pj5uMzMzCjMzMwUAAAAAQkJCjTU1Nf+rq6v7//////v9+//iyPj/rWfu/6dc7f/u4/n//////9PT1P82Njb9AAAAAf///wE+"
    s = s & "Pj7INzc3/9XV1v7/////+fT7/6FN7P9wAOL/dwDk/3YD4/+6gfH////7/vv7/P+EhIT0MDAwy0hISDUAAAAAPz8/BP///wFGRkYo"
    s = s & "OTk52kRERP/Hx8f7///////////////////+///////9/f3/enp5/TExMf8AAAADAAAAAEpKSj41NTXvZWVl+/Ly8v3/////8eb6"
    s = s & "/8SP8v+7gfH/2bf3//n4+v//////s7Oz/TAvMP9BQUHVODg4CQAAAAEAAAAAVVVVAwAAAABCQkJFPDw89jg4OP+Kior73tzf/Pz7"
    s = s & "/f79/P/+6+vs/IKCgvo0NDT/QEBAwgAAAAEAAAACAAAAAEJCQpE2Njb/aGho+vr6+/7//////////////////////////+Li4vte"
    s = s & "Xl7+NjY25U1NTSEAAAAAf39/AgAAAAAAAAAAAAAAAwAAAAE+Pj4xPz8/vDU1Nf9AQED/ZWVk/2xsbP9HR0f/NjY2/0FBQbo9PT0Z"
    s = s & "AAAAAAAAAAE/Pz8EMzMzBUNDQ5g2Njb/b29v/ri3ufvl5Ob88fDy/dTT1vyjoqP6QUFB/zw8PP9BQUGFAAAAAFVVVQMAAAAAAAAA"
    s = s & "AAAAAAAAAAAAf39/AgAAAAA/Pz8MREREXT09Pa4xMTHrMDAw7Tw8PL9ERERzOjo6DQAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAEFB"
    s = s & "QZc3NzfzNDQ0/0ZGRv9NTU3/PDw8/zU1Nf87OzvcQkJCWAAAAAEAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAA"
    s = s & "AAAAAAAAAFpaWhFxcXESf39/AgAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAUdHRzJDQ0NxOjo6tjk5ObpAQECb"
    s = s & "RkZGbURERB4AAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAT8/PwQAAAACAAAAAAAAAAAAAAAAPz8/"
    s = s & "BAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8BAAAAAAAAAAAAAAABf39/AgAAAAAAAAAAAAAAAFVVVQMAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAABVVVUDVVVVAwAAAAEAAAABf39/Aj8/PwQAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD//AP/"
    s = s & "//gB///4AP//8AB//+AAf//gAD//4AA//8AAP//AAB/AwAAfgAAAHwAAAA8AAAAPAAAAJwAAAAcAAAADAAAAAQAAIACAACAAgABg"
    s = s & "AMAAcADAAPAAwAD4AOAA/ADgAf4B8AH/h/gH////H////////////////////////ygAAAAQAAAAIAAAAAEAIAAAAAAAQAQAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAE/Pz8EPz8/BAAAAAEAAAAAPz8/MD09PToAAAAAAAAAAT8/PwQ/Pz8EAAAAAQAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAPT09Tz4+PvM+Pj79PDw8aQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAB"
    s = s & "SEhIBz09PYA/Pz+QKioqBjw8PJdBQUH/Pz8//z09PbY/Pz8IPj4+jz09PYBISEgHAAAAAQAAAAE/Pz8EAAAAAD09PYFDQ0P/QUFB"
    s = s & "/zw8PMU9PT3VQEBA/kBAQP89PT3fPDw8zUFBQf9DQ0P/PT09gQAAAAA/Pz8EPz8/BH9/fwI9PT2SQUFB/zo6Ovg/Pz//Pj4+/zAv"
    s = s & "MP8uLS//PT09/z8/P/86Ojr4QUFB/z09PZJVVVUDPz8/BAAAAAEAAAAAPz8/BD09PcA/Pz//Ojo6+jg3OP16fXf/goZ+/z09Pf44"
    s = s & "ODj6QEBA/z09PcszMzMFAAAAAAAAAAEAAAAAPj4+Uj8/P7A8PDzWPj4+/zc2N/65vLb/9+j///Xi///Mzsj/PTw9/j09Pf89PT3g"
    s = s & "Pz8/sj09PWcAAAAAPT09ITw8PPVAQED/Pz8//zIxM/9qbGj//fP//4QW5P93At//+eX//4OIgP8uLS//QEBA/z8/P/8+Pj7/PT09"
    s = s & "OkJCQhs8PDzoQUFB/z8/P/8zMzT/ZGVi//31//+RLun/gxjk//vs//98f3n/Ly8w/0BAQP9AQED/PT09+D09PTIAAAAAPz8/PEBA"
    s = s & "QJM9PT3IPz8//zQzNP2kpqH/+vP///ju//+4u7b/ODc4/T4+Pv89PT3VPz8/lTw8PEwAAAAAAAAAAgAAAAAAAAAAPj4+tz8/P/87"
    s = s & "Ozv6NDM0/mNkYf9oamb/NzY3/jo6Ovo/Pz//PDw8xQAAAAAAAAAAAAAAAj8/PwQAAAACPj4+k0FBQf86Ojr4Pz8//z8/P/8zMzT/"
    s = s & "MjIz/z8/Pv8/Pz//Ojo6+EFBQf8+Pj6Tf39/Aj8/PwQ/Pz8EAAAAAD09PYFDQ0P/QUFB/zw8PLk8PDzPPz8//j8/P/49PT3VPT09"
    s = s & "wEFBQf9DQ0P/PT09gQAAAAA/Pz8EAAAAAQAAAAFISEgHPT09gEBAQI9ISEgHPDw8skBAQP9AQED/PT09skhISAc+Pj6PPT09gEhI"
    s = s & "SAcAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD8/P1A+Pj7rPj4+6z8/P1AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAE/Pz8EPz8/BAAAAAEAAAAAQUFBH0FBQR8AAAAAAAAAAT8/PwQ/Pz8EAAAAAQAAAAAAAAAA//8AAP5/AADkJwAA"
    s = s & "wAMAAMADAADgBwAAwAMAAIABAACAAQAAwAMAAOAHAADAAwAAwAMAAOQnAAD+fwAA//8AACgAAAAYAAAAMAAAAAEAIAAAAAAAYAkA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAEzMzMFAAAAAAAA"
    s = s & "AAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVQM/Pz8EAAAAAQAAAAIA"
    s = s & "AAAAPT09Qj8/P8BAQEDKPT09ZAAAAAAAAAABAAAAAT8/PwRVVVUDAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB"
    s = s & "AAAAAAAAAAAAAAAAAAAAAH9/fwI9PT0yPj4+9T8/P/8+Pj7/QEBA/zw8PFV/f38CAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAEAAAAAPz8/CD4+Pl4/Pz9pPz8/HAAAAAA7OztvQUFB/zs7O/k7Ozv5Pz8//j4+PpwAAAAAPz8/HD8/"
    s = s & "P2k+Pj5ePz8/CAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAQAAAAFEREQPPT09tkFBQf9BQUH/Pj4+1zw8PCI+Pj5vPz8//js7O/w7"
    s = s & "Ozv8QEBA/z4+PpQ8PDwqPT091kFBQf9BQUH/PT09tkRERA8AAAABAAAAAQAAAAAAAAAAPz8/BAAAAAA9PT1gQUFB/zs7O/o7Ozv7"
    s = s & "Pj4+/z09Peo9PT30PT09/zw8PP88PDz/PDw8/zw8PPg9PT32Pj4+/zs7O/s7Ozv6QUFB/z09PWAAAAAAPz8/BAAAAAAAAAAAPz8/"
    s = s & "BAAAAAA7OztsQUFB/zs7O/s8PDz/Ozs7/T09Pf88PDz/Pj4+/z8/P/9AQED/Pj4+/zw8PP88PDz/PDw8/Tw8PP87Ozv7QUFB/zs7"
    s = s & "O2wAAAAAPz8/BAAAAAAAAAAAAAAAAVVVVQM9PT0dPT092D8/P/87Ozv+PDw8/zw8PP4+Pj7+NTU1/y4uLv8tLS3/MjIy/z4+Pv88"
    s = s & "PDz+PDw8/zs7O/4+Pj7/PT091j09PR0/Pz8EAAAAAQAAAAAAAAAAAAAAAQAAAAAAAAAANTU1Ez4+Pt09PT3/Ozs7/j4+Pv8yMjL/"
    s = s & "Xl1f/6Ghof+oqKj/bWxu/zIyMv88PDz/PDw8/jw8PP89PT3yPz8/HAAAAAAAAAAAAAAAAQAAAAAAAAABAAAAADs7OzxAQECXPDw8"
    s = s & "jDs7O+s9PT3/Pj4+/jMzM/93dnf/+/74/////////////////5SUlP8zMzP/Pj4+/jw8PP88PDz4PDw8mD8/P5g8PDxUAAAAAAAA"
    s = s & "AAEAAAAAOTk5PkBAQPNBQUH/QEBA/z09Pf89PT3/Nzc3/1FQUv/5+Pj/8OH9/5xI6P+SO+b/48z4//////9tbG7/MjIy/z4+Pv88"
    s = s & "PDz/QEBA/0BAQP9AQED/Ozs7YgAAAAAAAAAAPT09pkBAQP87Ozv3Ozs7+zw8PP4/Pz//Ly4v/4KBg///////q2Ls/3MA4v93AOP/"
    s = s & "kzbm//7+/v+qqar/LS0u/z8/P/88PDz/Ozs7/Dw8PPk+Pj79PT090DMzMwUAAAAAPDw8m0FBQf87Ozv4Ozs7/Dw8PP8/Pz//Ly8w"
    s = s & "/3x7ff//////tXbu/3IA4v90AOL/nUno//////+jo6T/LS0t/z8/P/88PDz/Ozs7/Ts7O/k/Pz//Pj4+xT8/PwQAAAAAPDw8Jj8/"
    s = s & "P9xBQUH/QEBA/z09Pf89PT3/OTk5/0dGR//p6ej//fb//7V27v+qY+z/8eP+//r6+v9eXV7/NTU0/z4+Pv88PDz/Pj4+/0FBQf9A"
    s = s & "QEDvPDw8QwAAAAAAAAABAAAAADs7Ox49PT1sPT09YDw8POI+Pj7/Pj4+/jQ1NP9cW1z/5ufk////////////9vjz/3d3d/8xMTH/"
    s = s & "Pj4+/jw8PP88PDz0PDw8bj4+Pm87OzsvAAAAAAAAAAEAAAAAAAAAAQAAAAAAAAAAREREDz4+PtA9PT3/Ozs7/T4+Pv80NDT/R0ZH"
    s = s & "/3p5e/+BgIL/UlFT/zMzM/8+Pj7/Ozs7/j09Pf89PT3nNTU1GAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAQAAAAI7OzsePj4+2D8/"
    s = s & "P/87Ozv+PDw8/zs7O/0+Pj7+OTk5/zAwMP8vLy//Nzc3/z4+Pv47Ozv+PDw8/zs7O/4+Pj7/PT091zs7Ox5VVVUDAAAAAQAAAAAA"
    s = s & "AAAAPz8/BAAAAAA7OztsQUFB/zs7O/o8PDz/Ozs7/T09Pf89PT3/PT09/z8/P/8/Pz//PT09/z09Pf89PT3/Ozs7/Tw8PP87Ozv7"
    s = s & "QUFB/zs7O2wAAAAAPz8/BAAAAAAAAAAAPz8/BAAAAAA9PT1gQUFB/zs7O/o8PDz6Pz8//zw8PNI8PDzfPT09/zw8PP48PDz+PT09"
    s = s & "/z09Pek9PT3hPz8//zs7O/s7Ozv6QUFB/z09PWAAAAAAPz8/BAAAAAAAAAAAAAAAAQAAAAFEREQPPT09tkFBQf9BQUH/Pz8/1j8/"
    s = s & "Pxg7OzuEQEBA/js7O/s7Ozv7QEBA/jw8PIo/Pz8cPz8/1kFBQf9BQUH/PT09tkRERA8AAAABAAAAAQAAAAAAAAAAAAAAAAAAAAEA"
    s = s & "AAAAPz8/CD4+Pl4/Pz9pODg4GwAAAAA9PT2dQEBA/zs7O/c7Ozv3Pz8//zw8PJsAAAAAODg4Gz8/P2k+Pj5ePz8/CAAAAAAAAAAB"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAI9PT02Pj4+80BAQP9AQED/Pj4+8z09PTYAAAACAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVQM/Pz8EAAAAAQAAAAEAAAAAPDw8N0BA"
    s = s & "QJ9AQECfPDw8NwAAAAAAAAABAAAAAT8/PwRVVVUDAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8A"
    s = s & "/+f/AP/D/wD/wf8A8MEPAPAADwDwAA8A8AAPAPgAHwDgAAcAwAADAIAAAQCAAAEAwAADAPgAHwD4AB8A8AAPAPAADwDwAA8A8IEP"
    s = s & "AP+B/wD/w/8A/+f/AP///wAoAAAAIAAAAEAAAAABACAAAAAAAIAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAAAAAAAAAAAAAAAf39/AgAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAB/f38CAAAAAD09PSE+Pj53QEBAfz4+Pj0AAAAA////AQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVVVUDVVVVAwAAAAIAAAAAAAAAAgAAAAA8PDw7Pz8/5EBA"
    s = s & "QP9AQED/Pz8//zw8PGoAAAAA////AQAAAAAAAAACVVVVA1VVVQMAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAMAAAABNTU1GD8/P9k+Pj7/Ozs7+js7O/o8PDz9Pz8//Ts7OzgAAAAA"
    AddonBinPart1 = s
End Function

Private Function AddonBinPart2() As String
    Dim s As String
    s = s & "MzMzBQAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf39/AgAAAABISEgHOjo6"
    s = s & "PT09PUI3NzcXAAAAAFVVVQM7OztJPz8//zs7O/s8PDz/PDw8/zw8PPpAQED/Pz8/gVVVVQMAAAAANzc3Fz09PUI6Ojo9SEhIBwAA"
    s = s & "AAB/f38CAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAPT09GT4+Prk+Pj79Pz8//z8/P9g9PT0+AAAAADg4ODI+"
    s = s & "Pj71PDw8/jw8PP48PDz/Ozs7+0BAQP4+Pj5iAAAAAD4+Pj0/Pz/YPz8//z4+Pv0+Pj65PT09GQAAAAAAAAABAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAABAAAAATMzMwo+Pj68QUFB/zw8PPs7Ozv7Pz8//z4+PuY6OjpOPDw8ij09Pfc8PDz/PDw8/jw8PP88PDz+Pj4+"
    s = s & "/z4+PqQ+Pj5iPj4+5j8/P/87Ozv7PDw8+0FBQf8+Pj68MzMzCgAAAAEAAAABAAAAAAAAAAAAAAAAAAAAAFVVVQMAAAAAOjo6PT4+"
    s = s & "Pv08PDz7PDw8/jw8PP87Ozv8PT09/z4+Pv9AQED/PDw8/zw8PP88PDz/PDw8/zw8PP88PDz/Pz8//0BAQP89PT3/Ozs7/Dw8PP88"
    s = s & "PDz+PDw8+z4+Pv06Ojo9AAAAAFVVVQMAAAAAAAAAAAAAAAAAAAAAPz8/BAAAAAA8PDxEPz8//zs7O/s8PDz/PDw8/zw8PP88PDz+"
    s = s & "Ozs7/Ts7O/w8PDz/PDw8/zw8PP89PT3/PDw8/zw8PP87Ozv8Ozs7+zw8PP48PDz/PDw8/zw8PP87Ozv7Pz8//zw8PEQAAAAAPz8/"
    s = s & "BAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAADU1NRg/Pz/aPz8//zs7O/w8PDz/PDw8/zw8PP88PDz+PDw8/z09Pf9AQED/Pj4+/z09"
    s = s & "Pf8/Pz//Pj4+/zw8PP88PDz/PDw8/zw8PP88PDz/Ozs7/T8/P/8/Pz/aNTU1GAAAAAAAAAACAAAAAAAAAAAAAAAAAAAAAAAAAAA/"
    s = s & "Pz8Ef39/Ajo6Oj0/Pz/mPDw8/zw8PP48PDz/PDw8/zw8PP8/Pz//Ojo6/y8vL/8zMzP/NDQ0/y4uLv81NTX/Pz8//zw8PP88PDz/"
    s = s & "PDw8/zw8PP89PT3/Pj4+4z09PT5/f38CMzMzBQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAD09PTo9PT3zPDw8"
    s = s & "/jw8PP48PDz/Pz8//zY2Nv9DQ0P/k5KT/8bEx//Lycz/p6eo/1VVVf8yMjL/Pz8//zw8PP88PDz/Ozs7/EBAQP88PDxYAAAAAAAA"
    s = s & "AAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAIAAAAAPDw8JkBAQHM9PT1wPDw8gj4+Pv08PDz9PDw8/z4+Pv84ODj/RERE/9HR0f//"
    s = s & "////////////////////6urr/2BgYP8xMTH/Pz8//zw8PP87Ozv8Pz8//zw8PKg9PT1oQUFBfDo6OjkAAAAAAAAAAAAAAAB/f38C"
    s = s & "AAAAAD09PTo+Pj7nQEBA/0BAQP8/Pz//PDw8/zw8PP89PT3/PDw8/zw8PP/R0dH///////j1+v/ImvL/vYXw//Hn+f//////7u3u"
    s = s & "/1RUVP81NTX/Pj4+/zw8PP88PDz/Pj4+/0BAQP9BQUH/QEBA+Dw8PGkAAAAAf39/AgAAAAA/Pz8QOzs74z8/P/87Ozv6Ozs7+zw8"
    s = s & "PPw8PDz/PDw8/z8/P/8wMDD/eXh5///////8+/z/pVTt/3MA4/91AOP/jino//Hm+f//////qamq/y0tLf9AQED/PDw8/zw8PP88"
    s = s & "PDz9Ozs7+zo6Ovs8PDz/Pj4+/zs7OzwAAAAAAAAAADo6OkY+Pj7+Ojo6/Dw8PP88PDz/PDw8/zw8PP88PDz/QEBA/y0tLf+enZ//"
    s = s & "/v7+/9zA9f95A+T/ggzl/4QR5f91AOP/vILv//7+/v/Mys7/NDQ0/z09Pf89PT3/PDw8/zw8PP88PDz/PDw8/zs7O/o/Pz/+Ozs7"
    s = s & "hAAAAAAAAAAAPz8/PD4+Pv86Ojr8PDw8/zw8PP48PDz+PDw8/zw8PP9AQED/LCws/5qZmv/+/v7/5c/3/30K5f99B+T/gg3l/3MA"
    s = s & "4//ImfH//f39/8jGyf8zMzP/Pj4+/z09Pf88PDz/PDw8/jw8PP48PDz/Ojo6+UFBQf87Ozt4AAAAAP///wEzMzMFPj4+wUFBQf87"
    s = s & "Ozv7PDw8/jw8PP88PDz/PDw8/z8/P/8yMjL/Z2dn//////////3/vIHx/3wJ5P96BeT/pVTt//n3+v//////l5aX/y0tLf9AQED/"
    s = s & "PDw8/zw8PP88PDz/PDw8/js7O/s/Pz//Pj4+6T4+PiUAAAAA////AQAAAAA9PT0ZQEBAuT4+Pv49PT39PT099T09Pf88PDz/PDw8"
    s = s & "/z09Pf80NDT/rKys////////////5c/3/9zB9v/+/vz//////9TU1P9CQkL/Ojo6/z09Pf88PDz+PT09/z09Pfg9PT33Pz8//0BA"
    s = s & "QNY+Pj45AAAAAH9/fwIAAAAAf39/AgAAAAAzMzMKPDw8Pzw8PD88PDxdPj4++zw8PP08PDz/PT09/zs7O/80NDT/q6ur//39/v//"
    s = s & "///////////////MzM3/RUVE/zY2Nv8+Pj7/PDw8/zo6Ovw/Pz//Pj4+hzo6Ojk9PT1HPDw8FQAAAAAAAAACAAAAAAAAAAAAAAAA"
    s = s & "AAAAAQAAAAAAAAAAAAAAADw8PC49PT3lPT09/zs7O/08PDz/PT09/zs7O/80NDT/ZWVl/5eXmP+dnJ3/d3d3/z09Pf83Nzf/Pz8/"
    s = s & "/zw8PP88PDz+PDw8/T09Pf05OTlLAAAAAAAAAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVA39/fwI9PT0+Pz8/5Dw8"
    s = s & "PP88PDz/PDw8/zw8PP88PDz/PT09/z4+Pv8zMzP/LS0t/y0tLP8wMDD/PDw8/z4+Pv88PDz/PDw8/zw8PP88PDz+PDw8/z4+PuM5"
    s = s & "OTk+f39/AlVVVQMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAANTU1GD8/P9o/Pz//Ozs7/Dw8PP88PDz/PDw8/zw8PP08PDz+"
    s = s & "PDw8/z8/P/9AQED/QEBA/z8/P/89PT3/PDw8/zw8PP48PDz/PDw8/zw8PP87Ozv9Pz8//z8/P9o1NTUYAAAAAAAAAAIAAAAAAAAA"
    s = s & "AAAAAAAAAAAAPz8/BAAAAAA8PDxEPz8//zs7O/s8PDz/PDw8/zw8PP88PDz+PT09/zw8PP48PDz/PDw8/zw8PP88PDz/PDw8/zw8"
    s = s & "PP88PDz+PDw8/zw8PP48PDz/PDw8/zw8PP87Ozv7Pz8//zw8PEQAAAAAPz8/BAAAAAAAAAAAAAAAAAAAAABVVVUDAAAAADo6Oj0+"
    s = s & "Pj79PDw8+zw8PP48PDz/Ozs7/D09Pf88PDznPDw8/Dw8PP88PDz/PDw8/zw8PP88PDz/PDw8/z4+Pv89PT33PT09/zs7O/w8PDz/"
    s = s & "PDw8/jw8PPs+Pj79Ojo6PQAAAABVVVUDAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABMzMzCj4+PrxBQUH/PDw8+zs7O/s/Pz//Pj4+"
    s = s & "6Tc3Ny46OjptPj4+/zw8PP08PDz/PDw8/zw8PP0/Pz//PDw8fzk5OT49PT3oPz8//zs7O/s8PDz7QUFB/z4+PrwzMzMKAAAAAQAA"
    s = s & "AAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAPT09GT4+Prk+Pj79Pz8//z8/P9g+Pj49AAAAAD4+PnI/Pz/+Ozs7+zw8PP88"
    s = s & "PDz/Ozs7+z8/P/4/Pz9wAAAAAD4+Pj0/Pz/YPz8//z4+Pv0+Pj65PT09GQAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAH9/fwIAAAAASEhIBzo6Oj09PT1CNzc3FwAAAAAAAAACPT09dUBAQP87Ozv6PDw8/zw8PP87Ozv6QEBA/zw8PHZ/f38CAAAA"
    s = s & "ADc3Nxc9PT1COjo6PUhISAcAAAAAf39/AgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAA"
    s = s & "AAAAAAAAPz8/BAAAAAE5OTkfPj4+30BAQP87Ozv8Ozs7/EBAQP8+Pj7fOTk5HwAAAAE/Pz8EAAAAAAAAAAAAAAAAAAAAAAAAAAEA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAFVVVUDVVVVAwAAAAIAAAAAAAAAAgAAAAA8PDw3"
    s = s & "Pz8/0j4+Pv0+Pj79Pz8/0jw8PDcAAAAAAAAAAgAAAAAAAAACVVVVA1VVVQMAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAAA6OjoNQUFBQkFBQUI6OjoNAAAAAAAA"
    s = s & "AAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVAwAAAAAAAAAAAAAAAAAAAABVVVUDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP////////////w////4H///+A///DgcP/gQCB/4AAAf+AAAH/gAAB/8AAA//gAA"
    s = s & "f/wAAD/gAAAHwAAAA8AAAAHAAAADwAAAA+AAAAf+AAA//gAAf/wAAD/4AAAf+AAAH/gAAB/4GBgf/DgcP//4H///+B////w/////"
    s = s & "////////DgBOAGUAbwBuAEMAZQBuAHQAZQByAGwAaQBuAGUACgAdBDUEPgQ9BDoAIAAcBD4EPQQ+BAwAHQQ1BD4EPQQ6ACAAGgRD"
    s = s & "BEAEQQQ4BDIEDwAdBDUEPgQ9BDoAIAA9BDAEQQRCBEAEPgQ5BDoEOAQaAB0ENQQ+BD0EOgAgAD8EQAQ+BDIENQRABDgEQgRMBCAA"
    s = s & "PgQxBD0EPgQyBDsENQQ9BDgETwQRAB0ENQQ+BD0EOgAgAD4EIAA/BEAEPgQzBEAEMAQ8BDwENQQAAAAAAAAAAAAAAAAAAAAAAAAA"
    s = s & "AAAAAAAAAAEAAwAQEAAAAQAgAGgEAAABABgYAAABACAAiAkAAAIAICAAAAEAIACoEAAAAwAAAAEAAwAQEAAAAQAgAGgEAAAEABgY"
    s = s & "AAABACAAiAkAAAUAICAAAAEAIACoEAAABgAAAAEAAwAQEAAAAQAgAGgEAAAHABgYAAABACAAiAkAAAgAICAAAAEAIACoEAAACQA="
    AddonBinPart2 = s
End Function

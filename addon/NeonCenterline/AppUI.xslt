<?xml version="1.0"?>
<!--
  NeonCenterline — команды и панель инструментов (CorelDRAW Addons).
  Команды вызывают макросы проекта NeonCenterline.gms (пользовательская папка GMS).
  GUID не менять: по ним CorelDRAW узнаёт кнопки в настроенных рабочих пространствах.
-->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data">
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  <frmwrk:uiconfig>
    <frmwrk:applicationInfo userConfiguration="true" />
  </frmwrk:uiconfig>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/items">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <itemData guid="4e6d43bf-0b2d-500c-b154-b553b6087991"
                icon="guid://4e6d43bf-0b2d-500c-b154-b553b6087991"
                dynamicCommand="NeonCenterline.NeonUI.RunMono"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: Моно"
                enable="true"/>
      <itemData guid="c68c5d4f-d668-5fb4-937e-adb7430cae58"
                icon="guid://c68c5d4f-d668-5fb4-937e-adb7430cae58"
                dynamicCommand="NeonCenterline.NeonUI.RunScript"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: Курсив"
                enable="true"/>
      <itemData guid="ce37ed0a-544f-575c-ac35-ef4d4aab66c9"
                icon="guid://ce37ed0a-544f-575c-ac35-ef4d4aab66c9"
                dynamicCommand="NeonCenterline.NeonUI.ShowSettings"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: настройки"
                enable="true"/>
      <itemData guid="1025a06f-bc40-5284-9292-2c3cf477d6a2"
                icon="guid://1025a06f-bc40-5284-9292-2c3cf477d6a2"
                dynamicCommand="NeonCenterline.NeonUI.CheckUpdates"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: проверить обновления"
                enable="true"/>
      <itemData guid="c5cc733d-ee55-5173-b5ef-17a954845da7"
                icon="guid://c5cc733d-ee55-5173-b5ef-17a954845da7"
                dynamicCommand="NeonCenterline.NeonUI.ShowAbout"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: о программе"
                enable="true"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/commandBars">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <commandBarData guid="772fbf90-5bc1-5c4d-93a0-d1fced7f98b5"
                      nonLocalizableName="NeonCenterline"
                      userCaption="NeonCenterline"
                      type="toolbar">
        <toolbar>
          <item guidRef="4e6d43bf-0b2d-500c-b154-b553b6087991"/>
          <item guidRef="c68c5d4f-d668-5fb4-937e-adb7430cae58"/>
          <item guidRef="ce37ed0a-544f-575c-ac35-ef4d4aab66c9"/>
        </toolbar>
      </commandBarData>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>

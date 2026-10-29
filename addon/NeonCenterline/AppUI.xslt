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
                userCaption="Неон: Курсив (бета)"
                userToolTip="Неон: Курсив — в разработке, работает нестабильно. Проверяйте линии перед резкой."
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
      <itemData guid="0b33ce77-31ad-5c46-bb8b-2cd2e362084f"
                icon="guid://0b33ce77-31ad-5c46-bb8b-2cd2e362084f"
                dynamicCommand="NeonCenterline.NeonTools.MeasureLength"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: длина контуров и линий"
                enable="true"/>
      <itemData guid="c913505c-324e-543b-b2b1-a55c9bcbc9eb"
                icon="guid://c913505c-324e-543b-b2b1-a55c9bcbc9eb"
                dynamicCommand="NeonCenterline.NeonTools.SelectNeonLines"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: выделить центральные линии"
                enable="true"/>
      <itemData guid="ce9120be-eca9-5ed0-b1c8-c43624951305"
                icon="guid://ce9120be-eca9-5ed0-b1c8-c43624951305"
                dynamicCommand="NeonCenterline.NeonCut.FitToCut"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: подогнать под кратность реза"
                userToolTip="Укоротить свободные концы выделенных линий до кратности реза (из настроек)"
                enable="true"/>
      <itemData guid="18048cca-273b-50eb-9d03-54111cc920ce"
                icon="guid://18048cca-273b-50eb-9d03-54111cc920ce"
                dynamicCommand="NeonCenterline.NeonTools.SimplifyContours"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: упростить контур"
                userToolTip="Меньше узлов у выделенных кривых: отход от исходной не больше допуска (по умолчанию 0.1 мм), углы остаются углами. Отмена — Ctrl+Z"
                enable="true"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="uiConfig/commandBars">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <commandBarData guid="3d22040b-4ee7-596d-91ae-1653ffb39f77"
                      nonLocalizableName="NeonCenterline2"
                      userCaption="Неон (NeonCenterline)"
                      type="toolbar">
        <toolbar>
          <item guidRef="4e6d43bf-0b2d-500c-b154-b553b6087991"/>
          <item guidRef="c68c5d4f-d668-5fb4-937e-adb7430cae58"/>
          <item guidRef="0b33ce77-31ad-5c46-bb8b-2cd2e362084f"/>
          <item guidRef="c913505c-324e-543b-b2b1-a55c9bcbc9eb"/>
          <item guidRef="ce9120be-eca9-5ed0-b1c8-c43624951305"/>
          <item guidRef="18048cca-273b-50eb-9d03-54111cc920ce"/>
          <item guidRef="ce37ed0a-544f-575c-ac35-ef4d4aab66c9"/>
        </toolbar>
      </commandBarData>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>

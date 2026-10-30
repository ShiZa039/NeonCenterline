<?xml version="1.0"?>
<!--
  NeonCenterline — команды и панель инструментов (CorelDRAW Addons).
  Команды вызывают макросы проекта NeonCenterline.gms (пользовательская папка GMS).
  Панель: Моно, Курсив | Контур ▾, Линии ▾ | Настройки ▾. «▾» — кнопка-группа:
  itemData type="flyout" с flyoutBarRef на commandBarData type="menu" flyout="true" (как кнопки
  со списком в самом CorelDRAW, DrawUI.items.xml; без type="flyout" CorelDRAW 2026 кнопку не рисует).
  GUID не менять: по ним CorelDRAW узнаёт кнопки в настроенных рабочих пространствах.
  Набор кнопок панели: [NeonToolbar 5]. При изменении состава панели поднять номер здесь и
  NEON_TOOLBAR_VER в src/NeonUI.bas — после обновления плагин сам сбросит панель к этому составу.
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
      <itemData guid="67333f2d-18b9-5925-b8bc-24d97a6dff93"
                icon="guid://67333f2d-18b9-5925-b8bc-24d97a6dff93"
                dynamicCommand="NeonCenterline.NeonTools.PlaceGrommets"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Неон: люверсы"
                userToolTip="Точки под люверсы по краю выделенного баннера: в каждом углу и равномерно по сторонам, не реже заданного расстояния"
                enable="true"/>
      <!-- кнопки-группы: открывают выпадающий список команд -->
      <itemData guid="dfad10fb-4f88-57cc-89aa-60840ed7bfd6"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="2061e8a0-b6a1-56bb-bd4f-3c8546a0e4a2"
                icon="guid://dfad10fb-4f88-57cc-89aa-60840ed7bfd6"
                userCaption="Неон: контур"
                userToolTip="Упростить контур, люверсы"
                enable="true"/>
      <itemData guid="305a1621-a721-5784-9de2-b2338e83336c"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="b14254e7-44de-5f90-bb10-6319042600d8"
                icon="guid://305a1621-a721-5784-9de2-b2338e83336c"
                userCaption="Неон: линии"
                userToolTip="Длина контуров и линий, выделить центральные линии, подогнать под кратность реза"
                enable="true"/>
      <itemData guid="0392cc33-564d-537f-b0df-9d90ebfd4177"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="54b0af94-56f6-588f-b329-fe0cd1ba2e6b"
                icon="guid://0392cc33-564d-537f-b0df-9d90ebfd4177"
                userCaption="Неон: настройки и помощь"
                userToolTip="Настройки, проверить обновления, о программе"
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
          <item guidRef="266435b4-6e53-460f-9fa7-f45be187d400"/>
          <item guidRef="dfad10fb-4f88-57cc-89aa-60840ed7bfd6"/>
          <item guidRef="305a1621-a721-5784-9de2-b2338e83336c"/>
          <item guidRef="266435b4-6e53-460f-9fa7-f45be187d400"/>
          <item guidRef="0392cc33-564d-537f-b0df-9d90ebfd4177"/>
        </toolbar>
      </commandBarData>
      <!-- выпадающие списки кнопок-групп -->
      <commandBarData guid="2061e8a0-b6a1-56bb-bd4f-3c8546a0e4a2"
                      type="menu"
                      nonLocalizableName="NeonCenterline Contour"
                      flyout="true">
        <menu>
          <item guidRef="18048cca-273b-50eb-9d03-54111cc920ce"/>
          <item guidRef="67333f2d-18b9-5925-b8bc-24d97a6dff93"/>
        </menu>
      </commandBarData>
      <commandBarData guid="b14254e7-44de-5f90-bb10-6319042600d8"
                      type="menu"
                      nonLocalizableName="NeonCenterline Lines"
                      flyout="true">
        <menu>
          <item guidRef="0b33ce77-31ad-5c46-bb8b-2cd2e362084f"/>
          <item guidRef="c913505c-324e-543b-b2b1-a55c9bcbc9eb"/>
          <item guidRef="ce9120be-eca9-5ed0-b1c8-c43624951305"/>
        </menu>
      </commandBarData>
      <commandBarData guid="54b0af94-56f6-588f-b329-fe0cd1ba2e6b"
                      type="menu"
                      nonLocalizableName="NeonCenterline Settings"
                      flyout="true">
        <menu>
          <item guidRef="ce37ed0a-544f-575c-ac35-ef4d4aab66c9"/>
          <item guidRef="266435b4-6e53-460f-9fa7-f45be187d400"/>
          <item guidRef="1025a06f-bc40-5284-9292-2c3cf477d6a2"/>
          <item guidRef="c5cc733d-ee55-5173-b5ef-17a954845da7"/>
        </menu>
      </commandBarData>
    </xsl:copy>
  </xsl:template>
</xsl:stylesheet>

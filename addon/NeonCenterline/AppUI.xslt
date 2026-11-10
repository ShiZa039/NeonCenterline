<?xml version="1.0"?>
<!--
  NeonCenterline — команды и панель инструментов (CorelDRAW Addons).
  Команды вызывают макросы проекта NeonCenterline.gms (пользовательская папка GMS).
  Панель: Моно, Курсив | Контур ▾, Линии ▾, Раскладка ▾ | Настройки ▾. «▾» — кнопка-группа:
  itemData type="flyout" с flyoutBarRef на commandBarData type="menu" flyout="true" (как кнопки
  со списком в самом CorelDRAW, DrawUI.items.xml; без type="flyout" CorelDRAW 2026 кнопку не рисует).
  GUID не менять: по ним CorelDRAW узнаёт кнопки в настроенных рабочих пространствах.
  Набор кнопок панели: [NeonToolbar 12]. При изменении состава панели поднять номер здесь и
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
                userCaption="Моно"
                enable="true"/>
      <itemData guid="c68c5d4f-d668-5fb4-937e-adb7430cae58"
                icon="guid://c68c5d4f-d668-5fb4-937e-adb7430cae58"
                dynamicCommand="NeonCenterline.NeonUI.RunScript"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Курсив (бета)"
                userToolTip="Курсив — в разработке, работает нестабильно. Проверяйте линии перед резкой."
                enable="true"/>
      <itemData guid="ce37ed0a-544f-575c-ac35-ef4d4aab66c9"
                icon="guid://ce37ed0a-544f-575c-ac35-ef4d4aab66c9"
                dynamicCommand="NeonCenterline.NeonUI.ShowSettings"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Настройки"
                enable="true"/>
      <itemData guid="1025a06f-bc40-5284-9292-2c3cf477d6a2"
                icon="guid://1025a06f-bc40-5284-9292-2c3cf477d6a2"
                dynamicCommand="NeonCenterline.NeonUI.CheckUpdates"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Проверить обновления"
                enable="true"/>
      <itemData guid="c5cc733d-ee55-5173-b5ef-17a954845da7"
                icon="guid://c5cc733d-ee55-5173-b5ef-17a954845da7"
                dynamicCommand="NeonCenterline.NeonUI.ShowAbout"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="О программе"
                enable="true"/>
      <itemData guid="0b33ce77-31ad-5c46-bb8b-2cd2e362084f"
                icon="guid://0b33ce77-31ad-5c46-bb8b-2cd2e362084f"
                dynamicCommand="NeonCenterline.NeonTools.MeasureLength"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Длина контуров и линий"
                enable="true"/>
      <itemData guid="c913505c-324e-543b-b2b1-a55c9bcbc9eb"
                icon="guid://c913505c-324e-543b-b2b1-a55c9bcbc9eb"
                dynamicCommand="NeonCenterline.NeonTools.SelectNeonLines"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Выделить центральные линии"
                enable="true"/>
      <itemData guid="ce9120be-eca9-5ed0-b1c8-c43624951305"
                icon="guid://ce9120be-eca9-5ed0-b1c8-c43624951305"
                dynamicCommand="NeonCenterline.NeonCut.FitToCut"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Подогнать под кратность реза"
                userToolTip="Укоротить свободные концы выделенных линий до кратности реза (из настроек)"
                enable="true"/>
      <itemData guid="18048cca-273b-50eb-9d03-54111cc920ce"
                icon="guid://18048cca-273b-50eb-9d03-54111cc920ce"
                dynamicCommand="NeonCenterline.NeonTools.SimplifyContours"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Упростить контур"
                userToolTip="Меньше узлов у выделенных кривых: отход от исходной не больше допуска (по умолчанию 0.1 мм), углы остаются углами. Отмена — Ctrl+Z"
                enable="true"/>
      <itemData guid="f694ba63-40ea-5adc-b912-0cea2e6c287a"
                icon="guid://f694ba63-40ea-5adc-b912-0cea2e6c287a"
                dynamicCommand="NeonCenterline.NeonTools.CleanContours"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Очистить контур"
                userToolTip="Убрать наложения и петли: перекрывающиеся части букв и фигур сливаются в один контур, внутренние линии уходят, дырки остаются. Текст переводится в кривые. Отмена — Ctrl+Z"
                enable="true"/>
      <itemData guid="61d74f9e-be33-5689-a10a-7913d24a91f5"
                icon="guid://61d74f9e-be33-5689-a10a-7913d24a91f5"
                dynamicCommand="NeonCenterline.NeonTools.SplitLetters"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Разбить на буквы"
                userToolTip="Разбить текст в кривых или объединённые буквы на отдельные буквы: дырки остаются в своих буквах (не нужно вычитать меньшее из большего). Текст переводится в кривые сам. Отмена — Ctrl+Z"
                enable="true"/>
      <itemData guid="da18bf83-4d0f-5e8b-8834-e23137080009"
                icon="guid://da18bf83-4d0f-5e8b-8834-e23137080009"
                dynamicCommand="NeonCenterline.NeonTools.MakeNoodles"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Лапша (полосы)"
                userToolTip="Нарезка полос заданной длины и ширины для плоттера, лазера, фрезера: горизонтальные резы зигзагом сверху вниз, вертикальные отрезы — в конце списка, резы с выходом за край. Отмена — Ctrl+Z"
                enable="true"/>
      <itemData guid="67333f2d-18b9-5925-b8bc-24d97a6dff93"
                icon="guid://67333f2d-18b9-5925-b8bc-24d97a6dff93"
                dynamicCommand="NeonCenterline.NeonTools.PlaceGrommets"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Люверсы"
                userToolTip="Точки под люверсы по краю выделенного баннера: в каждом углу и равномерно по сторонам, не реже заданного расстояния"
                enable="true"/>
      <itemData guid="b4b6f389-35cc-5b46-b3db-ffb4e200250f"
                icon="guid://b4b6f389-35cc-5b46-b3db-ffb4e200250f"
                dynamicCommand="NeonCenterline.NeonTools.NestSheets"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Раскладка на листы (бета)"
                userToolTip="В разработке: пока только по габаритам и без поворота. Разложить выделенные детали на листы: по габаритам, без поворота, с полем от края и зазором между деталями; каждый цвет — на свои листы; длина и ширина листа — любые, длина может быть по раскладке. Отмена — Ctrl+Z"
                enable="true"/>
      <itemData guid="3ef0efdf-6719-53ff-a444-924cb4d86b5e"
                icon="guid://3ef0efdf-6719-53ff-a444-924cb4d86b5e"
                dynamicCommand="NeonCenterline.NeonTools.GroupSame"
                dynamicCategory="2cc24a3e-fe24-4708-9a74-9c75406eebcd"
                userCaption="Одинаковые — в группы"
                userToolTip="Собрать одинаковые объекты (одна форма с точностью 0,1 мм и один цвет) в группы «N шт» и показать ведомость форм. Отмена — Ctrl+Z"
                enable="true"/>
      <!-- кнопки-группы: открывают выпадающий список команд -->
      <itemData guid="58b78302-39ef-5436-8097-5392406fe15a"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="5adecbd6-e92f-5962-9db2-01885295268a"
                icon="guid://58b78302-39ef-5436-8097-5392406fe15a"
                userCaption="Контур"
                userToolTip="Упростить контур, очистить наложения и петли, разбить на буквы"
                enable="true"/>
      <itemData guid="d32b7178-2e37-5391-aca2-12628ac3838b"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="5731901c-4115-58a7-9466-9da0ca279658"
                icon="guid://d32b7178-2e37-5391-aca2-12628ac3838b"
                userCaption="Раскладка"
                userToolTip="Раскладка на листы (бета), одинаковые — в группы, лапша, люверсы"
                enable="true"/>
      <itemData guid="305a1621-a721-5784-9de2-b2338e83336c"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="b14254e7-44de-5f90-bb10-6319042600d8"
                icon="guid://305a1621-a721-5784-9de2-b2338e83336c"
                userCaption="Линии"
                userToolTip="Длина контуров и линий, выделить центральные линии, подогнать под кратность реза"
                enable="true"/>
      <itemData guid="0392cc33-564d-537f-b0df-9d90ebfd4177"
                type="flyout"
                arrowStyle="corner"
                flyoutBarRef="54b0af94-56f6-588f-b329-fe0cd1ba2e6b"
                icon="guid://0392cc33-564d-537f-b0df-9d90ebfd4177"
                userCaption="Настройки и помощь"
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
          <item guidRef="58b78302-39ef-5436-8097-5392406fe15a"/>
          <item guidRef="305a1621-a721-5784-9de2-b2338e83336c"/>
          <item guidRef="d32b7178-2e37-5391-aca2-12628ac3838b"/>
          <item guidRef="266435b4-6e53-460f-9fa7-f45be187d400"/>
          <item guidRef="0392cc33-564d-537f-b0df-9d90ebfd4177"/>
        </toolbar>
      </commandBarData>
      <!-- выпадающие списки кнопок-групп -->
      <commandBarData guid="5adecbd6-e92f-5962-9db2-01885295268a"
                      type="menu"
                      nonLocalizableName="NeonCenterline Outline"
                      flyout="true">
        <menu>
          <item guidRef="18048cca-273b-50eb-9d03-54111cc920ce"/>
          <item guidRef="f694ba63-40ea-5adc-b912-0cea2e6c287a"/>
          <item guidRef="61d74f9e-be33-5689-a10a-7913d24a91f5"/>
        </menu>
      </commandBarData>
      <commandBarData guid="5731901c-4115-58a7-9466-9da0ca279658"
                      type="menu"
                      nonLocalizableName="NeonCenterline Layout"
                      flyout="true">
        <menu>
          <item guidRef="b4b6f389-35cc-5b46-b3db-ffb4e200250f"/>
          <item guidRef="3ef0efdf-6719-53ff-a444-924cb4d86b5e"/>
          <item guidRef="da18bf83-4d0f-5e8b-8834-e23137080009"/>
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

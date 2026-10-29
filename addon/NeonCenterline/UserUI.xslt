<?xml version="1.0"?>
<!--
  NeonCenterline — размещение панели в рабочем пространстве. Применяется один раз на
  рабочее пространство; после изменения — запуск CorelDRAW с зажатой F8 (сброс рабочего
  пространства!) или включить панель вручную: Окно > Панели инструментов > NeonCenterline.
-->
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:frmwrk="Corel Framework Data"
                exclude-result-prefixes="frmwrk">
  <xsl:output method="xml" encoding="UTF-8" indent="yes"/>

  <frmwrk:uiconfig>
    <frmwrk:compositeNode xPath="/uiConfig/commandBars/commandBarData[@guid='3d22040b-4ee7-596d-91ae-1653ffb39f77']"/>
    <frmwrk:compositeNode xPath="/uiConfig/frame"/>
  </frmwrk:uiconfig>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- недостающие кнопки — в уже стоящую панель -->
  <xsl:template match="commandBarData[@guid='3d22040b-4ee7-596d-91ae-1653ffb39f77']/toolbar">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
      <xsl:if test="not(./item[@guidRef='4e6d43bf-0b2d-500c-b154-b553b6087991'])">
        <item guidRef="4e6d43bf-0b2d-500c-b154-b553b6087991"/>
      </xsl:if>
      <xsl:if test="not(./item[@guidRef='c68c5d4f-d668-5fb4-937e-adb7430cae58'])">
        <item guidRef="c68c5d4f-d668-5fb4-937e-adb7430cae58"/>
      </xsl:if>
      <xsl:if test="not(./item[@guidRef='0b33ce77-31ad-5c46-bb8b-2cd2e362084f'])">
        <item guidRef="0b33ce77-31ad-5c46-bb8b-2cd2e362084f"/>
      </xsl:if>
      <xsl:if test="not(./item[@guidRef='c913505c-324e-543b-b2b1-a55c9bcbc9eb'])">
        <item guidRef="c913505c-324e-543b-b2b1-a55c9bcbc9eb"/>
      </xsl:if>
      <xsl:if test="not(./item[@guidRef='ce9120be-eca9-5ed0-b1c8-c43624951305'])">
        <item guidRef="ce9120be-eca9-5ed0-b1c8-c43624951305"/>
      </xsl:if>
      <xsl:if test="not(./item[@guidRef='18048cca-273b-50eb-9d03-54111cc920ce'])">
        <item guidRef="18048cca-273b-50eb-9d03-54111cc920ce"/>
      </xsl:if>
      <xsl:if test="not(./item[@guidRef='ce37ed0a-544f-575c-ac35-ef4d4aab66c9'])">
        <item guidRef="ce37ed0a-544f-575c-ac35-ef4d4aab66c9"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <!-- панель — рядом со стандартной панелью инструментов (соседний элемент) -->
  <xsl:template match="uiConfig/containers/container[@guid='bee85f91-3ad9-dc8d-48b5-d2a87c8b2109']/container[@guid='Framework_MainFrame-layout']/dockHost[@guid='894bf987-2ec1-8f83-41d8-68f6797d0db4']/toolbar[@guidRef='c2b44f69-6dec-444e-a37e-5dbf7ff43dae']">
    <xsl:copy-of select="."/>
    <xsl:if test="not(../toolbar[@guidRef='3d22040b-4ee7-596d-91ae-1653ffb39f77'])">
      <toolbar guidRef="3d22040b-4ee7-596d-91ae-1653ffb39f77"/>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>

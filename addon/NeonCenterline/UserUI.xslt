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
    <frmwrk:compositeNode xPath="/uiConfig/commandBars/commandBarData[@guid='772fbf90-5bc1-5c4d-93a0-d1fced7f98b5']"/>
    <frmwrk:compositeNode xPath="/uiConfig/frame"/>
  </frmwrk:uiconfig>

  <xsl:template match="node()|@*">
    <xsl:copy>
      <xsl:apply-templates select="node()|@*"/>
    </xsl:copy>
  </xsl:template>

  <!-- панель — рядом со стандартной панелью инструментов (соседний элемент) -->
  <xsl:template match="uiConfig/containers/container[@guid='bee85f91-3ad9-dc8d-48b5-d2a87c8b2109']/container[@guid='Framework_MainFrame-layout']/dockHost[@guid='894bf987-2ec1-8f83-41d8-68f6797d0db4']/toolbar[@guidRef='c2b44f69-6dec-444e-a37e-5dbf7ff43dae']">
    <xsl:copy-of select="."/>
    <xsl:if test="not(../toolbar[@guidRef='772fbf90-5bc1-5c4d-93a0-d1fced7f98b5'])">
      <toolbar guidRef="772fbf90-5bc1-5c4d-93a0-d1fced7f98b5"/>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>

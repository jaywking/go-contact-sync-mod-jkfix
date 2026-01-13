<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:wix="http://schemas.microsoft.com/wix/2006/wi"
xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="msxsl">

  <xsl:output method="xml" indent="yes"/>

  <!--Identity Transform-->
  <xsl:template match="@*|node()">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()"/>
    </xsl:copy>
  </xsl:template>

  <!--Set up keys of component Ids from Snapshot.xml-->
  <!--NB: Reinstate the line below when everyone has upgraded to Wix 3.6-->
  <!--<xsl:key name="snapshot-search" match="wix:Component[@Id = document('Snapshot.xml')//wix:Component/@Id]" use="@Id"/>-->

  <!--Set up keys for ignoring various file types-->
  <xsl:key name="xml-search" match="wix:Component[contains(wix:File/@Source, '.xml')]" use="@Id"/>
  <!--<xsl:key name="pdb-search" match="wix:Component[contains(wix:File/@Source, '.pdb')]" use="@Id"/>-->
  <!--
  <xsl:key name="svn-search" match="wix:Component[ancestor::wix:Directory/@Name = '.svn']" use="@Id"/>
  -->
  <!--Match and ignore .xml files-->
  <xsl:template match="wix:Component[key('xml-search', @Id)]"/>
  <xsl:template match="wix:ComponentRef[key('xml-search', @Id)]"/>

  <!--Match and ignore leftover .pdb files-->
  <!--
  <xsl:template match="wix:Component[key('pdb-search', @Id)]"/>
  <xsl:template match="wix:ComponentRef[key('pdb-search', @Id)]"/>
  -->
  
  <!--Match and ignore “.svn" directories on build machines -->
  <!--
  <xsl:template match=“wix:Directory[@Name = '.svn']“/>
  <xsl:template match="wix:ComponentRef[key('svn-search', @Id)]"/>
  -->
  <!--Match Components that also exist in Snapshot.xml, and use the snapshot version-->
  <!--NB: Reinstate the 4 lines below when everyone has upgraded to Wix 3.6-->
  <!--<xsl:template match="wix:Component[key('snapshot-search', @Id)]">
    <xsl:variable name="component" select="."/>
    <xsl:copy-of select="document(‘Snapshot.xml’)//wix:Component[@Id = $component/@Id]"/>
  </xsl:template>-->

</xsl:stylesheet>
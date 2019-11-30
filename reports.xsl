<?xml version="1.0" encoding="windows-1251" standalone="yes"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40" xmlns:s="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882" xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" xmlns:rs="urn:schemas-microsoft-com:rowset" xmlns:z="#RowsetSchema">
  <xsl:output indent="yes" method="xml" encoding="windows-1251"></xsl:output>
  <xsl:template name="head">
    <xsl:for-each select="/xml/rs:data">
      <xsl:variable name="row" select="current()"></xsl:variable>
      <ss:Row>
      <xsl:for-each select="/xml/s:Schema/s:ElementType/s:AttributeType">
        <xsl:variable name="col" select="current()"></xsl:variable>
        <xsl:variable name="colname" select="$col/@*[name()='name']"></xsl:variable>
        <xsl:variable name="coltype" select="$col/s:datatype/@*[name()='dt:type']"></xsl:variable>
        <xsl:variable name="colvalue" select="$col/@*[name()='rs:name']"></xsl:variable>
        <xsl:variable name="type">
          <xsl:choose>
            <xsl:when test="$coltype='float'">
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='number'">
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='string'">
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='dateTime'">
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <ss:Cell ss:StyleID="{$colname}">
          <ss:Data ss:Type="{$type}">
            <xsl:value-of select="$colvalue"></xsl:value-of>
          </ss:Data>
        </ss:Cell>
      </xsl:for-each>
      </ss:Row>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="data">
    <xsl:for-each select="/xml/rs:data/z:row">
      <xsl:variable name="row" select="current()"></xsl:variable>
      <ss:Row>
      <xsl:for-each select="/xml/s:Schema/s:ElementType/s:AttributeType">
        <xsl:variable name="col" select="current()"></xsl:variable>
        <xsl:variable name="colname" select="$col/@*[name()='name']"></xsl:variable>
        <xsl:variable name="coltype" select="$col/s:datatype/@*[name()='dt:type']"></xsl:variable>
        <xsl:variable name="colvalue" select="$row/@*[name()=$colname]"></xsl:variable>
        <xsl:variable name="type">
          <xsl:choose>
            <xsl:when test="$coltype='float'">
              <xsl:value-of select="'Number'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='number'">
              <xsl:value-of select="'Number'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='string'">
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='dateTime'">
              <xsl:value-of select="'DateTime'"></xsl:value-of>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'String'"></xsl:value-of>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <ss:Cell ss:StyleID="{$colname}">
          <ss:Data ss:Type="{$type}">
            <xsl:value-of select="$colvalue"></xsl:value-of>
          </ss:Data>
        </ss:Cell>
      </xsl:for-each>
      </ss:Row>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="style">
    <xsl:for-each select="/xml/rs:data">
      <xsl:variable name="row" select="current()"></xsl:variable>
      <ss:Styles>
      <xsl:for-each select="/xml/s:Schema/s:ElementType/s:AttributeType">
        <xsl:variable name="col" select="current()"></xsl:variable>
        <xsl:variable name="colname" select="$col/@*[name()='name']"></xsl:variable>
        <xsl:variable name="coltype" select="$col/s:datatype/@*[name()='dt:type']"></xsl:variable>
        <xsl:variable name="colvalue" select="$col/@*[name()='rs:name']"></xsl:variable>
        <xsl:variable name="type">
          <xsl:choose>
            <xsl:when test="$coltype='float'">
              <xsl:value-of select="'Fixed'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='number'">
              <xsl:value-of select="'General Number'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='string'">
              <xsl:value-of select="'@'"></xsl:value-of>
            </xsl:when>
            <xsl:when test="$coltype='dateTime'">
              <xsl:value-of select="'Short Date'"></xsl:value-of>
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="'General'"></xsl:value-of>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <ss:Style ss:ID="{$colname}">
          <ss:NumberFormat ss:Format="{$type}"></ss:NumberFormat>
        </ss:Style>
      </xsl:for-each>
      </ss:Styles>
    </xsl:for-each>
  </xsl:template>
  <xsl:template name="sheet">
    <xsl:for-each select="/xml/rs:data">
      <ss:Worksheet ss:Name="Report">
        <x:WorksheetOptions>
          <x:ActivePane>2</x:ActivePane>
          <x:FreezePanes></x:FreezePanes>
          <x:SplitHorizontal>1</x:SplitHorizontal>
          <x:TopRowBottomPane>1</x:TopRowBottomPane>
        </x:WorksheetOptions>
        <ss:Table>
          <xsl:call-template name="head"></xsl:call-template>
          <xsl:call-template name="data"></xsl:call-template>
        </ss:Table>
      </ss:Worksheet>
    </xsl:for-each>
  </xsl:template>
  <xsl:template match="/">
    <xsl:processing-instruction name="mso-application">progid="Excel.Sheet"</xsl:processing-instruction>
    <ss:Workbook>
      <xsl:call-template name="style"></xsl:call-template>
      <xsl:call-template name="sheet"></xsl:call-template>
    </ss:Workbook>
  </xsl:template>
</xsl:stylesheet>

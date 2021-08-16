<?xml version="1.0" encoding="windows-1251" standalone="yes"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:rs="urn:schemas-microsoft-com:rowset"
	xmlns:xs="urn:schemas-microsoft-com:office:excel"
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
	xmlns:st="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882"
	xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"
	xmlns:zt="#RowsetSchema">
	<xsl:output indent="yes" method="xml" encoding="windows-1251"/>
	<xsl:template match="/xml/st:Schema/st:ElementType" mode="head">
		<ss:Row>
			<xsl:for-each select="current()/st:AttributeType">
				<xsl:variable name="colvalue">
					<xsl:value-of select="normalize-space(current()/@rs:name)"/>
				</xsl:variable>
				<xsl:variable name="colname">
					<xsl:value-of select="normalize-space(current()/@name)"/>
				</xsl:variable>
				<xsl:variable name="coltype">
					<xsl:choose>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='float'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='number'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='string'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='dateTime'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'String'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<ss:Cell ss:StyleID="{$colname}">
					<ss:Data ss:Type="{$coltype}">
						<xsl:value-of select="$colvalue"/>
					</ss:Data>
				</ss:Cell>
			</xsl:for-each>
		</ss:Row>
	</xsl:template>
	<xsl:template match="/xml/rs:data/zt:row" mode="data">
		<ss:Row>
			<xsl:for-each select="current()/@*">
				<xsl:variable name="colvalue">
					<xsl:value-of select="normalize-space(current())"/>
				</xsl:variable>
				<xsl:variable name="colname">
					<xsl:value-of select="name(current())"/>
				</xsl:variable>
				<xsl:variable name="coltype">
					<xsl:choose>
						<xsl:when test="normalize-space(/xml/st:Schema/st:ElementType/st:AttributeType[@name=$colname]/st:datatype/@dt:type)='float'">
							<xsl:value-of select="'Number'"/>
						</xsl:when>
						<xsl:when test="normalize-space(/xml/st:Schema/st:ElementType/st:AttributeType[@name=$colname]/st:datatype/@dt:type)='number'">
							<xsl:value-of select="'Number'"/>
						</xsl:when>
						<xsl:when test="normalize-space(/xml/st:Schema/st:ElementType/st:AttributeType[@name=$colname]/st:datatype/@dt:type)='string'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="normalize-space(/xml/st:Schema/st:ElementType/st:AttributeType[@name=$colname]/st:datatype/@dt:type)='dateTime'">
							<xsl:value-of select="'DateTime'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'String'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<ss:Cell ss:StyleID="{$colname}">
					<ss:Data ss:Type="{$coltype}">
						<xsl:value-of select="$colvalue"/>
					</ss:Data>
				</ss:Cell>
			</xsl:for-each>
		</ss:Row>
	</xsl:template>
	<xsl:template match="/xml/st:Schema/st:ElementType" mode="style">
		<ss:Styles>
			<xsl:for-each select="current()/st:AttributeType">
				<xsl:variable name="colvalue">
					<xsl:value-of select="normalize-space(current()/@rs:name)"/>
				</xsl:variable>
				<xsl:variable name="colname">
					<xsl:value-of select="normalize-space(current()/@name)"/>
				</xsl:variable>
				<xsl:variable name="coltype">
					<xsl:choose>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='float'">
							<xsl:value-of select="'Fixed'"/>
						</xsl:when>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='number'">
							<xsl:value-of select="'General Number'"/>
						</xsl:when>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='string'">
							<xsl:value-of select="'@'"/>
						</xsl:when>
						<xsl:when test="normalize-space(current()/st:datatype/@dt:type)='dateTime'">
              						<xsl:value-of select="'Short Date'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'General'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:variable>
				<ss:Style ss:ID="{$colname}">
					<ss:NumberFormat ss:Format="{$coltype}"/>
				</ss:Style>
			</xsl:for-each>
		</ss:Styles>
	</xsl:template>
	<xsl:template match="/xml/rs:data" mode="sheet">
		<ss:Worksheet ss:Name="Report">
			<xs:WorksheetOptions>
				<xs:ActivePane>2</xs:ActivePane>
				<xs:FreezePanes></xs:FreezePanes>
				<xs:SplitHorizontal>1</xs:SplitHorizontal>
				<xs:TopRowBottomPane>1</xs:TopRowBottomPane>
			</xs:WorksheetOptions>
			<ss:Table>
				<xsl:apply-templates select="/xml/st:Schema/st:ElementType" mode="head"/>
				<xsl:apply-templates select="/xml/rs:data/zt:row" mode="data"/>
			</ss:Table>
		</ss:Worksheet>
	</xsl:template>
	<xsl:template match="/">
		<xsl:processing-instruction name="mso-application">progid="Excel.Sheet"</xsl:processing-instruction>
		<ss:Workbook>
			<xsl:apply-templates select="/xml/st:Schema/st:ElementType" mode="style"/>
			<xsl:apply-templates select="/xml/rs:data" mode="sheet"/>
		</ss:Workbook>
	</xsl:template>
</xsl:stylesheet>

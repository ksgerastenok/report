<?xml version="1.0" encoding="windows-1251"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
	xmlns:rs="urn:schemas-microsoft-com:rowset"
	xmlns:xs="urn:schemas-microsoft-com:office:excel"
	xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
	xmlns:st="uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882"
	xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882"
	xmlns:zt="#RowsetSchema">

	<xsl:output method="xml" version="1.0" encoding="windows-1251" indent="yes"/>
	<xsl:template name="head">
		<xsl:for-each select="/xml/st:Schema/st:ElementType">
			<ss:Row>
				<xsl:for-each select="current()/st:AttributeType">
					<xsl:variable name="name">
						<xsl:value-of select="normalize-space(current()/@name)"/>
					</xsl:variable>
					<xsl:variable name="text">
						<xsl:value-of select="normalize-space(current()/@rs:name)"/>
					</xsl:variable>
					<xsl:variable name="data">
						<xsl:value-of select="normalize-space(current()/st:datatype/@dt:type)"/>
					</xsl:variable>
					<xsl:variable name="type">
						<xsl:choose>
							<xsl:when test="$data='float'">
								<xsl:value-of select="'String'"/>
							</xsl:when>
							<xsl:when test="$data='number'">
								<xsl:value-of select="'String'"/>
							</xsl:when>
							<xsl:when test="$data='string'">
								<xsl:value-of select="'String'"/>
							</xsl:when>
							<xsl:when test="$data='dateTime'">
								<xsl:value-of select="'String'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'String'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<ss:Cell ss:StyleID="{$name}">
						<ss:Data ss:Type="{$type}">
							<xsl:value-of select="$text"/>
						</ss:Data>
					</ss:Cell>
				</xsl:for-each>
			</ss:Row>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="data">
		<xsl:for-each select="/xml/rs:data/zt:row">
			<ss:Row>
				<xsl:for-each select="current()/@*">
					<xsl:variable name="name">
						<xsl:value-of select="normalize-space(name())"/>
					</xsl:variable>
					<xsl:variable name="text">
						<xsl:value-of select="normalize-space(current())"/>
					</xsl:variable>
					<xsl:variable name="data">
						<xsl:value-of select="normalize-space(/xml/st:Schema/st:ElementType/st:AttributeType[@name=$name]/st:datatype/@dt:type)"/>
					</xsl:variable>
					<xsl:variable name="type">
						<xsl:choose>
							<xsl:when test="$data='float'">
								<xsl:value-of select="'Number'"/>
							</xsl:when>
							<xsl:when test="$data='number'">
								<xsl:value-of select="'Number'"/>
							</xsl:when>
							<xsl:when test="$data='string'">
								<xsl:value-of select="'String'"/>
							</xsl:when>
							<xsl:when test="$data='dateTime'">
								<xsl:value-of select="'DateTime'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'String'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<ss:Cell ss:StyleID="{$name}">
						<ss:Data ss:Type="{$type}">
							<xsl:value-of select="$text"/>
						</ss:Data>
					</ss:Cell>
				</xsl:for-each>
			</ss:Row>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="style">
		<xsl:for-each select="/xml/st:Schema/st:ElementType">
			<ss:Styles>
				<xsl:for-each select="current()/st:AttributeType">
					<xsl:variable name="name">
						<xsl:value-of select="normalize-space(current()/@name)"/>
					</xsl:variable>
					<xsl:variable name="text">
						<xsl:value-of select="normalize-space(current()/@rs:name)"/>
					</xsl:variable>
					<xsl:variable name="data">
						<xsl:value-of select="normalize-space(current()/st:datatype/@dt:type)"/>
					</xsl:variable>
					<xsl:variable name="type">
						<xsl:choose>
							<xsl:when test="$data='float'">
								<xsl:value-of select="'Fixed'"/>
							</xsl:when>
							<xsl:when test="$data='number'">
								<xsl:value-of select="'General Number'"/>
							</xsl:when>
							<xsl:when test="$data='string'">
								<xsl:value-of select="'@'"/>
							</xsl:when>
							<xsl:when test="$data='dateTime'">
              							<xsl:value-of select="'Short Date'"/>
							</xsl:when>
							<xsl:otherwise>
								<xsl:value-of select="'General'"/>
							</xsl:otherwise>
						</xsl:choose>
					</xsl:variable>
					<ss:Style ss:ID="{$name}">
						<ss:NumberFormat ss:Format="{$type}"/>
					</ss:Style>
				</xsl:for-each>
			</ss:Styles>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="sheet">
		<ss:Worksheet ss:Name="Report">
			<xs:WorksheetOptions>
				<xs:ActivePane>2</xs:ActivePane>
				<xs:FreezePanes></xs:FreezePanes>
				<xs:SplitHorizontal>1</xs:SplitHorizontal>
				<xs:TopRowBottomPane>1</xs:TopRowBottomPane>
			</xs:WorksheetOptions>
			<ss:Table>
				<xsl:call-template name="head"/>
				<xsl:call-template name="data"/>
			</ss:Table>
		</ss:Worksheet>
	</xsl:template>
	<xsl:template match="/">
		<xsl:processing-instruction name="mso-application">progid="Excel.Sheet"</xsl:processing-instruction>
		<ss:Workbook>
			<xsl:call-template name="style"/>
			<xsl:call-template name="sheet"/>
		</ss:Workbook>
	</xsl:template>
</xsl:stylesheet>

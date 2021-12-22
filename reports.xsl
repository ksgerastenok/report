<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template match="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="style">
		<xsl:element name="Styles" namespace="urn:schemas-microsoft-com:office:spreadsheet">	
			<xsl:apply-templates select="current()/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="style"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="style">
		<xsl:element name="Style" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:attribute name="ID" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:value-of select="current()/@*[local-name() = 'name']"/>
			</xsl:attribute>
			<xsl:element name="NumberFormat" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:attribute name="Format" namespace="urn:schemas-microsoft-com:office:spreadsheet">
					<xsl:choose>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'float'">
							<xsl:value-of select="'Fixed'"/>
						</xsl:when>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'number'">
							<xsl:value-of select="'General Number'"/>
						</xsl:when>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'string'">
							<xsl:value-of select="'@'"/>
						</xsl:when>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'dateTime'">
              						<xsl:value-of select="'Short Date'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'General'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
			</xsl:element>
		</xsl:element>
	</xsl:template>

	<xsl:template match="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="head">
		<xsl:element name="Row" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:apply-templates select="current()/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="head"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="head">
		<xsl:element name="Cell" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:attribute name="StyleID" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:value-of select="current()/@*[local-name() = 'name']"/>
			</xsl:attribute>
			<xsl:element name="Data" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:attribute name="Type" namespace="urn:schemas-microsoft-com:office:spreadsheet">
					<xsl:choose>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'float'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'number'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'string'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="current()/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'dateTime'">
              						<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'String'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:value-of select="current()/@*[local-name() = 'name'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']"/>
			</xsl:element>
		</xsl:element>
	</xsl:template>

	<xsl:template match="/*[local-name() = 'xml']/*[local-name() = 'data'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']/*[local-name() = 'row'][namespace-uri() = '#RowsetSchema']" mode="data">
		<xsl:element name="Row" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:apply-templates select="current()/@*" mode="data"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="/*[local-name() = 'xml']/*[local-name() = 'data'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']/*[local-name() = 'row'][namespace-uri() = '#RowsetSchema']/@*" mode="data">
		<xsl:element name="Cell" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:attribute name="StyleID" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:value-of select="name(current())"/>
			</xsl:attribute>
			<xsl:element name="Data" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:attribute name="Type" namespace="urn:schemas-microsoft-com:office:spreadsheet">
					<xsl:choose>
						<xsl:when test="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'][name(current()) = @name]/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'float'">
							<xsl:value-of select="'Number'"/>
						</xsl:when>
						<xsl:when test="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'][name(current()) = @name]/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'number'">
							<xsl:value-of select="'Number'"/>
						</xsl:when>
						<xsl:when test="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'][name(current()) = @name]/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'string'">
							<xsl:value-of select="'String'"/>
						</xsl:when>
						<xsl:when test="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882'][name(current()) = @name]/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882'] = 'dateTime'">
							<xsl:value-of select="'DateTime'"/>
						</xsl:when>
						<xsl:otherwise>
							<xsl:value-of select="'String'"/>
						</xsl:otherwise>
					</xsl:choose>
				</xsl:attribute>
				<xsl:value-of select="current()"/>
			</xsl:element>
		</xsl:element>
	</xsl:template>

	<xsl:template match="/">
		<xsl:processing-instruction name="mso-application">
			<xsl:value-of select="'progid=&quot;Excel.Sheet&quot;'"/>
		</xsl:processing-instruction>
		<xsl:element name="Workbook" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:apply-templates select="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="style"/>
			<xsl:element name="Worksheet" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:attribute name="Name" namespace="urn:schemas-microsoft-com:office:spreadsheet">
					<xsl:value-of select="'Report'"/>
				</xsl:attribute>
				<xsl:element name="WorksheetOptions" namespace="urn:schemas-microsoft-com:office:excel">
					<xsl:element name="ActivePane" namespace="urn:schemas-microsoft-com:office:excel">
						<xsl:value-of select="'2'"/>
					</xsl:element>
					<xsl:element name="FreezePanes" namespace="urn:schemas-microsoft-com:office:excel">
						<xsl:value-of select="''"/>
					</xsl:element>
					<xsl:element name="SplitHorizontal" namespace="urn:schemas-microsoft-com:office:excel">
						<xsl:value-of select="'1'"/>
					</xsl:element>
					<xsl:element name="TopRowBottomPane" namespace="urn:schemas-microsoft-com:office:excel">
						<xsl:value-of select="'1'"/>
					</xsl:element>
				</xsl:element>
				<xsl:element name="Table" namespace="urn:schemas-microsoft-com:office:spreadsheet">
					<xsl:apply-templates select="/*[local-name() = 'xml']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']" mode="head"/>
					<xsl:apply-templates select="/*[local-name() = 'xml']/*[local-name() = 'data'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']/*[local-name() = 'row'][namespace-uri() = '#RowsetSchema']" mode="data"/>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>

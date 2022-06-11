<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
	<xsl:output method="xml" version="1.0" encoding="UTF-8" indent="yes"/>
	<xsl:template name="style">
		<xsl:for-each select="/*[local-name() = 'xml'][namespace-uri() = '']/*[local-name() = 'data'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']">
			<xsl:variable name="row" select="current()"/>
			<xsl:element name="Styles" namespace="urn:schemas-microsoft-com:office:spreadsheet">	
				<xsl:for-each select="/*[local-name() = 'xml'][namespace-uri() = '']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']">
					<xsl:variable name="col" select="current()"/>
					<xsl:variable name="name" select="$col/@*[local-name() = 'name'][namespace-uri() = '']"/>
					<xsl:variable name="type" select="$col/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882']"/>
					<xsl:variable name="value" select="$col/@*[local-name() = 'name'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']"/>
					<xsl:element name="Style" namespace="urn:schemas-microsoft-com:office:spreadsheet">
						<xsl:attribute name="ID" namespace="urn:schemas-microsoft-com:office:spreadsheet">
							<xsl:value-of select="$name"/>
						</xsl:attribute>
						<xsl:element name="NumberFormat" namespace="urn:schemas-microsoft-com:office:spreadsheet">
							<xsl:attribute name="Format" namespace="urn:schemas-microsoft-com:office:spreadsheet">
								<xsl:choose>
									<xsl:when test="$type='float'">
										<xsl:value-of select="'Fixed'"/>
									</xsl:when>
									<xsl:when test="$type='number'">
										<xsl:value-of select="'General Number'"/>
									</xsl:when>
									<xsl:when test="$type='string'">
										<xsl:value-of select="'@'"/>
									</xsl:when>
									<xsl:when test="$type='dateTime'">
              									<xsl:value-of select="'Short Date'"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="'General'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
						</xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="head">
		<xsl:for-each select="/*[local-name() = 'xml'][namespace-uri() = '']/*[local-name() = 'data'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']">
			<xsl:variable name="row" select="current()"/>
			<xsl:element name="Row" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:for-each select="/*[local-name() = 'xml'][namespace-uri() = '']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']">
					<xsl:variable name="col" select="current()"/>
					<xsl:variable name="name" select="$col/@*[local-name() = 'name'][namespace-uri() = '']"/>
					<xsl:variable name="type" select="$col/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882']"/>
					<xsl:variable name="value" select="$col/@*[local-name() = 'name'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']"/>
					<xsl:element name="Cell" namespace="urn:schemas-microsoft-com:office:spreadsheet">
						<xsl:attribute name="StyleID" namespace="urn:schemas-microsoft-com:office:spreadsheet">
							<xsl:value-of select="$name"/>
						</xsl:attribute>
						<xsl:element name="Data" namespace="urn:schemas-microsoft-com:office:spreadsheet">
							<xsl:attribute name="Type" namespace="urn:schemas-microsoft-com:office:spreadsheet">
								<xsl:choose>
									<xsl:when test="$type='float'">
										<xsl:value-of select="'String'"/>
									</xsl:when>
									<xsl:when test="$type='number'">
										<xsl:value-of select="'String'"/>
									</xsl:when>
									<xsl:when test="$type='string'">
										<xsl:value-of select="'String'"/>
									</xsl:when>
									<xsl:when test="$type='dateTime'">
              									<xsl:value-of select="'String'"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="'String'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:value-of select="$value"/>
						</xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>
	<xsl:template name="data">
		<xsl:for-each select="/*[local-name() = 'xml'][namespace-uri() = '']/*[local-name() = 'data'][namespace-uri() = 'urn:schemas-microsoft-com:rowset']/*[local-name() = 'row'][namespace-uri() = '#RowsetSchema']">
			<xsl:variable name="row" select="current()"/>
			<xsl:element name="Row" namespace="urn:schemas-microsoft-com:office:spreadsheet">
				<xsl:for-each select="/*[local-name() = 'xml'][namespace-uri() = '']/*[local-name() = 'Schema'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'ElementType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/*[local-name() = 'AttributeType'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']">
					<xsl:variable name="col" select="current()"/>
					<xsl:variable name="name" select="$col/@*[local-name() = 'name'][namespace-uri() = '']"/>
					<xsl:variable name="type" select="$col/*[local-name() = 'datatype'][namespace-uri() = 'uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882']/@*[local-name() = 'type'][namespace-uri() = 'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882']"/>
					<xsl:variable name="value" select="$row/@*[local-name() = $name][namespace-uri() = '']"/>
					<xsl:element name="Cell" namespace="urn:schemas-microsoft-com:office:spreadsheet">
						<xsl:attribute name="StyleID" namespace="urn:schemas-microsoft-com:office:spreadsheet">
							<xsl:value-of select="$name"/>
						</xsl:attribute>
						<xsl:element name="Data" namespace="urn:schemas-microsoft-com:office:spreadsheet">
							<xsl:attribute name="Type" namespace="urn:schemas-microsoft-com:office:spreadsheet">
								<xsl:choose>
									<xsl:when test="$type='float'">
										<xsl:value-of select="'Number'"/>
									</xsl:when>
									<xsl:when test="$type='number'">
										<xsl:value-of select="'Number'"/>
									</xsl:when>
									<xsl:when test="$type='string'">
										<xsl:value-of select="'String'"/>
									</xsl:when>
									<xsl:when test="$type='dateTime'">
										<xsl:value-of select="'DateTime'"/>
									</xsl:when>
									<xsl:otherwise>
										<xsl:value-of select="'String'"/>
									</xsl:otherwise>
								</xsl:choose>
							</xsl:attribute>
							<xsl:value-of select="$value"/>
						</xsl:element>
					</xsl:element>
				</xsl:for-each>
			</xsl:element>
		</xsl:for-each>
	</xsl:template>

	<xsl:template match="/">
		<xsl:processing-instruction name="mso-application">
			<xsl:value-of select="'progid=&quot;Excel.Sheet&quot;'"/>
		</xsl:processing-instruction>
		<xsl:element name="Workbook" namespace="urn:schemas-microsoft-com:office:spreadsheet">
			<xsl:call-template name="style"/>
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
					<xsl:call-template name="head"/>
					<xsl:call-template name="data"/>
				</xsl:element>
			</xsl:element>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>

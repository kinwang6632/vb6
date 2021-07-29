<?xml version="1.0"?>
<xsl:stylesheet xmlns:xsl="uri:xsl">
  <xsl:template match="/">
    <HTML>
      <BODY>
        <TABLE BORDER="1">
          <TR STYLE="font-weight:bold">
            <TD>Common Name</TD>
            <TD>Botanical Name</TD>
            <TD>Zone</TD>
            <TD>Light</TD>
            <TD>Price</TD>
            <TD>Availability</TD>
          </TR>
          <xsl:for-each select="CATALOG/PLANT">
            <TR>
              <xsl:apply-templates/>
            </TR>
          </xsl:for-each>
        </TABLE>
      </BODY>
    </HTML>
  </xsl:template>
   
  <xsl:template match="NAME">
    <TD><xsl:value-of select="COMMON"/></TD>
    <TD><xsl:value-of select="BOTAN"/></TD>
  </xsl:template>

  <xsl:template match="GROWTH">
    <TD><xsl:value-of select="ZONE"/></TD>
    <TD><xsl:value-of select="LIGHT"/></TD>
  </xsl:template>

  <xsl:template match="SALESINFO">
    <TD><xsl:value-of select="PRICE"/></TD>
    <TD><xsl:value-of select="AVAILABILITY"/></TD>
  </xsl:template>   
</xsl:stylesheet>

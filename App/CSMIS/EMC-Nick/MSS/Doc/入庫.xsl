<?xml version="1.0" encoding="Big5"?>
<xsl:stylesheet xmlns:xsl="uri:xsl">
  <xsl:template match="/">
    <HTML>
      <BODY>
        <TABLE BORDER="1">
          <TR STYLE="font-weight:bold">
            <TD>日期</TD>
            <TD>時間</TD>
            <TD>廠商代碼</TD>
            <TD>廠商名稱</TD>
            <TD>狀況</TD>
            <TD>公司代碼</TD>
          </TR>
          <xsl:for-each select="入庫/批號">
            <TR>
              <xsl:apply-templates/>
            </TR>
          </xsl:for-each>

        </TABLE>
      </BODY>
    </HTML>
  </xsl:template>

  <xsl:template match="入庫/批號">
    <TD><xsl:value-of select="日期"/></TD>
    <TD><xsl:value-of select="時間"/></TD>
    <TD><xsl:value-of select="廠商代碼"/></TD>
    <TD><xsl:value-of select="廠商名稱"/></TD>
    <TD><xsl:value-of select="狀況"/></TD>
    <TD><xsl:value-of select="公司代碼"/></TD>
  </xsl:template>


</xsl:stylesheet>

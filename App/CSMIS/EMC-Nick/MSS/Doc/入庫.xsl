<?xml version="1.0" encoding="Big5"?>
<xsl:stylesheet xmlns:xsl="uri:xsl">
  <xsl:template match="/">
    <HTML>
      <BODY>
        <TABLE BORDER="1">
          <TR STYLE="font-weight:bold">
            <TD>���</TD>
            <TD>�ɶ�</TD>
            <TD>�t�ӥN�X</TD>
            <TD>�t�ӦW��</TD>
            <TD>���p</TD>
            <TD>���q�N�X</TD>
          </TR>
          <xsl:for-each select="�J�w/�帹">
            <TR>
              <xsl:apply-templates/>
            </TR>
          </xsl:for-each>

        </TABLE>
      </BODY>
    </HTML>
  </xsl:template>

  <xsl:template match="�J�w/�帹">
    <TD><xsl:value-of select="���"/></TD>
    <TD><xsl:value-of select="�ɶ�"/></TD>
    <TD><xsl:value-of select="�t�ӥN�X"/></TD>
    <TD><xsl:value-of select="�t�ӦW��"/></TD>
    <TD><xsl:value-of select="���p"/></TD>
    <TD><xsl:value-of select="���q�N�X"/></TD>
  </xsl:template>


</xsl:stylesheet>

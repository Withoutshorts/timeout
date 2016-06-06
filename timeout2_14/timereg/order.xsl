<?xml version="1.0" encoding="utf-8"?>

<xsl:stylesheet version="1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  
  

<xsl:template match="/">

  <html>
    <body>
      Hej der
      <xsl:apply-templates />
    </body>
    
  </html>
  
  
  
  <order>
    <date>
      <xsl:value-of select="/order/date/year"/>
      <xsl:value-of select="/order/date/day"/>
      <xsl:value-of select="/order/date/month"/>
    </date>
    <customer>CompanyA</customer>
    <item>

      <xsl:apply-templates select="/order/item"></xsl:apply-templates>
      <quantity>
        <xsl:apply-templates select="/order/quantity"></xsl:apply-templates>
      </quantity>
    </item>
  </order>
  </xsl:template>

  <xsl:template match="/">
    <part-number>

      <xsl:choose>
        <xsl:when test=". = 'Production-Class Widget'">E16-25A</xsl:when>
        <xsl:when test=". = Economy-Class Widget'">E16-25B</xsl:when>

        <xsl:otherwise>00</xsl:otherwise>
      </xsl:choose>
    </part-number>
    <description>
      <xsl:value-of select="."></xsl:value-of>
    </description>
  </xsl:template>
  
  
  

</xsl:stylesheet> 


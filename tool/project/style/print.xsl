<?xml version="1.0" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<xsl:template match="/">

<html>
<head><title>QML - Source view</title>
<link rel="stylesheet" href="../style/story_print.css" type="text/css" media="all" />
</head>
<body>

    <h1><xsl:value-of select="quest/about/title"/><br />
        <span class="author">by <xsl:value-of select="quest/about/author"/></span>
    </h1>
    <xsl:for-each select="quest/station" order-by="@id">
        <h2><xsl:value-of select="@id"/></h2>
        <p><xsl:value-of select=".//text"/></p>
        <ul><xsl:for-each select="choice">
            <li>
                <p><xsl:value-of select="."/><br />
                <span class="choice">Continue at <span class="reference"><xsl:value-of select="@station"/></span></span>
                </p>
             </li>
             </xsl:for-each>
        </ul>
    </xsl:for-each>

</body>
</html>

</xsl:template>
</xsl:stylesheet>
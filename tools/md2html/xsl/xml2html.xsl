<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:loc="http://documentat.io"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xpath-default-namespace="http://docbook.org/ns/docbook" exclude-result-prefixes="#all"
    version="2.0">

    <xsl:output method="html" indent="yes"/>


    <!-- 
        Utilities
    -->

    <xsl:template name="LFCR">
        <xsl:text>&#10;&#13;</xsl:text>
    </xsl:template>


    <!-- 
        Detecting element levels 
    -->

    <xsl:template match="section/title" mode="level">
        <xsl:value-of select="count(ancestor::section) - 1"/>
    </xsl:template>

    <xsl:template match="*" mode="level">
        <xsl:value-of select="1"/>
    </xsl:template>

    <xsl:function name="loc:level" as="xs:integer">

        <xsl:param name="element"/>

        <xsl:apply-templates select="$element" mode="level"/>

    </xsl:function>


    <!-- 
        Detecting output style names 
    -->

    <xsl:template match="para" mode="styleName">
        <xsl:text>Normal</xsl:text>
    </xsl:template>

    <xsl:template match="itemizedlist" mode="styleName">
        <xsl:text>Scroll List Bullet</xsl:text>
    </xsl:template>

    <xsl:template match="orderedlist" mode="styleName">
        <xsl:text>Scroll List Number</xsl:text>
    </xsl:template>

    <xsl:template match="section/title" mode="styleName">
        <xsl:value-of select="concat('Heading ', loc:level(.))"/>
    </xsl:template>

    <xsl:template match="*" mode="styleName"/>


    <!-- 
        Assembling anchore names for output styles 
    -->

    <xsl:template name="anchorPair">

        <xsl:param name="anchorContent"/>

        <xsl:param name="anchorName"/>

        <a name="{$anchorName}">@#$$#@</a>
        <xsl:copy-of select="$anchorContent"/>
        <a name="{concat('_', $anchorName)}">@#$$#@</a>

    </xsl:template>


    <xsl:function name="loc:escapeStyleName">

        <xsl:param name="styleName"/>

        <xsl:variable name="stage1" select="replace($styleName, '_', '__')"/>

        <xsl:variable name="stage2" select="replace($stage1, ' ', '_s')"/>

        <xsl:value-of select="$stage2"/>

    </xsl:function>


    <xsl:function name="loc:styleAnchorName">

        <xsl:param name="element"/>

        <xsl:param name="styleName"/>

        <xsl:variable name="escapedStyleName" select="loc:escapeStyleName($styleName)"/>

        <xsl:value-of select="concat('style_', $escapedStyleName, '_', generate-id($element))"/>

    </xsl:function>


    <!-- 
        Wrapping an output element into an anchor that delivers a style name 
    -->

    <xsl:template match="*" mode="hasStyleName" as="xs:boolean">

        <xsl:variable name="styleName">
            <xsl:apply-templates select="." mode="styleName"/>
        </xsl:variable>

        <xsl:sequence select="$styleName != ''"/>

    </xsl:template>


    <xsl:function name="loc:hasStyleName" as="xs:boolean">

        <xsl:param name="element"/>

        <xsl:apply-templates select="$element" mode="hasStyleName"/>

    </xsl:function>


    <xsl:template match="*" mode="styleAnchor">

        <xsl:param name="anchorContent"/>

        <xsl:param name="styleName"/>

        <xsl:choose>

            <xsl:when test="loc:hasStyleName(.)">

                <xsl:variable name="styleAnchorName" select="loc:styleAnchorName(., $styleName)"/>

                <xsl:call-template name="anchorPair">

                    <xsl:with-param name="anchorContent">
                        <xsl:copy-of select="$anchorContent"/>
                    </xsl:with-param>

                    <xsl:with-param name="anchorName" select="$styleAnchorName"/>

                </xsl:call-template>

            </xsl:when>

            <xsl:otherwise>
                <xsl:copy-of select="$anchorContent"/>
            </xsl:otherwise>

        </xsl:choose>

    </xsl:template>


    <!-- 
        Assembling output elements
    -->

    <xsl:template match="node()[name() = '']">
        <xsl:copy-of select="."/>
        
        <!--<xsl:value-of select="normalize-space(.)"/>-->
    </xsl:template>


    <xsl:template match="emphasis">

        <b>
            <xsl:apply-templates/>
        </b>

    </xsl:template>


    <xsl:template match="listitem/para">

        <xsl:apply-templates/>

    </xsl:template>


    <xsl:template match="listitem">

        <li>
            <xsl:apply-templates/>
        </li>

    </xsl:template>


    <xsl:template match="orderedlist">

        <xsl:apply-templates select="." mode="styleAnchor">

            <xsl:with-param name="anchorContent">
                <ol>
                    <xsl:apply-templates/>
                </ol>
            </xsl:with-param>

            <xsl:with-param name="styleName">
                <xsl:apply-templates select="." mode="styleName"/>
            </xsl:with-param>

        </xsl:apply-templates>

    </xsl:template>


    <xsl:template match="itemizedlist">

        <xsl:apply-templates select="." mode="styleAnchor">

            <xsl:with-param name="anchorContent">
                <ul>
                    <xsl:apply-templates/>
                </ul>
            </xsl:with-param>

            <xsl:with-param name="styleName">
                <xsl:apply-templates select="." mode="styleName"/>
            </xsl:with-param>

        </xsl:apply-templates>

    </xsl:template>


    <xsl:template match="para">

        <xsl:apply-templates select="." mode="styleAnchor">

            <xsl:with-param name="anchorContent">
                <p>
                    <xsl:apply-templates/>
                </p>
            </xsl:with-param>

            <xsl:with-param name="styleName">
                <xsl:apply-templates select="." mode="styleName"/>
            </xsl:with-param>

        </xsl:apply-templates>

    </xsl:template>


    <xsl:template match="section/title">

        <xsl:apply-templates select="." mode="styleAnchor">

            <xsl:with-param name="anchorContent">
                <xsl:element name="{concat('h', loc:level(.))}">
                    <xsl:apply-templates/>
                </xsl:element>
            </xsl:with-param>

            <xsl:with-param name="styleName">
                <xsl:apply-templates select="." mode="styleName"/>
            </xsl:with-param>

        </xsl:apply-templates>

    </xsl:template>


    <xsl:template match="tbody/row/entry">

        <td>
            <xsl:apply-templates/>
        </td>

    </xsl:template>


    <xsl:template match="thead/row/entry">

        <th>
            <xsl:apply-templates/>
        </th>

    </xsl:template>


    <xsl:template match="row">

        <tr>
            <xsl:apply-templates/>
        </tr>

    </xsl:template>


    <xsl:template match="colspec"/>


    <xsl:template match="tgroup | thead | tbody">

        <xsl:apply-templates/>

    </xsl:template>


    <xsl:template match="informaltable | table">

        <table>
            <xsl:apply-templates/>
        </table>

    </xsl:template>


    <xsl:template match="section/section">

        <xsl:apply-templates/>

    </xsl:template>


    <xsl:function name="loc:isFrontmatter" as="xs:boolean">

        <xsl:param name="sectionElement"/>

        <xsl:choose>
            <xsl:when test="upper-case($sectionElement/title) = 'ЛИСТ СОГЛАСОВАНИЯ'">
                <xsl:sequence select="true()"/>
            </xsl:when>
            <xsl:when test="upper-case($sectionElement/title) = 'ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ'">
                <xsl:sequence select="true()"/>
            </xsl:when>
            <xsl:when test="upper-case($sectionElement/title) = 'АННОТАЦИЯ'">
                <xsl:sequence select="true()"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:sequence select="false()"/>
            </xsl:otherwise>
        </xsl:choose>

    </xsl:function>


    <xsl:template match="article/section">

        <xsl:call-template name="anchorPair">

            <xsl:with-param name="anchorContent">
                <xsl:apply-templates select="section[loc:isFrontmatter(.)]"/>
            </xsl:with-param>

            <xsl:with-param name="anchorName" select="'part_frontmatter'"/>

        </xsl:call-template>

        <xsl:call-template name="anchorPair">

            <xsl:with-param name="anchorContent">
                <xsl:apply-templates select="section[not(loc:isFrontmatter(.))]"/>
            </xsl:with-param>

            <xsl:with-param name="anchorName" select="'part_main'"/>

        </xsl:call-template>

    </xsl:template>


    <xsl:template match="info"/>


    <xsl:template name="metaProps">

        <xsl:variable name="metaUri" select="
                resolve-uri(
                replace(tokenize(base-uri(/), '/')[last()], '\.xml$', '-meta.xml'),
                base-uri(/))"/>

        <xsl:variable name="meta" select="document($metaUri)"/>

        <xsl:for-each select="$meta//*[@name]">

            <xsl:call-template name="LFCR"/>

            <xsl:call-template name="anchorPair">

                <xsl:with-param name="anchorContent">
                    <xsl:value-of select="."/>
                </xsl:with-param>

                <xsl:with-param name="anchorName" select="concat('part_', @name)"/>

            </xsl:call-template>

        </xsl:for-each>

    </xsl:template>


    <xsl:template match="article">

        <html>
            <head>
                <title/>
            </head>
            <body>
                <xsl:call-template name="metaProps"/>
                <xsl:apply-templates/>
            </body>
        </html>

    </xsl:template>


    <xsl:template match="*">

        <xsl:copy>

            <xsl:copy-of select="@*"/>

            <xsl:apply-templates/>

        </xsl:copy>

    </xsl:template>


    <!--
        Creating an output document
    -->

    <xsl:template name="doctypeHtml5">
        <xsl:text disable-output-escaping="yes"><![CDATA[<!DOCTYPE html>]]></xsl:text>
    </xsl:template>


    <xsl:template match="/">

        <xsl:call-template name="doctypeHtml5"/>
        <xsl:call-template name="LFCR"/>

        <xsl:apply-templates/>

    </xsl:template>

</xsl:stylesheet>

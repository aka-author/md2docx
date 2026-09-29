<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xpath-default-namespace="http://docbook.org/ns/docbook" exclude-result-prefixes="#all"
    version="2.0">

    <xsl:output method="html" indent="yes"/>


    <xsl:template match="emphasis">

        <b>
            <xsl:apply-templates/>
        </b>

    </xsl:template>


    <xsl:template match="listitem">
        
        <li>
            <xsl:apply-templates/>
        </li>
        
    </xsl:template>
    
    
    <xsl:template match="orderedlist">
        
        <ol>
            <xsl:apply-templates/>
        </ol>
        
    </xsl:template>
    
    
    <xsl:template match="itemizedlist">
        
        <ul>
            <xsl:apply-templates/>
        </ul>
        
    </xsl:template>


    <xsl:template match="para">
        
        <p>
            <xsl:apply-templates/>
        </p>
        
    </xsl:template>


    <xsl:template match="title">

        <p>
            <xsl:apply-templates/>
        </p>

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
    

    <xsl:template match="tgroup|thead|tbody">

        <xsl:apply-templates/>

    </xsl:template>


    <xsl:template match="informaltable">

        <table>
            <xsl:apply-templates/>
        </table>

    </xsl:template>

    <xsl:template match="section">

        <div>
            <xsl:apply-templates/>
        </div>

    </xsl:template>


    <xsl:template match="info"/>


    <xsl:template match="article">

        <html>
            <head>
                <title/>
            </head>
            <body>
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


    <xsl:template name="LFCR">
        <xsl:text>&#10;&#13;</xsl:text>
    </xsl:template>


    <xsl:template name="doctypeHtml5">
        <xsl:text disable-output-escaping="yes"><![CDATA[<!DOCTYPE html>]]></xsl:text>
    </xsl:template>


    <xsl:template match="/">

        <xsl:call-template name="doctypeHtml5"/>
        <xsl:call-template name="LFCR"/>

        <xsl:apply-templates/>

    </xsl:template>

</xsl:stylesheet>

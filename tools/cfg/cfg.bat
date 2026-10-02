@echo off

set root=%~dp0..\..

set tools=%root%\tools
set tmp=%root%\tmp

set pandocExe=C:\Program Files\Pandoc\pandoc.exe

set mdmeta2xmlVbs=%tools%\md2html\vbs\mdmeta2xml.vbs 
set md2htmlBat=%tools%\md2html\md2html.bat
set saxonJar=c:\Saxon\saxon9.jar 
set xml2htmlXsl=%tools%\md2html\xsl\xml2html.xsl

set html2docxVbs=%tools%\html2docx\vbs\bookmarx.vbs
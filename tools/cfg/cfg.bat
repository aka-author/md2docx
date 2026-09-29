@echo off

set root=%~dp0..\..

set tools=%root%\tools
set tmp=%root%\tmp

set pandocExe=C:\Program Files\Pandoc\pandoc.exe

set md2htmlBat=%tools%\md2html\md2html.bat
set saxonJar=c:\Saxon\saxon9.jar 
set xml2htmlXsl=%tools%\md2html\xsl\xml2html.xsl
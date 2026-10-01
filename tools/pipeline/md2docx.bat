@echo off

call %~dp0..\cfg\cfg.bat

set srcDocFilePath=%~dpnx1
set srcDocFolderPath=%~dp1
set srcDocFileName=%~n1

echo. 
echo Publishing %srcDocFileName% to DOCX

set htmlDocFilePath=%tmp%\%srcDocFileName%.html

call "%md2htmlBat%" "%srcDocFilePath%" "%htmlDocFilePath%"

set docxTemplateFilePath=%~dpnx2
set "docxDocFolderPath=%~3\"
set docxDocFilePath=%docxDocFolderPath%%srcDocFileName%.docx

if not exist "%docxDocFolderPath%" md "%docxDocFolderPath%"

cscript "%html2docxVbs%" "%htmlDocFilePath%" "%docxTemplateFilePath%" "%docxDocFilePath%"

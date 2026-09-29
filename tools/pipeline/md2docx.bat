@echo off

call %~dp0..\cfg\cfg.bat

set srcDocFilePath=%~dpnx1
set srcDocFolderPath=%~dp1
set srcDocFileName=%~n1

echo. 
echo Publishing %srcDocFileName% to DOCX

set htmlDocFilePath=%tmp%\%srcDocName%.html

call "%md2htmlBat%" "%srcDocFilePath%" "%htmlDocFilePath%"
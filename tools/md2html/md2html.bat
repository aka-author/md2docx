@echo off

call %~dp0..\cfg\cfg.bat

set mdDocFilePath=%~dpnx1
set mdDocFileName=%~n1


echo Producing an interim HTML... 

rem Producing an interim XML

set xmlDocFilePath=%tmp%\%mdDocFileName%.xml

"%pandocExe%" "%mdDocFilePath%" -f markdown -t docbook -s -o "%xmlDocFilePath%"


rem Producing an interim HTML

set htmlDocFilePath=%tmp%\%mdDocFileName%.html

java -XX:-UsePerfData -jar "%saxonJar%" -s:"%xmlDocFilePath%" -xsl:"%xml2htmlXsl%" -o:"%htmlDocFilePath%"

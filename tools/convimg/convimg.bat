@echo off

call %~dp0..\cfg\cfg.bat

set srcImgFolder=%~1
if "%srcImgFolder%"=="" set srcImgFolder=%root%\content\img

set outImgFolder=%tmp%\img

if not exist "%outImgFolder%" md "%outImgFolder%"

echo Converting diagrams in "%srcImgFolder%" to SVG in "%outImgFolder%"...

for /r "%srcImgFolder%" %%F in (*.puml) do (
  java -XX:-UsePerfData %plantumlJvmEncoding% -DGRAPHVIZ_DOT="%graphvizDotExe%" -jar "%plantumlJar%" -charset UTF-8 -tsvg -o "%outImgFolder%" "%%F"
)

for /r "%srcImgFolder%" %%F in (*.gv) do (
  "%graphvizDotExe%" -Tsvg "%%F" -o "%outImgFolder%\%%~nF.svg"
)

@echo off
chcp 65001 >nul
set "EDGE_EXE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if exist "%EDGE_EXE%" (
  start "" "%EDGE_EXE%" "%~dp0index.html"
  exit /b
)
set "EDGE_EXE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
if exist "%EDGE_EXE%" (
  start "" "%EDGE_EXE%" "%~dp0index.html"
  exit /b
)
start "" "%~dp0index.html"

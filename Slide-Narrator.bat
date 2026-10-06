@echo off
rem Double-click to open Slide Narrator.
rem You can also drag one or more PowerPoint files onto this .bat file.
powershell -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Slide-Narrator.ps1" %*
if errorlevel 1 (
  echo.
  echo Slide Narrator stopped with an error. Please copy the red text above.
  pause
)

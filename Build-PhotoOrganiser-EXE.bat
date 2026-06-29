@echo off
REM ============================================
REM  Builds "Photo Organiser.exe"
REM  Double-click THIS file to run the build.
REM  (Runs the .ps1 with -ExecutionPolicy Bypass,
REM   the same way Run-Organise-Photos.bat does.)
REM ============================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build-PhotoOrganiser-EXE.ps1"

pause

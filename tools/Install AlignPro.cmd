@echo off
rem  AlignPro - install from an extracted release zip, without using a terminal.
rem
rem  -ExecutionPolicy Bypass is what makes this work: a .ps1 extracted by Explorer carries the mark
rem  of the web, and running one by path is otherwise refused outright - even with the execution
rem  policy at Unrestricted. Bypass skips that check for this one invocation only; nothing on the
rem  machine is reconfigured.
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-AlignPro.ps1"
echo.
pause

@echo off
call "%~dp0Start-Browser.cmd"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -Command "Start-Process 'http://127.0.0.1:8765/reset.html'"

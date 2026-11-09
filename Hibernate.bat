@echo off
rem Log END + push to GitHub first, then hibernate
powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\TimeLogger.ps1" -EventType "END"
shutdown /h

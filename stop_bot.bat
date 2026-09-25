@echo off
echo Stopping python bot processes...
taskkill /F /IM python.exe 2>nul
echo Bot stopped successfully.
timeout /t 3

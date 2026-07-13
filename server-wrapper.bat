@echo off
cd /d C:\Users\prati\finance-tracker
:restart
echo [%date% %time%] Starting finance tracker...
C:\Python314\python.exe app.py
echo [%date% %time%] Server exited with code %errorlevel%. Restarting in 2s...
timeout /t 2 /nobreak >nul
goto restart

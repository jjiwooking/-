@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 goto usepy
where python >nul 2>nul
if %errorlevel%==0 goto usepython
echo Python was not found. Install Python 3.10 or newer and try again.
pause
exit /b 1
:usepy
py launcher.py
if errorlevel 1 pause
exit /b %errorlevel%
:usepython
python launcher.py
if errorlevel 1 pause
exit /b %errorlevel%

@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  py tests\test_engine.py
) else (
  python tests\test_engine.py
)
pause

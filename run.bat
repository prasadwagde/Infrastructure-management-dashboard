@echo off
cd /d "%~dp0"
where py >nul 2>nul && set PY=py || set PY=python
%PY% -m pip install -r requirements.txt
if errorlevel 1 pause & exit /b 1
%PY% app.py
pause

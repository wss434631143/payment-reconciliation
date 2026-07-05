@echo off
setlocal
set "APP_DIR=%~dp0"
set "PYTHON=C:\Users\liudan\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHONPATH=%APP_DIR%..\..\work\pyside6_pkg;%PYTHONPATH%"
"%PYTHON%" "%APP_DIR%qt_app.py"
endlocal

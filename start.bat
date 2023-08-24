@echo off

set VENV_NAME=env
set VENV_PATH=%~dp0%VENV_NAME%


call %VENV_PATH%\Scripts\activate

python %~dp0%\src\main.py


pause
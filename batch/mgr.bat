@echo off

@REM @REM Pull all updates from github first
@REM git pull --rebase origin main
cd "%~dp"

@REM install packages if necessary
if not exist "%~dp0\packages\" (
    mkdir "%~dp0\packages
    python -m ensurepip --upgrade
    pip install --target=%~dp0\mgr_files\packages -r "%~dp0\..\requirements.txt"
)

@REM @REM Run the Python script
py %~dp0\mgr_files\gui.py

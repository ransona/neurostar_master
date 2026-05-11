@echo off
setlocal

cd /d C:\code\repos\neurostar_master

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on PATH.
    echo Install Python and tick "Add python.exe to PATH", then try again.
    pause
    exit /b 1
)

python -m pip show PySide6 >nul 2>nul
if errorlevel 1 (
    echo PySide6 is not installed. Installing it now...
    python -m pip install PySide6
    if errorlevel 1 (
        echo Failed to install PySide6.
        pause
        exit /b 1
    )
)

python .\tools\craniotomy_qt.py
if errorlevel 1 (
    echo.
    echo Craniotomy Planner exited with an error.
    pause
    exit /b 1
)

endlocal

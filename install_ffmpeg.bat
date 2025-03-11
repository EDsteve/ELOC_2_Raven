@echo off
echo ELOC Audio Processor - FFmpeg Installer
echo ======================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH.
    echo Please install Python 3.7 or higher from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo Running FFmpeg setup script...
python setup_ffmpeg.py

echo.
echo If FFmpeg was installed successfully, you can now run the ELOC Audio Processor.
echo.
pause

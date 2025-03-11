@echo off
echo ELOC Audio Processor Launcher
echo ============================
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

echo Installing required dependencies...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Failed to install dependencies.
    pause
    exit /b 1
)

REM Check if FFmpeg is installed
where ffmpeg >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo WARNING: FFmpeg is not found in your PATH.
    echo FFmpeg is required for audio processing.
    echo.
    echo Please install FFmpeg:
    echo 1. Download from https://ffmpeg.org/download.html
    echo 2. Extract the downloaded archive
    echo 3. Add the bin folder to your system PATH or place ffmpeg.exe in this directory
    echo.
    echo The application will still start, but audio processing may not work correctly.
    echo.
    pause
)

echo.
echo Starting ELOC Audio Processor...
echo.
python eloc_audio_processor.py

pause

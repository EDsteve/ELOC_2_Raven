@echo off
echo ELOC Audio Processor - Setup and Run
echo ==================================
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

echo Step 1: Installing Python dependencies...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Failed to install Python dependencies.
    pause
    exit /b 1
)

echo.
echo Step 2: Checking for FFmpeg...

REM Check if FFmpeg is already installed
where ffmpeg >nul 2>&1
if %errorlevel% equ 0 (
    echo FFmpeg is already installed and available in your PATH.
    goto start_app
)

REM Check if ffmpeg.exe exists in the current directory
if exist "ffmpeg.exe" (
    echo FFmpeg is already present in the current directory.
    goto start_app
)

echo FFmpeg is not found. It needs to be installed.
echo.
echo Would you like to automatically download and install FFmpeg? (Y/N)
choice /c YN /m "Download FFmpeg now?"

if %errorlevel% equ 1 (
    echo.
    echo Running FFmpeg setup script...
    python setup_ffmpeg.py
) else (
    echo.
    echo Please install FFmpeg manually:
    echo 1. Download from https://ffmpeg.org/download.html
    echo 2. Extract the downloaded archive
    echo 3. Add the bin folder to your system PATH or place ffmpeg.exe in this directory
    echo.
    echo The application will still start, but audio processing may not work correctly.
    pause
)

:start_app
echo.
echo Step 3: Starting ELOC Audio Processor...
echo.
echo NEW FEATURE: You can now drag and drop folders from Windows Explorer directly into the application!
echo.
python eloc_audio_processor.py

pause

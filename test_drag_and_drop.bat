@echo off
echo ELOC Audio Processor - Drag and Drop Test
echo =======================================
echo.
echo This script will launch the ELOC Audio Processor with drag and drop support.
echo.
echo Instructions:
echo 1. Drag folders from Windows Explorer into the application window
echo 2. Drag individual WAV or CSV files (will add their parent folder)
echo 3. Drag multiple folders or files at once
echo.
echo The application will automatically scan the dropped folders and select those with CSV files.
echo.
echo Press any key to start the test...
pause > nul

python test_eloc_processor.py

pause

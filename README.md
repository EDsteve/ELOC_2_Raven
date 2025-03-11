# ELOC Audio Processor

A Windows application for processing ELOC audio data from SD cards. This application provides a modern, futuristic GUI for selecting and processing audio files and CSV data from ELOC folders.

## Features

- Select SD card drive from a dropdown menu
- Automatically scan for ELOC folders and subfolders
- Display WAV and CSV file counts for each subfolder
- Select one or more subfolders for processing
- One-click selection of all folders containing CSV files
- Drag and drop support for folders and files from Windows Explorer
- Adjustable time offset and audio segment length parameters
- Background processing with status updates
- Creates Raven Selection Tables and extracts audio segments

## Requirements

- Windows operating system
- Python 3.7 or higher
- Required Python packages (see requirements.txt)
- FFmpeg (required for audio processing)

## Installation

### Quick Start (Recommended)

For the easiest setup, simply run:

```
setup_and_run.bat
```

This all-in-one batch file will:
1. Install required Python dependencies
2. Check for FFmpeg and offer to install it if not found
3. Launch the application

### Manual Installation

If you prefer to install components separately:

1. Clone or download this repository
2. Install the required Python dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Install FFmpeg (required for audio processing):
   
   **Option 1:** Use the automatic installer:
   ```
   install_ffmpeg.bat
   ```
   This will download and set up FFmpeg in the current directory.
   
   **Option 2:** Manual installation:
   - Download FFmpeg from [ffmpeg.org](https://ffmpeg.org/download.html)
   - Extract the downloaded archive
   - Add the bin folder to your system PATH or place the ffmpeg.exe file in the same directory as this application

4. Run the application:
   ```
   run_eloc_processor.bat
   ```

## Usage

1. Run the application:

```
python eloc_audio_processor.py
```

2. Select your SD card from the dropdown menu
3. The application will scan for ELOC folders and display them in the list
4. Adjust the time offset and segment length parameters if needed
5. Select the folders you want to process (or use "Select All Folders with CSV")
6. Click "Process Selected Folders" to start processing
7. Monitor progress in the status bar at the bottom of the window

### Drag and Drop Support

The application now supports drag and drop from Windows Explorer:

- Drag and drop folders containing WAV and CSV files directly into the application window
- Drag and drop individual WAV or CSV files (the application will add their parent folder)
- Drag and drop multiple folders or files at once
- The application will automatically scan the dropped folders and select those with CSV files

## Output

For each processed folder, the application creates:

- `output/Raven_Selection_Tables/` - Contains selection tables for Raven software
- `output/Audio_Segments/` - Contains extracted audio segments based on the selection tables

## Parameters

- **Time Offset**: Adjusts the begin time of audio segments relative to the detected event (default: -2 seconds)
- **Segment Length**: Sets the duration of extracted audio segments (default: 5 seconds)

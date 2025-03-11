"""
Test script for ELOC Audio Processor with drag and drop functionality.
This script simply launches the application so you can test the drag and drop feature.
"""

from eloc_audio_processor import ElocAudioProcessor

if __name__ == "__main__":
    print("Starting ELOC Audio Processor with drag and drop support...")
    print("You can now drag folders from Windows Explorer into the application window.")
    print("Drag and drop works for:")
    print("- Folders containing WAV and CSV files")
    print("- Individual WAV or CSV files (will add their parent folder)")
    print("- Multiple folders or files at once")
    
    app = ElocAudioProcessor()
    app.mainloop()

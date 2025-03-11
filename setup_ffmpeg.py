import os
import sys
import zipfile
import urllib.request
import shutil
from pathlib import Path

def download_progress(count, block_size, total_size):
    """Show download progress"""
    percent = int(count * block_size * 100 / total_size)
    sys.stdout.write(f"\rDownloading FFmpeg: {percent}% [{count * block_size} / {total_size} bytes]")
    sys.stdout.flush()

def main():
    print("FFmpeg Setup Utility for ELOC Audio Processor")
    print("=============================================")
    print()
    
    # Check if FFmpeg is already in PATH
    if shutil.which("ffmpeg"):
        print("FFmpeg is already installed and available in your PATH.")
        print("No further action is needed.")
        input("Press Enter to exit...")
        return
    
    # Check if ffmpeg.exe exists in the current directory
    if os.path.exists("ffmpeg.exe"):
        print("FFmpeg is already present in the current directory.")
        print("No further action is needed.")
        input("Press Enter to exit...")
        return
    
    print("This script will download and set up FFmpeg for Windows.")
    print("FFmpeg will be installed in the current directory.")
    print()
    
    # URL for FFmpeg (using a static build for Windows)
    ffmpeg_url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip"
    
    try:
        # Create a temporary directory
        temp_dir = Path("temp_ffmpeg")
        os.makedirs(temp_dir, exist_ok=True)
        
        # Download FFmpeg
        zip_path = temp_dir / "ffmpeg.zip"
        print(f"Downloading FFmpeg from {ffmpeg_url}")
        print("This may take a few minutes depending on your internet connection...")
        
        urllib.request.urlretrieve(ffmpeg_url, zip_path, reporthook=download_progress)
        print("\nDownload complete!")
        
        # Extract the zip file
        print("Extracting FFmpeg...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Find the bin directory in the extracted files
        bin_dirs = list(temp_dir.glob("*/bin"))
        if not bin_dirs:
            print("Error: Could not find bin directory in the extracted files.")
            return
        
        bin_dir = bin_dirs[0]
        
        # Copy ffmpeg.exe to the current directory
        print("Installing FFmpeg to the current directory...")
        shutil.copy(bin_dir / "ffmpeg.exe", "ffmpeg.exe")
        
        # Clean up
        print("Cleaning up temporary files...")
        shutil.rmtree(temp_dir)
        
        print()
        print("FFmpeg has been successfully installed!")
        print("The ffmpeg.exe file is now in the current directory.")
        print("You can now run the ELOC Audio Processor.")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        print("\nFFmpeg installation failed.")
        print("Please download FFmpeg manually from https://ffmpeg.org/download.html")
        print("and place ffmpeg.exe in this directory or add it to your system PATH.")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()

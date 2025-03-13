import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import string
import ctypes
import concurrent.futures
from datetime import datetime
from pydub import AudioSegment
import csv
import time
import logging
import sys
from tkinterdnd2 import DND_FILES, TkinterDnD

class ElocAudioProcessor(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        
        # Set window properties
        self.title("ELOC to Raven")
        self.geometry("700x800")
        self.configure(bg="#54613b")  # Background for main window
        
        # Set default values
        self.time_offset = -2
        self.segment_length = 5
        self.selected_folders = []
        
        # Set up logging to file
        self.log_file_path = "eloc_progress_log.txt"
        # Create or clear the log file
        with open(self.log_file_path, 'w') as f:
            f.write(f"ELOC Audio Processor Log - Started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("-" * 80 + "\n\n")
        
        # Create a style for ttk widgets
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use clam theme as base
        
        # Configure styles with new color scheme
        # Color scheme explanation:
        # #626F47 - Dark olive green - Used for frames, labels, and standard buttons backgrounds
        # #FEFAE0 - Off-white/cream - Used for text on buttons and labels
        # #e7e8a6 - Light green/beige - Used for combobox, treeview, and spinbox backgrounds
        # #da9432 - Golden amber - Used for accent buttons (like "Process Selected Folders")
        # #000000 -  - Used for combobox text
        # #000000 - Black - Used for combobox readonly field background
        
        self.style.configure('TFrame', background='#54613b')  # Dark olive green frames
        self.style.configure('TLabel', background='#54613b', foreground='#f3faeb', font=('Segoe UI', 10))  # Dark green labels with cream text
        
        # Standard buttons 
        self.style.configure('TButton', background='#424d2f', foreground='#f3faeb', font=('Segoe UI', 10), 
                            borderwidth=0, relief='flat', padding=(15, 10))  # Add padding (horizontal, vertical)
        # Add hover (active) state for standard buttons 
        self.style.map('TButton', 
                      background=[('active', '#424d2f')],
                      foreground=[('active', '#da9432')])
        
        # Accent buttons (like "Process Selected Folders") - Golden amber with cream text
        self.style.configure('Accent.TButton', background='#da9432', foreground='#FEFAE0', font=('Segoe UI', 16, 'bold'), 
                            borderwidth=0, relief='flat')
        # Add hover (active) state for accent buttons - brighter amber when hovered
        self.style.map('Accent.TButton', 
                      background=[('active', '#f5a93a')],

                      foreground=[('active', '#FFFFFF')])
        
        # Checkbuttons - Light green/beige background with cream text
        self.style.configure('TCheckbutton', background='#54613b', foreground='#FEFAE0', font=('Segoe UI', 10))
        # Remove hover effect for checkbuttons by setting the same background color for active state
        self.style.map('TCheckbutton', 
                      background=[('active', '#54613b')],
                      foreground=[('active', '#FEFAE0')])
        
        # Configure Treeview (folder list) with light green background
        self.style.configure('Treeview', background='#e7e8a6', fieldbackground='#e7e8a6')
        self.style.map('Treeview', background=[('selected', '#da9432')])  # Amber for selected items
        
        # Combobox - White background with brown text, light green field background
        self.style.configure('TCombobox', background='white', foreground='#8c4511', fieldbackground='#e7e8a6')
        self.style.map('TCombobox', fieldbackground=[('readonly', '#e7e8a6')])  # Light green background when readonly
        self.style.map('TCombobox', selectbackground=[('readonly', '#e7e8a6')])  # Light green selection background
        self.style.map('TCombobox', selectforeground=[('readonly', '#e7e8a6')])  # Light green selection text
        
        # Configure Spinbox (for offset and segment length) with light green background
        self.style.configure('TSpinbox', fieldbackground='#e7e8a6')
        
        # Create main frame
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create and place widgets
        self.create_widgets()
        
    def create_widgets(self):
        # Header section with logo and title
        header_frame = ttk.Frame(self.main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 35))  # Increased bottom margin
        
        # Load the ELOC icon with size reduction and transparency
        try:
            from PIL import Image, ImageTk
            
            # Open the image with PIL
            original_image = Image.open("ELOC-icon.png")
            
            # Calculate the resize ratio to get 150px width while maintaining aspect ratio
            width_percent = (80 / float(original_image.size[0]))
            new_height = int((float(original_image.size[1]) * float(width_percent)))
            
            # Resize the image
            resized_image = original_image.resize((80, new_height), Image.LANCZOS)
            
            # Add 30% transparency (70% opacity)
            if resized_image.mode != 'RGBA':
                resized_image = resized_image.convert('RGBA')
            
            # Create a new image with transparency
            data = resized_image.getdata()
            new_data = []
            for item in data:
                # Change all pixels to have 70% of their original opacity
                new_data.append((item[0], item[1], item[2], int(item[3] * 0.7) if len(item) > 3 else int(255 * 0.7)))
            
            resized_image.putdata(new_data)
            
            # Convert to PhotoImage for tkinter
            icon_image = ImageTk.PhotoImage(resized_image)
            
            # Display the image
            icon_label = ttk.Label(header_frame, image=icon_image, background='#54613b')
            icon_label.image = icon_image  # Keep a reference to prevent garbage collection
            icon_label.pack(side=tk.LEFT, padx=(0, 30))  # Increased right margin
        except Exception as e:
            print(f"Error loading or processing icon: {str(e)}")
        
        # Title and subtitle in a vertical frame
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # Main title "ELOC to Raven" in big, bold letters
        title_label = ttk.Label(title_frame, text="ELOC to Raven conveter", 
                               font=('Segoe UI', 24, 'bold'), background='#54613b', foreground='#20241d')
        title_label.pack(anchor=tk.W)
        
        # Subtitle in smaller letters
        subtitle_label = ttk.Label(title_frame, text="Creates Selection Tables & Extracts Audio Snippets", 
                                  font=('Segoe UI', 16), background='#54613b', foreground='#20241d')
        subtitle_label.pack(anchor=tk.W)
        
        # Add help icon on the right side
        try:
            from PIL import Image, ImageTk
            
            # Open the help icon image
            help_image = Image.open("help.png")
            
            # Resize the help icon to appropriate size
            help_image = help_image.resize((30, 30), Image.LANCZOS)
            
            # Convert to PhotoImage for tkinter
            help_icon = ImageTk.PhotoImage(help_image)
            
            # Create a label for the help icon and place it on the right side of the header
            help_label = ttk.Label(header_frame, image=help_icon, background='#54613b', cursor="hand2")
            help_label.image = help_icon  # Keep a reference to prevent garbage collection
            help_label.pack(side=tk.RIGHT, padx=10)
            
            # Bind click event to open README
            help_label.bind("<Button-1>", lambda e: self.open_readme())
        except Exception as e:
            print(f"Error loading help icon: {str(e)}")
        
        # Drive selection section
        drive_frame = ttk.Frame(self.main_frame)
        drive_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(drive_frame, text="Select SD Card Drive:", font=('Segoe UI', 12, 'bold')).pack(side=tk.LEFT, padx=(0, 10))
        
        # Dropdown for drive selection
        self.drive_var = tk.StringVar()
        self.drive_combo = ttk.Combobox(drive_frame, textvariable=self.drive_var, state='readonly', width=10)
        self.drive_combo.pack(side=tk.LEFT)
        self.drive_combo.bind('<<ComboboxSelected>>', self.on_drive_selected)
        
        # Refresh button for drives
        ttk.Button(drive_frame, text="Refresh Drives", command=self.refresh_drives).pack(side=tk.LEFT, padx=10)
        
        # Select custom folder button
        ttk.Button(drive_frame, text="Select Custom Folder", command=self.select_custom_folder).pack(side=tk.LEFT, padx=10)
        
        # Folder list section
        folder_frame = ttk.Frame(self.main_frame)
        folder_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        ttk.Label(folder_frame, text="Select Folders to Process or Drag & Drop from Windows Explorer", 
                 font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
        # Create a frame for the folder list with scrollbar
        list_frame = ttk.Frame(folder_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview for folder list
        columns = ("Folder", "WAV Files", "CSV Files")
        self.folder_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        self.folder_tree.pack(fill=tk.BOTH, expand=True)
        
        # Set up drag and drop for the folder list
        self.folder_tree.drop_target_register(DND_FILES)
        self.folder_tree.dnd_bind('<<Drop>>', self.on_drop)
        
        # Configure scrollbar
        scrollbar.config(command=self.folder_tree.yview)
        self.folder_tree.config(yscrollcommand=scrollbar.set)
        
        # Configure columns
        self.folder_tree.heading("Folder", text="Folder")
        self.folder_tree.heading("WAV Files", text="WAV Files")
        self.folder_tree.heading("CSV Files", text="Valid CSV File")
        
        self.folder_tree.column("Folder", width=400)
        self.folder_tree.column("WAV Files", width=100, anchor=tk.CENTER)
        self.folder_tree.column("CSV Files", width=100, anchor=tk.CENTER)
        
        # Button frame
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Button(button_frame, text="Select Valid Folders", 
                  command=self.select_folders_with_csv).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Clear Selection", 
                  command=self.clear_selection).pack(side=tk.LEFT)
        
        # Parameters section (moved above the process button)
        param_frame = ttk.Frame(self.main_frame)
        param_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(param_frame, text="Begin-Time Offset (seconds):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.time_offset_var = tk.DoubleVar(value=self.time_offset)
        ttk.Spinbox(param_frame, from_=-10, to=10, increment=0.5, textvariable=self.time_offset_var, width=10, 
                   style='TSpinbox').grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(param_frame, text="Snippet Length (seconds):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.segment_length_var = tk.DoubleVar(value=self.segment_length)
        ttk.Spinbox(param_frame, from_=1, to=30, increment=1, textvariable=self.segment_length_var, width=10,
                   style='TSpinbox').grid(row=1, column=1, padx=5, pady=5)
        
        # Processing options
        ttk.Label(param_frame, text="Processing Options:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=5)
        
        # Checkbuttons for processing options
        self.create_tables_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(param_frame, text="Create Raven Selection Tables", 
                       variable=self.create_tables_var, state='disabled').grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        self.extract_audio_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(param_frame, text="Cut and Copy Detected Soundfiles", 
                       variable=self.extract_audio_var).grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Process button
        process_button = ttk.Button(self.main_frame, text="Process Selected Folders", 
                                   command=self.process_folders, style='Accent.TButton')
        process_button.pack(fill=tk.X, ipady=10)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, relief='flat', anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(20, 0))
        
        # Initialize drives list
        self.refresh_drives()
        
    def get_removable_drives(self):
        """Get a list of removable drives using ctypes"""
        drives = []
        bitmask = ctypes.windll.kernel32.GetLogicalDrives()
        
        for letter in string.ascii_uppercase:
            if bitmask & 1:
                drive_path = f"{letter}:\\"
                drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive_path)
                # DRIVE_REMOVABLE = 2
                if drive_type == 2:  # Removable drive
                    drives.append(drive_path)
            bitmask >>= 1
            
        return drives
    
    def refresh_drives(self):
        """Refresh the list of available drives"""
        removable_drives = self.get_removable_drives()
        
        if removable_drives:
            self.drive_combo['values'] = removable_drives
            self.drive_combo.current(0)
            self.on_drive_selected(None)
        else:
            self.drive_combo['values'] = ["No SD Card"]
            self.drive_combo.current(0)
            self.folder_tree.delete(*self.folder_tree.get_children())
            messagebox.showinfo("No Removable Drives", "No SD card or removable drives detected.")
    
    def on_drive_selected(self, event):
        """Handle drive selection event"""
        selected_drive = self.drive_var.get()
        self.scan_eloc_folders(selected_drive)
    
    def select_custom_folder(self):
        """Allow user to select a custom folder instead of using SD card"""
        folder_path = filedialog.askdirectory(title="Select Parent Folder")
        if folder_path:
            # Clear current list
            self.folder_tree.delete(*self.folder_tree.get_children())
            
            # Update drive dropdown to show custom path
            self.drive_combo['values'] = ["Custom Folder"]
            self.drive_var.set("Custom Folder")
            
            # Store the custom folder path for processing
            self.custom_folder_path = folder_path
            
            # Check if the selected folder directly contains wav and csv files
            wav_count = len(glob.glob(os.path.join(folder_path, "*.wav")))
            csv_count = len(glob.glob(os.path.join(folder_path, "*.csv")))
            
            if wav_count > 0 and csv_count > 0:
                # The selected folder directly contains wav and csv files
                folder_name = os.path.basename(folder_path)
                if not folder_name:  # In case the path ends with a separator
                    folder_name = os.path.basename(os.path.dirname(folder_path))
                
                # Use a special marker to indicate this is the root folder itself
                self.folder_tree.insert("", tk.END, values=(f"[ROOT] {folder_name}", wav_count, csv_count))
                self.status_var.set(f"Selected folder with {wav_count} WAV files and {csv_count} CSV files")
                
                # Flag to indicate we're using the root folder directly
                self.using_root_folder = True
                return
            
            # If we get here, check for subfolders
            try:
                subfolders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
                
                if not subfolders:
                    # If no subfolders and no wav/csv files in root, show message
                    if wav_count == 0 or csv_count == 0:
                        self.status_var.set(f"No valid ELOC data found in {folder_path}")
                        messagebox.showinfo("No Data Found", 
                                           "The selected folder doesn't contain WAV and CSV files or valid subfolders.")
                    else:
                        # If no subfolders but has wav/csv files, add the selected folder itself
                        folder_name = os.path.basename(folder_path)
                        if not folder_name:  # In case the path ends with a separator
                            folder_name = os.path.basename(os.path.dirname(folder_path))
                        
                        # Use a special marker to indicate this is the root folder itself
                        self.folder_tree.insert("", tk.END, values=(f"[ROOT] {folder_name}", wav_count, csv_count))
                        self.status_var.set(f"Selected folder: {folder_path}")
                        
                        # Flag to indicate we're using the root folder directly
                        self.using_root_folder = True
                else:
                    # Count files in each subfolder
                    for folder in subfolders:
                        subfolder_path = os.path.join(folder_path, folder)
                        subfolder_wav_count = len(glob.glob(os.path.join(subfolder_path, "*.wav")))
                        subfolder_csv_count = len(glob.glob(os.path.join(subfolder_path, "*.csv")))
                        
                        # Add to treeview
                        self.folder_tree.insert("", tk.END, values=(folder, subfolder_wav_count, subfolder_csv_count))
                    
                    self.status_var.set(f"Found {len(subfolders)} subfolders in {folder_path}")
            except Exception as e:
                self.status_var.set(f"Error scanning folder: {str(e)}")
    
    def scan_eloc_folders(self, drive_path):
        """Scan for ELOC folders on the selected drive"""
        # Clear current list
        self.folder_tree.delete(*self.folder_tree.get_children())
        
        # Check if 'eloc' folder exists
        eloc_path = os.path.join(drive_path, "eloc")
        if not os.path.exists(eloc_path):
            self.status_var.set(f"No 'eloc' folder found on {drive_path}")
            return
        
        # Get all subfolders
        try:
            subfolders = [f for f in os.listdir(eloc_path) if os.path.isdir(os.path.join(eloc_path, f))]
            
            if not subfolders:
                self.status_var.set(f"No subfolders found in {eloc_path}")
                return
            
            # Count files in each subfolder
            for folder in subfolders:
                folder_path = os.path.join(eloc_path, folder)
                wav_count = len(glob.glob(os.path.join(folder_path, "*.wav")))
                csv_count = len(glob.glob(os.path.join(folder_path, "*.csv")))
                
                # Add to treeview
                self.folder_tree.insert("", tk.END, values=(folder, wav_count, csv_count))
            
            self.status_var.set(f"Found {len(subfolders)} folders in {eloc_path}")
            
        except Exception as e:
            self.status_var.set(f"Error scanning folders: {str(e)}")
    
    def select_folders_with_csv(self):
        """Select all folders that contain CSV files"""
        self.folder_tree.selection_set()  # Clear current selection
        
        for item in self.folder_tree.get_children():
            values = self.folder_tree.item(item, "values")
            if int(values[2]) > 0:  # If CSV count > 0
                self.folder_tree.selection_add(item)
    
    def clear_selection(self):
        """Clear the folder selection"""
        self.folder_tree.selection_set()  # Clear current selection
    
    def process_folders(self):
        """Process the selected folders"""
        selected_items = self.folder_tree.selection()
        if not selected_items:
            messagebox.showinfo("Selection Required", "Please select at least one folder to process.")
            return
        
        # Get selected folders
        selected_folders = []
        for item in selected_items:
            folder_name = self.folder_tree.item(item, "values")[0]
            selected_folders.append(folder_name)
        
        # Get parameters
        time_offset = self.time_offset_var.get()
        segment_length = self.segment_length_var.get()
        
        # Get drive path
        drive_path = self.drive_var.get()
        
        # Start processing in a separate thread
        self.status_var.set("Processing started...")
        threading.Thread(target=self.run_processing, 
                        args=(drive_path, selected_folders, time_offset, segment_length),
                        daemon=True).start()
    
    def run_processing(self, drive_path, selected_folders, time_offset, segment_length):
        """Run the processing in a background thread with parallel processing"""
        try:
            total_folders = len(selected_folders)
            self.update_status(f"Starting to process {total_folders} folders... Please wait.")
            
            # Prepare folder paths and output directories
            folder_paths = []
            selection_tables_dirs = []
            audio_segments_dirs = []
            
            for folder in selected_folders:
                # Check if we're using a custom folder or SD card
                if drive_path == "Custom Folder" and hasattr(self, 'custom_folder_path'):
                    # Check if this is the root folder (marked with [ROOT])
                    if folder.startswith("[ROOT]"):
                        # Use the custom folder path directly
                        folder_path = self.custom_folder_path
                    else:
                        # For custom folder, join the custom folder path with the subfolder name
                        folder_path = os.path.join(self.custom_folder_path, folder)
                else:
                    # For SD card, construct path as before
                    folder_path = os.path.join(drive_path, "eloc", folder)
                
                # Create output directories
                output_base = os.path.join(folder_path, "output")
                selection_tables_dir = os.path.join(output_base, "Raven_Selection_Tables")
                audio_segments_dir = os.path.join(output_base, "Audio_Segments")
                
                os.makedirs(selection_tables_dir, exist_ok=True)
                os.makedirs(audio_segments_dir, exist_ok=True)
                
                folder_paths.append(folder_path)
                selection_tables_dirs.append(selection_tables_dir)
                audio_segments_dirs.append(audio_segments_dir)
            
            # Process folders with parallel execution
            self.update_status(f"Processing {total_folders} folders in parallel... Please wait.")
            
            # Use a ThreadPoolExecutor for parallel processing
            # Limit the number of workers to avoid overloading the system
            max_workers = min(os.cpu_count() or 4, total_folders)
            self.update_status(f"Using {max_workers} parallel workers for processing.")
            
            start_time = time.time()
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Submit all folder processing tasks
                futures = []
                for i, (folder_path, selection_tables_dir, audio_segments_dir) in enumerate(
                    zip(folder_paths, selection_tables_dirs, audio_segments_dirs), 1
                ):
                    self.update_status(f"Submitting folder {i}/{total_folders}: {os.path.basename(folder_path)}")
                    future = executor.submit(
                        self.process_folder_parallel,
                        folder_path, 
                        selection_tables_dir, 
                        audio_segments_dir, 
                        time_offset, 
                        segment_length,
                        i,
                        total_folders
                    )
                    futures.append(future)
                
                # Wait for all tasks to complete and handle results
                for i, future in enumerate(concurrent.futures.as_completed(futures), 1):
                    try:
                        result = future.result()
                        self.update_status(f"Completed folder {i}/{total_folders}: {result}")
                    except Exception as e:
                        self.update_status(f"Error processing folder {i}: {str(e)}")
            
            end_time = time.time()
            processing_time = end_time - start_time
            
            self.update_status(f"Processing complete! All folders processed in {processing_time:.2f} seconds.")
            messagebox.showinfo("Processing Complete", 
                               f"All selected folders have been processed successfully in {processing_time:.2f} seconds.")
            
        except Exception as e:
            self.update_status(f"Error during processing: {str(e)}")
            messagebox.showerror("Processing Error", f"An error occurred: {str(e)}")
    
    def process_folder_parallel(self, folder_path, selection_tables_dir, audio_segments_dir, 
                               time_offset, segment_length, folder_index, total_folders):
        """Process a single folder in a parallel thread"""
        try:
            # Update status with thread-safe method
            self.update_status(f"Processing folder {folder_index}/{total_folders}: {os.path.basename(folder_path)}...")
            
            # Process the folder using the existing method
            self.process_folder(folder_path, selection_tables_dir, audio_segments_dir, time_offset, segment_length)
            
            # Return the folder name for status updates
            return os.path.basename(folder_path)
        except Exception as e:
            # Re-raise the exception to be caught by the executor
            raise Exception(f"Error processing {os.path.basename(folder_path)}: {str(e)}")
    
    def update_status(self, message):
        """Update status bar from a background thread and write to log file"""
        # Update the status bar
        self.after(0, lambda: self.status_var.set(message))
        
        # Write to log file
        try:
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with open(self.log_file_path, 'a') as f:
                f.write(f"[{timestamp}] {message}\n")
        except Exception as e:
            print(f"Error writing to log file: {str(e)}")
    
    def process_folder(self, folder_path, selection_tables_dir, audio_segments_dir, time_offset, segment_length):
        """Process a single folder (similar to the original scripts but adapted)"""
        # Get WAV files and their start times
        self.update_status(f"Scanning for WAV files in {os.path.basename(folder_path)}... Please wait.")
        wav_files = glob.glob(os.path.join(folder_path, "*.wav"))
        
        if not wav_files:
            self.update_status(f"No WAV files found in {folder_path}")
            return
        
        self.update_status(f"Found {len(wav_files)} WAV files. Extracting timestamps... Please wait.")
            
        wav_start_times = {}
        month_map = {
            '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
            '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
        }
        
        for wav_file in wav_files:
            datetime_str = self.extract_datetime_from_filename(os.path.basename(wav_file))
            if datetime_str:
                date_parts = datetime_str.split()[0].split('-')
                if len(date_parts) == 3:
                    year, month_num, day = date_parts
                    month_text = month_map.get(month_num, month_num)
                    formatted_date = f"{year}-{month_text}-{day}"
                    recording_id = formatted_date + '_' + datetime_str.split()[1][:2] + ':00:00'
                    seconds = self.datetime_to_seconds(datetime_str)
                    wav_start_times[recording_id] = seconds
        
        # Process CSV files
        self.update_status(f"Scanning for CSV files in {os.path.basename(folder_path)}... Please wait.")
        csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
        
        if not csv_files:
            self.update_status(f"No CSV files found in {folder_path}")
            return
            
        self.update_status(f"Found {len(csv_files)} CSV files. Processing detection data... Please wait.")
        
        for csv_file in csv_files:
            try:
                # Load the CSV file
                data = pd.read_csv(csv_file)
                
                # Strip spaces from column names
                data.columns = data.columns.str.strip()
                
                # Split the 'Month Date Year' column into separate columns
                date_columns = data['Month Date Year'].str.strip().str.split(expand=True)
                
                if date_columns.shape[1] == 3:
                    data['Month'] = date_columns[0]
                    data['Date'] = date_columns[1]
                    data['Year'] = date_columns[2]
                else:
                    self.update_status(f"Unexpected date format in file {os.path.basename(csv_file)}")
                    continue
                
                # Extract start time and group detections by one-hour recordings
                data['Recording_Start_Time'] = data['Hour:Min:Sec Day'].apply(lambda x: x.split()[0])
                data['Recording_Seconds'] = data['Recording_Start_Time'].apply(self.time_to_seconds)
                
                # Create Recording_ID
                data['Recording_ID'] = data.apply(
                    lambda row: f"{row['Year']}-{row['Month']}-{row['Date']}_{row['Recording_Start_Time'][:2]}:00:00", 
                    axis=1
                )
                
                # Group events by recording ID
                grouped = data.groupby('Recording_ID')
                
                for recording_id, group in grouped:
                    try:
                        date_part, start_time = recording_id.split('_')
                        year, month, date = date_part.split('-')
                    except ValueError:
                        self.update_status(f"Skipping incorrectly formatted Recording_ID: {recording_id}")
                        continue
                    
                    # Determine the start of the recording in seconds from the WAV file
                    if recording_id in wav_start_times:
                        start_seconds = wav_start_times[recording_id]
                    else:
                        self.update_status(f"No matching WAV file found for {recording_id}, using minimum event time instead.")
                        start_seconds = group['Recording_Seconds'].min()
                    
                    # Check if selection tables should be created
                    if self.create_tables_var.get():
                        self.update_status(f"Creating Raven selection table for recording {recording_id}... Please wait.")
                        # Initialize selection table content
                        raven_table_content = "Selection\tView\tChannel\tBegin Time (s)\tEnd Time (s)\tLow Freq (Hz)\tHigh Freq (Hz)\n"
                        
                        # Iterate over all detected events in this recording
                        for i, row in group.iterrows():
                            # Calculate begin time with adjustable offset
                            event_start_seconds = (row['Recording_Seconds'] - start_seconds) + time_offset
                            event_end_seconds = event_start_seconds + segment_length
                            
                            raven_table_content += (
                                f"{i+1}\tSpectrogram 1\t1\t{event_start_seconds:.2f}\t{event_end_seconds:.2f}\t"
                                f"{row['background'] * 1000:.2f}\t{row['trumpet'] * 5000:.2f}\n"
                            )
                        
                        # Fix Windows filename issue (replace ':' with '-')
                        file_name = f"{recording_id.replace(':', '-')}_SelectionTable.txt"
                        output_path = os.path.join(selection_tables_dir, file_name)
                        
                        with open(output_path, 'w') as f:
                            f.write(raven_table_content)
                        
                        self.update_status(f"Selection table created for {recording_id}")
                
                    # Check if audio segments should be extracted
                    if self.extract_audio_var.get():
                        self.update_status(f"Starting audio segment extraction for {recording_id}... Please wait.")
                        self.extract_audio_segments(folder_path, selection_tables_dir, audio_segments_dir)
                
            except Exception as e:
                self.update_status(f"Error processing CSV file {os.path.basename(csv_file)}: {str(e)}")
    
    def extract_audio_segments(self, folder_path, selection_tables_dir, audio_segments_dir):
        """Extract audio segments based on selection tables using optimized approach with parallel processing"""
        self.update_status(f"Scanning for selection tables in {os.path.basename(selection_tables_dir)}... Please wait.")
        selection_tables = glob.glob(os.path.join(selection_tables_dir, "*.txt"))
        
        if not selection_tables:
            self.update_status("No selection tables found to extract audio segments from.")
            return
            
        self.update_status(f"Found {len(selection_tables)} selection tables. Starting audio extraction... Please wait.")
        
        # Group segments by WAV file to avoid loading the same file multiple times
        segments_by_wav = {}
        
        # First pass: Parse all selection tables and organize segments by WAV file
        for selection_table in selection_tables:
            # Find the corresponding WAV file
            wav_file = self.find_wav_file(selection_table, folder_path)
            if not wav_file:
                self.update_status(f"No matching WAV file found for {os.path.basename(selection_table)}, skipping.")
                continue
            
            # Parse the selection table
            try:
                with open(selection_table, 'r') as f:
                    # Skip the header line
                    header = f.readline()
                    
                    # Read the rest of the lines
                    reader = csv.reader(f, delimiter='\t')
                    for i, row in enumerate(reader, 1):
                        if len(row) >= 5:  # Ensure we have enough columns
                            try:
                                # Extract begin and end times in seconds
                                begin_time = float(row[3])
                                end_time = float(row[4])
                                
                                # Store segment info
                                if wav_file not in segments_by_wav:
                                    segments_by_wav[wav_file] = []
                                
                                segments_by_wav[wav_file].append({
                                    'begin_time': begin_time,
                                    'end_time': end_time,
                                    'segment_id': i,
                                    'selection_table': os.path.basename(selection_table)
                                })
                            except Exception as e:
                                self.update_status(f"Error parsing segment {i} in {os.path.basename(selection_table)}: {e}")
            except Exception as e:
                self.update_status(f"Error reading selection table {os.path.basename(selection_table)}: {e}")
        
        # Second pass: Process WAV files in parallel
        total_wav_files = len(segments_by_wav)
        if total_wav_files == 0:
            self.update_status("No segments to extract.")
            return
            
        self.update_status(f"Processing {total_wav_files} WAV files in parallel... Please wait.")
        
        # Determine the number of workers for parallel processing
        # Use fewer workers for WAV processing to avoid memory issues
        max_workers = min(os.cpu_count() or 2, total_wav_files, 4)  # Limit to 4 max to avoid memory issues
        self.update_status(f"Using {max_workers} parallel workers for audio extraction.")
        
        start_time = time.time()
        
        # Process WAV files in parallel
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Prepare arguments for each WAV file
            wav_tasks = []
            for wav_file, segments in segments_by_wav.items():
                # Sort segments by begin time to optimize sequential access
                segments.sort(key=lambda x: x['begin_time'])
                wav_tasks.append((wav_file, segments, audio_segments_dir))
            
            # Submit all WAV processing tasks
            futures = {executor.submit(self.process_wav_file, *task, wav_index, total_wav_files): task 
                      for wav_index, task in enumerate(wav_tasks, 1)}
            
            # Process results as they complete
            for future in concurrent.futures.as_completed(futures):
                try:
                    wav_file, num_segments = future.result()
                    self.update_status(f"Completed processing {os.path.basename(wav_file)} with {num_segments} segments.")
                except Exception as e:
                    wav_task = futures[future]
                    wav_file = wav_task[0]
                    self.update_status(f"Error processing WAV file {os.path.basename(wav_file)}: {str(e)}")
        
        end_time = time.time()
        processing_time = end_time - start_time
        self.update_status(f"Audio extraction complete! Processed {total_wav_files} WAV files in {processing_time:.2f} seconds.")
    
    def process_wav_file(self, wav_file, segments, audio_segments_dir, wav_index, total_wav_files):
        """Process a single WAV file and extract all its segments"""
        try:
            self.update_status(f"Processing WAV file {wav_index}/{total_wav_files}: {os.path.basename(wav_file)}...")
            
            # Get total segments for this WAV file
            total_segments = len(segments)
            self.update_status(f"Found {total_segments} segments to extract from {os.path.basename(wav_file)}.")
            
            # Use pydub's segment extraction with frame-accurate seeking
            audio = AudioSegment.from_file(wav_file, format="wav")
            base_name = os.path.splitext(os.path.basename(wav_file))[0]
            
            # Track how many segments were actually processed
            processed_segments = 0
            
            # Process all segments for this WAV file
            for segment_index, segment_info in enumerate(segments, 1):
                begin_time = segment_info['begin_time']
                end_time = segment_info['end_time']
                segment_id = segment_info['segment_id']
                
                # Convert to milliseconds for pydub
                begin_ms = int(begin_time * 1000)
                end_ms = int(end_time * 1000)
                
                # Generate output filename
                segment_filename = f"{base_name}_segment_{segment_id:03d}_{begin_time:.2f}s-{end_time:.2f}s.wav"
                segment_path = os.path.join(audio_segments_dir, segment_filename)
                
                # Check if segment already exists (to avoid reprocessing)
                if os.path.exists(segment_path):
                    if segment_index % 10 == 0:  # Only update status every 10 segments to reduce UI updates
                        self.update_status(f"Segment {segment_index}/{total_segments} already exists, skipping.")
                    continue
                
                # Extract and export the segment directly
                if segment_index % 10 == 0:  # Only update status every 10 segments
                    self.update_status(f"Exporting segment {segment_index}/{total_segments} from {os.path.basename(wav_file)}...")
                
                # Extract the segment and export it in one operation
                segment = audio[begin_ms:end_ms]
                segment.export(segment_path, format="wav")
                processed_segments += 1
            
            # Free memory
            del audio
            
            # Return the WAV file name and number of segments processed for status updates
            return wav_file, processed_segments
            
        except Exception as e:
            # Re-raise the exception to be caught by the executor
            raise Exception(f"Error processing {os.path.basename(wav_file)}: {str(e)}")
    
    # Helper functions
    def extract_datetime_from_filename(self, filename):
        """Extract datetime from WAV filename"""
        parts = filename.split('_')
        if len(parts) >= 4:
            date_part = parts[2]  # "2025-03-10"
            time_part = parts[3].replace('.wav', '')  # "18-14-52"
            time_part = time_part.replace('-', ':')  # Convert to "18:14:52"
            return f"{date_part} {time_part}"
        return None
    
    def datetime_to_seconds(self, datetime_str):
        """Convert datetime string to seconds since midnight"""
        dt = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')
        return dt.hour * 3600 + dt.minute * 60 + dt.second
    
    def time_to_seconds(self, time_str):
        """Convert time to seconds since midnight"""
        time_obj = datetime.strptime(time_str, '%H:%M:%S')
        return time_obj.hour * 3600 + time_obj.minute * 60 + time_obj.second
    
    def month_name_to_number(self, month_name):
        """Convert month name to number"""
        month_map = {
            'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
            'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
        }
        return month_map.get(month_name, month_name)
    
    def open_readme(self):
        """Open the README.md file when the help icon is clicked"""
        try:
            readme_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "README.md")
            if os.path.exists(readme_path):
                # Use the default system application to open the README file
                if sys.platform == 'win32':
                    os.startfile(readme_path)
                elif sys.platform == 'darwin':  # macOS
                    import subprocess
                    subprocess.call(['open', readme_path])
                else:  # Linux
                    import subprocess
                    subprocess.call(['xdg-open', readme_path])
                
                self.status_var.set("Opened README file")
            else:
                messagebox.showinfo("README Not Found", "The README.md file could not be found.")
                self.status_var.set("README file not found")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open README file: {str(e)}")
            self.status_var.set(f"Error opening README: {str(e)}")
    
    def on_drop(self, event):
        """Handle dropped files/folders from Windows Explorer"""
        # Get the dropped data
        data = event.data
        
        # Debug the raw data format
        self.update_status(f"Raw drop data: {data}")
        
        # Parse the paths from the data
        paths = []
        
        # Handle different possible formats
        if data.startswith('{') and data.endswith('}'):
            # Format: {path1} {path2} {path3}
            # Remove the outer braces
            data = data[1:-1]
            
            # Split by '} {' for paths with spaces
            if '} {' in data:
                paths = data.split('} {')
            else:
                # Try to split by space, but only if the space is between paths
                # This is a heuristic and might need adjustment
                parts = data.split()
                current_path = ""
                
                for part in parts:
                    if os.path.exists(part) or (current_path == "" and part.endswith(':')):
                        # This is likely the start of a new path
                        if current_path:
                            paths.append(current_path)
                        current_path = part
                    else:
                        # This is likely part of the current path
                        if current_path:
                            current_path += " " + part
                        else:
                            current_path = part
                
                # Add the last path
                if current_path:
                    paths.append(current_path)
                
                # If we couldn't parse any paths, treat the whole string as one path
                if not paths:
                    paths = [data]
        else:
            # For space-separated paths without braces
            # First, try to split by space and check if each part is a valid path
            parts = data.split()
            
            # Check if each part is a valid path
            valid_paths = []
            for part in parts:
                if os.path.exists(part):
                    valid_paths.append(part)
            
            if valid_paths:
                paths = valid_paths
            else:
                # If no valid paths found, try to split by space and check if the resulting paths exist
                # This is for cases like "D:/path1 D:/path2"
                potential_paths = []
                current_path = ""
                
                for part in parts:
                    if part.startswith("D:") or part.startswith("C:") or part.startswith("/"):
                        # This looks like the start of a new path
                        if current_path:
                            potential_paths.append(current_path)
                        current_path = part
                    else:
                        # This is likely part of the current path
                        if current_path:
                            current_path += " " + part
                        else:
                            current_path = part
                
                # Add the last path
                if current_path:
                    potential_paths.append(current_path)
                
                # Check if the potential paths exist
                valid_potential_paths = []
                for path in potential_paths:
                    if os.path.exists(path):
                        valid_potential_paths.append(path)
                
                if valid_potential_paths:
                    paths = valid_potential_paths
                else:
                    # If still no valid paths found, treat the whole string as one path
                    paths = [data]
        
        # Debug the parsed paths
        for i, path in enumerate(paths):
            self.update_status(f"Parsed path {i+1}: {path}")
        
        # Update drive dropdown to show custom path
        self.drive_combo['values'] = ["Custom Folder"]
        self.drive_var.set("Custom Folder")
        
        # Clear current list only once
        self.folder_tree.delete(*self.folder_tree.get_children())
        
        # Store all parent directories for files
        parent_dirs = set()
        root_folders = []
        
        # First pass: collect all directories to process
        for path in paths:
            # Remove any quotes or braces
            path = path.strip('"{}')
            
            # Check if the path exists
            if not os.path.exists(path):
                self.update_status(f"Invalid path: {path}")
                continue
            
            # Check if it's a directory
            if os.path.isdir(path):
                # Add to the list of directories to process
                root_folders.append(path)
            else:
                # It's a file, check if it's a WAV or CSV file
                if path.lower().endswith('.wav') or path.lower().endswith('.csv'):
                    # Get the parent directory
                    parent_dir = os.path.dirname(path)
                    parent_dirs.add(parent_dir)
                else:
                    self.update_status(f"Unsupported file type: {path}")
        
        # Add parent directories of files to the list of directories to process
        root_folders.extend(list(parent_dirs))
        
        # Set the custom folder path to the first valid directory
        if root_folders:
            self.custom_folder_path = root_folders[0]
        
        # Process all collected directories
        folders_added = 0
        
        for path in root_folders:
            # Check if the folder directly contains WAV and CSV files
            wav_count = len(glob.glob(os.path.join(path, "*.wav")))
            csv_count = len(glob.glob(os.path.join(path, "*.csv")))
            
            if wav_count > 0 and csv_count > 0:
                # The folder directly contains WAV and CSV files
                folder_name = os.path.basename(path)
                if not folder_name:  # In case the path ends with a separator
                    folder_name = os.path.basename(os.path.dirname(path))
                
                # Use a special marker to indicate this is the root folder itself
                self.folder_tree.insert("", tk.END, values=(f"[ROOT] {folder_name}", wav_count, csv_count))
                folders_added += 1
                
                # Flag to indicate we're using the root folder directly
                self.using_root_folder = True
            else:
                # Check for subfolders
                try:
                    subfolders = [f for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))]
                    
                    if not subfolders:
                        # If no subfolders and no WAV/CSV files in root, show message
                        if wav_count == 0 or csv_count == 0:
                            self.update_status(f"No valid ELOC data found in {path}")
                        else:
                            # If no subfolders but has WAV/CSV files, add the folder itself
                            folder_name = os.path.basename(path)
                            if not folder_name:  # In case the path ends with a separator
                                folder_name = os.path.basename(os.path.dirname(path))
                            
                            # Use a special marker to indicate this is the root folder itself
                            self.folder_tree.insert("", tk.END, values=(f"[ROOT] {folder_name}", wav_count, csv_count))
                            folders_added += 1
                            
                            # Flag to indicate we're using the root folder directly
                            self.using_root_folder = True
                    else:
                        # Count files in each subfolder
                        subfolder_count = 0
                        for folder in subfolders:
                            subfolder_path = os.path.join(path, folder)
                            subfolder_wav_count = len(glob.glob(os.path.join(subfolder_path, "*.wav")))
                            subfolder_csv_count = len(glob.glob(os.path.join(subfolder_path, "*.csv")))
                            
                            # Add to treeview
                            self.folder_tree.insert("", tk.END, values=(folder, subfolder_wav_count, subfolder_csv_count))
                            subfolder_count += 1
                            folders_added += 1
                        
                        self.update_status(f"Found {subfolder_count} subfolders in {path}")
                except Exception as e:
                    self.update_status(f"Error scanning folder: {str(e)}")
        
        # Update status with total count
        if folders_added > 0:
            self.status_var.set(f"Added {folders_added} folders from drag and drop")
        else:
            self.status_var.set("No valid folders found from drag and drop")
        
        # Select all folders with CSV files
        self.select_folders_with_csv()
        
        return "break"  # Prevent further handling of the drop event
    
    def find_wav_file(self, selection_table_filename, folder_path):
        """Find the corresponding WAV file for a selection table"""
        parts = os.path.basename(selection_table_filename).split('_')
        if len(parts) >= 2:
            date_part = parts[0]  # "2025-Mar-10"
            hour_part = parts[1].split('-')[0]  # "18"
            
            # Convert month name to number
            date_parts = date_part.split('-')
            if len(date_parts) == 3:
                year, month_name, day = date_parts
                month_num = self.month_name_to_number(month_name)
                numeric_date = f"{year}-{month_num}-{day}"
                
                # Look for WAV files that match this date and hour
                wav_pattern = f"*_{numeric_date}_{hour_part}-*-*.wav"
                wav_files = glob.glob(os.path.join(folder_path, wav_pattern))
                
                if wav_files:
                    return wav_files[0]  # Return the first matching WAV file
                
                # If no match found with the specific pattern, try a more general search
                # This is useful when the WAV file naming convention might be different
                self.update_status(f"No exact match found, trying general search for hour {hour_part}")
                for wav_file in glob.glob(os.path.join(folder_path, "*.wav")):
                    filename = os.path.basename(wav_file)
                    if numeric_date in filename and f"_{hour_part}-" in filename:
                        return wav_file
        
        return None

if __name__ == "__main__":
    app = ElocAudioProcessor()
    app.mainloop()

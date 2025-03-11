import os
import glob
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import string
import ctypes
from datetime import datetime
from pydub import AudioSegment
import csv

class ElocAudioProcessor(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Set window properties
        self.title("ELOC Audio Processor")
        self.geometry("900x700")
        self.configure(bg="#54613b")  # Background for main window
        
        # Set default values
        self.time_offset = -2
        self.segment_length = 5
        self.selected_folders = []
        
        # Create a style for ttk widgets
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use clam theme as base
        
        # Configure styles with new color scheme
        # Color scheme explanation:
        # #626F47 - Dark olive green - Used for frames, labels, and standard buttons backgrounds
        # #FEFAE0 - Off-white/cream - Used for text on buttons and labels
        # #e7e8a6 - Light green/beige - Main window background and checkbutton background
        # #da9432 - Golden amber - Used for accent buttons (like "Process Selected Folders")
        # #8c4511 - Brown - Used for combobox text
        # #000000 - Black - Used for combobox readonly field background
        
        self.style.configure('TFrame', background='#54613b')  # Dark olive green frames
        self.style.configure('TLabel', background='#54613b', foreground='#FEFAE0', font=('Segoe UI', 10))  # Dark green labels with cream text
        
        # Standard buttons 
        self.style.configure('TButton', background='#424d2f', foreground='#FEFAE0', font=('Segoe UI', 10), 
                            borderwidth=0, relief='flat', padding=(15, 10))  # Add padding (horizontal, vertical)
        # Add hover (active) state for standard buttons - lighter olive green when hovered
        self.style.map('TButton', 
                      background=[('active', '#515c3b')],
                      foreground=[('active', '#da9432')])
        
        # Accent buttons (like "Process Selected Folders") - Golden amber with cream text
        self.style.configure('Accent.TButton', background='#da9432', foreground='#FEFAE0', font=('Segoe UI', 12, 'bold'), 
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
        
        # Combobox - White background with brown text, light green field background
        self.style.configure('TCombobox', background='white', foreground='#8c4511', fieldbackground='#e7e8a6')
        self.style.map('TCombobox', fieldbackground=[('readonly', '#e7e8a6')])  # Black background when readonly
        self.style.map('TCombobox', selectbackground=[('readonly', '#e7e8a6')])  # Light green selection background
        self.style.map('TCombobox', selectforeground=[('readonly', '#e7e8a6')])  # Light green selection text
        
        # Create main frame
        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create and place widgets
        self.create_widgets()
        
    def create_widgets(self):
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
        
        # Parameters section
        param_frame = ttk.Frame(self.main_frame)
        param_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(param_frame, text="Time Offset (seconds):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.time_offset_var = tk.DoubleVar(value=self.time_offset)
        ttk.Spinbox(param_frame, from_=-10, to=10, increment=0.5, textvariable=self.time_offset_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(param_frame, text="Segment Length (seconds):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.segment_length_var = tk.DoubleVar(value=self.segment_length)
        ttk.Spinbox(param_frame, from_=1, to=30, increment=1, textvariable=self.segment_length_var, width=10).grid(row=1, column=1, padx=5, pady=5)
        
        # Processing options
        ttk.Label(param_frame, text="Processing Options:").grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=5)
        
        # Checkbuttons for processing options
        self.create_tables_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(param_frame, text="Create Raven Selection Tables", 
                       variable=self.create_tables_var, state='disabled').grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        
        self.extract_audio_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(param_frame, text="Cut and Copy Detected Soundfiles", 
                       variable=self.extract_audio_var).grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)
        
        # Folder list section
        folder_frame = ttk.Frame(self.main_frame)
        folder_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        ttk.Label(folder_frame, text="Available ELOC Folders:", font=('Segoe UI', 12, 'bold')).pack(anchor=tk.W, pady=(0, 10))
        
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
        
        # Configure scrollbar
        scrollbar.config(command=self.folder_tree.yview)
        self.folder_tree.config(yscrollcommand=scrollbar.set)
        
        # Configure columns
        self.folder_tree.heading("Folder", text="Folder")
        self.folder_tree.heading("WAV Files", text="WAV Files")
        self.folder_tree.heading("CSV Files", text="CSV Files")
        
        self.folder_tree.column("Folder", width=400)
        self.folder_tree.column("WAV Files", width=100, anchor=tk.CENTER)
        self.folder_tree.column("CSV Files", width=100, anchor=tk.CENTER)
        
        # Button frame
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Button(button_frame, text="Select All Folders with CSV", 
                  command=self.select_folders_with_csv).pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Clear Selection", 
                  command=self.clear_selection).pack(side=tk.LEFT)
        
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
            
            # Scan for subfolders
            try:
                subfolders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
                
                if not subfolders:
                    # If no subfolders, add the selected folder itself
                    wav_count = len(glob.glob(os.path.join(folder_path, "*.wav")))
                    csv_count = len(glob.glob(os.path.join(folder_path, "*.csv")))
                    
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
                        wav_count = len(glob.glob(os.path.join(subfolder_path, "*.wav")))
                        csv_count = len(glob.glob(os.path.join(subfolder_path, "*.csv")))
                        
                        # Add to treeview
                        self.folder_tree.insert("", tk.END, values=(folder, wav_count, csv_count))
                    
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
        """Run the processing in a background thread"""
        try:
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
                
                self.update_status(f"Processing folder: {os.path.basename(folder_path)}")
                
                # Create output directories
                output_base = os.path.join(folder_path, "output")
                selection_tables_dir = os.path.join(output_base, "Raven_Selection_Tables")
                audio_segments_dir = os.path.join(output_base, "Audio_Segments")
                
                os.makedirs(selection_tables_dir, exist_ok=True)
                os.makedirs(audio_segments_dir, exist_ok=True)
                
                # Process CSV and WAV files
                self.process_folder(folder_path, selection_tables_dir, audio_segments_dir, time_offset, segment_length)
            
            self.update_status("Processing complete!")
            messagebox.showinfo("Processing Complete", "All selected folders have been processed successfully.")
            
        except Exception as e:
            self.update_status(f"Error during processing: {str(e)}")
            messagebox.showerror("Processing Error", f"An error occurred: {str(e)}")
    
    def update_status(self, message):
        """Update status bar from a background thread"""
        self.after(0, lambda: self.status_var.set(message))
    
    def process_folder(self, folder_path, selection_tables_dir, audio_segments_dir, time_offset, segment_length):
        """Process a single folder (similar to the original scripts but adapted)"""
        # Get WAV files and their start times
        wav_files = glob.glob(os.path.join(folder_path, "*.wav"))
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
        csv_files = glob.glob(os.path.join(folder_path, "*.csv"))
        
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
                        
                        self.update_status(f"Selection tables created for {os.path.basename(csv_file)}")
                
                # Check if audio segments should be extracted
                if self.extract_audio_var.get():
                    self.extract_audio_segments(folder_path, selection_tables_dir, audio_segments_dir)
                
            except Exception as e:
                self.update_status(f"Error processing CSV file {os.path.basename(csv_file)}: {str(e)}")
    
    def extract_audio_segments(self, folder_path, selection_tables_dir, audio_segments_dir):
        """Extract audio segments based on selection tables"""
        selection_tables = glob.glob(os.path.join(selection_tables_dir, "*.txt"))
        
        for selection_table in selection_tables:
            # Find the corresponding WAV file
            wav_file = self.find_wav_file(selection_table, folder_path)
            if not wav_file:
                self.update_status(f"No matching WAV file found for {os.path.basename(selection_table)}, skipping.")
                continue
            
            # Load the WAV file
            try:
                audio = AudioSegment.from_wav(wav_file)
            except Exception as e:
                self.update_status(f"Error loading audio file: {e}")
                continue
            
            # Parse the selection table
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
                            
                            # Convert to milliseconds for pydub
                            begin_ms = int(begin_time * 1000)
                            end_ms = int(end_time * 1000)
                            
                            # Extract the segment
                            segment = audio[begin_ms:end_ms]
                            
                            # Generate output filename
                            base_name = os.path.splitext(os.path.basename(wav_file))[0]
                            segment_filename = f"{base_name}_segment_{i:03d}_{begin_time:.2f}s-{end_time:.2f}s.wav"
                            segment_path = os.path.join(audio_segments_dir, segment_filename)
                            
                            # Export the segment
                            segment.export(segment_path, format="wav")
                            
                        except Exception as e:
                            self.update_status(f"Error processing segment {i}: {e}")
                            continue
    
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
        
        return None

if __name__ == "__main__":
    app = ElocAudioProcessor()
    app.mainloop()

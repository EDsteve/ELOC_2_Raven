import pandas as pd
import os
from datetime import datetime, timedelta

# Variable to adjust begin time relative to actual timestamp
# Negative values will decrease the begin time, positive values will increase it
time_offset = -2  # Default is 0 (no adjustment)

# Get the current directory where the script is located
current_directory = os.path.dirname(os.path.abspath(__file__))

# Function to extract datetime from WAV filename
def extract_datetime_from_filename(filename):
    # Expected format: test2_1741605292068_2025-03-10_18-14-52.wav
    parts = filename.split('_')
    if len(parts) >= 4:
        date_part = parts[2]  # "2025-03-10"
        time_part = parts[3].replace('.wav', '')  # "18-14-52"
        time_part = time_part.replace('-', ':')  # Convert to "18:14:52"
        return f"{date_part} {time_part}"
    return None

# Function to convert datetime string to seconds since midnight
def datetime_to_seconds(datetime_str):
    # Expected format: "2025-03-10 18:14:52"
    dt = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M:%S')
    return dt.hour * 3600 + dt.minute * 60 + dt.second

# Get WAV files and their start times
wav_files = [f for f in os.listdir(current_directory) if f.endswith('.wav')]
wav_start_times = {}
month_map = {
    '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
    '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
}

print(f"Found {len(wav_files)} WAV files in the directory.")
for wav_file in wav_files:
    datetime_str = extract_datetime_from_filename(wav_file)
    if datetime_str:
        # Parse the date to convert numeric month to text month
        date_parts = datetime_str.split()[0].split('-')
        if len(date_parts) == 3:
            year, month_num, day = date_parts
            month_text = month_map.get(month_num, month_num)
            formatted_date = f"{year}-{month_text}-{day}"
            recording_id = formatted_date + '_' + datetime_str.split()[1][:2] + ':00:00'
            seconds = datetime_to_seconds(datetime_str)
            wav_start_times[recording_id] = seconds

# List all CSV files in the current directory
csv_files = [f for f in os.listdir(current_directory) if f.endswith('.csv')]

# Check if there are any CSV files in the directory
if not csv_files:
    print("No CSV files found in the directory.")
else:
    for csv_file in csv_files:
        # Construct the full file path
        file_path = os.path.join(current_directory, csv_file)
        
        # Load the CSV file
        data = pd.read_csv(file_path)

        # Strip spaces from column names
        data.columns = data.columns.str.strip()

        # Split the 'Month Date Year' column into separate columns
        date_columns = data['Month Date Year'].str.strip().str.split(expand=True)

        if date_columns.shape[1] == 3:
            data['Month'] = date_columns[0]
            data['Date'] = date_columns[1]
            data['Year'] = date_columns[2]
        else:
            print(f"Unexpected date format in file {csv_file}")
            continue

        # Create a folder for Raven selection tables
        output_folder = 'Raven_Selection_Tables'
        os.makedirs(output_folder, exist_ok=True)

        # Function to convert time to seconds since midnight
        def time_to_seconds(time_str):
            time_obj = datetime.strptime(time_str, '%H:%M:%S')
            return time_obj.hour * 3600 + time_obj.minute * 60 + time_obj.second

        # Extract start time and group detections by one-hour recordings
        data['Recording_Start_Time'] = data['Hour:Min:Sec Day'].apply(lambda x: x.split()[0])  # Extract time part
        data['Recording_Seconds'] = data['Recording_Start_Time'].apply(time_to_seconds)

        # ✅ Corrected `Recording_ID` formatting
        data['Recording_ID'] = data.apply(lambda row: f"{row['Year']}-{row['Month']}-{row['Date']}_{row['Recording_Start_Time'][:2]}:00:00", axis=1)

        # Process the data by recording ID

        # Group events by recording ID (one file per one-hour recording)
        grouped = data.groupby('Recording_ID')

        for recording_id, group in grouped:
            # ✅ Fix parsing of `recording_id`
            try:
                date_part, start_time = recording_id.split('_')
                year, month, date = date_part.split('-')
            except ValueError:
                print(f"Skipping incorrectly formatted Recording_ID: {recording_id}")
                continue

            # Determine the start of the recording in seconds from the WAV file
            if recording_id in wav_start_times:
                start_seconds = wav_start_times[recording_id]
            else:
                print(f"Warning: No matching WAV file found for {recording_id}, using minimum event time instead.")
                start_seconds = group['Recording_Seconds'].min()  # Fallback to minimum event time

            # Initialize selection table content
            raven_table_content = "Selection\tView\tChannel\tBegin Time (s)\tEnd Time (s)\tLow Freq (Hz)\tHigh Freq (Hz)\n"

            # Iterate over all detected events in this recording
            for i, row in group.iterrows():
                # Calculate begin time with adjustable offset
                event_start_seconds = (row['Recording_Seconds'] - start_seconds) + time_offset  # Apply time offset
                event_end_seconds = event_start_seconds + 5  # Fixed event duration of 5 sec

                raven_table_content += (
                    f"{i+1}\tSpectrogram 1\t1\t{event_start_seconds:.2f}\t{event_end_seconds:.2f}\t"
                    f"{row['background'] * 1000:.2f}\t{row['trumpet'] * 5000:.2f}\n"
                )

            # ✅ Fix Windows filename issue (replace `:` with `-`)
            file_name = f"{recording_id.replace(':', '-')}_SelectionTable.txt"
            output_path = os.path.join(output_folder, file_name)

            with open(output_path, 'w') as f:
                f.write(raven_table_content)

        print(f"Selection tables created for {csv_file}.")

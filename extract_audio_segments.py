import os
import glob
import csv
from pydub import AudioSegment

# Get the current directory where the script is located
current_directory = os.path.dirname(os.path.abspath(__file__))

# Define paths
selection_tables_dir = os.path.join(current_directory, "Raven_Selection_Tables")
output_dir = os.path.join(current_directory, "Audio_Segments")

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

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

# Function to convert month name to number
def month_name_to_number(month_name):
    month_map = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
        'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
    }
    return month_map.get(month_name, month_name)

# Function to find the corresponding WAV file for a selection table
def find_wav_file(selection_table_filename):
    # Extract date and hour from selection table filename
    # Expected format: 2025-Mar-10_18-00-00_SelectionTable.txt
    parts = os.path.basename(selection_table_filename).split('_')
    if len(parts) >= 2:
        date_part = parts[0]  # "2025-Mar-10"
        hour_part = parts[1].split('-')[0]  # "18"
        
        # Convert month name to number
        date_parts = date_part.split('-')
        if len(date_parts) == 3:
            year, month_name, day = date_parts
            month_num = month_name_to_number(month_name)
            numeric_date = f"{year}-{month_num}-{day}"
            
            # Look for WAV files that match this date and hour
            wav_pattern = f"*_{numeric_date}_{hour_part}-*-*.wav"
            wav_files = glob.glob(os.path.join(current_directory, wav_pattern))
            
            if wav_files:
                return wav_files[0]  # Return the first matching WAV file
    
    return None

# Process all selection tables
selection_tables = glob.glob(os.path.join(selection_tables_dir, "*.txt"))
print(f"Found {len(selection_tables)} selection tables.")

for selection_table in selection_tables:
    print(f"Processing {os.path.basename(selection_table)}...")
    
    # Find the corresponding WAV file
    wav_file = find_wav_file(selection_table)
    if not wav_file:
        print(f"  No matching WAV file found for {os.path.basename(selection_table)}, skipping.")
        continue
    
    print(f"  Found matching WAV file: {os.path.basename(wav_file)}")
    
    # Load the WAV file
    try:
        audio = AudioSegment.from_wav(wav_file)
        print(f"  Loaded audio file: {os.path.basename(wav_file)}")
    except Exception as e:
        print(f"  Error loading audio file: {e}")
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
                    segment_path = os.path.join(output_dir, segment_filename)
                    
                    # Export the segment
                    segment.export(segment_path, format="wav")
                    print(f"  Extracted segment {i}: {begin_time:.2f}s - {end_time:.2f}s")
                    
                except Exception as e:
                    print(f"  Error processing segment {i}: {e}")
                    continue
    
    print(f"  Finished processing {os.path.basename(selection_table)}")

print("Audio segment extraction complete!")

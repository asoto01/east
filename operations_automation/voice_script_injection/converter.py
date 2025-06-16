# convert.py

from pydub import AudioSegment
import os

# Define the input and output filenames
m4a_file = "english_vm.m4a"
wav_file = "english_vm.wav"

# Check if the input file exists
if not os.path.exists(m4a_file):
    print(f"Error: The file '{m4a_file}' was not found in this directory.")
else:
    try:
        # Load the M4A file
        print(f"Loading '{m4a_file}'...")
        audio = AudioSegment.from_file(m4a_file, format="m4a")

        # Export it as a WAV file
        print(f"Converting and exporting to '{wav_file}'...")
        audio.export(wav_file, format="wav")

        print("\nConversion successful!")
        print(f"A new file named '{wav_file}' has been created.")
    except Exception as e:
        print(f"An error occurred during conversion: {e}")

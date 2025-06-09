import sounddevice as sd
import soundfile as sf
from pydub import AudioSegment
from pydub.playback import play # Keep if you use it elsewhere, otherwise it's not strictly necessary for sounddevice playback
import os
import time

# --- Configuration ---
ENGLISH_VOICEMAIL_PATH = "english_vm.wav"
SPANISH_VOICEMAIL_DUMMY = AudioSegment.silent(duration=5000) # 5 seconds of silence
VIRTUAL_DEVICE_NAME = "BlackHole 2ch"

def get_default_device_info():
    """Gets the names of the current default audio input and output devices."""
    try:
        input_device_index = sd.default.device[0]
        output_device_index = sd.default.device[1]
        input_info = sd.query_devices(input_device_index)
        output_info = sd.query_devices(output_device_index)
        return input_info['name'], output_info['name']
    except Exception as e:
        print(f"Error getting default device info: {e}")
        return None, None

def play_audio(file_path=None, dummy_segment=None, output_device_id=None):
    """Plays an audio file or a pydub audio segment to a specific output device ID."""
    try:
        data = None
        fs = None

        if file_path and os.path.exists(file_path):
            data, fs = sf.read(file_path, dtype='float32')
        elif dummy_segment:
            # Export pydub AudioSegment to numpy array for sounddevice
            # Ensure correct sample width and channels for conversion
            data = dummy_segment.get_array_of_samples()
            fs = dummy_segment.frame_rate
            
            # Convert to float32 as sounddevice often prefers it
            if dummy_segment.sample_width == 2: # 16-bit
                data = (data / 32768.0).astype('float32')
            elif dummy_segment.sample_width == 4: # 32-bit
                 data = (data / 2147483648.0).astype('float32') # Max for 32-bit int
            
            # Reshape for stereo if necessary (sounddevice expects (samples, channels) for stereo)
            if dummy_segment.channels == 2:
                data = data.reshape(-1, 2)
            elif dummy_segment.channels == 1 and data.ndim == 1:
                # If mono, ensure it's treated as mono (1D array is fine)
                pass 
            else:
                print(f"Warning: Unexpected channel/dimension combination for pydub segment: channels={dummy_segment.channels}, data.ndim={data.ndim}")

        else:
            print("Audio file not found or no audio to play.")
            return

        if data is not None and fs is not None:
            print(f"Attempting to play audio to device ID: {output_device_id}")
            # Use OutputStream for more control and explicit device selection
            with sd.OutputStream(samplerate=fs, channels=data.shape[1] if data.ndim > 1 else 1, device=output_device_id) as stream:
                stream.write(data)
            print("Playback finished.")
    except Exception as e:
        print(f"Error playing audio: {e}")
        print("Available devices (for debugging):")
        print(sd.query_devices())


def main_loop():
    """Main function to run the voicemail player tool."""
    original_input_device_name, original_output_device_name = get_default_device_info()
    if not original_input_device_name or not original_output_device_name:
        print("Could not determine the original audio devices. Exiting.")
        return

    print("--- Voicemail Player ---")
    print(f"Original audio input device: {original_input_device_name}")
    print(f"Original audio output device: {original_output_device_name}")

    while True:
        print("\n--- MENU ---")
        action = input("Enter 'p' to play a voicemail, or 'q' to quit: ").lower()

        if action == 'q':
            print("Exiting program.")
            break

        if action == 'p':
            lang_choice = input("Play voicemail: Enter '1' for English, '2' for Spanish: ")

            if lang_choice in ['1', '2']:
                # Find the index of BlackHole for output
                blackhole_output_index = None
                devices = sd.query_devices()
                for i, device in enumerate(devices):
                    # Look for BlackHole with output capabilities
                    if VIRTUAL_DEVICE_NAME in device['name'] and device['max_output_channels'] > 0:
                        blackhole_output_index = i
                        print(f"Found BlackHole 2ch output device at index: {i}")
                        break

                if blackhole_output_index is None:
                    print(f"Error: Could not find virtual output device '{VIRTUAL_DEVICE_NAME}'. Make sure BlackHole is installed and has output channels.")
                    print("Available devices:")
                    print(sd.query_devices()) # Print all devices for debugging
                    continue

                print(f"-> Routing audio output to: {VIRTUAL_DEVICE_NAME} (Device ID: {blackhole_output_index})")

                if lang_choice == '1':
                    print("Playing English voicemail...")
                    play_audio(file_path=ENGLISH_VOICEMAIL_PATH, output_device_id=blackhole_output_index)
                elif lang_choice == '2':
                    print("Playing Spanish voicemail (dummy)...")
                    play_audio(dummy_segment=SPANISH_VOICEMAIL_DUMMY, output_device_id=blackhole_output_index)

                print("Audio playback complete. Remember to set your recording app's INPUT to BlackHole 2ch.")
            else:
                print("Invalid choice. Returning to main menu.")
        else:
            print("Invalid command. Please enter 'p' or 'q'.")

if __name__ == "__main__":
    main_loop()

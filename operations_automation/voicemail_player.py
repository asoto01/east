import subprocess
import sounddevice as sd
import soundfile as sf
import os
import time

# --- Configuration ---
# Name of the virtual audio device (BlackHole 2ch)
VIRTUAL_DEVICE_NAME = "BlackHole 2ch"
# Path to your voicemail audio file
VOICEMAIL_FILE = "english_vm.wav"
# Small delay in seconds to allow the system to register device changes
DEVICE_SWITCH_DELAY = 0.5

def run_command(command):
    """Executes a shell command and returns the output, handling errors."""
    try:
        # Using shell=True for simpler command execution, but be cautious with untrusted input
        result = subprocess.run(command, check=True, capture_output=True, text=True, shell=True)
        return result.stdout.strip()
    except FileNotFoundError:
        print(f"❌ Error: The command '{command.split()[0]}' was not found.")
        print("Please ensure 'switchaudio-osx' is installed (e.g., 'brew install switchaudio-osx').")
        return None
    except subprocess.CalledProcessError as e:
        print(f"❌ Error executing command: {e}")
        print(f"Stderr: {e.stderr.strip()}")
        return None
    except Exception as e:
        print(f"❌ An unexpected error occurred in run_command: {e}")
        return None

def get_device_index(name, kind='output'):
    """Finds the index of a device by name and kind ('input' or 'output')."""
    try:
        devices = sd.query_devices()
        for i, device in enumerate(devices):
            # Check if the device name contains the specified name (case-insensitive for robustness)
            if name.lower() in device['name'].lower():
                if kind == 'output' and device['max_output_channels'] > 0:
                    print(f"Found {kind} device '{device['name']}' at index {i}")
                    return i
                elif kind == 'input' and device['max_input_channels'] > 0:
                    print(f"Found {kind} device '{device['name']}' at index {i}")
                    return i
        print(f"⚠️ Could not find a '{kind}' device containing '{name}'.")
        return None
    except Exception as e:
        print(f"❌ Error querying devices: {e}")
        return None

def get_current_input_device():
    """Gets the system's current default input device."""
    print("ℹ️ Getting current system input device...")
    return run_command("SwitchAudioSource -c -t input")

def set_input_device(device_name):
    """Sets the system's default input device."""
    print(f"🎤 Setting system microphone to: '{device_name}'...")
    run_command(f"SwitchAudioSource -t input -s '{device_name}'")
    time.sleep(DEVICE_SWITCH_DELAY) # Give the system a moment to recognize the change

def play_audio_to_device(file_path, output_device_index):
    """Plays a WAV file to a SPECIFIC output device and waits for it to finish."""
    if output_device_index is None:
        print("❌ Error: Cannot play audio, invalid output device index provided.")
        return

    try:
        data, fs = sf.read(file_path, dtype='float32')
        # Ensure 'channels' matches the audio file's channel count
        num_channels = data.shape[1] if data.ndim > 1 else 1

        # Use an explicit OutputStream to direct audio to the correct device
        with sd.OutputStream(samplerate=fs,
                             device=output_device_index,
                             channels=num_channels) as stream:
            print(f"▶️ Playing '{os.path.basename(file_path)}' to device {output_device_index} ({sd.query_devices(output_device_index)['name']})...")
            stream.write(data)
        print("✅ Playback finished.")

    except FileNotFoundError:
        print(f"❌ Error: Audio file not found at '{file_path}'")
    except Exception as e:
        print(f"❌ Error playing audio to device {output_device_index}: {e}")
        print("💡 Ensure the audio file is valid and the device supports the sample rate/channels.")

def main():
    """Main application loop."""
    print("--- 🎙️ Voicemail Player Initializing ---")

    if not os.path.exists(VOICEMAIL_FILE):
        print(f"❌ Fatal Error: The voicemail file '{VOICEMAIL_FILE}' was not found.")
        print("Please ensure the file is in the same directory as this script, or provide its full path.")
        return

    # 1. Get the device index for BlackHole output
    blackhole_output_index = get_device_index(VIRTUAL_DEVICE_NAME, kind='output')
    if blackhole_output_index is None:
        print(f"❌ Fatal Error: Could not find virtual audio device '{VIRTUAL_DEVICE_NAME}' output.")
        print("Please ensure BlackHole is installed and working. Check your macOS 'Audio MIDI Setup' utility.")
        return
    print(f"✅ Found virtual audio device '{VIRTUAL_DEVICE_NAME}' output at index: {blackhole_output_index}")

    # 2. Get the user's original, physical microphone.
    original_mic = get_current_input_device()
    if original_mic is None or original_mic.lower() == VIRTUAL_DEVICE_NAME.lower():
        print(f"⚠️ Warning: Your current system microphone is '{original_mic}'.")
        print("It should NOT be BlackHole at the start. Please set your default input to your actual microphone in System Settings -> Sound -> Input, and then restart the script.")
        # If it's already BlackHole, we can't reliably restore the original later.
        # Forcing exit to prevent issues.
        return
    print(f"✅ Detected original microphone: '{original_mic}'")
    print("-" * 40)

    try:
        while True:
            print("\nWaiting for your command...")
            print("  [1] Play Voicemail (Will temporarily switch your mic to BlackHole)")
            print("  [q] Quit")
            choice = input("Enter your choice: ").strip()

            if choice == '1':
                print("\n--- Initiating Voicemail Playback Sequence ---")
                # Step 1: Switch system input to BlackHole
                set_input_device(VIRTUAL_DEVICE_NAME)
                print(f"🕒 Waiting {DEVICE_SWITCH_DELAY} seconds for system to adjust...")

                # Step 2: Play audio to BlackHole's output
                play_audio_to_device(VOICEMAIL_FILE, blackhole_output_index)

                # Step 3: Switch system input back to original mic
                print("--- Voicemail Playback Complete ---")
                set_input_device(original_mic)
                print(f"✅ Switched system microphone back to '{original_mic}'. You can speak again.")

            elif choice.lower() == 'q':
                print("Exiting application.")
                break
            else:
                print("Invalid choice. Please try again.")

    except KeyboardInterrupt:
        print("\nInterrupted by user.")
    finally:
        # Always attempt to restore the original microphone on exit
        current_mic = get_current_input_device()
        if current_mic is not None and current_mic.lower() == VIRTUAL_DEVICE_NAME.lower():
            print("\n🚨 Cleaning up: restoring original microphone...")
            set_input_device(original_mic)
        print("--- Application Closed ---")

if __name__ == "__main__":
    main()


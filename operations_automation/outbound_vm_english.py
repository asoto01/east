#!/usr/bin/env python3
"""
GoTo Voicemail Automation Script
Automatically plays voicemail messages through BlackHole virtual microphone
"""

import sounddevice as sd
import soundfile as sf
import subprocess
import time
import threading
from pathlib import Path

class GoToVoicemailAutomator:
    def __init__(self):
        self.audio_file = "english_vm.m4a"
        self.blackhole_device = "BlackHole 2ch"
        self.original_input_device = None
        self.blackhole_device_id = None
        self.original_device_id = None
        self.is_running = True
        
        # Initialize and find devices
        self.setup_devices()
    
    def setup_devices(self):
        """Find and setup audio devices"""
        try:
            devices = sd.query_devices()
            
            # Find BlackHole device
            for i, device in enumerate(devices):
                if "BlackHole" in device['name'] and device['max_output_channels'] > 0:
                    self.blackhole_device_id = i
                    print(f"✅ Found BlackHole: {device['name']} (ID: {i})")
                    break
            
            if self.blackhole_device_id is None:
                raise Exception("BlackHole 2ch not found! Please install BlackHole first.")
            
            # Get current default input device
            self.original_device_id = sd.default.device[0]  # Input device
            if self.original_device_id is not None:
                original_device = devices[self.original_device_id]
                print(f"✅ Current input device: {original_device['name']} (ID: {self.original_device_id})")
            
            # Check if audio file exists
            if not Path(self.audio_file).exists():
                raise Exception(f"Audio file '{self.audio_file}' not found!")
            
            print(f"✅ Audio file found: {self.audio_file}")
            
        except Exception as e:
            print(f"❌ Setup error: {e}")
            raise
    
    def get_system_input_device(self):
        """Get current system input device using AppleScript"""
        try:
            script = '''
            tell application "System Preferences"
                activate
                set current pane to pane "com.apple.preference.sound"
                delay 0.5
            end tell
            
            tell application "System Events"
                tell process "System Preferences"
                    click radio button "Input" of tab group 1 of window "Sound"
                    delay 0.2
                    set selectedDevice to name of (first row of table 1 of scroll area 1 of tab group 1 of window "Sound" whose selected is true)
                end tell
            end tell
            
            tell application "System Preferences" to quit
            return selectedDevice
            '''
            
            result = subprocess.run(['osascript', '-e', script], 
                                  capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                return "Unknown"
        except:
            return "Unknown"
    
    def set_system_input_device(self, device_name):
        """Set system input device using AppleScript"""
        try:
            script = f'''
            tell application "System Preferences"
                activate
                set current pane to pane "com.apple.preference.sound"
                delay 0.5
            end tell
            
            tell application "System Events"
                tell process "System Preferences"
                    click radio button "Input" of tab group 1 of window "Sound"
                    delay 0.2
                    
                    repeat with theRow in rows of table 1 of scroll area 1 of tab group 1 of window "Sound"
                        if name of theRow contains "{device_name}" then
                            select theRow
                            delay 0.2
                            exit repeat
                        end if
                    end repeat
                end tell
            end tell
            
            tell application "System Preferences" to quit
            '''
            
            result = subprocess.run(['osascript', '-e', script], 
                                  capture_output=True, text=True, timeout=15)
            return result.returncode == 0
        except Exception as e:
            print(f"Error setting input device: {e}")
            return False
    
    def play_voicemail(self):
        """Play the voicemail message through BlackHole"""
        try:
            print("🎵 Playing voicemail message...")
            
            # Read the audio file
            data, sample_rate = sf.read(self.audio_file)
            
            # Play to BlackHole device
            sd.play(data, samplerate=sample_rate, device=self.blackhole_device_id)
            
            # Wait for playback to complete
            sd.wait()
            
            print("✅ Voicemail message completed!")
            return True
            
        except Exception as e:
            print(f"❌ Error playing voicemail: {e}")
            return False
    
    def switch_to_blackhole(self):
        """Switch system input to BlackHole"""
        print("🔄 Switching to BlackHole input...")
        success = self.set_system_input_device("BlackHole 2ch")
        if success:
            print("✅ Switched to BlackHole input")
            time.sleep(1)  # Give system time to switch
            return True
        else:
            print("❌ Failed to switch to BlackHole")
            return False
    
    def restore_original_input(self):
        """Restore original input device"""
        print("🔄 Restoring original input device...")
        
        # Get the original device name
        devices = sd.query_devices()
        if self.original_device_id and self.original_device_id < len(devices):
            original_name = devices[self.original_device_id]['name']
            success = self.set_system_input_device(original_name)
            if success:
                print(f"✅ Restored to: {original_name}")
                return True
        
        print("❌ Failed to restore original input")
        return False
    
    def automated_voicemail_sequence(self):
        """Complete automated voicemail sequence"""
        print("\n" + "="*50)
        print("🎯 STARTING AUTOMATED VOICEMAIL SEQUENCE")
        print("="*50)
        
        try:
            # Step 1: Switch to BlackHole
            if not self.switch_to_blackhole():
                return False
            
            # Step 2: Small delay to ensure GoTo picks up the change
            print("⏳ Waiting 2 seconds for GoTo to detect input change...")
            time.sleep(2)
            
            # Step 3: Play the voicemail
            success = self.play_voicemail()
            
            # Step 4: Wait a moment then restore original input
            time.sleep(1)
            self.restore_original_input()
            
            if success:
                print("🎉 Automated voicemail sequence completed successfully!")
                return True
            else:
                print("❌ Voicemail sequence failed")
                return False
                
        except Exception as e:
            print(f"❌ Sequence error: {e}")
            self.restore_original_input()  # Ensure we restore input
            return False
    
    def show_menu(self):
        """Show the main menu"""
        print("\n" + "="*50)
        print("🎙️  GOTO VOICEMAIL AUTOMATOR")
        print("="*50)
        print("1. Play automated voicemail")
        print("2. Test audio playback")
        print("3. Check current input device")
        print("4. Quit")
        print("-" * 50)
    
    def test_audio(self):
        """Test audio playback without switching inputs"""
        print("\n🧪 Testing audio playback...")
        try:
            data, sample_rate = sf.read(self.audio_file)
            print(f"📊 Audio info: {len(data)} samples, {sample_rate}Hz")
            
            print("🎵 Playing test (you should hear this through your speakers)...")
            sd.play(data, samplerate=sample_rate)
            sd.wait()
            print("✅ Audio test completed!")
            
        except Exception as e:
            print(f"❌ Audio test failed: {e}")
    
    def check_input_device(self):
        """Check current system input device"""
        print("\n🔍 Checking current input device...")
        current_device = self.get_system_input_device()
        print(f"📱 Current system input: {current_device}")
    
    def run(self):
        """Main program loop"""
        print("🚀 GoTo Voicemail Automator Started!")
        print("💡 Make sure GoTo is running and ready for calls")
        
        while self.is_running:
            try:
                self.show_menu()
                choice = input("Enter your choice (1-4): ").strip()
                
                if choice == "1":
                    confirm = input("🎯 Ready to play automated voicemail? (y/n): ").strip().lower()
                    if confirm in ['y', 'yes']:
                        self.automated_voicemail_sequence()
                    else:
                        print("❌ Voicemail cancelled")
                
                elif choice == "2":
                    self.test_audio()
                
                elif choice == "3":
                    self.check_input_device()
                
                elif choice == "4":
                    print("👋 Goodbye!")
                    self.is_running = False
                
                else:
                    print("❌ Invalid choice. Please enter 1-4.")
                    
            except KeyboardInterrupt:
                print("\n\n👋 Program interrupted. Restoring original input...")
                self.restore_original_input()
                break
            except Exception as e:
                print(f"❌ Error: {e}")
        
        # Cleanup
        self.restore_original_input()
        print("🔧 Cleanup completed.")

def main():
    """Main entry point"""
    try:
        automator = GoToVoicemailAutomator()
        automator.run()
    except Exception as e:
        print(f"❌ Failed to start: {e}")
        print("\n🔧 Troubleshooting:")
        print("1. Make sure BlackHole 2ch is installed")
        print("2. Check that english_vm.m4a exists in this directory")
        print("3. Grant microphone permissions to Terminal")

if __name__ == "__main__":
    main()

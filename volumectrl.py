import speech_recognition as sr
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

# Setting up audio library
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# String to Integer
textIntegers = ["zero","one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"]

def main():
    r = sr.Recognizer()
    while True:
        try: 
            # Creating microphone instance, empty parameters (I only have 1 microphone).
            with sr.Microphone() as mic:
                # Handling ambient noise and capturing microphone input
                r.adjust_for_ambient_noise(mic, duration=0.5)
                audio = r.listen(mic)
                text = r.recognize_google(audio).lower()
                # Stop the script
                if "cancel script" in text:
                    break
                elif "set volume" in text:
                    # Iterate through the integer list
                    for nums in textIntegers:
                        if nums in text:
                            digits = textIntegers.index(nums)
                            # Break out of the for-loop
                            break
                        else: 
                            digits = ''.join([i for i in text if i.isdigit()])
                    volume.SetMasterVolumeLevelScalar(int(digits)/100, None)
                
        # Unrecognizable speech handler
        except sr.UnknownValueError: 
            r = sr.Recognizer()
            continue

if __name__ == "__main__":
    main()
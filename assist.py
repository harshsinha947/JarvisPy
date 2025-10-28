import pvporcupine
import pyaudio
import struct
import threading
import speech_recognition as sr
import pyttsx3
import pywhatkit
import wikipedia
import pyjokes
import webbrowser
import os
import requests
import datetime
import time
from plyer import notification

ACCESS_KEY = "deZFqozbWA+rt+EkxDwqEfVCMjLi1D6yGK63CQThKyFJRtn+T/zrVg=="  # Picovoice key
WEATHER_KEY = "48b3dab91be8d3452287ae0cd682768f"                        # Your OpenWeatherMap API key


# -------------------------- BASIC FUNCTIONS --------------------------

def show_popup(msg):
    notification.notify(
        title="Jarvis Assistant",
        message=msg,
        app_name="Jarvis",
        timeout=4
    )

def speak(text):
    print(text)
    show_popup(text)
    engine = pyttsx3.init()
    engine.setProperty('rate', 175)
    engine.setProperty('voice', engine.getProperty('voices')[0].id)
    engine.say(text)
    engine.runAndWait()
    engine.stop()

def get_weather(city):
    url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={WEATHER_KEY}&units=metric"
    try:
        response = requests.get(url)
        data = response.json()
        if data["cod"] != "404":
            temp = data["main"]["temp"]
            desc = data["weather"][0]["description"]
            speak(f"The temperature in {city} is {temp} degrees Celsius with {desc}.")
        else:
            speak("Sorry, I couldn't find the weather for that city.")
    except Exception:
        speak("Couldn't fetch weather. Please check your connection.")


# -------------------------- VOLUME CONTROL --------------------------

def change_volume(direction):
    try:
        # Try Pycaw first
        try:
            from comtypes import CLSCTX_ALL
            from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
            from ctypes import cast, POINTER

            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            current = volume.GetMasterVolumeLevelScalar()

            if "up" in direction:
                new_volume = min(1.0, current + 0.1)
                volume.SetMasterVolumeLevelScalar(new_volume, None)
                speak("Volume increased.")
            elif "down" in direction:
                new_volume = max(0.0, current - 0.1)
                volume.SetMasterVolumeLevelScalar(new_volume, None)
                speak("Volume decreased.")
            return

        except Exception as e:
            print("Pycaw method failed, using fallback:", e)

        # Fallback method using Windows API
        import ctypes

        def get_current_volume():
            value = ctypes.c_uint()
            ctypes.windll.winmm.waveOutGetVolume(0, ctypes.byref(value))
            return value.value & 0xFFFF

        def set_volume(val):
            val = max(0, min(0xFFFF, val))
            ctypes.windll.winmm.waveOutSetVolume(0, val | (val << 16))

        current = get_current_volume()
        step = int(0xFFFF * 0.1)
        if "up" in direction:
            set_volume(min(0xFFFF, current + step))
            speak("Volume increased.")
        elif "down" in direction:
            set_volume(max(0, current - step))
            speak("Volume decreased.")

    except Exception as e:
        speak("Sorry, I couldn’t change the volume.")
        print("Volume control error:", e)


# -------------------------- BRIGHTNESS CONTROL --------------------------

def change_brightness(direction):
    try:
        import screen_brightness_control as sbc
        current = sbc.get_brightness(display=0)[0]
        if "up" in direction:
            new_brightness = min(100, current + 10)
            sbc.set_brightness(new_brightness)
            speak("Brightness increased.")
        elif "down" in direction:
            new_brightness = max(0, current - 10)
            sbc.set_brightness(new_brightness)
            speak("Brightness decreased.")
    except Exception as e:
        speak("Sorry, I couldn’t adjust the brightness.")
        print("Brightness error:", e)


# -------------------------- SYSTEM CONTROLS --------------------------

def system_control(action):
    try:
        if "shutdown" in action:
            speak("Shutting down your computer.")
            os.system("shutdown /s /t 1")
        elif "restart" in action:
            speak("Restarting your computer.")
            os.system("shutdown /r /t 1")
        elif "sleep" in action:
            speak("Putting the system to sleep.")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        elif "lock" in action:
            speak("Locking your computer.")
            os.system("rundll32.exe user32.dll,LockWorkStation")
    except Exception as e:
        speak("Unable to perform that system action.")
        print("System control error:", e)


# -------------------------- MAIN JARVIS LOGIC --------------------------

def run_jarvis(command):
    try:
        if "time" in command:
            now = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The current time is {now}")

        elif "date" in command:
            date_now = datetime.datetime.now().strftime("%B %d, %Y")
            speak(f"Today's date is {date_now}")

        elif "open youtube" in command:
            speak("Opening YouTube")
            threading.Thread(target=lambda: webbrowser.open("https://www.youtube.com")).start()

        elif "open google" in command:
            speak("Opening Google")
            threading.Thread(target=lambda: webbrowser.open("https://www.google.com")).start()

        elif "open instagram" in command:
            speak("Opening Instagram")
            threading.Thread(target=lambda: webbrowser.open("https://www.instagram.com")).start()

        elif "play" in command:
            song = command.replace("play", "").strip()
            if song:
                speak(f"Playing {song} on YouTube")
                threading.Thread(target=lambda: pywhatkit.playonyt(song)).start()
            else:
                speak("Please tell me a song name to play.")

        elif "wikipedia" in command:
            speak("Searching Wikipedia...")
            query = command.replace("wikipedia", "").strip()
            try:
                info = wikipedia.summary(query, sentences=2)
                speak(info)
            except Exception:
                speak("Sorry, could not find information on Wikipedia.")

        elif "weather" in command:
            speak("Please tell me the city name for weather.")
            r = sr.Recognizer()
            with sr.Microphone() as source:
                r.adjust_for_ambient_noise(source, duration=0.8)
                audio = r.listen(source, timeout=8)
            try:
                city = r.recognize_google(audio, language='en-in')
                get_weather(city)
            except Exception:
                speak("Sorry, I didn't catch the city name.")

        elif "joke" in command:
            speak(pyjokes.get_joke())

        elif "open notepad" in command:
            speak("Opening Notepad")
            threading.Thread(target=lambda: os.system("notepad")).start()

        elif "open whatsapp" in command:
            speak("Opening WhatsApp Desktop")
            try:
                whatsapp_path = r"C:\Users\Harsh Sinha\Desktop\WhatsApp - Shortcut.lnk"
                os.system(f'start "" "{whatsapp_path}"')
            except Exception as e:
                speak("Sorry, I couldn't open WhatsApp Desktop.")
                print("Error opening WhatsApp:", e)

        elif "open chrome" in command:
            speak("Opening Google Chrome")
            threading.Thread(target=lambda: os.system("start chrome")).start()

        elif "volume up" in command or "increase volume" in command:
            change_volume("up")

        elif "volume down" in command or "decrease volume" in command:
            change_volume("down")

        elif "brightness up" in command or "increase brightness" in command:
            change_brightness("up")

        elif "brightness down" in command or "decrease brightness" in command:
            change_brightness("down")

        elif any(x in command for x in ["shutdown", "restart", "sleep", "lock"]):
            system_control(command)

        elif command.strip() in ["exit", "quit", "stop"]:
            speak("Goodbye, have a great day!")
            os._exit(0)

        else:
            speak("Sorry, I didn’t understand that. Please say again.")

    except Exception as e:
        print("run_jarvis error:", e)
        speak("Sorry, something went wrong.")


# -------------------------- WAKE WORD LOOP --------------------------

def jarvis_loop():
    porcupine = pvporcupine.create(keywords=["jarvis"], access_key=ACCESS_KEY)
    pa = pyaudio.PyAudio()
    audio_stream = pa.open(rate=porcupine.sample_rate, channels=1, format=pyaudio.paInt16,
                           input=True, frames_per_buffer=porcupine.frame_length)
    speak("Always listening for 'Jarvis'...")

    recognizer = sr.Recognizer()
    while True:
        try:
            pcm = audio_stream.read(porcupine.frame_length)
            pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
            keyword_index = porcupine.process(pcm)
            if keyword_index >= 0:
                speak("How can I help you?")
                with sr.Microphone() as source:
                    recognizer.adjust_for_ambient_noise(source, duration=0.8)
                    audio = recognizer.listen(source, timeout=8)
                    try:
                        command = recognizer.recognize_google(audio, language='en-in').lower()
                        print("Heard:", command)
                        threading.Thread(target=run_jarvis, args=(command,)).start()
                        time.sleep(1.2)
                    except Exception as ex:
                        print("Command error:", ex)
                        speak("Sorry, I didn't catch that.")
        except Exception as e:
            print("Main loop error:", e)
            speak("An error occurred, restarting listening.")
            time.sleep(1)
            break

def run_forever():
    while True:
        try:
            jarvis_loop()
        except Exception as big_error:
            print("Restarting Jarvis (outer loop) error:", big_error)
            time.sleep(2)

if __name__ == "__main__":
    run_forever()

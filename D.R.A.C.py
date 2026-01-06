import speech_recognition as sr
import datetime
import webbrowser
import random
import requests
import wikipedia
import pyjokes
import subprocess
import threading
import time
import re
import os
import glob
import queue
import pyaudio

import config


def calculate_expression(expr: str) -> float:
    expr = expr.strip()
    if not re.fullmatch(r"[0-9+\-*/().\s]+", expr):
        raise ValueError("Invalid characters")

    tokens = re.findall(r"\d+\.\d+|\d+|[()+\-*/]", expr.replace(" ", ""))

    def prec(op):
        return 1 if op in ("+", "-") else 2 if op in ("*", "/") else 0

    output = []
    ops = []

    def is_unary_minus(i):
        if tokens[i] != "-":
            return False
        if i == 0:
            return True
        return tokens[i - 1] in ("(", "+", "-", "*", "/")

    i = 0
    while i < len(tokens):
        t = tokens[i]
        if re.fullmatch(r"\d+\.\d+|\d+", t):
            output.append(float(t))
        elif t in ("+", "-", "*", "/"):
            if is_unary_minus(i):
                output.append(0.0)
                ops.append("-")
            else:
                while ops and ops[-1] in ("+", "-", "*", "/") and prec(ops[-1]) >= prec(t):
                    output.append(ops.pop())
                ops.append(t)
        elif t == "(":
            ops.append(t)
        elif t == ")":
            while ops and ops[-1] != "(":
                output.append(ops.pop())
            if not ops:
                raise ValueError("Mismatched parentheses")
            ops.pop()
        else:
            raise ValueError("Bad token")
        i += 1

    while ops:
        op = ops.pop()
        if op in ("(", ")"):
            raise ValueError("Mismatched parentheses")
        output.append(op)

    stack = []
    for item in output:
        if isinstance(item, float):
            stack.append(item)
        else:
            if len(stack) < 2:
                raise ValueError("Invalid expression")
            b = stack.pop()
            a = stack.pop()
            if item == "+":
                stack.append(a + b)
            elif item == "-":
                stack.append(a - b)
            elif item == "*":
                stack.append(a * b)
            elif item == "/":
                stack.append(a / b)

    if len(stack) != 1:
        raise ValueError("Invalid expression")
    return stack[0]


def clean_query(text: str) -> str:
    return re.sub(r"\s+", " ", text.strip())


def youtube_search_url(query: str) -> str:
    return f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"


def google_search_url(query: str) -> str:
    return f"https://www.google.com/search?q={query.replace(' ', '+')}"


def try_open_exe(exe_path: str) -> bool:
    try:
        subprocess.Popen(exe_path)
        return True
    except Exception:
        return False


def find_pycharm_exe() -> str | None:
    candidates = []
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

    patterns = [
        os.path.join(program_files, "JetBrains", "PyCharm*", "bin", "pycharm64.exe"),
        os.path.join(program_files, "JetBrains", "PyCharm*", "bin", "pycharm.exe"),
        os.path.join(program_files_x86, "JetBrains", "PyCharm*", "bin", "pycharm64.exe"),
        os.path.join(program_files_x86, "JetBrains", "PyCharm*", "bin", "pycharm.exe"),
    ]
    for p in patterns:
        candidates.extend(glob.glob(p))
    return candidates[0] if candidates else None


class WindowsTTSWorker(threading.Thread):
    def __init__(self, rate: int = 0, volume: int = 100, voice: str | None = None):
        super().__init__(daemon=True)
        self.q = queue.Queue()
        self.rate = rate
        self.volume = volume
        self.voice = voice

    def say(self, text: str):
        self.q.put(str(text))

    @staticmethod
    def _ps_escape(s: str) -> str:
        return s.replace("'", "''")

    def _speak_powershell(self, text: str):
        safe = self._ps_escape(text)

        voice_block = ""
        if self.voice:
            v = self._ps_escape(self.voice)
            voice_block = (
                f"$v = $synth.GetInstalledVoices() | "
                f"Where-Object {{$_.VoiceInfo.Name -like '*{v}*'}} | "
                f"Select-Object -First 1;"
                f"if ($v) {{ $synth.SelectVoice($v.VoiceInfo.Name) }};"
            )

        ps = (
            "Add-Type -AssemblyName System.Speech;"
            "$synth = New-Object System.Speech.Synthesis.SpeechSynthesizer;"
            f"$synth.Rate = {int(self.rate)};"
            f"$synth.Volume = {int(self.volume)};"
            f"{voice_block}"
            f"$synth.Speak('{safe}');"
        )

        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
        )

    def run(self):
        while True:
            text = self.q.get()
            if text is None:
                break
            try:
                self._speak_powershell(text)
            except Exception as e:
                print("TTS error:", repr(e))


class DRAC:
    def __init__(self):
        self.user_name = getattr(config, "USER_NAME", "Sir")
        self.ai_name = getattr(config, "AI_NAME", "DRAC")
        self.wake_word = getattr(config, "WAKE_WORD", "hey").lower().strip()

        self.extra_wake_words = ["drac", "hey", "hello", "what's up", "wake up"]

        self.weather_api_key = getattr(config, "WEATHER_API_KEY", "")
        self.news_api_key = getattr(config, "NEWS_API_KEY", "")
        self.news_country = getattr(config, "NEWS_COUNTRY", "us")
        self.default_location = getattr(config, "DEFAULT_LOCATION", "Baku")

        self.tts = WindowsTTSWorker(
            rate=getattr(config, "TTS_RATE", 0),
            volume=getattr(config, "TTS_VOLUME", 100),
            voice=getattr(config, "TTS_VOICE", None),
        )
        self.tts.start()

        self.recognizer = sr.Recognizer()
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.pause_threshold = 0.8
        self.recognizer.non_speaking_duration = 0.5

        mic_index = getattr(config, "MIC_DEVICE_INDEX", None)
        self.microphone = sr.Microphone() if mic_index is None else sr.Microphone(device_index=mic_index)

        self.system_online = True
        self._calibrated = False

        self.legendary_jokes = [
            "I tried to catch fog yesterday. Mist.",
            "Why do programmers prefer dark mode? Because light attracts bugs.",
            "I would tell you a construction joke… but I’m still working on it.",
            "Your problems are temporary. Your potential is permanent.",
            "If stress had a face, I would uninstall it.",
        ]

        self.hype_lines = [
            "Alright, listen. You are built different. Let’s do this.",
            "Boss mode activated. What’s the mission?",
            "You’re not behind. You’re loading… and it’s going to be legendary.",
            "One step at a time. You’ve got this.",
        ]

        self.comfort_lines = [
            "I hear you. Breathe with me for a second.",
            "It’s okay to feel this way. You’re not alone.",
            "Let’s reduce the chaos: one small win right now.",
        ]

        print(f"{self.ai_name} Initializing...")
        self.speak(f"{self.ai_name} online. Say {self.wake_word} and tell me what you need, {self.user_name}.")
        self.calibrate_mic_once()
        self.speak("Voice system ready.")

    def speak(self, text: str):
        self.tts.say(text)

    def calibrate_mic_once(self):
        if self._calibrated:
            return
        try:
            with self.microphone as source:
                print("Adjusting for ambient noise (one time)...")
                self.recognizer.adjust_for_ambient_noise(source, duration=2)
            self._calibrated = True
        except Exception as e:
            print("Mic calibration error:", repr(e))

    def listen(self, timeout=6, phrase_time_limit=10):
        with self.microphone as source:
            print(f"Listening... (Say '{self.wake_word}' then command)")
            try:
                audio = self.recognizer.listen(source, timeout=timeout, phrase_time_limit=phrase_time_limit)
                text = self.recognizer.recognize_google(audio, language="en-US").lower().strip()

                print("\n==============================")
                print("YOU SAID:")
                print(text)
                print("==============================\n")

                if self.wake_word in text:
                    cmd = clean_query(text.replace(self.wake_word, ""))
                    return cmd if cmd else None

                for w in self.extra_wake_words:
                    if w in text:
                        cmd = clean_query(text.replace(w, ""))
                        return cmd if cmd else None

                return None

            except sr.WaitTimeoutError:
                return None
            except sr.UnknownValueError:
                return None
            except sr.RequestError as e:
                print("Speech service error:", repr(e))
                self.speak("Speech recognition is not available right now.")
                return None
            except Exception as e:
                print("Listen error:", repr(e))
                return None

    def process_command(self, command: str) -> bool:
        if not command:
            return False

        c = command.lower().strip()

        if any(x in c for x in ["exit", "quit", "goodbye", "shutdown", "bye"]):
            self.speak(f"Goodbye, {self.user_name}. Stay legendary.")
            self.system_online = False
            return True

        if "how are you" in c:
            self.speak("All systems green. How are you feeling?")
            return True

        if "tell me something interesting" in c or "something interesting" in c:
            self.speak(random.choice([
                "Octopuses have three hearts and blue blood.",
                "Honey never spoils.",
                "The Eiffel Tower can grow taller in summer due to heat expansion.",
            ]))
            return True

        if any(k in c for k in ["i'm sad", "i am sad", "sad", "depressed", "stressed", "anxious", "tired", "bored"]):
            self.speak(random.choice(self.comfort_lines))
            self.speak(random.choice(self.hype_lines))
            return True

        if "motivate me" in c or "hype me up" in c or "give me motivation" in c:
            self.speak(random.choice(self.hype_lines))
            return True

        if any(x in c for x in ["hello", "hi"]):
            self.speak(f"Hey {self.user_name}. What’s the mission?")
            return True

        if "time" in c:
            now = datetime.datetime.now()
            self.speak(f"It is {now.strftime('%H:%M')} right now.")
            return True

        if "date" in c or "today" in c:
            today = datetime.datetime.now().strftime("%B %d, %Y")
            self.speak(f"Today's date is {today}.")
            return True

        if "joke" in c or "make me laugh" in c or "cheer me up" in c:
            try:
                self.speak(pyjokes.get_joke())
            except Exception:
                self.speak(random.choice(self.legendary_jokes))
            return True

        if c.startswith("calculate") or c.startswith("what is") or c.startswith("how much is"):
            self.perform_calculation(c)
            return True

        if c.startswith("play "):
            q = c.replace("play", "").replace("on youtube", "").strip()
            if not q:
                self.speak("Tell me what you want to play.")
                return True
            self.speak(f"Opening YouTube for: {q}")
            webbrowser.open(youtube_search_url(q))
            return True

        if c.startswith("open ") or c.startswith("launch ") or c.startswith("start "):
            self.open_application(c)
            return True

        if "weather" in c or "forecast" in c:
            self.get_weather_forecast(c)
            return True

        if "news" in c or "headlines" in c:
            self.get_news()
            return True

        if any(t in c for t in ["search", "google", "internet", "on the web"]):
            q = self.extract_search_query(c)
            if q:
                self.speak(f"Searching the web for: {q}")
                webbrowser.open(google_search_url(q))
            else:
                self.speak("Tell me what you want me to search.")
            return True

        if c.startswith("who is ") or c.startswith("tell me about "):
            topic = c.replace("who is", "").replace("tell me about", "").strip()
            if topic:
                self.quick_info(topic)
            else:
                self.speak("Tell me the topic.")
            return True

        self.speak("I’m not sure. Say: hey search <topic> on the internet.")
        return True

    def extract_search_query(self, c: str) -> str:
        q = c
        for k in ["search", "on the web", "on the internet", "internet", "google", "web browser", "browser search", "for"]:
            q = q.replace(k, "")
        return clean_query(q)

    def quick_info(self, topic: str):
        try:
            self.speak(f"Here is a short summary about {topic}.")
            summary = wikipedia.summary(topic, sentences=3)
            self.speak(summary)
        except Exception:
            self.speak("I couldn't get a summary. Say: hey search it on the internet.")

    def open_application(self, c: str):
        name = c.replace("open", "").replace("launch", "").replace("start", "").strip()

        if "browser" in name or "chrome" in name:
            self.speak("Opening browser.")
            webbrowser.open("https://www.google.com")
            return

        if "instagram" in name:
            self.speak("Opening Instagram.")
            webbrowser.open("https://www.instagram.com/")
            return

        if "pycharm" in name:
            self.speak("Opening PyCharm.")
            exe = find_pycharm_exe()
            if exe and try_open_exe(exe):
                return
            try:
                subprocess.Popen(["cmd", "/c", "start", "", "pycharm"])
                return
            except Exception:
                self.speak("I couldn't find PyCharm on this PC.")
                return

        apps = {
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "cmd": "cmd.exe",
            "task manager": "taskmgr.exe",
            "paint": "mspaint.exe",
        }

        for k, exe in apps.items():
            if k in name:
                self.speak(f"Opening {k}.")
                try_open_exe(exe)
                return

        self.speak("I didn't recognize that app.")

    def get_weather_forecast(self, command: str):
        if not self.weather_api_key or str(self.weather_api_key).startswith("YOUR_"):
            self.speak("Weather API key is not configured.")
            return

        city = command.replace("weather", "").replace("forecast", "").replace("in", "").strip()
        if not city:
            city = self.default_location

        try:
            url = f"http://api.openweathermap.org/data/2.5/forecast?q={city}&appid={self.weather_api_key}&units=metric"
            data = requests.get(url, timeout=10).json()

            if str(data.get("cod")) != "200":
                self.speak(f"I couldn't get the forecast for {city}.")
                return

            city_name = data["city"]["name"]
            items = data["list"][:3]

            self.speak(f"Forecast for {city_name}. Next hours:")
            for it in items:
                hour = it["dt_txt"][11:16]
                temp = it["main"]["temp"]
                desc = it["weather"][0]["description"]
                self.speak(f"At {hour}, {desc}, {temp} degrees.")
                time.sleep(0.12)

        except Exception:
            self.speak("Sorry, I couldn't retrieve the forecast right now.")

    def get_news(self):
        if not self.news_api_key or str(self.news_api_key).startswith("YOUR_"):
            self.speak("News API key is not configured.")
            return

        try:
            url = f"https://newsapi.org/v2/top-headlines?country={self.news_country}&apiKey={self.news_api_key}"
            data = requests.get(url, timeout=10).json()

            if data.get("status") != "ok":
                self.speak("Could not retrieve news at this time.")
                return

            articles = data.get("articles", [])[:4]
            if not articles:
                self.speak("No headlines found.")
                return

            self.speak("Top headlines:")
            for i, a in enumerate(articles, 1):
                self.speak(f"{i}. {a.get('title', 'Untitled')}")
                time.sleep(0.15)

        except Exception:
            self.speak("Sorry, I couldn't retrieve the news.")

    def perform_calculation(self, command: str):
        try:
            expr = command.replace("calculate", "").replace("what is", "").replace("how much is", "").strip()
            result = calculate_expression(expr)
            if abs(result - int(result)) < 1e-9:
                result = int(result)
            self.speak(f"The answer is {result}.")
        except Exception:
            self.speak("I couldn't calculate that. Example: hey calculate 12*(3+4).")

    def run(self):
        print(f"\n{self.ai_name} is now active. Say '{self.wake_word}' then your command.")
        print("Say 'exit', 'quit', or 'goodbye' to shut down.\n")

        while self.system_online:
            cmd = self.listen()
            if cmd:
                self.process_command(cmd)
            time.sleep(0.08)


if __name__ == "__main__":
    print("=" * 60)
    print("DRAC")
    print("=" * 60)

    assistant = DRAC()
    assistant.run()

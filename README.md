# 🧠 D.R.A.C. – Deep Responsive Artificial Cognition

**DRAC** is a Windows-based AI voice assistant built with **Python**, designed to interact with users through natural voice commands.  
Inspired by Jarvis-style assistants, DRAC can listen, understand, and respond **out loud**, while also controlling system tasks and searching the web.

---

## ✨ Features

- 🎙️ **Real-time Speech Recognition**
  - Uses Google Speech Recognition (English – `en-US`)

- 🔊 **Reliable Text-to-Speech (TTS)**
  - Native Windows voice output
  - Speaks responses instead of only printing to console

- 🖥️ **System Control**
  - Open applications (browser, PyCharm, calculator, etc.)
  - Execute basic system commands

- 🌐 **Web Capabilities**
  - Google search
  - YouTube search & playback
  - Wikipedia summaries

- 🧮 **Utilities**
  - Voice-based calculator
  - Current time & date
  - Jokes & motivation responses

- 🧠 **Conversational Logic**
  - Context-aware responses
  - Mood detection (basic emotional feedback)

---

## 🛠️ Tech Stack

- **Python 3**
- speech_recognition
- install libraries from "requirements.txt"

---

## 🚀 Installation

### 1️⃣ Clone the repository
```bash
git clone https://github.com/RaufAlizada/drac-ai-voice-assistant.git
cd drac-ai-voice-assistant
```

### 2️⃣ Install requirements
```bash
pip install -r requirements.txt
```

### 3️⃣ Configure config.py file
- Get your weather api key -> https://openweathermap.org/api
- Get your news api key -> https://newsapi.org/

### 4️⃣ Run the assistant
```bash
python D.R.A.C.py
```

---

## 📜 License

This project is MIT licensed.

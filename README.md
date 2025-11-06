# Welding Shop Manager

A modern, bilingual (English/Arabic) desktop application for managing welding shop operations. Built with Python, CustomTkinter for a sleek UI, OpenAI Whisper for voice transcription, and pyttsx3 for text-to-speech. It supports voice input for fields, data export to Excel templates, and a responsive layout for low-resolution screens.

---

## ✨ Features

- 🎙 **Voice-Enabled Form Filling**
  - Record and transcribe speech for any field using Whisper (English/Arabic).
  - Confirm inputs via voice (“Yes/No”) or GUI popup.
  - Fields lock on confirmation; mic button disables to prevent accidental edits.

- 🌐 **Bilingual Support**
  - Instantly switch between English and Arabic.
  - UI labels and prompts update dynamically.

- 🌓 **Dark/Light Theme**
  - Toggle sleek modern themes using CustomTkinter.

- 🧾 **Data Management**
  - Add, edit, delete, and view records in a scrollable table.
  - Double-click rows for quick actions.

- 📊 **Excel Export**
  - Generate reports using a custom template:
    ATNM-ODC-MF-014-Daily Welding Production – Visual Inspection Report – Rev 01.xlsx
  - Preserves logos, merged cells, and formatting.

- 🔢 **Number Recognition**
  - Converts spoken digits (e.g., “five three one” → 531) for precise numeric input.

- 💻 **Responsive UI**
  - Scrollable and resizable; works well at 800×600 and above.

---

## 📦 Installation

1) **Clone the repository**  
   git clone https://github.com/Ashu-14-oct/product-vt.git

2) **Install dependencies**  
   pip install customtkinter openai-whisper pyttsx3 pyaudio openpyxl pillow wave

   Notes:
   • Whisper downloads models automatically on first run, or pre-download with whisper.load_model("medium").  
   • macOS: brew install portaudio → then pip install pyaudio.  
   • Linux (Debian/Ubuntu): sudo apt-get install portaudio19-dev → then pip install pyaudio.

3) **Add template (recommended)**
   • Place the Excel template file:
     ATNM-ODC-MF-014-Daily Welding Production – Visual Inspection Report – Rev 01.xlsx
     in the project root.
   • Optionally add logo.png in the project root for header branding.

4) **Run the app**  
   python welding_app.py

---

## ▶️ Usage

- **Launch**: Form on the left, records table on the right.  
- **Voice input**:
  1. Click the 🎤 next to a field; the app plays a prompt (e.g., “What is your Job ID?”).
  2. Speak your response; it is transcribed and read back.
  3. Confirm with “Yes” to lock the field or “No” to retry. Up to 2 retries before GUI fallback.
- **Add Record**: Fill fields → click **Add Entry** (validates Job ID & Welder Name).  
- **Edit/Delete**: Double-click table rows for actions.  
- **Export**: Click **Submit** to generate an Excel report for all records using the template.  
- **Language/Theme**: Use header controls to switch language and dark/light theme.

---

## 💡 Voice & TTS Tips

- Speak clearly; pause briefly between fields.  
- For numbers, say digits individually for exact input (“five three one”).  
- On macOS, native `say` is used for TTS when available; pyttsx3 is the fallback.

---

---

## 📋 Dependencies (key)

| Package         | Version (example) | Purpose                          |
|-----------------|-------------------|----------------------------------|
| customtkinter   | ^5.2.0            | Modern UI widgets                |
| openai-whisper  | ^20231117         | Speech-to-text (EN/AR)           |
| pyttsx3         | ^2.90             | Text-to-speech (fallback)        |
| pyaudio         | ^0.2.11           | Audio recording                  |
| openpyxl        | ^3.1.2            | Excel read/write                 |
| pillow          | ^10.0.0           | Image handling (logo)            |
| wave            | stdlib            | Audio I/O                        |

Generate a lockfile with: pip freeze > requirements.txt

---

## 🧑‍💻 Contributing

1. Fork the repo.  
2. Create a feature branch: git checkout -b feature/voice-enhance  
3. Commit: git commit -m "Add voice retry logic"  
4. Push: git push origin feature/voice-enhance  
5. Open a Pull Request.

Report bugs or request features via GitHub Issues. Contributions are welcome!

---

## ⚖️ License

MIT License — see `LICENSE` for details.

---

## 🙏 Acknowledgments

- CustomTkinter for the modern desktop UI.  
- OpenAI Whisper for robust speech transcription (EN/AR).  
- Inspired by welding production reporting needs at Al Tasnim Enterprises LLC.

---

⭐ If this project helps you, please consider starring the repository!

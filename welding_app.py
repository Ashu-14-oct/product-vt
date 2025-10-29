import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime
import threading
import pyaudio
import wave  # For saving WAV file
import whisper  # OpenAI Whisper for transcription
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pyttsx3
from tkinter import messagebox
import re
import tempfile


def safe_json_loads(data):
    try:
        return json.loads(data)
    except Exception:
        return {}


class WeldingShopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Welding Shop Manager")
        self.root.geometry("1000x700")
        self.root.configure(bg="#f0f0f0")

        # No shared TTS engine anymore - init fresh each time to avoid silent state bug

        # Data storage
        self.records = []
        self.data_file = "welding_data.json"
        self.load_data()

        # Language settings
        self.current_lang = "en"
        self.translations = {
            "en": {
                "title": "Welding Shop Manager",
                "job_id": "Job ID:",
                "weld_id": "Weld ID:",
                "kp_sec": "KP Sec:",
                "wps_no": "WPS No:",
                "material_gr": "Material Gr:",
                "heat_no": "Heat No:",
                "size": "Size(Inches):",
                "thk": "Thk(mm):",
                "weld_side": "Weld Side:",
                "root": "Root:",
                "material_comb": "Mtrl. Comb:",
                "pipe_no": "Pipe No:",
                "pipe_length": "Pipe length(mtrs):",
                "welder_name": "Welder Name:",
                "material": "Material:",
                "weld_type": "Weld Type:",
                "description": "Description:",
                "date": "Date:",
                "add_entry": "Add Entry",
                "clear_form": "Clear Form",
                "download_excel": "Download Excel",
                "language": "Language:",
                "records_title": "Records",
                "delete": "Delete",
                "success_add": "Entry added successfully!",
                "success_export": "Excel file exported successfully!",
                "error_fill": "Please fill Job ID and Welder Name",
                "recording": "Recording... Speak now",
                "mic_tooltip": "Click to record voice",
                "confirm_text": "Is this correct: '{}'? (Yes to confirm and lock field, No to re-record)",
            },
            "ar": {
                "title": "إدارة ورشة اللحام",
                "job_id": "رقم العمل:",
                "weld_id": "رقم اللحام:",
                "kp_sec": "قسم KP:",
                "wps_no": "رقم WPS:",
                "material_gr": "درجة المادة:",
                "heat_no": "رقم الحرارة:",
                "size": "الحجم (بوصة):",
                "thk": "السماكة (مم):",
                "weld_side": "جانب اللحام:",
                "root": "الجذر:",
                "material_comb": "تركيبة المادة:",
                "pipe_no": "رقم الأنبوب:",
                "pipe_length": "طول الأنبوب (متر):",
                "welder_name": "اسم اللحام:",
                "material": "المادة:",
                "weld_type": "نوع اللحام:",
                "description": "الوصف:",
                "date": "التاريخ:",
                "add_entry": "إضافة سجل",
                "clear_form": "مسح النموذج",
                "download_excel": "تحميل Excel",
                "language": "اللغة:",
                "records_title": "السجلات",
                "delete": "حذف",
                "success_add": "تمت إضافة السجل بنجاح!",
                "success_export": "تم تصدير ملف Excel بنجاح!",
                "error_fill": "يرجى ملء رقم العمل واسم اللحام",
                "recording": "جاري التسجيل... تحدث الآن",
                "mic_tooltip": "انقر للتسجيل الصوتي",
                "confirm_text": "هل هذا صحيح: '{}'؟ (نعم للتأكيد وتثبيت الحقل، لا لإعادة التسجيل)",
            },
        }

        # Number word dictionaries for conversion
        self.number_dicts = {
            "en": {
                "zero": 0,
                "one": 1,
                "two": 2,
                "three": 3,
                "four": 4,
                "five": 5,
                "six": 6,
                "seven": 7,
                "eight": 8,
                "nine": 9,
                "ten": 10,
                "eleven": 11,
                "twelve": 12,
                "thirteen": 13,
                "fourteen": 14,
                "fifteen": 15,
                "sixteen": 16,
                "seventeen": 17,
                "eighteen": 18,
                "nineteen": 19,
                "twenty": 20,
                "thirty": 30,
                "forty": 40,
                "fifty": 50,
                "sixty": 60,
                "seventy": 70,
                "eighty": 80,
                "ninety": 90,
            },
            "ar": {
                "صفر": 0,
                "واحد": 1,
                "اثنان": 2,
                "ثلاثة": 3,
                "أربعة": 4,
                "خمسة": 5,
                "ستة": 6,
                "سبعة": 7,
                "ثمانية": 8,
                "تسعة": 9,
                "عشرة": 10,
                "أحد عشر": 11,
                "اثنا عشر": 12,
                "ثلاثة عشر": 13,
                "أربعة عشر": 14,
                "خمسة عشر": 15,
                "ستة عشر": 16,
                "سبعة عشر": 17,
                "ثمانية عشر": 18,
                "تسعة عشر": 19,
                "عشرون": 20,
                "ثلاثون": 30,
                "أربعون": 40,
                "خمسون": 50,
                "ستون": 60,
                "سبعون": 70,
                "ثمانون": 80,
                "تسعون": 90,
            },
        }

        # Whisper model (load once; use 'base' for speed, 'small' for better accuracy)
        self.whisper_model = None
        self.load_whisper_model()

        # Voice recognition
        self.is_recording = False
        self.audio = None
        self.stream = None

        # Build UI
        self.create_ui()
        self.update_language()

    def load_whisper_model(self):
        """Load Whisper model for speech recognition"""
        try:
            # Load Whisper model (change to 'small', 'medium', etc. for better accuracy; 'base' is fast)
            self.whisper_model = whisper.load_model("medium")
            print("Loaded Whisper model: base")
        except Exception as e:
            print(
                f"Error loading Whisper model: {e}. Install with: pip install openai-whisper"
            )
            self.whisper_model = None

    def words_to_digits(self, text, lang):
        """Convert number words in text to digits if the entire text represents a number.
        Handles sequences (concat) vs compounds (sum)."""
        if not text:
            return None

        # First, check if it's already digits (spoken as digits, e.g., "5 3 1")
        cleaned = re.sub(r"\s+", "", text.strip())
        if re.match(r"^\d+$", cleaned):
            return cleaned  # Already a number string

        # Normalize text
        if lang == "en":
            normalized = text.lower().strip()
        else:
            normalized = text.strip()  # Arabic doesn't have case

        # Split into words
        words = re.split(r"\s+and\s+|and\s+|,\s*|\s+", normalized)
        words = [w.strip() for w in words if w.strip()]

        if not words:
            return None

        # Check if all words are number words
        num_values = []
        single_digits_only = True
        for word in words:
            num = self.number_dicts[lang].get(word, None)
            if num is None:
                return None  # Not a pure number
            num_values.append(num)
            if num > 9:  # Tens or higher
                single_digits_only = False

        if not num_values:
            return None

        # Logic: If all single digits (0-9), concatenate as sequence (e.g., "five three one" → "531")
        # Else, sum for compounds (e.g., "twenty three" → 23)
        if single_digits_only:
            return "".join(str(num) for num in num_values)
        else:
            return str(sum(num_values))

    def create_ui(self):
        """Create the user interface"""
        # Header
        header_frame = tk.Frame(self.root, bg="#2a5298", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        self.title_label = tk.Label(
            header_frame,
            text="Welding Shop Manager",
            font=("Arial", 20, "bold"),
            bg="#2a5298",
            fg="white",
        )
        self.title_label.pack(side=tk.LEFT, padx=20, pady=20)

        # Language selector
        lang_frame = tk.Frame(header_frame, bg="#2a5298")
        lang_frame.pack(side=tk.RIGHT, padx=20)

        self.lang_label = tk.Label(
            lang_frame, text="Language:", font=("Arial", 11), bg="#2a5298", fg="white"
        )
        self.lang_label.pack(side=tk.LEFT, padx=5)

        self.lang_var = tk.StringVar(value="English")
        lang_combo = ttk.Combobox(
            lang_frame,
            textvariable=self.lang_var,
            values=["English", "العربية"],
            state="readonly",
            width=15,
            font=("Arial", 10),
        )
        lang_combo.pack(side=tk.LEFT)
        lang_combo.bind("<<ComboboxSelected>>", self.change_language)

        # Main container
        main_container = tk.Frame(self.root, bg="#f0f0f0")
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Left panel - Form
        form_frame = tk.LabelFrame(
            main_container,
            text="Add Welding Entry",
            font=("Arial", 12, "bold"),
            bg="white",
            padx=20,
            pady=20,
        )
        form_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))

        # Form fields
        self.fields = {}
        field_config = [
            ("job_id", "Job ID:"),
            ("weld_id", "Weld ID:"),
            ("kp_sec", "KP Sec:"),
            ("wps_no", "WPS No:"),
            ("material_gr", "Material Gr:"),
            ("heat_no", "Heat No:"),
            ("size", "Size(Inches):"),
            ("weld_side", "Weld Side:"),
            ("thk", "Thk(mm):"),
            ("root", "Root:"),
            ("material_comb", "Mtrl. Comb:"),
            ("pipe_no", "Pipe No:"),
            ("pipe_length", "Pipe length(mtrs):"),
            ("welder_name", "Welder Name:"),
            ("material", "Material:"),
            ("weld_type", "Weld Type:"),
            ("description", "Description:"),
            ("date", "Date:"),
        ]

        for i, (field_id, label_text) in enumerate(field_config):
            field_frame = tk.Frame(form_frame, bg="white")
            field_frame.pack(fill=tk.X, pady=8)

            label = tk.Label(
                field_frame,
                text=label_text,
                font=("Arial", 10, "bold"),
                bg="white",
                width=15,
                anchor="w",
            )
            label.pack(side=tk.LEFT)

            if field_id == "description":
                entry = tk.Text(field_frame, height=3, width=30, font=("Arial", 10))
            elif field_id == "date":
                entry = tk.Entry(field_frame, width=30, font=("Arial", 10))
                entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
            else:
                entry = tk.Entry(field_frame, width=30, font=("Arial", 10))

            entry.pack(side=tk.LEFT, padx=5)

            # Mic button
            mic_btn = tk.Button(
                field_frame,
                text="🎤",
                font=("Arial", 14),
                bg="#2a5298",
                fg="white",
                width=3,
                cursor="hand2",
                command=lambda f=field_id: self.record_voice(f),
            )
            mic_btn.pack(side=tk.LEFT)

            self.fields[field_id] = {"label": label, "entry": entry, "mic": mic_btn}

        # Buttons
        button_frame = tk.Frame(form_frame, bg="white")
        button_frame.pack(pady=20)

        self.add_btn = tk.Button(
            button_frame,
            text="Add Entry",
            font=("Arial", 11, "bold"),
            bg="#2a5298",
            fg="white",
            width=15,
            height=2,
            cursor="hand2",
            command=self.add_entry,
        )
        self.add_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = tk.Button(
            button_frame,
            text="Clear Form",
            font=("Arial", 11, "bold"),
            bg="#95a5a6",
            fg="white",
            width=15,
            height=2,
            cursor="hand2",
            command=self.clear_form,
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        # Status label
        self.status_label = tk.Label(
            form_frame, text="", font=("Arial", 10), bg="white", fg="green"
        )
        self.status_label.pack(pady=10)

        # Right panel - Records
        records_frame = tk.LabelFrame(
            main_container,
            text="Records",
            font=("Arial", 12, "bold"),
            bg="white",
            padx=10,
            pady=10,
        )
        records_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Export button
        self.export_btn = tk.Button(
            records_frame,
            text="Download Excel",
            font=("Arial", 11, "bold"),
            bg="#27ae60",
            fg="white",
            width=20,
            height=2,
            cursor="hand2",
            command=self.export_excel,
        )
        self.export_btn.pack(pady=10)

        # Records table
        table_frame = tk.Frame(records_frame, bg="white")
        table_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        v_scrollbar = tk.Scrollbar(table_frame, orient=tk.VERTICAL)
        h_scrollbar = tk.Scrollbar(table_frame, orient=tk.HORIZONTAL)

        self.tree = ttk.Treeview(
            table_frame,
            columns=("Job ID", "Welder", "Material", "Type", "Date", "Actions"),
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            height=15,
        )

        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)

        # Column headings
        for col in ("Job ID", "Welder", "Material", "Type", "Date", "Actions"):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        # Bind delete
        self.tree.bind("<Double-Button-1>", self.delete_record)

        self.refresh_table()

    def change_language(self, event=None):
        """Change application language"""
        lang = self.lang_var.get()
        self.current_lang = "ar" if lang == "العربية" else "en"
        self.update_language()
        # Reload Whisper model if needed (Whisper auto-detects language)

    def update_language(self):
        """Update UI text based on selected language"""
        t = self.translations[self.current_lang]

        self.title_label.config(text=t["title"])
        self.lang_label.config(text=t["language"])

        # Update field labels
        for field_id in self.fields:
            self.fields[field_id]["label"].config(text=t[field_id])

        # Update buttons
        self.add_btn.config(text=t["add_entry"])
        self.clear_btn.config(text=t["clear_form"])
        self.export_btn.config(text=t["download_excel"])

    def record_voice(self, field_id):
        """
        Speak the question (and show it in UI) then start recording.
        Re-inits TTS engine each time to fix 'silent after first call' bug.
        """
        print(f"[DEBUG] record_voice called for field: {field_id}")
        print(f"[DEBUG] is_recording flag: {self.is_recording}")

        if self.is_recording:
            print("[DEBUG] Recording already in progress, exiting")
            return

        if not self.whisper_model:
            print("[DEBUG] No Whisper model loaded")
            messagebox.showwarning(
                "Warning",
                "Whisper model not loaded. Install with: pip install openai-whisper",
            )
            return

        # Build language-specific question text
        field_label_text = self.fields[field_id]["label"].cget("text")
        question_en = f"What is your {field_label_text.replace(':','').lower()}?"
        question_ar = f"ما هو {field_label_text.replace(':','')}؟"
        question = question_ar if self.current_lang == "ar" else question_en
        print(f"[DEBUG] Question to speak: '{question}'")

        # Show the question in the status label immediately and force UI refresh
        try:
            self.status_label.config(text=question, fg="blue")
            self.root.update_idletasks()
            print("[DEBUG] UI updated with question")
        except Exception as e:
            print(f"[DEBUG] UI update failed: {e}")

        # Create FRESH TTS engine each time to avoid silent bug
        print("[DEBUG] Initializing fresh TTS engine...")
        tts_engine = None
        try:
            tts_engine = pyttsx3.init()
            voices = tts_engine.getProperty("voices")
            print(f"[DEBUG] Available voices: {len(voices)}")
            if len(voices) == 0:
                print("[DEBUG] WARNING: No TTS voices detected - install OS voices")
                messagebox.showwarning(
                    "TTS Warning", "No speech voices installed. Check OS settings."
                )
                # Still proceed to recording
            else:
                # Set a default voice (index 0 usually English; adjust for Arabic if needed)
                tts_engine.setProperty("voice", voices[0].id)
            tts_engine.setProperty("rate", 170)
            tts_engine.setProperty("volume", 1.0)
            print("[DEBUG] Engine initialized with voice")

            print("[DEBUG] Starting TTS sequence...")
            tts_engine.say(question)
            print("[DEBUG] After say(), calling runAndWait()")
            tts_engine.runAndWait()
            print("[DEBUG] runAndWait() completed")
            tts_engine.endLoop()  # Flush queue fully
            print("[DEBUG] endLoop() called")
            tts_engine.stop()
            print("[DEBUG] TTS sequence completed successfully")
        except Exception as e:
            print(f"[DEBUG] TTS Error in sequence: {e}")
            # If TTS fails entirely, still proceed to recording
        finally:
            # Clean up engine
            if tts_engine:
                try:
                    tts_engine.stop()
                    tts_engine = None
                except:
                    pass

        # Longer pause so TTS hardware finishes fully
        print("[DEBUG] Sleeping 0.3s after TTS")
        time.sleep(0.3)

        # Update status to recording then start recording in background thread
        t = self.translations[self.current_lang]
        try:
            self.status_label.config(text=t["recording"], fg="red")
            self.root.update_idletasks()
            print("[DEBUG] UI updated to recording status")
        except Exception as e:
            print(f"[DEBUG] Status update failed: {e}")

        # Start recording in background thread
        print("[DEBUG] Starting recording thread")
        threading.Thread(
            target=self._record_audio, args=(field_id,), daemon=True
        ).start()

    def _record_audio(self, field_id):
        """Record audio to WAV file for ~3 seconds, then transcribe with Whisper."""
        print(f"[DEBUG] _record_audio started for field: {field_id}")
        print(f"[DEBUG] Setting is_recording to True")

        if self.is_recording:
            print("[DEBUG] Re-entry detected, exiting")
            return

        self.is_recording = True
        audio_file = None
        text_output = ""

        audio = None
        stream = None
        try:
            print("[DEBUG] Initializing PyAudio")
            audio = pyaudio.PyAudio()

            # Create temp WAV file
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as temp_file:
                audio_file = temp_file.name

            # Open stream
            print("[DEBUG] Opening audio stream")
            stream = audio.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=16000,
                input=True,
                frames_per_buffer=4000,
            )

            # Record to file for 3s
            print("[DEBUG] Starting stream and recording loop (3s)")
            frames = []
            stream.start_stream()
            start = time.time()
            timeout_seconds = 7.0

            while time.time() - start < timeout_seconds:
                try:
                    data = stream.read(4000, exception_on_overflow=False)
                    frames.append(data)
                except Exception as e:
                    print(f"[DEBUG] Audio read error: {e}")
                    continue

            # Save to WAV
            wf = wave.open(audio_file, "wb")
            wf.setnchannels(1)
            wf.setsampwidth(audio.get_sample_size(pyaudio.paInt16))
            wf.setframerate(16000)
            wf.writeframes(b"".join(frames))
            wf.close()

            # Transcribe with Whisper
            print("[DEBUG] Transcribing with Whisper")
            if self.whisper_model:
                # Auto-detect language or specify
                lang_code = "ar" if self.current_lang == "ar" else "en"
                result = self.whisper_model.transcribe(audio_file, language=lang_code)
                text_output = result["text"].strip()

                # Remove punctuation and extra spaces
                clean_text = re.sub(r"(?<!\d)\.(?!\d)", "", text_output)  # remove dots not between numbers
                clean_text = re.sub(r"[^\w.\s]", "", clean_text)  # remove other punctuations but keep decimals
                clean_text = re.sub(r"\s+", " ", clean_text).strip() # normalize spaces

                print(f"[DEBUG] Whisper transcription: '{clean_text}'")
            else:
                print("[DEBUG] No Whisper model available")

        except Exception as e:
            print(f"[DEBUG] Recording/Transcription Error: {e}")
        finally:
            # Cleanup
            print("[DEBUG] Cleaning up audio resources")
            try:
                if stream:
                    stream.stop_stream()
                    stream.close()
                if audio:
                    audio.terminate()
            except Exception as e:
                print(f"[DEBUG] Cleanup error: {e}")

            # Delete temp file
            if audio_file and os.path.exists(audio_file):
                os.unlink(audio_file)

            # Reset status and flag
            try:
                t = self.translations[self.current_lang]
                self.root.after(
                    0, lambda: self.status_label.config(text="", fg="green")
                )
                print("[DEBUG] Scheduled status reset")
            except Exception as e:
                print(f"[DEBUG] Status reset error: {e}")

            self.is_recording = False

        print(f"[DEBUG] Raw text_output: '{clean_text}'")

        # Convert number words to digits if applicable
        converted = self.words_to_digits(clean_text, self.current_lang)
        if converted:
            clean_text = converted
            print(f"[DEBUG] Converted number: '{clean_text}'")
        else:
            print(f"[DEBUG] No number conversion applied")

        if clean_text:
            # Insert and confirm on main thread
            try:
                self.root.after(
                    0, lambda: self._insert_and_confirm(field_id, clean_text)
                )
                print("[DEBUG] Scheduled insert and confirm")
            except Exception as e:
                print(f"[DEBUG] Scheduling error: {e}")
                self._insert_and_confirm(field_id, clean_text)
            print(f"[DEBUG] Recognized for {field_id}: {clean_text}")
        else:
            print("[DEBUG] No speech detected or recognition returned empty text")

    def _insert_and_confirm(self, field_id, text):
        """Insert transcribed text, ask for confirmation, and lock if confirmed."""
        print(f"[DEBUG] _insert_and_confirm called for {field_id} with text: '{text}'")
        entry = self.fields[field_id]["entry"]
        mic_btn = self.fields[field_id]["mic"]

        # Insert the text
        if isinstance(entry, tk.Text):
            entry.insert(tk.END, text + " ")
            current_text = entry.get("1.0", tk.END).strip()
        else:
            current = entry.get()
            entry.delete(0, tk.END)
            entry.insert(0, current + " " + text if current else text)
            current_text = entry.get().strip()

        print(f"[DEBUG] Inserted text: '{current_text}'")

        # Ask for confirmation
        t = self.translations[self.current_lang]
        confirm_msg = t["confirm_text"].format(current_text)
        if messagebox.askyesno("Confirm Input", confirm_msg):
            # Confirmed: Make field uneditable and disable mic
            if isinstance(entry, tk.Text):
                entry.config(state="disabled")
            else:
                entry.config(state="readonly")
            mic_btn.config(state="disabled", bg="gray")
            print(f"[DEBUG] Field '{field_id}' locked after confirmation")
            self.status_label.config(text="Confirmed and locked!", fg="green")
            self.root.after(2000, lambda: self.status_label.config(text=""))
        else:
            # Not confirmed: Clear the field and re-enable mic
            if isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)
            else:
                entry.delete(0, tk.END)
            mic_btn.config(state="normal", bg="#2a5298")
            print(f"[DEBUG] Field '{field_id}' cleared for re-recording")
            self.status_label.config(text="Cleared - please re-record", fg="orange")
            self.root.after(2000, lambda: self.status_label.config(text=""))

    def _insert_text(self, field_id, text):
        """Insert transcribed text into field (deprecated, use _insert_and_confirm)"""
        pass  # Not used anymore

    def add_entry(self):
        """Add new welding entry"""
        t = self.translations[self.current_lang]

        # Get values
        job_id = self.fields["job_id"]["entry"].get().strip()
        weld_id = self.fields["weld_id"]["entry"].get().strip()
        kp_sec = self.fields["kp_sec"]["entry"].get().strip()
        wps_no = self.fields["wps_no"]["entry"].get().strip()
        material_gr = self.fields["material_gr"]["entry"].get().strip()
        heat_no = self.fields["heat_no"]["entry"].get().strip()
        size = self.fields["size"]["entry"].get().strip()
        thk = self.fields["thk"]["entry"].get().strip()
        weld_side = self.fields["weld_side"]["entry"].get().strip()
        root = self.fields["root"]["entry"].get().strip()
        material_comb = self.fields["material_comb"]["entry"].get().strip()
        pipe_no = self.fields["pipe_no"]["entry"].get().strip()
        pipe_length = self.fields["pipe_length"]["entry"].get().strip()
        welder_name = self.fields["welder_name"]["entry"].get().strip()
        material = self.fields["material"]["entry"].get().strip()
        weld_type = self.fields["weld_type"]["entry"].get().strip()
        description = self.fields["description"]["entry"].get("1.0", tk.END).strip()
        date = self.fields["date"]["entry"].get().strip()


        # Validation
        if not job_id or not welder_name:
            messagebox.showerror("Error", t["error_fill"])
            return

        # Create record
        # Create record
        record = {
            "id": datetime.now().timestamp(),
            "job_id": job_id,
            "weld_id": weld_id,
            "kp_sec": kp_sec,
            "wps_no": wps_no,
            "material_gr": material_gr,
            "heat_no": heat_no,
            "size": size,
            "thk": thk,
            "weld_side": weld_side,
            "root": root,
            "material_comb": material_comb,
            "pipe_no": pipe_no,
            "pipe_length": pipe_length,
            "welder_name": welder_name,
            "material": material,
            "weld_type": weld_type,
            "description": description,
            "date": date,
        }

        self.records.insert(0, record)
        self.save_data()
        self.refresh_table()
        self.clear_form()

        self.status_label.config(text=t["success_add"], fg="green")
        self.root.after(3000, lambda: self.status_label.config(text=""))

    def clear_form(self):
        """Clear all form fields"""
        for field_id in [
            "job_id",
            "weld_id",
            "kp_sec",
            "wps_no",
            "material_gr",
            "heat_no",
            "size",
            "thk",
            "weld_side",
            "root",
            "material_comb",
            "pipe_no",
            "pipe_length",
            "welder_name",
            "material",
            "weld_type",
            "description",
            "date",
        ]:
            field = self.fields[field_id]
            entry = field["entry"]
            mic = field["mic"]
            if isinstance(entry, tk.Text):
                entry.config(state="normal")
                entry.delete("1.0", tk.END)
            else:
                entry.config(state="normal")
                entry.delete(0, tk.END)
            mic.config(state="normal", bg="#2a5298")

        # Reset date
        self.fields["date"]["entry"].config(state="normal")
        self.fields["date"]["entry"].delete(0, tk.END)
        self.fields["date"]["entry"].insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.fields["date"]["mic"].config(state="normal", bg="#2a5298")

    def refresh_table(self):
        """Refresh the records table"""
        # Clear existing
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add records
        for record in self.records:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    record["job_id"],
                    record["welder_name"],
                    record["material"],
                    record["weld_type"],
                    record["date"],
                    "Double-click to delete",
                ),
            )

    def delete_record(self, event):
        """Delete selected record"""
        selection = self.tree.selection()
        if not selection:
            return

        if messagebox.askyesno("Confirm", "Delete this record?"):
            item = self.tree.item(selection[0])
            job_id = item["values"][0]

            self.records = [r for r in self.records if r["job_id"] != job_id]
            self.save_data()
            self.refresh_table()

    def export_excel(self):
        """Export records to Excel"""
        if not self.records:
            messagebox.showwarning("Warning", "No records to export")
            return

        t = self.translations[self.current_lang]

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"welding_records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )

        if not filename:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Welding Records"

            # Header style
            header_fill = PatternFill(
                start_color="2a5298", end_color="2a5298", fill_type="solid"
            )
            header_font = Font(bold=True, color="FFFFFF", size=12)
            border = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )

            # Headers
            headers = [
                "Job ID",
                "Weld ID",
                "KP Sec",
                "WPS No",
                "Material Gr",
                "Heat No",
                "Size",
                "Thk",
                "Weld Side",
                "Root",
                "Material Comb",
                "Pipe No",
                "Pipe Length",
                "Welder Name",
                "Material",
                "Weld Type",
                "Description",
                "Date",
            ]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

            # Data
            for row, record in enumerate(self.records, 2):
                ws.cell(row=row, column=1, value=record["job_id"]).border = border
                ws.cell(row=row, column=2, value=record["weld_id"]).border = border
                ws.cell(row=row, column=3, value=record["kp_sec"]).border = border
                ws.cell(row=row, column=4, value=record["wps_no"]).border = border
                ws.cell(row=row, column=5, value=record["material_gr"]).border = border
                ws.cell(row=row, column=6, value=record["heat_no"]).border = border
                ws.cell(row=row, column=7, value=record["size"]).border = border
                ws.cell(row=row, column=8, value=record["thk"]).border = border
                ws.cell(row=row, column=9, value=record["weld_side"]).border = border
                ws.cell(row=row, column=10, value=record["root"]).border = border
                ws.cell(row=row, column=11, value=record["material_comb"]).border = border
                ws.cell(row=row, column=12, value=record["pipe_no"]).border = border
                ws.cell(row=row, column=13, value=record["pipe_length"]).border = border
                ws.cell(row=row, column=14, value=record["welder_name"]).border = border
                ws.cell(row=row, column=15, value=record["material"]).border = border
                ws.cell(row=row, column=16, value=record["weld_type"]).border = border
                ws.cell(row=row, column=17, value=record["description"]).border = border
                ws.cell(row=row, column=18, value=record["date"]).border = border

            # Adjust column widths
            for col in range(1, 19):
                ws.column_dimensions[chr(64 + col)].width = 20

            wb.save(filename)

            self.status_label.config(text=t["success_export"], fg="green")
            self.root.after(3000, lambda: self.status_label.config(text=""))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {str(e)}")

    def load_data(self):
        """Load records from JSON file"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, "r", encoding="utf-8") as f:
                    self.records = json.load(f)
            except:
                self.records = []

    def save_data(self):
        """Save records to JSON file"""
        try:
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving data: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = WeldingShopApp(root)
    root.mainloop()
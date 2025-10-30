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
import customtkinter as ctk  # New import for modern UI
import platform
import subprocess
from PIL import Image


def safe_json_loads(data):
    try:
        return json.loads(data)
    except Exception:
        return {}


# Set CustomTkinter appearance (modern look)
ctk.set_appearance_mode("light")  # Modes: "light", "dark", "system"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"


class WeldingShopApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Al Tasnim Enterprises LLC")
        self.root.geometry("1000x700")

        # Data storage
        self.records = []
        self.data_file = "welding_data.json"
        self.load_data()

        # Language settings
        self.current_lang = "en"
        self.translations = {
            "en": {
                "title": "AL TASNIM ENTERPRISES LLC",
                "add_form_title": "Add Welding Entry",
                "job_id": "Job ID:",
                "contract_number": "Contract No.",
                "contract_title": "Contract Title:",
                "report_number": "Report No.",
                "po_wo_number": "PO / WO No.",
                "client_wps_number": "Client WPS No.",
                "project_title_wellID": "Project Title / Well ID:",
                "drawing_no": "Drawing/ISO No.",
                "line_no": "Line No.",
                "site_name": "Site Name:",
                "job_desc": "Job description:",
                "location": "Location:",
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
                "download_excel": "Submit",
                "language": "Language:",
                "theme": "Theme:",
                "records_title": "Records",
                "delete": "Delete",
                "success_add": "Entry added successfully!",
                "success_export": "Form exported successfully!",
                "error_fill": "Please fill Job ID and Welder Name",
                "recording": "Recording... Speak now",
                "mic_tooltip": "Click to record voice",
                "confirm_text": "Is this correct: '{}'? (Yes to confirm and lock field, No to re-record)",
                },
                "ar": {
                "title": "إدارة ورشة اللحام",
                "add_form_title": "إضافة سجل اللحام",
                "job_id": "رقم العمل:",
                "contract_number": "رقم العقد:",
                "contract_title": "عنوان العقد:",
                "report_number": "رقم التقرير:",
                "po_wo_number": "رقم PO/WO:",
                "client_wps_number": "رقم WPS العميل:",
                "project_title_wellID": "عنوان المشروع / معرف البئر:",
                "drawing_no": "رقم الرسم/ISO:",
                "line_no": "رقم الخط:",
                "site_name": "اسم الموقع:",
                "job_desc": "وصف العمل:",
                "location": "الموقع:",
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
                "download_excel": "إرسال النموذج",
                "language": "اللغة:",
                "theme": "الثيم:",
                "records_title": "السجلات",
                "delete": "حذف",
                "success_add": "تمت إضافة السجل بنجاح!",
                "success_export": "تم تصدير النموذج بنجاح!",
                "error_fill": "يرجى ملء رقم العمل واسم اللحام",
                "recording": "جاري التسجيل... تحدث الآن",
                "mic_tooltip": "انقر للتسجيل الصوتي",
                "confirm_text": "هل هذا صحيح: '{}'؟ (نعم للتأكيد وتثبيت الحقل، لا لإعادة التسجيل)",
                },
        }

        # Number word dictionaries for conversion
        self.number_dicts = {
            "en": {
                "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
                "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
                "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16,
                "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
                "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
                "eighty": 80, "ninety": 90,
            },
            "ar": {
                "صفر": 0, "واحد": 1, "اثنان": 2, "ثلاثة": 3, "أربعة": 4, "خمسة": 5,
                "ستة": 6, "سبعة": 7, "ثمانية": 8, "تسعة": 9, "عشرة": 10,
                "أحد عشر": 11, "اثنا عشر": 12, "ثلاثة عشر": 13, "أربعة عشر": 14,
                "خمسة عشر": 15, "ستة عشر": 16, "سبعة عشر": 17, "ثمانية عشر": 18,
                "تسعة عشر": 19, "عشرون": 20, "ثلاثون": 30, "أربعون": 40,
                "خمسون": 50, "ستون": 60, "سبعون": 70, "ثمانون": 80, "تسعون": 90,
            },
        }

        # Whisper model
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
        try:
            self.whisper_model = whisper.load_model("medium")
            print("Loaded Whisper model: medium")
        except Exception as e:
            print(f"Error loading Whisper model: {e}. Install with: pip install openai-whisper")
            self.whisper_model = None

    def words_to_digits(self, text, lang):
        if not text:
            return None
        cleaned = re.sub(r"\s+", "", text.strip())
        if re.match(r"^\d+$", cleaned):
            return cleaned
        if lang == "en":
            normalized = text.lower().strip()
        else:
            normalized = text.strip()
        words = re.split(r"\s+and\s+|and\s+|,\s*|\s+", normalized)
        words = [w.strip() for w in words if w.strip()]
        if not words:
            return None
        num_values = []
        single_digits_only = True
        for word in words:
            num = self.number_dicts[lang].get(word, None)
            if num is None:
                return None
            num_values.append(num)
            if num > 9:
                single_digits_only = False
        if not num_values:
            return None
        if single_digits_only:
            return "".join(str(num) for num in num_values)
        else:
            return str(sum(num_values))

    # -------------------- TTS helpers --------------------
    def speak_sync(self, text, timeout=20):
        """
        Speak text synchronously using platform-native method when possible.
        """
        try:
            system = platform.system()
            if system == "Windows":
                # Use PowerShell SAPI synchronously and pass text via stdin
                ps_script = (
                    "Add-Type -AssemblyName System.Speech;"
                    + "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer;"
                    + "$s.Speak([Console]::In.ReadToEnd());"
                )
                subprocess.run([
                    "powershell",
                    "-Command",
                    ps_script,
                ], input=text.encode("utf-8"), timeout=timeout)
                return

            if system == "Darwin":
                subprocess.run(["say", text], check=False, timeout=timeout)
                return

            # Fallback: pyttsx3
            local_tts = pyttsx3.init()
            try:
                local_tts.setProperty("rate", 160)
            except Exception:
                pass
            local_tts.say(text)
            local_tts.runAndWait()
            try:
                local_tts.stop()
            except Exception:
                pass
        except Exception as e:
            print(f"[speak_sync] error: {e}")

    def speak_async(self, text):
        """Run speak_sync in a background thread."""
        threading.Thread(target=lambda: self.speak_sync(text), daemon=True).start()

    # -------------------- UI / Flow --------------------
    def create_ui(self):
        """Build modern UI with CustomTkinter."""
        # ---------- HEADER ------------------------------------------------
        # header with orange background
        header = ctk.CTkFrame(self.root, height=70, corner_radius=0, fg_color="#FF8C00")  # or "#FFA500"
        header.pack(fill="x")
        header.pack_propagate(False)

        try:
            logo_path = os.path.join(os.getcwd(), "logo.png")   # adjust if logo is elsewhere
            print(f"[DEBUG] ------------------------------------------------ Loading logo from: {logo_path}")
            if os.path.exists(logo_path):
                pil_img = Image.open("logo.png")
                pil_img = pil_img.resize((140, 40), Image.LANCZOS)   # (width, height) forced — stretches
                self.logo_img = ctk.CTkImage(pil_img, size=(140, 40))
                self.logo_label = ctk.CTkLabel(header, image=self.logo_img, text="")
                self.logo_label.pack(side="left", padx=(6, 4), pady=12)

            else:
                print(f"[WARN] logo.png not found at {logo_path}")
        except Exception as e:
            print(f"[WARN] Failed to load logo: {e}")
        # ===============================================================

        self.title_label = ctk.CTkLabel(
            header,
            text="Al Tasnim Enterprises LLC",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color="white"
        )
        self.title_label.pack(side="left", padx=20, pady=15)

        # Theme toggle (new: dark/light switch)
        theme_fr = ctk.CTkFrame(header)
        theme_fr.pack(side="right", padx=20, pady=15)
        ctk.CTkLabel(theme_fr, text="Theme:").pack(side="left", padx=(0, 5))
        self.theme_switch = ctk.CTkSwitch(
            theme_fr, text="", command=self.toggle_theme, onvalue="dark", offvalue="light"
        )
        self.theme_switch.pack(side="left")
        self.theme_switch.select() if ctk.get_appearance_mode() == "dark" else None

        # Language selector
        lang_fr = ctk.CTkFrame(header, fg_color="transparent", corner_radius=0)
        lang_fr.pack(side="right", padx=(8, 10), pady=12)   # smaller right padding so it sits closer to title/logo

        # white label so it shows up on the orange header
        lang_label = ctk.CTkLabel(
            lang_fr,
            text="Language:",
            text_color="white",
            font=ctk.CTkFont(size=12, weight="bold")   # <-- bold and slightly larger
        )
        lang_label.pack(side="left", padx=(0, 6))

        self.lang_var = ctk.StringVar(value="English")
        lang_cb = ctk.CTkComboBox(
            lang_fr,
            values=["English", "العربية"],
            variable=self.lang_var,
            state="readonly",
            width=120,
            command=lambda v: self.change_language(v)   # CTkComboBox passes the selected value
        )

        lang_cb.pack(side="left")

        # ---------- MAIN SCROLLABLE CONTENT -------------------------------
        scrollable = ctk.CTkScrollableFrame(self.root, corner_radius=10)
        scrollable.pack(fill="both", expand=True, padx=20, pady=20)

        # ---------- LEFT: FORM --------------------------------------------
        form_frame = ctk.CTkFrame(scrollable, corner_radius=10)
        form_frame.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # Use grid for the entire form_frame
        form_frame.grid_columnconfigure(1, weight=1)
        form_frame.grid_rowconfigure(0, weight=1)

        form_title = ctk.CTkLabel(form_frame, text="Add Welding Entry", font=ctk.CTkFont(size=14, weight="bold"))
        form_title.grid(row=0, column=0, columnspan=3, pady=10, sticky="ew")

        # Two-column grid for fields
        self.fields = {}
        field_cfg = [
            ("job_id", "Job ID:"), ("contract_number", "Contract No."), ("contract_title", "Contract Title:"), ("report_number", "Report No.:"), ("po_wo_number", "PO / WO No.:"), ("client_wps_number", "Client WPS No.:"), ("project_title_wellID", "Project Title / Well ID:"), ("drawing_no", "Drawing/ISO No.:"), ("line_no", "Line No.:"), ("site_name", "Site Name:"), ("job_desc", "Job description:"), ("location", "Location:"),("weld_id", "Weld ID:"), ("kp_sec", "KP Sec:"), ("wps_no", "WPS No:"),
            ("material_gr", "Material Gr:"), ("heat_no", "Heat No:"), ("size", "Size(Inches):"), ("thk", "Thk(mm):"),
            ("weld_side", "Weld Side:"), ("root", "Root:"), ("material_comb", "Mtrl. Comb:"), ("pipe_no", "Pipe No:"),
            ("pipe_length", "Pipe length(mtrs):"), ("welder_name", "Welder Name:"), ("material", "Material:"),
            ("weld_type", "Weld Type:"), ("description", "Description:"), ("date", "Date:"),
        ]

        rows = (len(field_cfg) + 1) // 2
        for idx, (fid, label_txt) in enumerate(field_cfg):
            row = idx // 2 + 1  # Start after title
            col = (idx % 2) * 3  # Label | Entry | Mic

            # Label
            lbl = ctk.CTkLabel(form_frame, text=label_txt, anchor="w")
            lbl.grid(row=row, column=col, padx=10, pady=5, sticky="w")

            # Entry/Text
            if fid == "description":
                entry = ctk.CTkTextbox(form_frame, height=80, corner_radius=8)
            elif fid == "date":
                entry = ctk.CTkEntry(form_frame, width=200)
                entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
            else:
                entry = ctk.CTkEntry(form_frame, width=200)
            entry.grid(row=row, column=col + 1, padx=10, pady=5, sticky="ew")

            # Mic button
            mic = ctk.CTkButton(
                form_frame, text="🎤", width=40, height=30, font=ctk.CTkFont(size=14),
                command=lambda f=fid: self.record_voice(f), corner_radius=20
            )
            mic.grid(row=row, column=col + 2, padx=5, pady=5)

            self.fields[fid] = {"label": lbl, "entry": entry, "mic": mic}

        # Form buttons - use grid
        btn_fr = ctk.CTkFrame(form_frame)
        btn_fr.grid(row=rows + 1, column=0, columnspan=3, pady=15, sticky="ew")
        btn_fr.grid_columnconfigure(0, weight=1)
        btn_fr.grid_columnconfigure(1, weight=1)

        self.add_btn = ctk.CTkButton(btn_fr, text="Add Entry", command=self.add_entry)
        self.add_btn.grid(row=0, column=0, padx=5, sticky="ew")

        self.clear_btn = ctk.CTkButton(btn_fr, text="Clear Form", fg_color="gray", command=self.clear_form)
        self.clear_btn.grid(row=0, column=1, padx=5, sticky="ew")

        # Status label - use grid
        self.status_label = ctk.CTkLabel(form_frame, text="")
        self.status_label.grid(row=rows + 2, column=0, columnspan=3, pady=5, sticky="ew")

        # ---------- RIGHT: RECORDS ---------------------------------------
        rec_frame = ctk.CTkFrame(scrollable, corner_radius=10)
        rec_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        rec_title = ctk.CTkLabel(rec_frame, text="Records", font=ctk.CTkFont(size=14, weight="bold"))
        rec_title.pack(pady=10)

        # Export button
        self.export_btn = ctk.CTkButton(rec_frame, text="Submit", command=self.export_excel)
        self.export_btn.pack(pady=5, fill="x")

        # Treeview in a frame
        tree_fr = ctk.CTkFrame(rec_frame)
        tree_fr.pack(fill="both", expand=True, pady=5)

        self.tree = ttk.Treeview(
            tree_fr, columns=("Job ID", "Welder", "Material", "Type", "Date", "Actions"), show="headings", height=15
        )
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        v_scroll = ttk.Scrollbar(tree_fr, orient="vertical", command=self.tree.yview)
        h_scroll = ttk.Scrollbar(tree_fr, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")
        h_scroll.pack(side="bottom", fill="x")

        self.tree.bind("<Double-1>", self.delete_record)

        self.refresh_table()

    def toggle_theme(self):
        mode = "dark" if self.theme_switch.get() == "on" else "light"
        ctk.set_appearance_mode(mode)

    def change_language(self, value_or_event=None):
        # CTkComboBox passes a string (the selected value).
        # If called via an event (unlikely here), fall back to lang_var.
        if isinstance(value_or_event, str):
            sel = value_or_event
        else:
            sel = self.lang_var.get()

        print(f"[DEBUG] change_language called, selected={sel}")
        self.current_lang = "ar" if sel == "العربية" else "en"
        self.update_language()

    def update_language(self):
        t = self.translations[self.current_lang]
        self.title_label.configure(text=t["title"])
        # Update form/rec titles (assuming refs; in full impl, store them as self.form_title = form_title)
        # For now, hardcode or loop; here we skip dynamic title update for simplicity
        for fid in self.fields:
            self.fields[fid]["label"].configure(text=t[fid])
        self.add_btn.configure(text=t["add_entry"])
        self.clear_btn.configure(text=t["clear_form"])
        self.export_btn.configure(text=t["download_excel"])
        self.status_label.configure(text="")

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
        question_en = f"What is your {field_label_text.replace(':', '').lower()}?"
        question_ar = f"ما هو {field_label_text.replace(':', '')}؟"
        question = question_ar if self.current_lang == "ar" else question_en
        print(f"[DEBUG] Question to speak: '{question}'")
        # Show the question in the status label immediately and force UI refresh
        try:
            self.status_label.configure(text=question)
            self.root.update_idletasks()
            print("[DEBUG] UI updated with question")
        except Exception as e:
            print(f"[DEBUG] UI update failed: {e}")
        # Speak synchronously (ensures TTS finishes before recording starts)
        try:
            self.speak_sync(question)
        except Exception as e:
            print(f"[DEBUG] speak_sync failed: {e}")
        # Longer pause so TTS hardware finishes fully
        print("[DEBUG] Sleeping 0.2s after TTS")
        time.sleep(0.2)
        # Update status to recording then start recording in background thread
        t = self.translations[self.current_lang]
        try:
            self.status_label.configure(text=t["recording"])
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
        clean_text = ""
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
            timeout_seconds = 5.0
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
                print(f"[DEBUG] Raw Whisper transcription: '{text_output}'")
            else:
                print("[DEBUG] No Whisper model available")
                text_output = ""
            # Clean the text output
            clean_text = re.sub(r"(?<!\d)\.(?!\d)", "", text_output)  # remove dots not between numbers
            clean_text = re.sub(r"[^\w.\s-]", "", clean_text)         # keep hyphens and dots
            clean_text = re.sub(r"\s+", " ", clean_text).strip()  # normalize spaces
            print(f"[DEBUG] Cleaned transcription: '{clean_text}'")
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
                    0, lambda: self.status_label.configure(text="")
                )
                print("[DEBUG] Scheduled status reset")
            except Exception as e:
                print(f"[DEBUG] Status reset error: {e}")
            self.is_recording = False
        print(f"[DEBUG] Final text_output: '{text_output}'")
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
                threading.Thread(
                    target=lambda: self._voice_confirm(field_id, clean_text), daemon=True
                ).start()
                print("[DEBUG] Scheduled insert and confirm")
            except Exception as e:
                print(f"[DEBUG] Scheduling error: {e}")
            print(f"[DEBUG] Recognized for {field_id}: {clean_text}")
        else:
            print("[DEBUG] No speech detected or recognition returned empty text")

    def _voice_confirm(self, field_id, recognized_text, max_retries=2):
        """
        Voice-based confirmation after transcription with robust handling:
        - waits briefly after TTS so audio device frees
        - warms up and discards first frames (often silent)
        - retries listening a few times on empty/unrecognized responses
        - falls back to a messagebox if voice confirmation repeatedly fails
        """
        print(f"[DEBUG] Voice confirm started for {field_id} with text: {recognized_text}")

        # speak the confirmation synchronously so TTS output finishes before we record
        confirm_text = f"You said {recognized_text}. Do you confirm?"
        try:
            self.speak_sync(confirm_text)
        except Exception as e:
            print(f"[DEBUG] speak_sync failed: {e}")

        # tiny pause to ensure audio device is free after TTS (important on some systems)
        time.sleep(0.25)

        yes_keywords_en = {"yes", "yeah", "yup", "yep", "confirm", "correct"}
        no_keywords_en = {"no", "nah", "nope", "incorrect", "wrong"}
        yes_keywords_ar = {"نعم", "ايوه", "ايه", "نَعَم"}
        no_keywords_ar = {"لا", "لأ", "لاا"}

        attempt = 0
        while attempt <= max_retries:
            attempt += 1
            confirm_audio = None
            audio = None
            stream = None
            try:
                # create temp file
                with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                    confirm_audio = tmp.name

                audio = pyaudio.PyAudio()

                # open stream
                stream = audio.open(format=pyaudio.paInt16,
                                    channels=1,
                                    rate=16000,
                                    input=True,
                                    frames_per_buffer=4000)

                # warm-up: read & discard a short amount to avoid initial silence
                try:
                    for _ in range(0, 2):
                        stream.read(4000, exception_on_overflow=False)
                except Exception as e:
                    # not critical, continue
                    print(f"[DEBUG] warmup read error: {e}")

                # record for a bit (2.5s) to capture reply
                frames = []
                record_seconds = 2.5
                start = time.time()
                while time.time() - start < record_seconds:
                    try:
                        data = stream.read(4000, exception_on_overflow=False)
                        frames.append(data)
                    except Exception as e:
                        print(f"[DEBUG] confirmation read error: {e}")
                        # continue reading attempts until timeout
                        continue

                # save file
                wf = wave.open(confirm_audio, "wb")
                wf.setnchannels(1)
                wf.setsampwidth(audio.get_sample_size(pyaudio.paInt16))
                wf.setframerate(16000)
                wf.writeframes(b"".join(frames))
                wf.close()

                # transcribe
                lang_code = "ar" if self.current_lang == "ar" else "en"
                if not self.whisper_model:
                    print("[DEBUG] No Whisper model for confirmation")
                    response_text = ""
                else:
                    result = self.whisper_model.transcribe(confirm_audio, language=lang_code)
                    response_text = result.get("text", "").strip().lower()

                print(f"[DEBUG] Voice response detected: '{response_text}' (attempt {attempt})")

                # clean temp audio
                try:
                    if confirm_audio and os.path.exists(confirm_audio):
                        os.unlink(confirm_audio)
                        confirm_audio = None
                except Exception:
                    pass

                # analyze response
                if not response_text:
                    # nothing recognized — retry (with a short spoken prompt)
                    if attempt <= max_retries:
                        self.speak_async("I didn't hear you. Please say yes or no.")
                        time.sleep(0.15)
                        continue
                    else:
                        print("[DEBUG] No speech after retries, falling back to GUI confirm")
                        break

                # check for yes/no in either language
                # Arabic responses might include Arabic words; check both sets
                tokens = set(re.split(r"\s+|[^\w\u0600-\u06FF]+", response_text))
                if (tokens & yes_keywords_en) or (tokens & yes_keywords_ar):
                    print("[DEBUG] User confirmed by voice (YES)")
                    # insert (and lock) the recognized_text into field
                    try:
                        # use your helper that handles CTk widgets properly
                        self._insert_field_value(field_id, recognized_text)
                    except Exception as e:
                        print(f"[DEBUG] _insert_field_value failed: {e}")
                        # fallback manual locking:
                        entry = self.fields[field_id]["entry"]
                        mic_btn = self.fields[field_id]["mic"]
                        try:
                            if isinstance(entry, ctk.CTkTextbox):
                                entry.configure(state="normal")
                                entry.delete("1.0", "end")
                                entry.insert("end", recognized_text + " ")
                                entry.configure(state="disabled")
                            else:
                                entry.configure(state="normal")
                                entry.delete(0, "end")
                                entry.insert(0, recognized_text)
                                entry.configure(state="readonly")
                            mic_btn.configure(state="disabled", fg_color="gray")
                        except Exception as ee:
                            print(f"[DEBUG] manual insert/lock failed: {ee}")
                    self.status_label.configure(text="Confirmed and locked!")
                    self.root.after(2000, lambda: self.status_label.configure(text=""))
                    return

                if (tokens & no_keywords_en) or (tokens & no_keywords_ar):
                    print("[DEBUG] User rejected by voice (NO)")
                    # speak a retry prompt and clear field for re-recording
                    self.speak_async("Okay, please say it again.")
                    entry = self.fields[field_id]["entry"]
                    mic_btn = self.fields[field_id]["mic"]
                    try:
                        if isinstance(entry, ctk.CTkTextbox):
                            entry.configure(state="normal")
                            entry.delete("1.0", "end")
                        else:
                            entry.configure(state="normal")
                            entry.delete(0, "end")
                        mic_btn.configure(state="normal", fg_color="#1E90FF")
                        self.status_label.configure(text="Cleared - please re-record")
                        self.root.after(2000, lambda: self.status_label.configure(text=""))
                    except Exception as e:
                        print(f"[DEBUG] failed to clear field after NO: {e}")
                    return

                # unrecognized but not empty => ask to repeat
                if attempt <= max_retries:
                    self.speak_async("I didn't catch that. Please say yes or no.")
                    time.sleep(0.15)
                    continue
                else:
                    print("[DEBUG] Unclear confirmation after retries, falling back")
                    break

            except Exception as e:
                print(f"[DEBUG] Voice confirm error on attempt {attempt}: {e}")
                # cleanup and possibly retry
                try:
                    if stream:
                        stream.stop_stream()
                        stream.close()
                    if audio:
                        audio.terminate()
                except Exception:
                    pass
                time.sleep(0.1)
                continue
            finally:
                try:
                    if stream:
                        stream.stop_stream()
                        stream.close()
                    if audio:
                        audio.terminate()
                except Exception:
                    pass
                if confirm_audio and os.path.exists(confirm_audio):
                    try:
                        os.unlink(confirm_audio)
                    except Exception:
                        pass

        # fallback to GUI confirm if voice attempts didn't reach a decision
        try:
            t = self.translations[self.current_lang]
            confirm_msg = t["confirm_text"].format(recognized_text)
            if messagebox.askyesno("Confirm Input", confirm_msg):
                self._insert_field_value(field_id, recognized_text)
                self.status_label.configure(text="Confirmed and locked!")
                self.root.after(2000, lambda: self.status_label.configure(text=""))
            else:
                # user chose No via GUI fallback
                entry = self.fields[field_id]["entry"]
                mic_btn = self.fields[field_id]["mic"]
                if isinstance(entry, ctk.CTkTextbox):
                    entry.configure(state="normal")
                    entry.delete("1.0", "end")
                else:
                    entry.configure(state="normal")
                    entry.delete(0, "end")
                mic_btn.configure(state="normal", fg_color="#1E90FF")
                self.status_label.configure(text="Cleared - please re-record")
                self.root.after(2000, lambda: self.status_label.configure(text=""))
        except Exception as e:
            print(f"[DEBUG] Fallback GUI confirm error: {e}")

    def _insert_field_value(self, field_id, value):
        """Directly insert text into a field without popup confirmation."""
        try:
            entry_widget = self.fields[field_id]["entry"]
            # CTkTextbox vs CTkEntry handling
            if isinstance(entry_widget, ctk.CTkTextbox):
                entry_widget.configure(state="normal")
                entry_widget.delete("1.0", "end")
                entry_widget.insert("end", value)
                entry_widget.configure(state="disabled")
            else:
                entry_widget.configure(state="normal")
                entry_widget.delete(0, "end")
                entry_widget.insert(0, value)
                entry_widget.configure(state="readonly")

            # Also disable mic button for this field
            mic_button = self.fields[field_id].get("mic")
            if mic_button:
                mic_button.configure(state="disabled", fg_color="gray")

            print(f"[DEBUG] Inserted and locked field '{field_id}' with value: {value}")

        except Exception as e:
            print(f"[DEBUG] _insert_field_value error for {field_id}: {e}")

    def _insert_and_confirm(self, field_id, text):
        """Insert transcribed text, ask for confirmation, and lock if confirmed."""
        print(f"[DEBUG] _insert_and_confirm called for {field_id} with text: '{text}'")
        entry = self.fields[field_id]["entry"]
        mic_btn = self.fields[field_id]["mic"]
        # Insert the text
        if isinstance(entry, ctk.CTkTextbox):
            entry.insert("end", text + " ")
            current_text = entry.get("1.0", "end").strip()
        else:
            current = entry.get()
            entry.delete(0, "end")
            entry.insert(0, current + " " + text if current else text)
            current_text = entry.get().strip()
        print(f"[DEBUG] Inserted text: '{current_text}'")
        # Ask for confirmation
        t = self.translations[self.current_lang]
        confirm_msg = t["confirm_text"].format(current_text)
        if messagebox.askyesno("Confirm Input", confirm_msg):
            # Confirmed: Make field uneditable and disable mic
            if isinstance(entry, ctk.CTkTextbox):
                entry.configure(state="disabled")
            else:
                entry.configure(state="readonly")
            mic_btn.configure(state="disabled", fg_color="gray")
            print(f"[DEBUG] Field '{field_id}' locked after confirmation")
            self.status_label.configure(text="Confirmed and locked!")
            self.root.after(2000, lambda: self.status_label.configure(text=""))
        else:
            # Not confirmed: Clear the field and re-enable mic
            if isinstance(entry, ctk.CTkTextbox):
                entry.delete("1.0", "end")
            else:
                entry.delete(0, "end")
            mic_btn.configure(state="normal", fg_color="#1E90FF")
            print(f"[DEBUG] Field '{field_id}' cleared for re-recording")
            self.status_label.configure(text="Cleared - please re-record")
            self.root.after(2000, lambda: self.status_label.configure(text=""))

    def _insert_text(self, field_id, text):
        """Insert transcribed text into field (deprecated, use _insert_and_confirm)"""
        pass  # Not used anymore

    def add_entry(self):
        """Add new welding entry"""
        t = self.translations[self.current_lang]

        # Get values
        job_id = self.fields["job_id"]["entry"].get().strip()
        contract_number = self.fields["contract_number"]["entry"].get().strip()
        contract_title = self.fields["contract_title"]["entry"].get().strip()
        report_number = self.fields["report_number"]["entry"].get().strip()
        po_wo_number = self.fields["po_wo_number"]["entry"].get().strip()
        client_wps_number = self.fields["client_wps_number"]["entry"].get().strip()
        project_title_wellID = self.fields["project_title_wellID"]["entry"].get().strip()
        drawing_no = self.fields["drawing_no"]["entry"].get().strip()
        line_no = self.fields["line_no"]["entry"].get().strip()
        site_name = self.fields["site_name"]["entry"].get().strip()
        job_desc = self.fields["job_desc"]["entry"].get().strip()
        location = self.fields["location"]["entry"].get().strip()
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
        description = self.fields["description"]["entry"].get("1.0", "end").strip()
        date = self.fields["date"]["entry"].get().strip()

        # Validation
        if not job_id or not welder_name:
            messagebox.showerror("Error", t["error_fill"])
            return
        # Create record
        record = {
            "id": datetime.now().timestamp(),
            "job_id": job_id,
            "contract_number": contract_number,
            "contract_title": contract_title,
            "report_number": report_number,
            "po_wo_number": po_wo_number,
            "client_wps_number": client_wps_number,
            "project_title_wellID": project_title_wellID,
            "drawing_no": drawing_no,
            "line_no": line_no,
            "site_name": site_name,
            "job_desc": job_desc,
            "location": location,
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
        self.status_label.configure(text=t["success_add"])
        self.root.after(3000, lambda: self.status_label.configure(text=""))

    def clear_form(self):
        """Clear all form fields"""
        for field_id in [
            "job_id", "contract_number", "contract_title", "report_number", "po_wo_number",
            "client_wps_number", "project_title_wellID", "drawing_no", "line_no", "site_name",
            "job_desc", "location", "weld_id", "kp_sec", "wps_no", "material_gr", "heat_no",
            "size", "thk", "weld_side", "root", "material_comb", "pipe_no", "pipe_length",
            "welder_name", "material", "weld_type", "description", "date",
        ]:
            field = self.fields[field_id]
            entry = field["entry"]
            mic = field["mic"]
            if isinstance(entry, ctk.CTkTextbox):
                entry.delete("1.0", "end")
            else:
                entry.delete(0, "end")
            mic.configure(state="normal", fg_color="#1E90FF")
        # Reset date
        self.fields["date"]["entry"].insert(0, datetime.now().strftime("%Y-%m-%d"))

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
        """Export data into the given Excel template preserving all formatting and logo."""
        import shutil, tempfile, os
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image
        from openpyxl.styles import Alignment
        from tkinter import filedialog, messagebox

        if not self.records:
            messagebox.showwarning("Warning", "No records to export")
            return

        t = self.translations[self.current_lang]

        # filename = filedialog.asksaveasfilename(
        #     defaultextension=".xlsx",                 #for downloading window
        #     filetypes=[("Excel files", "*.xlsx")],
        #     initialfile=f"ATNM_Welding_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        # )

        filename = os.path.join(
            os.getcwd(), f"ATNM_Welding_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not filename:
            return

        try:
            template_path = "ATNM-ODC-MF-014-Daily Welding Production - Visual Inspection Report-Rev 01.xlsx"
            if not os.path.exists(template_path):
                messagebox.showerror("Error", f"Template not found:\n{template_path}")
                return

            shutil.copyfile(template_path, filename)

            # ✅ Try to extract and reinsert the logo safely
            temp_logo_path = None
            try:
                temp_wb = load_workbook(template_path)
                ws_t = temp_wb.active
                if ws_t._images:
                    logo_obj = ws_t._images[0]
                    temp_logo_path = tempfile.mktemp(suffix=".png")
                    logo_obj.image.save(temp_logo_path)
                temp_wb.close()
            except Exception as e:
                print(f"⚠️ Logo extraction failed: {e}")

            wb = load_workbook(filename)
            ws = wb.active

            # Helper to safely write while preserving labels
            def write_with_label(cell_ref, value):
                if not value:
                    return
                cell = ws[cell_ref]
                for merged in ws.merged_cells.ranges:
                    if cell.coordinate in merged:
                        cell = ws.cell(merged.min_row, merged.min_col)
                        break

                existing = str(cell.value).strip() if cell.value else ""
                if ":" in existing:
                    prefix = existing.split(":")[0].strip() + ":"
                    cell.value = f"{prefix} {value}"
                else:
                    cell.value = value
                cell.alignment = Alignment(vertical="center")

            # === HEADER MAPPING based on your layout ===
            header_map = {
                "contract_number": "E5",
                "contract_title": "K5",
                "report_number": "Q5",
                "activity_date": "U5",
                "po_wo_number": "E6",
                "client_wps_number": "K6",
                "project_title_wellID": "Q6",
                "drawing_no": "E7",
                "line_no": "Q7",
                "site_name": "U7",
                "job_desc": "E8",
                "location": "U8",
                "date": "U5"
            }

            header_source = self.records[0]
            for key, cell_ref in header_map.items():
                if header_source.get(key):
                    write_with_label(cell_ref, header_source[key])

            # === DATA TABLE SECTION ===
            start_row = 11
            row_gap = 2

            def write_safe(cell_ref, value):
                if not value or str(value).strip() == "":
                    return
                cell = ws[cell_ref]
                for merged in ws.merged_cells.ranges:
                    if cell.coordinate in merged:
                        cell = ws.cell(merged.min_row, merged.min_col)
                        break
                cell.value = value
                cell.alignment = Alignment(horizontal="center", vertical="center")

            for idx, record in enumerate(self.records, start=1):
                row = start_row + (idx - 1) * row_gap
                write_safe(f"B{row}", idx)
                write_safe(f"C{row}", record.get("kp_sec"))
                write_safe(f"D{row}", record.get("weld_id"))
                write_safe(f"E{row}", record.get("wps_no"))
                combo = f"{record.get('material_gr', '')} / {record.get('heat_no', '')}".strip(" / ")
                write_safe(f"F{row}", combo)
                write_safe(f"H{row}", record.get("size"))
                write_safe(f"I{row}", record.get("thk"))
                write_safe(f"J{row}", record.get("weld_side"))
                write_safe(f"K{row}", record.get("root"))
                write_safe(f"U{row}", record.get("material_comb"))
                write_safe(f"V{row}", record.get("pipe_no"))
                write_safe(f"W{row}", record.get("pipe_length"))
                remarks = f"Welder: {record.get('welder_name','')} | Date: {record.get('date','')}"
                write_safe(f"X{row}", remarks)

            # ✅ Reinsert logo if extracted
            if temp_logo_path and os.path.exists(temp_logo_path):
                try:
                    img = Image(temp_logo_path)
                    ws.add_image(img, "B1")
                    print("✅ Logo restored successfully.")
                except Exception as e:
                    print(f"⚠️ Failed to reinsert logo: {e}")

            wb.save(filename)
            wb.close()

            # self.status_label.config(text=t["success_export"], fg="green")
            # self.root.after(3000, lambda: self.status_label.config(text=""))
            messagebox.showinfo("Success", f"Form saved successfully:\n{filename}")

        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")

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
    root = ctk.CTk()  # Use CTk as root for full modern support
    app = WeldingShopApp(root)
    root.mainloop()

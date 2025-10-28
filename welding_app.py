import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime
import threading
import pyaudio
import vosk
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pyttsx3
from tkinter import messagebox

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
                "mic_tooltip": "Click to record voice"
            },
            "ar": {
                "title": "إدارة ورشة اللحام",
                "job_id": "رقم العمل:",
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
                "mic_tooltip": "انقر للتسجيل الصوتي"
            }
        }
        
        # Voice recognition
        self.is_recording = False
        self.recognizer = None
        self.audio = None
        self.stream = None
        self.load_vosk_model()
        
        # Build UI
        self.create_ui()
        self.update_language()
        
    def load_vosk_model(self):
        """Load Vosk model for speech recognition"""
        try:
            model_path = "models/vosk-model-en-us-0.22"
            if self.current_lang == "ar":
                model_path = "models/vosk-model-en-us-0.22"
            
            if os.path.exists(model_path):
                model = vosk.Model(model_path)
                self.recognizer = vosk.KaldiRecognizer(model, 16000)
                print(f"Loaded Vosk model: {model_path}")
            else:
                print(f"Warning: Vosk model not found at {model_path}")
                self.recognizer = None
        except Exception as e:
            print(f"Error loading Vosk model: {e}")
            self.recognizer = None

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
            fg="white"
        )
        self.title_label.pack(side=tk.LEFT, padx=20, pady=20)
        
        # Language selector
        lang_frame = tk.Frame(header_frame, bg="#2a5298")
        lang_frame.pack(side=tk.RIGHT, padx=20)
        
        self.lang_label = tk.Label(
            lang_frame,
            text="Language:",
            font=("Arial", 11),
            bg="#2a5298",
            fg="white"
        )
        self.lang_label.pack(side=tk.LEFT, padx=5)
        
        self.lang_var = tk.StringVar(value="English")
        lang_combo = ttk.Combobox(
            lang_frame,
            textvariable=self.lang_var,
            values=["English", "العربية"],
            state="readonly",
            width=15,
            font=("Arial", 10)
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
            pady=20
        )
        form_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        
        # Form fields
        self.fields = {}
        field_config = [
            ("job_id", "Job ID:"),
            ("welder_name", "Welder Name:"),
            ("material", "Material:"),
            ("weld_type", "Weld Type:"),
            ("description", "Description:"),
            ("date", "Date:")
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
                anchor="w"
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
                command=lambda f=field_id: self.record_voice(f)
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
            command=self.add_entry
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
            command=self.clear_form
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_label = tk.Label(
            form_frame,
            text="",
            font=("Arial", 10),
            bg="white",
            fg="green"
        )
        self.status_label.pack(pady=10)
        
        # Right panel - Records
        records_frame = tk.LabelFrame(
            main_container,
            text="Records",
            font=("Arial", 12, "bold"),
            bg="white",
            padx=10,
            pady=10
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
            command=self.export_excel
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
            height=15
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
        self.load_vosk_model()

    def update_language(self):
        """Update UI text based on selected language"""
        t = self.translations[self.current_lang]
        
        self.title_label.config(text=t["title"])
        self.lang_label.config(text=t["language"])
        
        # Update field labels
        self.fields["job_id"]["label"].config(text=t["job_id"])
        self.fields["welder_name"]["label"].config(text=t["welder_name"])
        self.fields["material"]["label"].config(text=t["material"])
        self.fields["weld_type"]["label"].config(text=t["weld_type"])
        self.fields["description"]["label"].config(text=t["description"])
        self.fields["date"]["label"].config(text=t["date"])
        
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

        if not self.recognizer:
            print("[DEBUG] No recognizer loaded")
            messagebox.showwarning("Warning", "Voice recognition model not loaded")
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
            voices = tts_engine.getProperty('voices')
            print(f"[DEBUG] Available voices: {len(voices)}")
            if len(voices) == 0:
                print("[DEBUG] WARNING: No TTS voices detected - install OS voices")
                messagebox.showwarning("TTS Warning", "No speech voices installed. Check OS settings.")
                # Still proceed to recording
            else:
                # Set a default voice (index 0 usually English; adjust for Arabic if needed)
                tts_engine.setProperty('voice', voices[0].id)
            tts_engine.setProperty('rate', 170)
            tts_engine.setProperty('volume', 1.0)
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
        threading.Thread(target=self._record_audio, args=(field_id,), daemon=True).start()

    def _record_audio(self, field_id):
        """Record and transcribe audio for ~3 seconds safely, then insert text."""
        print(f"[DEBUG] _record_audio started for field: {field_id}")
        print(f"[DEBUG] Setting is_recording to True")

        # Prevent re-entry
        if self.is_recording:
            print("[DEBUG] Re-entry detected, exiting")
            return

        self.is_recording = True
        text_output = ""

        audio = None
        stream = None
        try:
            print("[DEBUG] Initializing PyAudio")
            audio = pyaudio.PyAudio()

            # Open stream; handle device errors gracefully
            print("[DEBUG] Opening audio stream")
            stream = audio.open(
                format=pyaudio.paInt16,
                channels=1,
                rate=16000,
                input=True,
                frames_per_buffer=4000
            )

            # Start the stream and read for a fixed duration
            print("[DEBUG] Starting stream and recording loop (3s)")
            stream.start_stream()
            start = time.time()
            timeout_seconds = 3.0  # Fixed recording length; change if needed

            while time.time() - start < timeout_seconds:
                try:
                    data = stream.read(4000, exception_on_overflow=False)
                except Exception as e:
                    # Read error (overflow etc.) — skip and continue
                    print(f"[DEBUG] Audio read error: {e}")
                    continue

                if not data:
                    continue

                # Feed data to recognizer
                try:
                    if self.recognizer.AcceptWaveform(data):
                        res = safe_json_loads(self.recognizer.Result())
                        txt = res.get("text", "").strip()
                        if txt:
                            text_output += " " + txt
                            print(f"[DEBUG] Partial recognition: '{txt}'")
                    else:
                        # Partial result ignored for now
                        pass
                except Exception as e:
                    # If Vosk internal assertion happens, ignore and continue safely
                    print(f"[DEBUG] Vosk consume error: {e}")

            # Safe final result read
            print("[DEBUG] Getting final recognition result")
            try:
                final = safe_json_loads(self.recognizer.FinalResult())
                txt = final.get("text", "").strip()
                if txt:
                    text_output += " " + txt
                    print(f"[DEBUG] Final recognition: '{txt}'")
            except Exception as e:
                print(f"[DEBUG] Vosk final error: {e}")

        except Exception as e:
            print(f"[DEBUG] Recording Error: {e}")
        finally:
            # Cleanup audio resources safely
            print("[DEBUG] Cleaning up audio resources")
            try:
                if stream is not None:
                    stream.stop_stream()
                    stream.close()
            except Exception as e:
                print(f"[DEBUG] Stream cleanup error: {e}")
            try:
                if audio is not None:
                    audio.terminate()
            except Exception as e:
                print(f"[DEBUG] Audio cleanup error: {e}")

            # Reset status UI and recording flag on main thread
            try:
                t = self.translations[self.current_lang]
                self.root.after(0, lambda: self.status_label.config(text="", fg="green"))
                print("[DEBUG] Scheduled status reset")
            except Exception as e:
                print(f"[DEBUG] Status reset scheduling error: {e}")

            print("[DEBUG] Setting is_recording to False")
            self.is_recording = False

        text_output = (text_output or "").strip()
        print(f"[DEBUG] Final text_output: '{text_output}'")
        if text_output:
            # Insert text into UI on main thread
            try:
                self.root.after(0, lambda: self._insert_text(field_id, text_output))
                print("[DEBUG] Scheduled text insertion")
            except Exception as e:
                print(f"[DEBUG] Text insertion scheduling error: {e}")
                # Fallback directly if after fails
                self._insert_text(field_id, text_output)
            print(f"[DEBUG] Recognized for {field_id}: {text_output}")
        else:
            print("[DEBUG] No speech detected or recognition returned empty text")

    def _insert_text(self, field_id, text):
        """Insert transcribed text into field"""
        print(f"[DEBUG] _insert_text called for {field_id} with text: '{text}'")
        entry = self.fields[field_id]["entry"]
        
        if isinstance(entry, tk.Text):
            entry.insert(tk.END, text + " ")
            print("[DEBUG] Inserted into Text widget")
        else:
            current = entry.get()
            entry.delete(0, tk.END)
            entry.insert(0, current + " " + text if current else text)
            print("[DEBUG] Inserted into Entry widget")

    def add_entry(self):
        """Add new welding entry"""
        t = self.translations[self.current_lang]
        
        # Get values
        job_id = self.fields["job_id"]["entry"].get().strip()
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
        record = {
            "id": datetime.now().timestamp(),
            "job_id": job_id,
            "welder_name": welder_name,
            "material": material,
            "weld_type": weld_type,
            "description": description,
            "date": date
        }
        
        self.records.insert(0, record)
        self.save_data()
        self.refresh_table()
        self.clear_form()
        
        self.status_label.config(text=t["success_add"], fg="green")
        self.root.after(3000, lambda: self.status_label.config(text=""))

    def clear_form(self):
        """Clear all form fields"""
        for field_id in ["job_id", "welder_name", "material", "weld_type"]:
            self.fields[field_id]["entry"].delete(0, tk.END)
        
        self.fields["description"]["entry"].delete("1.0", tk.END)
        self.fields["date"]["entry"].delete(0, tk.END)
        self.fields["date"]["entry"].insert(0, datetime.now().strftime("%Y-%m-%d"))

    def refresh_table(self):
        """Refresh the records table"""
        # Clear existing
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add records
        for record in self.records:
            self.tree.insert("", tk.END, values=(
                record["job_id"],
                record["welder_name"],
                record["material"],
                record["weld_type"],
                record["date"],
                "Double-click to delete"
            ))

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
            initialfile=f"welding_records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not filename:
            return
        
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Welding Records"
            
            # Header style
            header_fill = PatternFill(start_color="2a5298", end_color="2a5298", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=12)
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Headers
            headers = ["Job ID", "Welder Name", "Material", "Weld Type", "Description", "Date"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Data
            for row, record in enumerate(self.records, 2):
                ws.cell(row=row, column=1, value=record["job_id"]).border = border
                ws.cell(row=row, column=2, value=record["welder_name"]).border = border
                ws.cell(row=row, column=3, value=record["material"]).border = border
                ws.cell(row=row, column=4, value=record["weld_type"]).border = border
                ws.cell(row=row, column=5, value=record["description"]).border = border
                ws.cell(row=row, column=6, value=record["date"]).border = border
            
            # Adjust column widths
            for col in range(1, 7):
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
                with open(self.data_file, 'r', encoding='utf-8') as f:
                    self.records = json.load(f)
            except:
                self.records = []

    def save_data(self):
        """Save records to JSON file"""
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving data: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WeldingShopApp(root)
    root.mainloop()
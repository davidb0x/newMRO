import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import sys
import os
import io
import json

# Konfiguracja ścieżek
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

class RedirectText(io.StringIO):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)
        
    def flush(self):
        pass

class MroPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MRO PDF - Ekstrakcja danych")
        self.root.geometry("600x550")
        self.root.resizable(False, False)
        
        # Zmienne
        self.input_dir = tk.StringVar()
        self.output_file = tk.StringVar()
        self.parser_type = tk.StringVar(value="Festo")
        
        # Słownik skryptów (dostawców) - można tu dodawać kolejne!
        self.parsers = {
            "Festo": {
                "folder": "MRO",
                "file": "Festo.py",
                "function": "process_all_pdfs_in_folder"
            },
            "CTS Technology": {
                "folder": "MRO",
                "file": "CTS Technology.py",
                "function": "process_all_pdfs_in_folder"
            }
        }
        
        self.load_settings()

        # Śledzenie zmian - auto-zapis po każdej zmianie wartości!
        self.input_dir.trace_add("write", self.save_settings)
        self.output_file.trace_add("write", self.save_settings)
        self.parser_type.trace_add("write", self.save_settings)
        
        self.create_widgets()

    def load_settings(self):
        """Wczytuje ostatnio użyte ścieżki i skrypt z pliku config.json"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.input_dir.set(data.get("input_dir", ""))
                    self.output_file.set(data.get("output_file", ""))
                    if data.get("parser_type") in self.parsers:
                        self.parser_type.set(data.get("parser_type"))
            except Exception:
                pass

    def save_settings(self, *args):
        """Zapisuje ścieżki i wybrany skrypt na bieżąco do pliku config.json"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    "input_dir": self.input_dir.get(),
                    "output_file": self.output_file.get(),
                    "parser_type": self.parser_type.get()
                }, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- Dostawca ---
        ttk.Label(main_frame, text="1. Wybierz dostawcę (skrypt):", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        parser_combo = ttk.Combobox(main_frame, textvariable=self.parser_type, values=list(self.parsers.keys()), state="readonly", width=30)
        parser_combo.pack(anchor=tk.W, pady=(0, 15))
        
        # --- Folder Wejściowy ---
        ttk.Label(main_frame, text="2. Wybierz folder z PDF-ami:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        in_frame = ttk.Frame(main_frame)
        in_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Entry(in_frame, textvariable=self.input_dir, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(in_frame, text="Wybierz...", command=self.browse_input).pack(side=tk.RIGHT)
        
        # --- Plik Wyjściowy ---
        ttk.Label(main_frame, text="3. Wybierz miejsce zapisu np. Excel (.xlsx):", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        out_frame = ttk.Frame(main_frame)
        out_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Entry(out_frame, textvariable=self.output_file, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(out_frame, text="Zapisz jako...", command=self.browse_output).pack(side=tk.RIGHT)
        
        # --- Przycisk START ---
        self.start_btn = ttk.Button(main_frame, text="Uruchom ekstrakcję", command=self.start_processing)
        self.start_btn.pack(fill=tk.X, pady=(10, 15))
        
        # --- Konsola (Logi) ---
        ttk.Label(main_frame, text="Logi z działania:", font=("Arial", 9)).pack(anchor=tk.W)
        self.log_text = tk.Text(main_frame, height=12, bg="#f4f4f4", wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def browse_input(self):
        folder = filedialog.askdirectory(title="Wybierz folder z ofertami PDF")
        if folder:
            self.input_dir.set(folder)

    def browse_output(self):
        file = filedialog.asksaveasfilename(
            title="Gdzie zapisać plik wynikowy?",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.output_file.set(file)

    def start_processing(self):
        if not self.input_dir.get():
            messagebox.showwarning("Braki", "Wybierz folder z plikami PDF!")
            return
        if not self.output_file.get():
            messagebox.showwarning("Braki", "Wybierz miejsce zapisu pliku wynikowego!")
            return
            
        self.start_btn.config(state=tk.DISABLED)
        self.log_text.delete(1.0, tk.END)
        print("Trwa uruchamianie procedury...\n")
        
        # Przekierowujemy wyjścia
        self.old_stdout = sys.stdout
        sys.stdout = RedirectText(self.log_text)
        
        # Uruchamiamy w nowym wątku
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            vendor = self.parser_type.get()
            config = self.parsers[vendor]
            
            # Formowanie ścieżki do odpowiedniego modułu
            module_dir = os.path.join(BASE_DIR, config['folder'])
            file_name = config['file']
            func_name = config['function']
            
            if module_dir not in sys.path:
                sys.path.insert(0, module_dir)
                
            module_name = file_name.replace('.py', '')
            
            print(f"--- Uruchomiono skrypt dla dostawcy: {vendor} ---")
            print(f"-> Z katalogu PDF: {self.input_dir.get()}")
            print(f"-> Do pliku: {self.output_file.get()}\n")
            
            # Dynamiczny import
            import importlib
            module = importlib.import_module(module_name)
            # Przeładowanie (jeśli modyfikowaliśmy niedawno skrypt)
            importlib.reload(module)
            
            # Pobranie odpowiedniej funkcji
            process_func = getattr(module, func_name)
            
            # Wywołanie funkcji z zachowaniem logiki poszczególnych skryptów
            process_func(self.input_dir.get(), self.output_file.get())
            
            print("\n--- Zakończono ---")
            messagebox.showinfo("Sukces", "Zakończono generowanie pliku Excel!")
            
        except Exception as e:
            print(f"\n❌ Wystąpił błąd krytyczny: {e}")
            messagebox.showerror("Błąd", f"Wystąpił błąd:\n{e}")
            
        finally:
            sys.stdout = self.old_stdout
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))

if __name__ == '__main__':
    root = tk.Tk()
    app = MroPdfApp(root)
    root.mainloop()
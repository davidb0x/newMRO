import contextlib
import io
import json
import os
import sys

import flet as ft

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")

PARSERS = {
    "Festo": {
        "folder": "MRO",
        "file": "Festo.py",
        "function": "process_all_pdfs_in_folder",
    },
    "CTS Technology": {
        "folder": "MRO",
        "file": "CTS Technology.py",
        "function": "process_all_pdfs_in_folder",
    },
}


def load_settings():
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass


class LogWriter(io.TextIOBase):
    def __init__(self, append_log):
        super().__init__()
        self.append_log = append_log

    def write(self, string):
        if not string:
            return 0
        self.append_log(string)
        return len(string)

    def flush(self):
        pass


def main(page: ft.Page):
    page.title = "MRO PDF - Ekstrakcja danych"
    page.window.width = 720
    page.window.height = 620
    page.window.resizable = False
    page.padding = 20

    settings = load_settings()
    parser_default = settings.get("parser_type")
    if parser_default not in PARSERS:
        parser_default = "Festo"

    input_dir_field = ft.TextField(
        label="Folder z PDF-ami",
        value=settings.get("input_dir", ""),
        read_only=True,
        expand=True,
    )
    output_file_field = ft.TextField(
        label="Plik wynikowy (.xlsx)",
        value=settings.get("output_file", ""),
        read_only=True,
        expand=True,
    )
    parser_dropdown = ft.Dropdown(
        label="Dostawca (skrypt)",
        value=parser_default,
        options=[ft.dropdown.Option(name) for name in PARSERS.keys()],
        width=320,
    )
    log_field = ft.TextField(
        label="Logi z działania",
        multiline=True,
        read_only=True,
        expand=True,
        min_lines=12,
        max_lines=12,
    )
    start_btn = ft.ElevatedButton("Uruchom ekstrakcję")

    def write_settings():
        save_settings(
            {
                "input_dir": input_dir_field.value,
                "output_file": output_file_field.value,
                "parser_type": parser_dropdown.value,
            }
        )

    def append_log(text):
        log_field.value = (log_field.value or "") + text
        log_field.update()

    def set_start_enabled(enabled):
        start_btn.disabled = not enabled
        start_btn.update()

    def show_dialog(title, message):
        dialog = ft.AlertDialog(
            title=ft.Text(title),
            content=ft.Text(message),
            actions=[
                ft.TextButton(
                    "OK",
                    on_click=lambda e: close_dialog(),
                )
            ],
        )
        page.show_dialog(dialog)

    def close_dialog():
        page.pop_dialog()

    def run_processing():
        try:
            vendor = parser_dropdown.value or parser_default
            config = PARSERS[vendor]

            module_dir = os.path.join(BASE_DIR, config["folder"])
            if module_dir not in sys.path:
                sys.path.insert(0, module_dir)

            module_name = os.path.splitext(config["file"])[0]

            append_log(f"--- Uruchomiono skrypt dla dostawcy: {vendor} ---\n")
            append_log(f"-> Z katalogu PDF: {input_dir_field.value}\n")
            append_log(f"-> Do pliku: {output_file_field.value}\n\n")

            import importlib

            module = importlib.import_module(module_name)
            importlib.reload(module)
            process_func = getattr(module, config["function"])

            writer = LogWriter(append_log)
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                process_func(input_dir_field.value, output_file_field.value)

            append_log("\n--- Zakończono ---\n")
            show_dialog("Sukces", "Zakończono generowanie pliku Excel!")
        except Exception as e:
            append_log(f"\n❌ Wystąpił błąd krytyczny: {e}\n")
            show_dialog("Błąd", f"Wystąpił błąd:\n{e}")
        finally:
            set_start_enabled(True)

    def start_processing(e):
        if not input_dir_field.value:
            show_dialog("Braki", "Wybierz folder z plikami PDF!")
            return
        if not output_file_field.value:
            show_dialog("Braki", "Wybierz miejsce zapisu pliku wynikowego!")
            return

        set_start_enabled(False)
        log_field.value = ""
        log_field.update()
        append_log("Trwa uruchamianie procedury...\n\n")
        page.run_thread(run_processing)

    async def pick_input_dir():
        path = await file_picker.get_directory_path(
            dialog_title="Wybierz folder z ofertami PDF"
        )
        if path:
            input_dir_field.value = path
            input_dir_field.update()
            write_settings()

    async def pick_output_file():
        path = await file_picker.save_file(
            dialog_title="Gdzie zapisać plik wynikowy?",
            file_name="wynik.xlsx",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path = f"{path}.xlsx"
            output_file_field.value = path
            output_file_field.update()
            write_settings()

    def on_parser_change(e):
        write_settings()

    file_picker = ft.FilePicker()

    parser_dropdown.on_select = on_parser_change
    start_btn.on_click = start_processing

    page.add(
        ft.Column(
            [
                parser_dropdown,
                ft.Row(
                    [
                        input_dir_field,
                        ft.ElevatedButton(
                            "Wybierz...",
                            on_click=lambda e: page.run_task(pick_input_dir),
                        ),
                    ],
                ),
                ft.Row(
                    [
                        output_file_field,
                        ft.ElevatedButton(
                            "Zapisz jako...",
                            on_click=lambda e: page.run_task(pick_output_file),
                        ),
                    ],
                ),
                start_btn,
                log_field,
            ],
            spacing=12,
            expand=True,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)
import os
import json
import sys
import re
from PySide6.QtCore import QThread, Signal
import xlwings as xw

class PresetManager:
    @staticmethod
    def save_preset(filepath, rules):
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump({"rules": rules}, f, indent=4)

    @staticmethod
    def load_preset(filepath):
        if not os.path.exists(filepath):
            return []
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data.get("rules", [])

class ExcelMappingWorker(QThread):
    progress = Signal(int)
    finished = Signal(str, int)
    error = Signal(str)

    def __init__(self, rules, auto_append=False):
        super().__init__()
        self.rules = rules
        self.auto_append = auto_append

    def run(self):
        if not self.rules:
            self.error.emit("Tidak ada rule untuk dieksekusi.")
            return

        try:
            src_file = self.rules[0]["src_file"]
            dest_file = self.rules[0]["dest_file"]
            src_name = os.path.basename(src_file)
            dest_name = os.path.basename(dest_file)

            if xw.apps.keys(): 
                app = xw.apps.active
            else: 
                app = xw.App(visible=True, add_book=False)
            
            app.visible = True 

            try: 
                wb_dest = app.books[dest_name]
            except KeyError: 
                wb_dest = app.books.open(dest_file)

            try: 
                wb_src = app.books[src_name]
            except KeyError: 
                wb_src = app.books.open(src_file, read_only=True)

            dest_sheets = {}
            for rule in self.rules:
                sheet_name = rule["dest_sheet"]
                if sheet_name not in dest_sheets:
                    dest_sheets[sheet_name] = []
                dest_sheets[sheet_name].append(rule)

            total = len(self.rules)
            processed = 0

            for sheet_name, sheet_rules in dest_sheets.items():
                ws_dest = wb_dest.sheets[sheet_name]
                
                next_row = None
                if self.auto_append:
                    # Cari baris paling bawah yang kosong (Radar)
                    ref_col = re.sub(r'[0-9]', '', sheet_rules[0]["dest_cell"])
                    last_row = ws_dest.range(f'{ref_col}1048576').end('up').row
                    next_row = last_row + 1

                for rule in sheet_rules:
                    val = wb_src.sheets[rule["src_sheet"]].range(rule["src_cell"]).value
                    
                    if self.auto_append:
                        col_letter = re.sub(r'[0-9]', '', rule["dest_cell"])
                        target_cell = f"{col_letter}{next_row}"
                    else:
                        target_cell = rule["dest_cell"]
                    
                    ws_dest.range(target_cell).value = val
                    processed += 1
                    self.progress.emit(int((processed / total) * 100))

            # Simpan langsung ke file aslinya (OneDrive)
            wb_dest.save()
            app.activate()
            wb_dest.activate()

            mode_teks = "Baris Baru" if self.auto_append else "Sel Spesifik"
            self.finished.emit(f"Data berhasil disuntikkan ke {mode_teks} dan di-Save ke OneDrive!", total)

        except Exception as e:
            self.error.emit(f"Terjadi kesalahan:\n{str(e)}")
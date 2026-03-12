import os
import json
import sys
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

    def __init__(self, rules):
        super().__init__()
        self.rules = rules

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

            total = len(self.rules)
            for i, rule in enumerate(self.rules):
                val = wb_src.sheets[rule["src_sheet"]].range(rule["src_cell"]).value
                wb_dest.sheets[rule["dest_sheet"]].range(rule["dest_cell"]).value = val
                self.progress.emit(int(((i + 1) / total) * 100))

            app.activate()
            wb_dest.activate()

            self.finished.emit("Silakan periksa Excel Anda. Jika sudah benar, tekan Save manual di Excel.", total)

        except Exception as e:
            self.error.emit(str(e))
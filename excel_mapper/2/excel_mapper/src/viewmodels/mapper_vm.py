from PySide6.QtCore import QObject, Signal
from models.excel_worker import ExcelMappingWorker, PresetManager

class MapperViewModel(QObject):
    rules_updated = Signal(list)
    mapping_started = Signal()
    mapping_progress = Signal(int)
    mapping_finished = Signal(str, int)
    mapping_error = Signal(str)

    def __init__(self):
        super().__init__()
        self.rules = []
        self.worker = None
        self.undo_stack = []

    def add_rule(self, src_file, src_sheet, src_cell, src_val, dest_file, dest_sheet, dest_cell):
        if not src_file or not dest_file: return
        rule = {
            "src_file": src_file, "src_sheet": src_sheet, "src_cell": src_cell, "src_val": src_val,
            "dest_file": dest_file, "dest_sheet": dest_sheet, "dest_cell": dest_cell
        }
        self.rules.append(rule)
        self.undo_stack.append(rule)
        self.rules_updated.emit(self.rules)

    def remove_rule(self, index):
        if 0 <= index < len(self.rules):
            removed = self.rules.pop(index)
            if removed in self.undo_stack:
                self.undo_stack.remove(removed)
            self.rules_updated.emit(self.rules)

    def run_mapping(self, auto_append=False):
        if not self.rules: 
            self.mapping_error.emit("Tidak ada rule untuk dieksekusi.")
            return
            
        self.mapping_started.emit()
        self.worker = ExcelMappingWorker(self.rules, auto_append)
        self.worker.progress.connect(self.mapping_progress.emit)
        self.worker.finished.connect(self.mapping_finished.emit)
        self.worker.error.connect(self.mapping_error.emit)
        self.worker.start()

    def save_preset(self, filepath):
        # Memastikan hanya menyimpan daftar rules ke PresetManager
        PresetManager.save_preset(filepath, self.rules)

    def load_preset(self, filepath):
        """Memuat preset dan membersihkan path agar dinamis mengikuti input user di UI"""
        try:
            # Ambil data dari PresetManager
            data = PresetManager.load_preset(filepath)
            
            # Logika penanganan jika data berupa LIST atau DICT (agar tidak AttributeError)
            if isinstance(data, list):
                loaded_rules = data
            elif isinstance(data, dict):
                loaded_rules = data.get("rules", [])
            else:
                loaded_rules = []

            # Proses Pembersihan Path:
            # Kita hapus path file lama agar saat tombol 'Run' ditekan di main.py,
            # aplikasi akan menggunakan file yang sedang dipilih user (Browse).
            for rule in loaded_rules:
                rule["src_file"] = ""  
                rule["dest_file"] = "" 
            
            self.rules = loaded_rules
            self.undo_stack.clear() # Reset undo stack karena ini data baru
            self.rules_updated.emit(self.rules)
            
        except Exception as e:
            self.mapping_error.emit(f"Gagal memuat preset: {str(e)}")

    def undo_last_rule(self):
        if self.undo_stack and self.rules:
            last_rule = self.undo_stack.pop()
            if last_rule in self.rules:
                self.rules.remove(last_rule)
                self.rules_updated.emit(self.rules)
            return True
        return False
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
        self.rules_updated.emit(self.rules)
        self.undo_stack.append(rule) # Catat untuk Undo

    def remove_rule(self, index):
        if 0 <= index < len(self.rules):
            self.rules.pop(index)
            self.rules_updated.emit(self.rules)

    def run_mapping(self):
        if not self.rules: return
        self.mapping_started.emit()
        self.worker = ExcelMappingWorker(self.rules)
        self.worker.progress.connect(self.mapping_progress.emit)
        self.worker.finished.connect(self.mapping_finished.emit)
        self.worker.error.connect(self.mapping_error.emit)
        self.worker.start()

    def save_preset(self, filepath):
        PresetManager.save_preset(filepath, self.rules)

    def load_preset(self, filepath):
        self.rules = PresetManager.load_preset(filepath)
        self.rules_updated.emit(self.rules)
        
    def undo_last_rule(self):
        if self.undo_stack and self.rules:
            last_rule = self.undo_stack.pop()
            if last_rule in self.rules:
                self.rules.remove(last_rule)
                self.rules_updated.emit(self.rules)
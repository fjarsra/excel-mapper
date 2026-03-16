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
        self.undo_stack = []
        self.worker = None

    def add_rule(self, src_file, src_sheet, src_cell, src_val, dest_file, dest_sheet, dest_cell):
        # Memastikan tidak ada data yang kosong sebelum ditambahkan ke tabel
        if not all([src_file, src_sheet, src_cell, dest_file, dest_sheet, dest_cell]):
            return
            
        rule = {
            "src_file": src_file, "src_sheet": src_sheet, "src_cell": src_cell, 
            "src_val": src_val, "dest_file": dest_file, "dest_sheet": dest_sheet, 
            "dest_cell": dest_cell
        }
        
        # Mencegah duplikasi: Jika user menarik sel yang sama ke tujuan yang sama, abaikan
        if any(r['src_cell'] == src_cell and r['dest_cell'] == dest_cell for r in self.rules):
            return

        self.rules.append(rule)
        self.undo_stack.append(("add", rule))
        self.rules_updated.emit(self.rules)

    def remove_rule(self, index):
        if 0 <= index < len(self.rules):
            removed_rule = self.rules.pop(index)
            self.undo_stack.append(("remove", removed_rule, index))
            self.rules_updated.emit(self.rules)

    def undo_last_rule(self):
        if not self.undo_stack: return
        
        action = self.undo_stack.pop()
        if action[0] == "add":
            rule_to_remove = action[1]
            if rule_to_remove in self.rules:
                self.rules.remove(rule_to_remove)
        elif action[0] == "remove":
            _, rule, index = action
            self.rules.insert(index, rule)
            
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
import sys, os, openpyxl
from openpyxl.utils import coordinate_to_tuple
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QHeaderView, QPushButton, QLabel, QSplitter, QTableWidget, 
    QTableWidgetItem, QMessageBox, QFileDialog, QLineEdit, QComboBox, QFrame, QProgressBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

from viewmodels.mapper_vm import MapperViewModel
from models.excel_handler import ExcelTableModel
from views.components.excel_grid import DraggableTableView, DroppableTableView

class MainWindow(QMainWindow):
    def __init__(self, view_model: MapperViewModel):
        super().__init__()
        self.vm = view_model
        self.setWindowTitle("NexusXL Mapper - V1.1 (Enterprise Edition)")
        self.resize(1300, 850)
        self.setup_ui()
        self.setup_connections()

    def setup_ui(self):
        # StyleSheet Global: Memaksa Light Mode dan Teks Hitam
        self.setStyleSheet("""
            QMainWindow, QWidget#central { background-color: #f1f5f9; }
            QFrame#card { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; }
            QTableView { 
                background-color: white; 
                gridline-color: #cbd5e1; 
                color: black; 
                border: none;
                selection-background-color: #bfdbfe;
            }
            QHeaderView::section { 
                background-color: #f8fafc; color: #475569; border: 1px solid #e2e8f0; font-weight: normal; 
            }
            QPushButton.primary { background-color: #2563eb; color: white; border-radius: 6px; padding: 8px 16px; font-weight: bold; }
            QPushButton.secondary { background-color: #ffffff; color: #475569; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 12px; font-weight: bold; }
            
            /* Style khusus untuk tombol zoom agar simbol + dan - terlihat jelas */
            QPushButton.zoom-btn { 
                background-color: #ffffff; 
                color: #1e293b; 
                border: 1px solid #cbd5e1; 
                border-radius: 4px; 
                padding: 2px; 
                font-weight: bold; 
                font-size: 16px;
                min-width: 35px; 
                max-width: 35px;
            }
            QPushButton.zoom-btn:hover { background-color: #f1f5f9; }
            
            QLineEdit, QComboBox { padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: white; color: black; }
            QLabel { color: #1e293b; }
            QLabel.title { font-size: 18px; font-weight: bold; color: #1e293b; }
        """)
        
        central_widget = QWidget()
        central_widget.setObjectName("central")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title_label = QLabel("🔗 NexusXL Mapper Pro")
        title_label.setProperty("class", "title")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        self.btn_load_preset = QPushButton("📂 Load Preset")
        self.btn_load_preset.setProperty("class", "secondary")
        self.btn_save_preset = QPushButton("💾 Save Preset")
        self.btn_save_preset.setProperty("class", "secondary")
        self.btn_run = QPushButton("▶ Run Mapping")
        self.btn_run.setProperty("class", "primary")
        for btn in [self.btn_load_preset, self.btn_save_preset, self.btn_run]: header_layout.addWidget(btn)
        main_layout.addLayout(header_layout)

        # Config Panel
        config_frame = QFrame(objectName="card")
        config_layout = QHBoxLayout(config_frame)
        self.src_file_input = QLineEdit(readOnly=True)
        self.src_sheet_combo = QComboBox()
        self.dest_file_input = QLineEdit(readOnly=True)
        self.dest_sheet_combo = QComboBox()
        self.btn_src_browse = QPushButton("...")
        self.btn_src_browse.setProperty("class", "secondary")
        self.src_sheet_combo = QComboBox()
        src_row.addWidget(self.src_file_input, stretch=3)
        src_row.addWidget(self.btn_src_browse)
        src_row.addWidget(self.src_sheet_combo, stretch=1)
        src_layout.addLayout(src_row)
        
        arrow_label = QLabel(" ➔ ")
        arrow_label.setStyleSheet("font-size: 40px; color: #94a3b8; font-weight: bold;")

        dest_layout = QVBoxLayout()
        lbl_dest = QLabel("DESTINATION DATA")
        lbl_dest.setProperty("class", "subtitle")
        dest_layout.addWidget(lbl_dest)
        dest_row = QHBoxLayout()
        self.dest_file_input = QLineEdit()
        self.dest_file_input.setPlaceholderText("C:\\...\\Template.xlsx")
        self.dest_file_input.setReadOnly(True)
        self.btn_dest_browse = QPushButton("...")
        self.btn_dest_browse.setProperty("class", "secondary")
        
        for label, line, btn, combo in [("SOURCE DATA", self.src_file_input, self.btn_src_browse, self.src_sheet_combo), 
                                      ("DESTINATION DATA", self.dest_file_input, self.btn_dest_browse, self.dest_sheet_combo)]:
            layout = QVBoxLayout()
            layout.addWidget(QLabel(label))
            row = QHBoxLayout()
            row.addWidget(line, 3); row.addWidget(btn); row.addWidget(combo, 1)
            layout.addLayout(row)
            config_layout.addLayout(layout, 1)
        main_layout.addWidget(config_frame)

        # Workspace with Zoom & Toggle Controls
        workspace_frame = QFrame(objectName="card")
        workspace_layout = QHBoxLayout(workspace_frame)
        workspace_layout.setContentsMargins(5, 5, 5, 5)
        splitter = QSplitter(Qt.Horizontal)

        # Source UI
        src_cont = QWidget()
        src_vbox = QVBoxLayout(src_cont)
        src_header = QHBoxLayout()
        src_header.addWidget(QLabel("🔵 Source Sheet"))
        src_header.addStretch()
        self.btn_src_toggle_view = QPushButton("👁️ Toggle View")
        self.btn_src_toggle_view.setProperty("class", "secondary")
        self.btn_src_zoom_out = QPushButton("-"); self.btn_src_zoom_out.setProperty("class", "zoom-btn")
        self.btn_src_zoom_in = QPushButton("+"); self.btn_src_zoom_in.setProperty("class", "zoom-btn")
        src_header.addWidget(self.btn_src_toggle_view)
        src_header.addWidget(self.btn_src_zoom_out); src_header.addWidget(self.btn_src_zoom_in)
        src_vbox.addLayout(src_header)
        self.source_view = DraggableTableView()
        src_vbox.addWidget(self.source_view)

        # Destination UI
        dest_cont = QWidget()
        dest_vbox = QVBoxLayout(dest_cont)
        dest_header = QHBoxLayout()
        dest_header.addWidget(QLabel("🟢 Destination Sheet"))
        dest_header.addStretch()
        self.btn_dest_toggle_view = QPushButton("👁️ Toggle View")
        self.btn_dest_toggle_view.setProperty("class", "secondary")
        self.btn_dest_zoom_out = QPushButton("-"); self.btn_dest_zoom_out.setProperty("class", "zoom-btn")
        self.btn_dest_zoom_in = QPushButton("+"); self.btn_dest_zoom_in.setProperty("class", "zoom-btn")
        dest_header.addWidget(self.btn_dest_toggle_view)
        dest_header.addWidget(self.btn_dest_zoom_out); dest_header.addWidget(self.btn_dest_zoom_in)
        dest_vbox.addLayout(dest_header)
        self.dest_view = DroppableTableView(self.handle_drop)
        dest_vbox.addWidget(self.dest_view)

        splitter.addWidget(src_cont); splitter.addWidget(dest_cont)
        workspace_layout.addWidget(splitter)
        main_layout.addWidget(workspace_frame, stretch=2)

        # Rules
        self.rules_table = QTableWidget(0, 4)
        self.rules_table.setHorizontalHeaderLabels(["Source", "Destination", "Type", "Action"])
        self.rules_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        main_layout.addWidget(self.rules_table, stretch=1)

    def setup_connections(self):
        self.btn_src_browse.clicked.connect(lambda: self.browse_file(self.src_file_input, self.src_sheet_combo))
        self.btn_dest_browse.clicked.connect(lambda: self.browse_file(self.dest_file_input, self.dest_sheet_combo))
        self.src_sheet_combo.currentTextChanged.connect(lambda s: self.load_sheet(self.src_file_input.text(), s, self.source_view))
        self.dest_sheet_combo.currentTextChanged.connect(lambda s: self.load_sheet(self.dest_file_input.text(), s, self.dest_view))
        
        # Koneksi Zoom & Toggle View
        self.btn_src_zoom_in.clicked.connect(lambda: self.source_view.apply_zoom(1))
        self.btn_src_zoom_out.clicked.connect(lambda: self.source_view.apply_zoom(-1))
        self.btn_src_toggle_view.clicked.connect(self.source_view.toggle_hidden_columns)
        
        self.btn_dest_zoom_in.clicked.connect(lambda: self.dest_view.apply_zoom(1))
        self.btn_dest_zoom_out.clicked.connect(lambda: self.dest_view.apply_zoom(-1))
        self.btn_dest_toggle_view.clicked.connect(self.dest_view.toggle_hidden_columns)
        
        self.btn_run.clicked.connect(self.vm.run_mapping)
        self.btn_save_preset.clicked.connect(self.save_preset_dialog)
        self.btn_load_preset.clicked.connect(self.load_preset_dialog)
        self.vm.rules_updated.connect(self.refresh_rules_ui)

    def browse_file(self, line_edit, combo_box):
        path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xlsx *.xlsm)")
        if path:
            line_edit.setText(os.path.normpath(path))
            wb = openpyxl.load_workbook(path, read_only=True)
            combo_box.clear(); combo_box.addItems(wb.sheetnames); wb.close()

    def load_sheet(self, filepath, sheetname, view):
        if filepath and sheetname:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                model = ExcelTableModel(filepath, sheetname)
                view.setModel(model)
                view.sync_with_excel() 
                self.refresh_highlights()
            finally:
                QApplication.restoreOverrideCursor()

    def handle_drop(self, source_cell, dest_cell):
        src_model = self.source_view.model()
        src_val = "N/A"
        if src_model:
            r, c = coordinate_to_tuple(source_cell)
            src_val = src_model.grid_data[r-1][c-1]["value"]
        self.vm.add_rule(self.src_file_input.text(), self.src_sheet_combo.currentText(), source_cell, src_val,
                         self.dest_file_input.text(), self.dest_sheet_combo.currentText(), dest_cell)

    def refresh_rules_ui(self, rules):
        self.rules_table.setRowCount(0)
        for i, rule in enumerate(rules):
            row = self.rules_table.rowCount()
            self.rules_table.insertRow(row)
            self.rules_table.setItem(row, 0, QTableWidgetItem(f"{rule['src_sheet']}!{rule['src_cell']}"))
            self.rules_table.setItem(row, 1, QTableWidgetItem(f"{rule['dest_sheet']}!{rule['dest_cell']}"))
            self.rules_table.setItem(row, 2, QTableWidgetItem("Live Write"))
            btn_delete = QPushButton("🗑️")
            btn_delete.clicked.connect(lambda _, idx=i: self.vm.remove_rule(idx))
            self.rules_table.setCellWidget(row, 3, btn_delete)
        self.refresh_highlights()

    def refresh_highlights(self):
        curr_src, curr_dest = self.src_sheet_combo.currentText(), self.dest_sheet_combo.currentText()
        src_model, dest_model = self.source_view.model(), self.dest_view.model()
        if src_model: src_model.clear_highlights()
        if dest_model: dest_model.clear_highlights()
        for rule in self.vm.rules:
            if src_model and rule["src_sheet"] == curr_src: src_model.add_highlight(rule["src_cell"], QColor("#dbeafe"))
            if dest_model and rule["dest_sheet"] == curr_dest: dest_model.add_highlight(rule["dest_cell"], QColor("#dcfce3"))

    def save_preset_dialog(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Preset", "", "JSON Files (*.json)")
        if path: self.vm.save_preset(path)

    def load_preset_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Preset", "", "JSON Files (*.json)")
        if path: self.vm.load_preset(path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    vm = MapperViewModel()
    window = MainWindow(vm)
    window.show()
    sys.exit(app.exec())
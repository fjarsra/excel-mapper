import sys
import os
import openpyxl
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
        self.setStyleSheet("""
            QMainWindow { background-color: #f1f5f9; } 
            QFrame#card { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; }
            QTableView { 
                background-color: white; color: #334155; border: 1px solid #e2e8f0; 
                border-radius: 4px; gridline-color: #cbd5e1; selection-background-color: transparent; 
            }
            QHeaderView::section { background-color: #f8fafc; padding: 4px; border: 1px solid #e2e8f0; font-weight: bold; color: #64748b; }
            QPushButton.primary { background-color: #2563eb; color: white; border-radius: 6px; padding: 8px 16px; font-weight: bold; }
            QPushButton.primary:hover { background-color: #1d4ed8; }
            QPushButton.secondary { background-color: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 12px; font-weight: bold; }
            QPushButton.secondary:hover { background-color: #e2e8f0; }
            QPushButton.icon-btn { background-color: #ffffff; color: #475569; border: 1px solid #cbd5e1; border-radius: 4px; padding: 4px 8px; font-weight: bold; }
            QLabel { color: #334155; }
            QLabel.title { font-size: 18px; font-weight: bold; color: #1e293b; }
            QLabel.subtitle { font-size: 12px; font-weight: bold; color: #64748b; text-transform: uppercase; }
            QLineEdit, QComboBox { padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; background: #f8fafc; color: #334155; }
        """)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title_label = QLabel("🔗 NexusXL Mapper Pro")
        title_label.setProperty("class", "title")
        header_layout.addWidget(title_label)
        
        self.btn_load_preset = QPushButton("📂 Load Preset")
        self.btn_load_preset.setProperty("class", "secondary")
        self.btn_save_preset = QPushButton("💾 Save Preset")
        self.btn_save_preset.setProperty("class", "secondary")
        self.btn_run = QPushButton("▶ Run Mapping")
        self.btn_run.setProperty("class", "primary")
        
        header_layout.addStretch()
        header_layout.addWidget(self.btn_load_preset)
        header_layout.addWidget(self.btn_save_preset)
        header_layout.addWidget(self.btn_run)
        main_layout.addLayout(header_layout)

        # Config Panel (Source & Dest)
        config_frame = QFrame()
        config_frame.setObjectName("card")
        config_layout = QHBoxLayout(config_frame)
        config_layout.setContentsMargins(15, 15, 15, 15)
        
        src_layout = QVBoxLayout()
        lbl_src = QLabel("SOURCE DATA")
        lbl_src.setProperty("class", "subtitle")
        src_layout.addWidget(lbl_src)
        src_row = QHBoxLayout()
        self.src_file_input = QLineEdit()
        self.src_file_input.setPlaceholderText("C:\\...\\Source.xlsx")
        self.src_file_input.setReadOnly(True)
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
        self.dest_sheet_combo = QComboBox()
        dest_row.addWidget(self.dest_file_input, stretch=3)
        dest_row.addWidget(self.btn_dest_browse)
        dest_row.addWidget(self.dest_sheet_combo, stretch=1)
        dest_layout.addLayout(dest_row)

        config_layout.addLayout(src_layout, stretch=1)
        config_layout.addWidget(arrow_label)
        config_layout.addLayout(dest_layout, stretch=1)
        main_layout.addWidget(config_frame)

        # Workspace Grids
        workspace_frame = QFrame()
        workspace_frame.setObjectName("card")
        workspace_layout = QHBoxLayout(workspace_frame)
        workspace_layout.setContentsMargins(0, 0, 0, 0)
        
        splitter = QSplitter(Qt.Horizontal)
        
        src_grid_widget = QWidget()
        src_grid_layout = QVBoxLayout(src_grid_widget)
        src_grid_layout.addWidget(QLabel("🔵 Source Sheet"))
        self.source_view = DraggableTableView()
        src_grid_layout.addWidget(self.source_view)
        
        dest_grid_widget = QWidget()
        dest_grid_layout = QVBoxLayout(dest_grid_widget)
        dest_grid_layout.addWidget(QLabel("🟢 Destination Sheet"))
        self.dest_view = DroppableTableView(self.handle_drop)
        dest_grid_layout.addWidget(self.dest_view)

        splitter.addWidget(src_grid_widget)
        splitter.addWidget(dest_grid_widget)
        workspace_layout.addWidget(splitter)
        main_layout.addWidget(workspace_frame, stretch=2)

        # Rules Table & Progress
        rules_frame = QFrame()
        rules_frame.setObjectName("card")
        rules_layout = QVBoxLayout(rules_frame)
        lbl_rules = QLabel("Mapping Rules")
        lbl_rules.setProperty("class", "title")
        rules_layout.addWidget(lbl_rules)

        self.rules_table = QTableWidget(0, 4)
        self.rules_table.setHorizontalHeaderLabels(["Source Preview", "Destination", "Type", "Action"])
        self.rules_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.rules_table.setStyleSheet("QTableView { border: none; }")
        rules_layout.addWidget(self.rules_table)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.hide()
        rules_layout.addWidget(self.progress_bar)
        main_layout.addWidget(rules_frame, stretch=1)

    def setup_connections(self):
        self.btn_src_browse.clicked.connect(lambda: self.browse_file(self.src_file_input, self.src_sheet_combo))
        self.btn_dest_browse.clicked.connect(lambda: self.browse_file(self.dest_file_input, self.dest_sheet_combo))
        self.src_sheet_combo.currentTextChanged.connect(lambda s: self.load_sheet(self.src_file_input.text(), s, self.source_view))
        self.dest_sheet_combo.currentTextChanged.connect(lambda s: self.load_sheet(self.dest_file_input.text(), s, self.dest_view))
        self.btn_run.clicked.connect(self.vm.run_mapping)
        self.btn_save_preset.clicked.connect(self.save_preset_dialog)
        self.btn_load_preset.clicked.connect(self.load_preset_dialog)

        # VM Connections
        self.vm.rules_updated.connect(self.refresh_rules_ui)
        self.vm.mapping_started.connect(self.on_mapping_start)
        self.vm.mapping_progress.connect(self.progress_bar.setValue)
        self.vm.mapping_finished.connect(self.on_mapping_finish)
        self.vm.mapping_error.connect(self.on_mapping_error)

    def browse_file(self, line_edit, combo_box):
        path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xlsx *.xlsm)")
        if path:
            line_edit.setText(os.path.normpath(path))
            try:
                wb = openpyxl.load_workbook(path, read_only=True)
                combo_box.clear()
                combo_box.addItems(wb.sheetnames)
                wb.close()
            except: pass

    def load_sheet(self, filepath, sheetname, view):
        if filepath and sheetname:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                # 1. Load data dari file asli
                model = ExcelTableModel(filepath, sheetname)
                view.setModel(model)
                
                # 2. Sinkronisasi visual otomatis (Sangat Penting!)
                view.sync_with_excel() 
                
                self.refresh_highlights()
            finally:
                QApplication.restoreOverrideCursor()

# Dan update StyleSheet di setup_ui agar terlihat profesional seperti Office:
        self.setStyleSheet("""
            QTableView { 
                background-color: white; 
                gridline-color: #d1d5db; 
                color: black; 
                border: none;
                selection-background-color: #bbdefb;
            }
            QHeaderView::section { 
                background-color: #f3f4f6; 
                border: 1px solid #d1d5db; 
                padding: 4px;
                color: #4b5563;
                font-weight: normal;
            }
        """)
    def handle_drop(self, source_cell, dest_cell):
        src_model = self.source_view.model()
        src_val = "N/A"
        if src_model:
            r, c = coordinate_to_tuple(source_cell)
            src_val = src_model.grid_data[r-1][c-1]["value"]

        self.vm.add_rule(
            self.src_file_input.text(), self.src_sheet_combo.currentText(), source_cell, src_val,
            self.dest_file_input.text(), self.dest_sheet_combo.currentText(), dest_cell
        )

    def refresh_rules_ui(self, rules):
        self.rules_table.setRowCount(0)
        for i, rule in enumerate(rules):
            row = self.rules_table.rowCount()
            self.rules_table.insertRow(row)
            
            preview_text = f"{rule['src_sheet']}!{rule['src_cell']}  ➔  \"{rule.get('src_val', 'N/A')}\""
            dest_text = f"{rule['dest_sheet']}!{rule['dest_cell']}"
            
            self.rules_table.setItem(row, 0, QTableWidgetItem(preview_text))
            self.rules_table.setItem(row, 1, QTableWidgetItem(dest_text))
            self.rules_table.setItem(row, 2, QTableWidgetItem("Live Write"))
            
            btn_delete = QPushButton("🗑️")
            btn_delete.setProperty("class", "icon-btn")
            btn_delete.clicked.connect(lambda _, idx=i: self.vm.remove_rule(idx))
            self.rules_table.setCellWidget(row, 3, btn_delete)
            
        self.refresh_highlights()

    def refresh_highlights(self):
        curr_src = self.src_sheet_combo.currentText()
        curr_dest = self.dest_sheet_combo.currentText()
        src_model = self.source_view.model()
        dest_model = self.dest_view.model()
        
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

    def on_mapping_start(self):
        self.btn_run.setEnabled(False)
        self.progress_bar.show()
        self.progress_bar.setValue(0)

    def on_mapping_finish(self, message, count):
        self.btn_run.setEnabled(True)
        self.progress_bar.hide()
        QMessageBox.information(self, "Berhasil!", message)

    def on_mapping_error(self, error_msg):
        self.btn_run.setEnabled(True)
        self.progress_bar.hide()
        QMessageBox.critical(self, "Error", f"Terjadi kesalahan:\n{error_msg}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    vm = MapperViewModel()
    window = MainWindow(vm)
    window.show()
    sys.exit(app.exec())
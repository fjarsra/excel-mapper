import sys, os, openpyxl
from openpyxl.utils import coordinate_to_tuple
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QHeaderView, QPushButton, QLabel, QSplitter, QTableWidget, 
    QTableWidgetItem, QMessageBox, QFileDialog, QLineEdit, QComboBox, QFrame
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QPainter

from viewmodels.mapper_vm import MapperViewModel
from models.excel_handler import ExcelTableModel
# Pastikan views ini tersedia di project Anda
from views.components.excel_grid import DraggableTableView, DroppableTableView

class LoadingOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.hide()

    def paintEvent(self, event):
        painter = QPainter(self)
        # Menggunakan self.rect() memastikan background hitam transparan 
        # menutupi seluruh area yang sudah di-resize
        painter.fillRect(self.rect(), QColor(0, 0, 0, 150)) 
        
        painter.setPen(Qt.white)
        font = painter.font()
        font.setPointSize(20)
        font.setBold(True)
        painter.setFont(font)
        
        # Qt.AlignCenter akan menaruh teks tepat di tengah rect()
        painter.drawText(self.rect(), Qt.AlignCenter, "⏳ Membaca Excel...\nMohon Tunggu")
class MainWindow(QMainWindow):
    def __init__(self, view_model: MapperViewModel):
        super().__init__()
        self.vm = view_model
        
        # Penampung hasil pencarian
        self.src_matches = []
        self.dest_matches = []
        self.src_search_idx = -1
        self.dest_search_idx = -1

        self.setWindowTitle("Project Magang - Excel Mapper Pro")
        self.resize(1300, 850)
        self.is_dark_mode = False
        
        self.setup_ui()
        self.setup_connections()
        self.overlay = LoadingOverlay(self)

    def setup_ui(self):
        # MENGUNCI WARNA DAN MEMASTIKAN TAMPILAN MODERN (ANTI DARK MODE SISTEM)
        self.setStyleSheet("""
            QMainWindow { background-color: #f8fafc; }
            QFrame#card { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; }

            /* TOMBOL UTILITAS POJOK KANAN (Abu-abu) */
            QPushButton.icon_btn {
                background-color: #e2e8f0;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 5px 12px;
                min-height: 35px;
                font-size: 13px;
                color: #1e293b;
                font-weight: 500;
            }
            QPushButton.icon_btn:hover { background-color: #cbd5e1; }

            /* TOMBOL RUN (Hijau Emerald) */
            QPushButton#btn_run { 
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #065f46, stop:1 #10b981);
                color: #ffffff; border-radius: 10px; padding: 8px 25px; font-weight: bold; border: none;
            }

            /* TOMBOL AKSI (Browse, Zoom, Navigasi) */
            QPushButton.action_style {
                background-color: #f1f5f9;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                color: #0f172a;
                font-weight: bold;
            }
            QPushButton.action_style:hover { background-color: #e2e8f0; }

            /* INPUT & DROPDOWN */
            QLineEdit, QComboBox { 
                padding: 8px; border: 1px solid #e2e8f0; border-radius: 8px; 
                background-color: #ffffff; color: #000000; min-width: 150px;
            }
            
            /* TABEL MAPPING RULES MODERN */
            QTableWidget#rules_table {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                gridline-color: transparent; /* Hilangkan garis kotak kaku */
                outline: none;
            }

            QTableWidget#rules_table::item {
                color: #334155;
                padding-left: 15px;
                border-bottom: 1px solid #f1f5f9; /* Garis pemisah horizontal saja */
            }

            /* Header Tabel yang Elegan */
            QHeaderView::section {
                background-color: #f8fafc;
                color: #64748b;
                font-weight: bold;
                font-size: 10px;
                padding: 10px;
                border: none;
                border-bottom: 2px solid #e2e8f0;
                text-transform: uppercase;
                letter-spacing: 1px;
            }

            /* Warna Baris Selang-seling */
            QTableWidget#rules_table {
                alternate-background-color: #fafafa;
            }
            /* FIX AREA TABEL (Paling sering jadi gelap) */
            QTableView, QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f8fafc; /* Zebra stripe terang */
                border: 1px solid #e2e8f0;
                gridline-color: transparent;
                color: #1e293b; /* Warna teks hitam keabuan */
            }

            /* Paksa area kosong di dalam tabel tetap putih */
            QAbstractItemView {
                background-color: #ffffff;
                alternate-background-color: #f8fafc;
            }
        """)
        
        central_widget = QWidget(objectName="central")
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(25, 20, 25, 25)

        # --- HEADER AREA ---
        header_layout = QHBoxLayout()
        title_label = QLabel("🔗 Project Magang")
        title_label.setStyleSheet("font-size: 20px; font-weight: bold; color: #0f172a;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()

        # Inisialisasi Tombol dengan Teks
        self.btn_undo = QPushButton("↩ Undo"); self.btn_undo.setProperty("class", "icon_btn")
        self.btn_load_preset = QPushButton("📂 Load Preset"); self.btn_load_preset.setProperty("class", "icon_btn")
        self.btn_save_preset = QPushButton("💾 Save Preset"); self.btn_save_preset.setProperty("class", "icon_btn")
        self.btn_theme = QPushButton("🌙 Dark Mode"); self.btn_theme.setProperty("class", "icon_btn")
        self.btn_run = QPushButton("▶ Run Mapping"); self.btn_run.setObjectName("btn_run")

        for btn in [self.btn_undo, self.btn_load_preset, self.btn_save_preset, self.btn_theme, self.btn_run]:
            header_layout.addWidget(btn)
        main_layout.addLayout(header_layout)

        # --- CONFIG PANEL ---
        config_frame = QFrame(objectName="card")
        config_layout = QHBoxLayout(config_frame)
        
        for panel_type in ["src", "dest"]:
            box = QVBoxLayout()
            label = "SOURCE DATA" if panel_type == "src" else "DESTINATION DATA"
            box.addWidget(QLabel(label, styleSheet="font-weight: bold; color: #64748b; font-size: 11px;"))
            
            row = QHBoxLayout()
            f_input = QLineEdit(readOnly=True, placeholderText=f"Select {panel_type}...")
            b_btn = QPushButton("Browse"); b_btn.setProperty("class", "action_style")
            s_combo = QComboBox()
            
            row.addWidget(f_input, stretch=3); row.addWidget(b_btn); row.addWidget(s_combo, stretch=2)
            box.addLayout(row)
            
            if panel_type == "src":
                self.src_file_input, self.btn_src_browse, self.src_sheet_combo = f_input, b_btn, s_combo
                config_layout.addLayout(box, 1)
                config_layout.addWidget(QLabel("➔", styleSheet="font-size: 20px; color: #cbd5e1;"))
            else:
                self.dest_file_input, self.btn_dest_browse, self.dest_sheet_combo = f_input, b_btn, s_combo
                config_layout.addLayout(box, 1)

        main_layout.addWidget(config_frame)

        # --- WORKSPACE ---
        workspace_frame = QFrame(objectName="card")
        workspace_layout = QVBoxLayout(workspace_frame)
        splitter = QSplitter(Qt.Horizontal)

        for panel_type in ["src", "dest"]:
            cont = QWidget()
            vbox = QVBoxLayout(cont)
            h_header = QHBoxLayout()
            title = "🔵 Source Sheet" if panel_type == "src" else "🟢 Destination Sheet"
            h_header.addWidget(QLabel(title, styleSheet="font-weight: bold; color: #0f172a;"))
            h_header.addStretch()
            
            toggle = QPushButton("👁️"); z_out = QPushButton("-"); z_in = QPushButton("+")
            for b in [toggle, z_out, z_in]:
                b.setProperty("class", "action_style"); b.setFixedSize(35, 30); h_header.addWidget(b)
            
            search_lay = QHBoxLayout()
            s_input = QLineEdit(placeholderText="🔍 Search...")
            b_prev = QPushButton("⇽"); b_prev.setProperty("class", "action_style"); b_prev.setFixedSize(30, 30)
            b_next = QPushButton("⇾"); b_next.setProperty("class", "action_style"); b_next.setFixedSize(30, 30)
            
            search_lay.addWidget(s_input); search_lay.addWidget(b_prev); search_lay.addWidget(b_next)
            vbox.addLayout(h_header); vbox.addLayout(search_lay)
            
            if panel_type == "src":
                self.btn_src_toggle_view, self.btn_src_zoom_out, self.btn_src_zoom_in = toggle, z_out, z_in
                self.src_search_input, self.btn_src_prev, self.btn_src_next = s_input, b_prev, b_next
                self.source_view = DraggableTableView(); vbox.addWidget(self.source_view)
            else:
                self.btn_dest_toggle_view, self.btn_dest_zoom_out, self.btn_dest_zoom_in = toggle, z_out, z_in
                self.dest_search_input, self.btn_dest_prev, self.btn_dest_next = s_input, b_prev, b_next
                self.dest_view = DroppableTableView(self.handle_drop); vbox.addWidget(self.dest_view)
            splitter.addWidget(cont)

        workspace_layout.addWidget(splitter); main_layout.addWidget(workspace_frame, stretch=3)

        # --- RULES TABLE (LOG AKTIVITAS) ---
        rules_frame = QFrame(objectName="card")
        rules_layout = QVBoxLayout(rules_frame)
        rules_layout.addWidget(QLabel("MAPPING RULES", styleSheet="font-weight: bold; color: #64748b; font-size: 11px;"))

        self.rules_table = QTableWidget(0, 4)
        self.rules_table.setObjectName("rules_table")
        self.rules_table.setHorizontalHeaderLabels(["Source", "Destination", "Type", "Action"])
        self.rules_table.setObjectName("rules_table")
        self.rules_table.setStyleSheet("background-color: white; color: black;") # Proteksi extra
        self.rules_table.setAlternatingRowColors(True)

        # Pengaturan visual agar tabel bersih
        self.rules_table.verticalHeader().setVisible(False)
        self.rules_table.setShowGrid(False)
        self.rules_table.setAlternatingRowColors(True)
        self.rules_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.rules_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        rules_layout.addWidget(self.rules_table)
        main_layout.addWidget(rules_frame, stretch=1)

    def setup_connections(self):
        # --- Browsing & Loading ---
        self.btn_src_browse.clicked.connect(lambda: self.browse_file(self.src_file_input, self.src_sheet_combo))
        self.btn_dest_browse.clicked.connect(lambda: self.browse_file(self.dest_file_input, self.dest_sheet_combo))
        self.src_sheet_combo.currentTextChanged.connect(lambda s: self.load_sheet(self.src_file_input.text(), s, self.source_view))
        self.dest_sheet_combo.currentTextChanged.connect(lambda s: self.load_sheet(self.dest_file_input.text(), s, self.dest_view))
        
        # --- Search Connections ---
        self.src_search_input.textChanged.connect(lambda: self.apply_highlight(self.source_view, "src"))
        self.dest_search_input.textChanged.connect(lambda: self.apply_highlight(self.dest_view, "dest"))
        self.btn_src_next.clicked.connect(lambda: self.navigate_search("src", True))
        self.btn_src_prev.clicked.connect(lambda: self.navigate_search("src", False))
        self.btn_dest_next.clicked.connect(lambda: self.navigate_search("dest", True))
        self.btn_dest_prev.clicked.connect(lambda: self.navigate_search("dest", False))

        # --- Utilitas Header (Save, Load, Undo, Theme) ---
        self.btn_save_preset.clicked.connect(self.save_preset_dialog) # Tambahkan ini
        self.btn_load_preset.clicked.connect(self.load_preset_dialog) # Tambahkan ini
        self.btn_undo.clicked.connect(self.vm.undo_last_rule)
        self.btn_theme.clicked.connect(self.toggle_theme)

        # --- Main Execution ---
        self.btn_run.clicked.connect(self.vm.run_mapping)
        self.vm.rules_updated.connect(self.refresh_rules_ui)
        
        # --- Zoom & Toggle ---
        self.btn_src_zoom_in.clicked.connect(lambda: self.source_view.apply_zoom(1))
        self.btn_src_zoom_out.clicked.connect(lambda: self.source_view.apply_zoom(-1))
        self.btn_src_toggle_view.clicked.connect(lambda: self.source_view.toggle_hidden_columns() if hasattr(self.source_view, 'toggle_hidden_columns') else None)
        self.btn_dest_zoom_in.clicked.connect(lambda: self.dest_view.apply_zoom(1))
        self.btn_dest_zoom_out.clicked.connect(lambda: self.dest_view.apply_zoom(-1))
        self.btn_dest_toggle_view.clicked.connect(lambda: self.dest_view.toggle_hidden_columns() if hasattr(self.dest_view, 'toggle_hidden_columns') else None)

    def apply_highlight(self, view, p_type):
        model = view.model()
        if not model: return
        
        text = self.src_search_input.text().lower() if p_type == "src" else self.dest_search_input.text().lower()
        matches = []
        
        model.layoutAboutToBeChanged.emit()
        for r in range(model.rowCount()):
            for c in range(model.columnCount()):
                idx = model.index(r, c)
                val = str(model.grid_data[r][c]["value"]).lower()
                if text and text in val:
                    model.setData(idx, QColor("#ccfbf1"), Qt.BackgroundRole)
                    matches.append(idx)
                else:
                    model.setData(idx, None, Qt.BackgroundRole)
        
        if p_type == "src":
            self.src_matches, self.src_search_idx = matches, -1
            self.btn_src_next.setEnabled(len(matches) > 0)
            self.btn_src_prev.setEnabled(len(matches) > 0)
        else:
            self.dest_matches, self.dest_search_idx = matches, -1
            self.btn_dest_next.setEnabled(len(matches) > 0)
            self.btn_dest_prev.setEnabled(len(matches) > 0)
            
        model.layoutChanged.emit()

        # TAMBAHKAN INI: Jika ditemukan hasil, langsung lompat ke hasil pertama
        if matches:
            self.navigate_search(p_type, forward=True)

    def navigate_search(self, panel_type, forward=True):
        matches = self.src_matches if panel_type == "src" else self.dest_matches
        view = self.source_view if panel_type == "src" else self.dest_view
        
        if not matches:
            return

        if panel_type == "src":
            self.src_search_idx = (self.src_search_idx + 1) % len(matches) if forward else (self.src_search_idx - 1) % len(matches)
            idx = matches[self.src_search_idx]
        else:
            self.dest_search_idx = (self.dest_search_idx + 1) % len(matches) if forward else (self.dest_search_idx - 1) % len(matches)
            idx = matches[self.dest_search_idx]
            
        # Perbaikan: Langsung panggil scrollTo(idx) tanpa parameter tambahan.
        # Secara default, PySide akan memastikan sel tersebut terlihat (EnsureVisible).
        view.scrollTo(idx)
        view.setCurrentIndex(idx)
    def browse_file(self, line_edit, combo_box):
        path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xlsx *.xlsm)")
        if path:
            line_edit.setText(os.path.normpath(path))
            wb = openpyxl.load_workbook(path, read_only=True)
            combo_box.clear(); combo_box.addItems(wb.sheetnames); wb.close()

    def load_sheet(self, filepath, sheetname, view):
        if filepath and sheetname:
            self.overlay.show(); QApplication.processEvents(); QApplication.setOverrideCursor(Qt.WaitCursor)
            try:
                model = ExcelTableModel(filepath, sheetname)
                view.setModel(model)
                # Sinkronisasi visual awal (highlight sel yang sudah dimapping jika ada)
                self.refresh_highlights()
            finally:
                QApplication.restoreOverrideCursor(); self.overlay.hide()
    
    def save_preset_dialog(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Preset", "", "JSON Files (*.json)")
        if path:
            self.vm.save_preset(path)
            QMessageBox.information(self, "Success", "Preset berhasil disimpan!")

    def load_preset_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load Preset", "", "JSON Files (*.json)")
        if path:
            self.vm.load_preset(path)
            # Setelah load, kita perlu refresh UI tabel dan highlight
            self.refresh_rules_ui(self.vm.rules)
            QMessageBox.information(self, "Success", "Preset berhasil dimuat!")
            
    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        if self.is_dark_mode:
            self.btn_theme.setText("☀️ Light Mode")
            # Di sini Anda bisa menambahkan logic untuk mengganti self.setStyleSheet 
            # menjadi versi gelap jika dibutuhkan di masa depan.
        else:
            self.btn_theme.setText("🌙 Dark Mode")

    def handle_drop(self, source_cell, dest_cell):
        src_model = self.source_view.model()
        src_val = "N/A"
        if src_model:
            try:
                r, c = coordinate_to_tuple(source_cell)
                src_val = src_model.grid_data[r-1][c-1]["value"]
            except: pass
        self.vm.add_rule(self.src_file_input.text(), self.src_sheet_combo.currentText(), source_cell, src_val,
                         self.dest_file_input.text(), self.dest_sheet_combo.currentText(), dest_cell)

    def refresh_rules_ui(self, rules):
        self.rules_table.setRowCount(0)
        # Matikan update visual sementara agar proses lebih cepat
        self.rules_table.setUpdatesEnabled(False)
        
        for i, rule in enumerate(rules):
            row = self.rules_table.rowCount()
            self.rules_table.insertRow(row)
            self.rules_table.setRowHeight(row, 45) # Buat baris lebih tinggi agar lega

            # Data Sel dengan Formatting
            src_text = f"📄 {rule['src_sheet']} › {rule['src_cell']}"
            dest_text = f"🎯 {rule['dest_sheet']} › {rule['dest_cell']}"
            
            item_src = QTableWidgetItem(src_text)
            item_dest = QTableWidgetItem(dest_text)
            item_type = QTableWidgetItem("⚡ Live Mapping")
            
            # Tambahkan sedikit warna pada teks type agar menarik
            item_type.setForeground(QColor("#0891b2")) 
            
            for col, item in enumerate([item_src, item_dest, item_type]):
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.rules_table.setItem(row, col, item)

            # --- Tombol Delete Custom ---
            btn_container = QWidget()
            btn_lay = QHBoxLayout(btn_container)
            btn_lay.setContentsMargins(0, 0, 0, 0)
            btn_lay.setAlignment(Qt.AlignCenter)

            btn_del = QPushButton("✕") # Gunakan silang tipis yang modern
            btn_del.setFixedSize(24, 24)
            btn_del.setStyleSheet("""
                QPushButton {
                    border-radius: 12px;
                    border: 1px solid #fecaca;
                    color: #ef4444;
                    background-color: #ffffff;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #ef4444;
                    color: white;
                }
            """)
            btn_del.clicked.connect(lambda _, idx=i: self.vm.remove_rule(idx))
            
            btn_lay.addWidget(btn_del)
            self.rules_table.setCellWidget(row, 3, btn_container)

        self.rules_table.setUpdatesEnabled(True)
        self.refresh_highlights()

    def refresh_highlights(self):
        src_model, dest_model = self.source_view.model(), self.dest_view.model()
        if src_model: src_model.clear_highlights()
        if dest_model: dest_model.clear_highlights()
        
        curr_src = self.src_sheet_combo.currentText()
        curr_dest = self.dest_sheet_combo.currentText()

        for rule in self.vm.rules:
            if src_model and rule["src_sheet"] == curr_src:
                src_model.add_highlight(rule["src_cell"], QColor("#f0fdfa"))
            if dest_model and rule["dest_sheet"] == curr_dest:
                dest_model.add_highlight(rule["dest_cell"], QColor("#f0fdfa"))

    def toggle_theme(self):
        self.is_dark_mode = not self.is_dark_mode
        self.btn_theme.setText("☀️" if self.is_dark_mode else "🌙")
       
    def resizeEvent(self, event):
        # Memastikan overlay selalu mengikuti ukuran jendela MainWindow
        if hasattr(self, 'overlay'):
            self.overlay.resize(self.size())
        super().resizeEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    vm = MapperViewModel()
    window = MainWindow(vm)
    window.show()
    sys.exit(app.exec())
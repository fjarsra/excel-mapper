from PySide6.QtWidgets import QTableView
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QDrag
from models.excel_handler import get_excel_col_name
from PySide6.QtGui import QDrag, QPixmap

class DraggableTableView(QTableView):
    def __init__(self):
        super().__init__()
        self.setSelectionMode(QTableView.SingleSelection)
        self.setSelectionBehavior(QTableView.SelectItems)
        self.setDragEnabled(True)
        self.zoom_level = 0
        self.is_compact_view = True # Default menyembunyikan kolom tersembunyi

    def sync_with_excel(self):
        """Menerapkan struktur sel asli (Merged Cells)."""
        model = self.model()
        if not model: return
        self.clearSpans()
        for m_range in model.merged_ranges:
            min_col, min_row, max_col, max_row = m_range.bounds
            self.setSpan(min_row - 1, min_col - 1, max_row - min_row + 1, max_col - min_col + 1)
        self.apply_zoom(0)
        self.update_column_visibility()

    def toggle_hidden_columns(self):
        """Beralih antara mode Ringkas (mengikuti Excel) dan mode Lengkap."""
        self.is_compact_view = not self.is_compact_view
        self.update_column_visibility()

    def update_column_visibility(self):
        """Memperbarui tampilan kolom berdasarkan status is_compact_view."""
        model = self.model()
        if not model: return
        for col_idx in model.hidden_cols:
            self.setColumnHidden(col_idx, self.is_compact_view)

    def apply_zoom(self, delta):
        self.zoom_level = max(-5, min(10, self.zoom_level + delta))
        model = self.model()
        if not model: return
        scale = 1 + (self.zoom_level * 0.15)
        model.set_font_size(max(6, int(10 * scale)))
        for r_idx, h in model.row_heights.items():
            self.setRowHeight(r_idx, int(h * 1.33 * scale))
        for c_idx, w in model.col_widths.items():
            self.setColumnWidth(c_idx, int(w * 8 * scale))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.LeftButton): return
        if (event.position().toPoint() - self.drag_start_position).manhattanLength() < 5: return
        index = self.indexAt(event.position().toPoint())
        if not index.isValid(): return
        excel_cell = f"{get_excel_col_name(index.column())}{index.row() + 1}"
        
        # --- FITUR GHOST IMAGE ---
        pixmap = self.viewport().grab(self.visualRect(index)) # Ambil gambar sel
        
        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setText(f"SOURCE:{excel_cell}")
        drag.setMimeData(mime_data)
        
        drag.setPixmap(pixmap) # Pasang gambar transparan ke kursor
        drag.setHotSpot(event.position().toPoint() - self.visualRect(index).topLeft())
        
        drag.exec_(Qt.CopyAction)
    # Tambahkan method ini di dalam kelas DraggableTableView dan DroppableTableView
def apply_zoom(self, delta):
    """Mengatur ukuran font dan tinggi baris/kolom berdasarkan delta (+1 atau -1)"""
    font = self.font()
    new_size = max(6, font.pointSize() + delta) # Ukuran minimal 6
    font.setPointSize(new_size)
    self.setFont(font)
    
    # Update ukuran header
    header_font = self.horizontalHeader().font()
    header_font.setPointSize(new_size)
    self.horizontalHeader().setFont(header_font)
    self.verticalHeader().setFont(header_font)
    
    # Sesuaikan tinggi baris dan lebar kolom otomatis agar pas dengan font baru
    self.verticalHeader().setDefaultSectionSize(new_size * 2.5)
    self.horizontalHeader().setDefaultSectionSize(new_size * 5)
class DroppableTableView(QTableView):
    def __init__(self, drop_callback):
        super().__init__()
        self.drop_callback = drop_callback
        self.setAcceptDrops(True)
        self.zoom_level = 0
        self.is_compact_view = True

    def sync_with_excel(self):
        model = self.model()
        if not model: return
        self.clearSpans()
        for m_range in model.merged_ranges:
            min_col, min_row, max_col, max_row = m_range.bounds
            self.setSpan(min_row - 1, min_col - 1, max_row - min_row + 1, max_col - min_col + 1)
        self.apply_zoom(0)
        self.update_column_visibility()

    def toggle_hidden_columns(self):
        self.is_compact_view = not self.is_compact_view
        self.update_column_visibility()

    def update_column_visibility(self):
        model = self.model()
        if not model: return
        for col_idx in model.hidden_cols:
            self.setColumnHidden(col_idx, self.is_compact_view)

    def apply_zoom(self, delta):
        """Menerapkan zoom dengan menjaga proporsi lebar kolom asli Excel"""
        self.zoom_level = max(-5, min(10, self.zoom_level + delta))
        model = self.model()
        if not model: return
        
        # Hitung skala (1.0 adalah ukuran normal)
        scale = 1 + (self.zoom_level * 0.15)
        new_font_size = max(6, int(10 * scale))
        
        # 1. Update Font Isi Tabel
        model.set_font_size(new_font_size)
        
        # 2. Update Font Header (A, B, C dan 1, 2, 3)
        h_font = self.horizontalHeader().font()
        h_font.setPointSize(new_font_size)
        self.horizontalHeader().setFont(h_font)
        self.verticalHeader().setFont(h_font)

        # 3. Update Tinggi Baris & Lebar Kolom Berdasarkan Data Excel
        for r_idx, h in model.row_heights.items():
            self.setRowHeight(r_idx, int(h * 1.33 * scale))
        
        for c_idx, w in model.col_widths.items():
            # Gunakan angka 9 sebagai pengali agar pas dengan pixel
            self.setColumnWidth(c_idx, int(w * 9 * scale))
            
        # 4. Update ukuran default untuk sel yang tidak punya ukuran khusus
        self.verticalHeader().setDefaultSectionSize(int(25 * scale))

    def wheelEvent(self, event):
        if event.modifiers() == Qt.ControlModifier:
            delta = 1 if event.angleDelta().y() > 0 else -1
            self.apply_zoom(delta)
        else:
            super().wheelEvent(event)

    # --- EVENT DRAG & DROP TETAP SAMA ---
    def dragEnterEvent(self, event):
        if event.mimeData().hasText() and event.mimeData().text().startswith("SOURCE:"):
            self.setStyleSheet("QTableView::item:selected { border: 2px dashed #10b981; background-color: rgba(16, 185, 129, 0.3); }")
            event.acceptProposedAction()

    def dragMoveEvent(self, event): 
        index = self.indexAt(event.position().toPoint())
        if index.isValid():
            self.setCurrentIndex(index)
        event.acceptProposedAction()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("")
        self.clearSelection()
        
    def dropEvent(self, event):
        self.setStyleSheet("")
        self.clearSelection() 
        text = event.mimeData().text()
        if text.startswith("SOURCE:"):
            source_cell = text.split(":")[1]
            index = self.indexAt(event.position().toPoint())
            if index.isValid():
                dest_cell = f"{get_excel_col_name(index.column())}{index.row() + 1}"
                self.drop_callback(source_cell, dest_cell)
                event.acceptProposedAction()
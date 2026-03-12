from PySide6.QtWidgets import QTableView
from PySide6.QtCore import Qt, QMimeData
from PySide6.QtGui import QDrag
from models.excel_handler import get_excel_col_name

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
        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setText(f"SOURCE:{excel_cell}")
        drag.setMimeData(mime_data)
        drag.exec_(Qt.CopyAction)

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
        self.zoom_level = max(-5, min(10, self.zoom_level + delta))
        model = self.model()
        if not model: return
        scale = 1 + (self.zoom_level * 0.15)
        model.set_font_size(max(6, int(10 * scale)))
        for r_idx, h in model.row_heights.items():
            self.setRowHeight(r_idx, int(h * 1.33 * scale))
        for c_idx, w in model.col_widths.items():
            self.setColumnWidth(c_idx, int(w * 8 * scale))

    def dragEnterEvent(self, event):
        if event.mimeData().hasText() and event.mimeData().text().startswith("SOURCE:"):
            event.acceptProposedAction()
    def dragMoveEvent(self, event): event.acceptProposedAction()
    def dropEvent(self, event):
        text = event.mimeData().text()
        if text.startswith("SOURCE:"):
            source_cell = text.split(":")[1]
            index = self.indexAt(event.position().toPoint())
            if index.isValid():
                dest_cell = f"{get_excel_col_name(index.column())}{index.row() + 1}"
                self.drop_callback(source_cell, dest_cell)
                event.acceptProposedAction()
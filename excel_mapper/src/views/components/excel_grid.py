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

    def apply_zoom(self, delta):
        self.zoom_level += delta
        font_size = max(8, 10 + self.zoom_level)
        self.horizontalHeader().setDefaultSectionSize(max(50, 100 + (self.zoom_level * 15)))
        self.verticalHeader().setDefaultSectionSize(max(20, 30 + (self.zoom_level * 5)))
        if self.model(): self.model().set_font_size(font_size)

class DroppableTableView(QTableView):
    def __init__(self, drop_callback):
        super().__init__()
        self.drop_callback = drop_callback
        self.setSelectionMode(QTableView.SingleSelection)
        self.setSelectionBehavior(QTableView.SelectItems)
        self.setAcceptDrops(True)
        self.zoom_level = 0

    def dragEnterEvent(self, event):
        if event.mimeData().hasText() and event.mimeData().text().startswith("SOURCE:"):
            event.acceptProposedAction()

    def dragMoveEvent(self, event): 
        event.acceptProposedAction()

    def dropEvent(self, event):
        text = event.mimeData().text()
        if text.startswith("SOURCE:"):
            source_cell = text.split(":")[1]
            index = self.indexAt(event.position().toPoint())
            if index.isValid():
                dest_cell = f"{get_excel_col_name(index.column())}{index.row() + 1}"
                self.drop_callback(source_cell, dest_cell)
                event.acceptProposedAction()

    def apply_zoom(self, delta):
        self.zoom_level += delta
        font_size = max(8, 10 + self.zoom_level)
        self.horizontalHeader().setDefaultSectionSize(max(50, 100 + (self.zoom_level * 15)))
        self.verticalHeader().setDefaultSectionSize(max(20, 30 + (self.zoom_level * 5)))
        if self.model(): self.model().set_font_size(font_size)
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMessageBox, QTableWidget, QPushButton
from app.utils.table_handler import re_sort_table
from app.ui.styles import AppStyles, AppColors

class StyledButton(QPushButton):
    """Standard Button with predefined styles."""
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(AppStyles.BUTTON_HEIGHT)


class MaterialTableWidget(QTableWidget):
    """
    Cosmetic Raw Material Table Widget
    Encapsulates logic for:
    - Right-click Context Menu (Row Splitting)
    - Validating edits (Prevent duplicate RM names with different percentages)
    - Propagating edits to merged cells
    - Auto-resorting
    """
    
    # Signal emitted when content changes significantly (requires external re-comparison)
    contentChanged = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_updating = False
        self.old_text_value = ""
        
        # Connect internal signals
        self.itemChanged.connect(self._on_item_changed)
        self.cellDoubleClicked.connect(self._on_cell_double_clicked)
        
        # Context Menu
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_context_menu)

    def _on_cell_double_clicked(self, row, col):
        """Save old value before editing."""
        item = self.item(row, col)
        if item:
            self.old_text_value = item.text()
        else:
            self.old_text_value = ""

    def _show_context_menu(self, pos):
        """Show context menu for splitting rows."""
        item = self.itemAt(pos)
        if not item:
            return
            
        # Only allow splitting on INCI columns (2, 3) to avoid ambiguity
        if item.column() not in [2, 3]:
            return

        menu = QtWidgets.QMenu(self)
        split_action = menu.addAction("이 행 분할하기 (Split Row)")
        action = menu.exec_(self.mapToGlobal(pos))
        
        if action == split_action:
            self.split_row(item.row())

    def split_row(self, row):
        """Slits the selected row into a new RM group."""
        if self.is_updating:
            return

        try:
            self.is_updating = True
            rm_item = self.item(row, 0)
            if rm_item:
                original_name = rm_item.text()
                # Create new name: "Original (1)"
                new_name = f"{original_name} (1)"
                rm_item.setText(new_name)
            
            # Re-sort immediately to reflect structure change
            re_sort_table(self)
            
            # Notify external world (Main Window) to re-compare
            self.contentChanged.emit()

        finally:
            self.is_updating = False

    def _on_item_changed(self, item):
        """Handle item edits: Validation, Propagation, and Re-sorting."""
        if self.is_updating:
            return

        try:
            self.is_updating = True
            new_text = item.text().strip()

            # 1. Empty Check
            if not new_text:
                QMessageBox.warning(self, "경고", "빈 값은 입력할 수 없습니다.")
                item.setText(self.old_text_value)
                return

            # 2. Duplicate Check (RM Name conflict with different percentages)
            if item.column() == 0:
                current_rm_name = new_text
                # Get current row's percentage
                pct_item = self.item(item.row(), 1)
                current_rm_pct = pct_item.text().strip() if pct_item else ""
                
                # Scan entire table for conflict
                for r in range(self.rowCount()):
                    if r == item.row():
                        continue
                    
                    other_rm = self.item(r, 0)
                    if other_rm and other_rm.text().strip() == current_rm_name:
                        other_pct_item = self.item(r, 1)
                        other_pct = other_pct_item.text().strip() if other_pct_item else ""
                        
                        if other_pct != current_rm_pct:
                            QMessageBox.warning(
                                self, 
                                "값 변경 불가", 
                                f"이미 존재하는 '{current_rm_name}' 원료와 함량({other_pct}%)이 다릅니다.\n"
                                f"현재 행의 함량({current_rm_pct}%)과 일치하지 않아 병합할 수 없습니다."
                            )
                            item.setText(self.old_text_value)
                            return

            # 3. Propagate Changes to Merged Cells
            self._propagate_merged_cell_change(item)

            # 4. Re-sort
            re_sort_table(self)
            
            # 5. Notify Main
            self.contentChanged.emit()

        finally:
            self.is_updating = False

    def _propagate_merged_cell_change(self, item):
        """If a merged cell (RM/Percent) is edited, update all cells in the span."""
        row = item.row()
        col = item.column()
        
        # Only apply to RM(0) or Percentage(1)
        if col not in [0, 1]:
            return

        span_height = self.rowSpan(row, col)
        if span_height > 1:
            new_text = item.text()
            for r in range(row + 1, row + span_height):
                target = self.item(r, col)
                if target:
                    target.setText(new_text)

    def reset_styles(self):
        """Reset all cell styles to default (White bg, Black text)."""
        white_brush = QtGui.QBrush(AppColors.WHITE)
        black_brush = QtGui.QBrush(AppColors.BLACK)
        
        for r in range(self.rowCount()):
            for c in range(self.columnCount()):
                item = self.item(r, c)
                if item:
                    item.setBackground(white_brush)
                    item.setForeground(black_brush)

    def apply_diff_report(self, diff_items):
        """Apply styling based on diff report."""
        from PyQt5 import QtGui
        from app.models import DiffType
        
        red_brush = QtGui.QBrush(AppColors.TEXT_RED)
        bg_red_brush = QtGui.QBrush(AppColors.BG_RED)
        
        # Block signals to prevent itemChanged recursion
        was_blocked = self.signalsBlocked()
        self.blockSignals(True)
        
        try:
            self.reset_styles()
            
            for diff in diff_items:
                item = self.item(diff.row, diff.col)
                if not item:
                    continue
                    
                if diff.diff_type == DiffType.CONTENT_MISMATCH:
                    item.setForeground(red_brush)
                elif diff.diff_type in (DiffType.MISSING_ROW, DiffType.MISSING_INCI):
                    item.setBackground(bg_red_brush)
                    
        finally:
            self.blockSignals(was_blocked)


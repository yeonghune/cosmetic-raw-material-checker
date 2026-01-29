import os
from PyQt5 import QtWidgets, QtCore, QtGui
from itertools import zip_longest

from app.ui.dialogs.text_input_dialog import TextInputDialog
from app.ui.widgets import StyledButton
from app.ui.styles import AppColors, AppStyles
from app.utils.text_parser import parse_ingredients
from app.utils.comparator import compare_ingredients
from app.utils.excel_handler import export_comparison_table

class TextComparatorPage(QtWidgets.QWidget):
    # Signal to request navigation to home
    navigate_home = QtCore.pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.list1_data = [] # List of strings
        self.list2_data = [] # List of strings
        self.is_updating = False
        self._init_ui()
        
    def _init_ui(self):
        self.setObjectName("TextComparatorPage")
        
        # Main Layout
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(AppStyles.LAYOUT_MARGIN, AppStyles.LAYOUT_MARGIN, AppStyles.LAYOUT_MARGIN, AppStyles.LAYOUT_MARGIN)
        layout.setSpacing(AppStyles.LAYOUT_SPACING)
        
        # Components
        layout.addLayout(self._create_header())
        self.table = self._create_table()
        layout.addWidget(self.table)
        
        self.summaryLabel = QtWidgets.QLabel("데이터를 업로드해주세요.")
        layout.addWidget(self.summaryLabel)
        
    def _create_header(self) -> QtWidgets.QHBoxLayout:
        layout = QtWidgets.QHBoxLayout()
        layout.setSpacing(AppStyles.HEADER_SPACING)
        
        # Home Button
        self.homeButton = StyledButton("🏠 Home")
        self.homeButton.clicked.connect(self.go_home)
        layout.addWidget(self.homeButton)
        
        # Action Buttons
        self.btnUpload1 = StyledButton("1열 업로드")
        self.btnUpload1.clicked.connect(lambda: self.on_upload_click(1))
        layout.addWidget(self.btnUpload1)
        
        self.btnUpload2 = StyledButton("2열 업로드")
        self.btnUpload2.clicked.connect(lambda: self.on_upload_click(2))
        layout.addWidget(self.btnUpload2)
        
        self.btnExport = StyledButton("데이터 추출 (.xlsx)")
        self.btnExport.clicked.connect(self.on_export_click)
        layout.addWidget(self.btnExport)
        
        self.btnReset = StyledButton("초기화")
        self.btnReset.clicked.connect(self.reset_ui)
        layout.addWidget(self.btnReset)
        
        layout.addStretch(1)
        return layout
        
    def _create_table(self) -> QtWidgets.QTableWidget:
        table = QtWidgets.QTableWidget()
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["A열 성분", "B열 성분"])
        table.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        # Allow editing
        table.itemChanged.connect(self.on_item_changed)
        return table

    def go_home(self):
        """Reset state and navigate home."""
        self.reset_ui()
        self.navigate_home.emit()

    def reset_ui(self):
        """Resets the UI and internal data to initial state."""
        self.list1_data = []
        self.list2_data = []
        self.is_updating = True # Block signals while clearing
        self.table.setRowCount(0)
        self.is_updating = False
        self.summaryLabel.setText("데이터를 업로드해주세요.")
            
    def on_upload_click(self, col_idx: int):
        dialog = TextInputDialog(f"{col_idx}열 데이터 입력", self)
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            text = dialog.get_text()
            ingredients = parse_ingredients(text)
            
            if col_idx == 1:
                self.list1_data = ingredients
            else:
                self.list2_data = ingredients
                
            self.update_comparison()
            
    def on_item_changed(self, item):
        """Handle user edits in the table."""
        if self.is_updating:
            return
            
        row = item.row()
        col = item.column()
        text = item.text().strip()
        
        # Update internal data
        target_list = self.list1_data if col == 0 else self.list2_data
        
        # Ensure list is long enough
        while len(target_list) <= row:
            target_list.append("")
            
        target_list[row] = text
        
        # Re-run comparison to update highlights
        self.update_comparison()

    def update_comparison(self):
        # Prevent recursion (since setItem triggers itemChanged)
        self.is_updating = True
        try:
            # 1. Compare using logic from utils
            rows = compare_ingredients(self.list1_data, self.list2_data)
            
            # 2. Render to Table
            self.table.setRowCount(0)
            self.table.setRowCount(len(rows))
            
            match_count = 0
            
            for r_idx, (val1, val2, status) in enumerate(rows):
                item1 = QtWidgets.QTableWidgetItem(val1)
                item2 = QtWidgets.QTableWidgetItem(val2)
                
                if status == "MATCH":
                    match_count += 1
                else:
                    # "DIFF" - Highlight both cells
                    item1.setBackground(AppColors.DIFF_BG_YELLOW)
                    item2.setBackground(AppColors.DIFF_BG_YELLOW)
                    
                self.table.setItem(r_idx, 0, item1)
                self.table.setItem(r_idx, 1, item2)
                
            # 3. Update Summary
            total = len(rows)
            if total > 0:
                self.summaryLabel.setText(f"총 {total}행 / 일치 {match_count}행")
            else:
                self.summaryLabel.setText("데이터를 업로드해주세요.")
        
        finally:
            self.is_updating = False

    def on_export_click(self):
        if self.table.rowCount() == 0:
            QtWidgets.QMessageBox.warning(self, "경고", "추출할 데이터가 없습니다.")
            return

        file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "엑셀 파일 저장", "comparison_result.xlsx", "Excel Files (*.xlsx)"
        )
        if not file_path:
            return
            
        try:
            export_comparison_table(self.table, file_path)
            
            if QtWidgets.QMessageBox.question(self, "완료", "파일이 저장되었습니다.\n열시겠습니까?") == QtWidgets.QMessageBox.Yes:
                os.startfile(os.path.dirname(file_path))
                
        except Exception as e:
            print(f"Export Error: {e}")
            QtWidgets.QMessageBox.critical(self, "에러", f"저장 중 오류가 발생했습니다.\n{e}")

import os
from pathlib import Path
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMessageBox

from app.ui.widgets import MaterialTableWidget, StyledButton
from app.ui.styles import AppStyles
from app.utils.excel_handler import (
    download_template_file, 
    export_to_excel
)
from app.utils.table_handler import (
    make_table,
    extract_data_from_table
)
from app.utils.diff_logic import generate_diff_report

class CheckerPage(QtWidgets.QWidget):
    """
    기존 MainWindow의 기능(원료 검증기)을 모두 포함하는 위젯입니다.
    """
    
    # 페이지 전환 요청 시그널 (부모인 Main에게 전달)
    navigate_home = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_updating = False
        self._init_ui()
        self._setup_connections()
        # Custom table setup moved to _init_ui
        self._setup_table_sync()

    def _init_ui(self):
        """UI 구성 (기존 ui/main_ui.py의 내용을 코드로 포팅)"""
        self.setObjectName("CheckerPage")
        
        # Main Layout
        self.verticalLayout = QtWidgets.QVBoxLayout(self)
        self.verticalLayout.setContentsMargins(AppStyles.LAYOUT_MARGIN, AppStyles.LAYOUT_MARGIN, AppStyles.LAYOUT_MARGIN, AppStyles.LAYOUT_MARGIN)
        self.verticalLayout.setSpacing(AppStyles.LAYOUT_SPACING)

        # ----------------------------------------------------------------
        # Header (Home Button + Download Template + Upload + Download Result)
        # ----------------------------------------------------------------
        self.headerLayout = QtWidgets.QHBoxLayout()
        self.headerLayout.setSpacing(AppStyles.HEADER_SPACING)

        # [NEW] Home Button
        self.homeButton = StyledButton("🏠 Home")
        self.headerLayout.addWidget(self.homeButton)

        # Download Template
        self.downloadButton = StyledButton("템플릿 다운로드")
        self.headerLayout.addWidget(self.downloadButton)

        # Upload
        self.uploadButton = StyledButton("템플릿 불러오기")
        self.headerLayout.addWidget(self.uploadButton)

        # Download Result
        self.downloadResultButton = StyledButton("검증 결과 다운로드")
        self.headerLayout.addWidget(self.downloadResultButton)

        # File Label
        self.fileLabel = QtWidgets.QLabel("템플릿이 로드되지 않았습니다.")
        font = QtGui.QFont("Arial", 8)
        self.fileLabel.setFont(font)
        self.fileLabel.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        
        # Stretch Item (Spacer alternative) to push label to right or fill space
        self.headerLayout.addWidget(self.fileLabel, 1)

        self.verticalLayout.addLayout(self.headerLayout)

        # ----------------------------------------------------------------
        # Splitter (Table 1 | Table 2)
        # ----------------------------------------------------------------
        self.tableSplitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        self.tableSplitter.setHandleWidth(6)

        # Table 1 Group
        self.table1Group = QtWidgets.QGroupBox("테이블 1")
        self.table1Layout = QtWidgets.QVBoxLayout(self.table1Group)
        
        # Direct Instantiation of MaterialTableWidget
        self.table1Table = MaterialTableWidget()
        # Connect change signal immediately
        self.table1Table.contentChanged.connect(self.on_tables_content_changed)
        self.table1Layout.addWidget(self.table1Table)
        self.tableSplitter.addWidget(self.table1Group)

        # Table 2 Group
        self.table2Group = QtWidgets.QGroupBox("테이블 2")
        self.table2Layout = QtWidgets.QVBoxLayout(self.table2Group)
        
        # Direct Instantiation of MaterialTableWidget
        self.table2Table = MaterialTableWidget()
        self.table2Table.contentChanged.connect(self.on_tables_content_changed)
        self.table2Layout.addWidget(self.table2Table)
        self.tableSplitter.addWidget(self.table2Group)

        self.verticalLayout.addWidget(self.tableSplitter, 1) # Stretch factor 1

        # ----------------------------------------------------------------
        # Summary Label
        # ----------------------------------------------------------------
        self.summaryLabel = QtWidgets.QLabel("불일치 0건 / 총 0건")
        self.summaryLabel.setMinimumHeight(20)
        self.verticalLayout.addWidget(self.summaryLabel)

    def _setup_connections(self):
        """기본 시그널 연결"""
        self.homeButton.clicked.connect(self.go_home)
        self.downloadButton.clicked.connect(self.on_download_template)
        self.uploadButton.clicked.connect(self.on_upload_file)
        self.downloadResultButton.clicked.connect(self.on_download_result)

    def go_home(self):
        self.reset_ui()
        self.navigate_home.emit()

    def reset_ui(self):
        """Resets the UI state."""
        self.table1Table.setRowCount(0)
        self.table2Table.setRowCount(0)
        self.summaryLabel.setText("불일치 0건 / 총 0건")
        self.fileLabel.setText("템플릿이 로드되지 않았습니다.")


    def _setup_table_sync(self):
        """Scroll synchronization for tables."""
        # Note: MaterialTableWidgets are already created in _init_ui
        self.table1Table.viewport().installEventFilter(self)
        self.table2Table.viewport().installEventFilter(self)

    # --------------------------------------------------------------------------
    # Logic (Ported from existing main.py)
    # --------------------------------------------------------------------------

    def on_tables_content_changed(self):
        if self.is_updating:
            return
            
        try:
            self.is_updating = True
            
            data1 = extract_data_from_table(self.table1Table)
            data2 = extract_data_from_table(self.table2Table)
            
            diff1 = generate_diff_report(data1, data2)
            diff2 = generate_diff_report(data2, data1)
            
            self.table1Table.apply_diff_report(diff1)
            self.table2Table.apply_diff_report(diff2)
            
            # Simple Summary Update
            count = len(diff1) + len(diff2)
            self.summaryLabel.setText(f"감지된 차이점: {count}건 (스타일링 갱신 완료)")
            
        finally:
            self.is_updating = False

    def eventFilter(self, source, event):
        if event.type() == QtCore.QEvent.Wheel and \
           event.modifiers() == QtCore.Qt.ShiftModifier:
            
            if source == self.table1Table.viewport():
                target = self.table2Table
                my_table = self.table1Table
            elif source == self.table2Table.viewport():
                target = self.table1Table
                my_table = self.table2Table
            else:
                return super().eventFilter(source, event)

            delta = event.angleDelta().y()
            v_bar = my_table.verticalScrollBar()
            t_bar = target.verticalScrollBar()
            
            step = v_bar.singleStep() * 3
            current_val = v_bar.value()
            
            if delta > 0:
                new_val = max(v_bar.minimum(), current_val - step)
            else:
                new_val = min(v_bar.maximum(), current_val + step)
            
            v_bar.setValue(new_val)
            t_bar.setValue(new_val)
            return True
            
        return super().eventFilter(source, event)

    def on_download_template(self):
        try:
            file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Save Template File", "template.xlsx", "Excel Files (*.xlsx)"
            )
            if not file_path:
                return

            saved_path = download_template_file(file_path)
            
            if QMessageBox.question(self, "완료", "템플릿이 저장된 폴더를 여시겠습니까?") == QMessageBox.Yes:
                os.startfile(saved_path.parent)
                
        except Exception as e:
            print(f"Template Download Error: {e}")
            QMessageBox.critical(self, "에러", "템플릿 다운로드 실패")

    def on_upload_file(self):
        try:
            file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "Select XLSX File", "", "Excel Files (*.xlsx)"
            )
            if not file_path:
                return
            
            self.fileLabel.setText(Path(file_path).name)

            self._set_tables_signal_blocked(True)
            
            try:
                make_table(self.table1Table, file_path, "Table1")
                make_table(self.table2Table, file_path, "Table2")
                self.on_tables_content_changed()
            finally:
                self._set_tables_signal_blocked(False)

        except Exception as e:
            print(f"File Upload Error: {e}")
            QMessageBox.critical(self, "에러", "파일 업로드 및 처리 실패")
    
    def on_download_result(self):
        try:
            file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Save Result File", "result.xlsx", "Excel Files (*.xlsx)"
            )
            if not file_path:
                return

            data1 = extract_data_from_table(self.table1Table)
            data2 = extract_data_from_table(self.table2Table)

            saved_path = export_to_excel(file_path, data1, data2)

            if QMessageBox.question(self, "완료", "결과 파일이 저장되었습니다.\n폴더를 여시겠습니까?") == QMessageBox.Yes:
                os.startfile(saved_path.parent)

        except Exception as e:
            print(f"Result Download Error: {e}")
            QMessageBox.critical(self, "에러", f"결과 다운로드 중 오류가 발생했습니다.\n{e}")

    def _set_tables_signal_blocked(self, blocked: bool):
        self.table1Table.blockSignals(blocked)
        self.table2Table.blockSignals(blocked)

# app/main.py
import sys
import os
from pathlib import Path

from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMessageBox
from app.ui.main_ui import Ui_MainWindow
from app.ui.widgets import MaterialTableWidget  # [신규] 커스텀 위젯
from app.excel import (
    download_template_file, 
    make_table, 
    generate_diff_report, 
    extract_data_from_table,
    export_to_excel
)

class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        # [신규] 기본 QTableWidget을 커스텀 MaterialTableWidget으로 교체
        self._setup_custom_tables()

        # 시그널 연결
        self.downloadButton.clicked.connect(self.on_download_template)
        self.uploadButton.clicked.connect(self.on_upload_file)
        self.downloadResultButton.clicked.connect(self.on_download_result)
        
        # 스크롤 동기화 등록
        self.table1Table.viewport().installEventFilter(self)
        self.table2Table.viewport().installEventFilter(self)

        # 상태 플래그
        self.is_updating = False

    def _setup_custom_tables(self):
        """UI 파일에서 생성된 기본 테이블을 MaterialTableWidget으로 교체합니다."""
        def replace_widget(old_widget):
            layout = old_widget.parentWidget().layout()
            # 같은 부모를 가지는 새 위젯 생성
            new_widget = MaterialTableWidget(old_widget.parentWidget())
            
            # 기존 속성 복사 (필요한 경우)
            new_widget.setObjectName(old_widget.objectName())
            
            # 레이아웃에서 교체
            layout.replaceWidget(old_widget, new_widget)
            
            # 기존 위젯 삭제
            old_widget.deleteLater()
            
            return new_widget

        self.table1Table = replace_widget(self.table1Table)
        self.table2Table = replace_widget(self.table2Table)

        # 변경 시그널 연결 (비교 로직 수행)
        self.table1Table.contentChanged.connect(self.on_tables_content_changed)
        self.table2Table.contentChanged.connect(self.on_tables_content_changed)

    def on_tables_content_changed(self):
        """테이블 내용이 변경되었을 때 (편집, 분할 등) 비교를 다시 수행합니다."""
        if self.is_updating:
            return
            
        try:
            self.is_updating = True
            
            # 1. 데이터 추출 (IngredientRow 리스트)
            data1 = extract_data_from_table(self.table1Table)
            data2 = extract_data_from_table(self.table2Table)
            
            # 2. 비교 리포트 생성 (Logic)
            diff1 = generate_diff_report(data1, data2)
            diff2 = generate_diff_report(data2, data1)
            
            # 3. 스타일 적용 (UI)
            self.table1Table.apply_diff_report(diff1)
            self.table2Table.apply_diff_report(diff2)
            
        finally:
            self.is_updating = False

    def eventFilter(self, source, event):
        """테이블 뷰포트의 이벤트를 가로챕니다 (스크롤 동기화)."""
        
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

            # 휠 방향 계산
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
        """템플릿 다운로드 버튼 핸들러"""
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
        """파일 업로드 버튼 핸들러"""
        try:
            file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
                self, "Select XLSX File", "", "Excel Files (*.xlsx)"
            )
            if not file_path:
                return
            
            self.fileLabel.setText(Path(file_path).name)

            self._set_tables_signal_blocked(True)
            
            try:
                # 테이블 생성
                make_table(self.table1Table, file_path, "Table1")
                make_table(self.table2Table, file_path, "Table2")
                
                # 초기 비교 수행
                self.on_tables_content_changed()
                
            finally:
                self._set_tables_signal_blocked(False)

        except Exception as e:
            print(f"File Upload Error: {e}")
            QMessageBox.critical(self, "에러", "파일 업로드 및 처리 실패")
    
    
    def on_download_result(self):
        """결과 다운로드 핸들러"""
        try:
            # 1. 저장 경로 확인
            file_path, _ = QtWidgets.QFileDialog.getSaveFileName(
                self, "Save Result File", "result.xlsx", "Excel Files (*.xlsx)"
            )
            if not file_path:
                return

            # 2. 데이터 추출
            data1 = extract_data_from_table(self.table1Table)
            data2 = extract_data_from_table(self.table2Table)

            # 3. 엑셀 내보내기
            saved_path = export_to_excel(file_path, data1, data2)

            # 4. 완료 알림 및 폴더 열기
            if QMessageBox.question(self, "완료", "결과 파일이 저장되었습니다.\n폴더를 여시겠습니까?") == QMessageBox.Yes:
                os.startfile(saved_path.parent)

        except Exception as e:
            print(f"Result Download Error: {e}")
            QMessageBox.critical(self, "에러", f"결과 다운로드 중 오류가 발생했습니다.\n{e}")

    def _set_tables_signal_blocked(self, blocked: bool):
        """두 테이블의 시그널 차단 여부를 일괄 설정"""
        self.table1Table.blockSignals(blocked)
        self.table2Table.blockSignals(blocked)


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

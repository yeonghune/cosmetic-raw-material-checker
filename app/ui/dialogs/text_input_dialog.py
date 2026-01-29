from PyQt5 import QtWidgets, QtCore, QtGui

class TextInputDialog(QtWidgets.QDialog):
    def __init__(self, title="성분 입력", parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(500, 400)
        self._init_ui()

    def _init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(12)

        # Instruction Label
        label = QtWidgets.QLabel("성분 리스트를 아래에 붙여넣으세요:")
        layout.addWidget(label)

        # Text Edit Area
        self.textEdit = QtWidgets.QTextEdit()
        self.textEdit.setPlaceholderText("예:\nWater,\nGlycerin,\n1,2-Hexanediol")
        layout.addWidget(self.textEdit)

        # Image Upload Button (Placeholder)
        self.btnImageUpload = QtWidgets.QPushButton("📷 이미지 업로드 (OCR)")
        self.btnImageUpload.setFixedHeight(36)
        self.btnImageUpload.clicked.connect(self._on_image_upload)
        layout.addWidget(self.btnImageUpload)

        # Buttons (OK / Cancel)
        buttonBox = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel
        )
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)
        layout.addWidget(buttonBox)

    def _on_image_upload(self):
        QtWidgets.QMessageBox.information(
            self, 
            "안내", 
            "이미지 인식 기능은 현재 개발 중입니다.\n(Coming Soon)"
        )

    def get_text(self):
        return self.textEdit.toPlainText()

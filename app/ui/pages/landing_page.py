from PyQt5 import QtWidgets, QtCore, QtGui

class LandingPage(QtWidgets.QWidget):
    """
    애플리케이션의 시작 화면입니다.
    사용자가 원하는 기능을 선택할 수 있습니다.
    """
    
    # 페이지 전환 요청 시그널
    navigate_to = QtCore.pyqtSignal(str) # 'checker', 'new_feature' 등

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()
        self._setup_connections()

    def _init_ui(self):
        self.setObjectName("LandingPage")
        
        # Main Layout (Centered)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.setAlignment(QtCore.Qt.AlignCenter)
        self.layout.setSpacing(30)
        
        # Title
        self.titleLabel = QtWidgets.QLabel("Cosmetic Raw Material Checker")
        title_font = QtGui.QFont("Arial", 24, QtGui.QFont.Bold)
        self.titleLabel.setFont(title_font)
        self.titleLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.layout.addWidget(self.titleLabel)
        
        # Subtitle
        self.subtitleLabel = QtWidgets.QLabel("사용할 기능을 선택해주세요")
        subtitle_font = QtGui.QFont("Arial", 12)
        self.subtitleLabel.setFont(subtitle_font)
        self.subtitleLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.layout.addWidget(self.subtitleLabel)
        
        # Buttons Container
        self.buttonContainer = QtWidgets.QWidget()
        self.buttonLayout = QtWidgets.QHBoxLayout(self.buttonContainer)
        self.buttonLayout.setSpacing(20)
        
        # Button 1: Checker
        self.checkerButton = self._create_card_button(
            "📝", "원료 검증기", "두 엑셀 테이블을 비교하고\n차이점을 분석합니다."
        )
        self.buttonLayout.addWidget(self.checkerButton)
        
        # Button 2: Text Comparator
        self.textComparatorButton = self._create_card_button(
            "📋", "성분 텍스트 비교", "텍스트 목록을 직접 입력하여\n빠르게 비교합니다."
        )
        self.buttonLayout.addWidget(self.textComparatorButton)
        
        self.layout.addWidget(self.buttonContainer)

    def _create_card_button(self, icon_text, title_text, desc_text):
        """카드 형태의 버튼을 생성합니다."""
        button = QtWidgets.QPushButton()
        button.setFixedSize(220, 180)
        button.setCursor(QtCore.Qt.PointingHandCursor)
        
        # Simple Layout inside button handling with text
        # PyQt PushButton can hold text, but for multi-line styled text, 
        # using a simple text set is easier than custom painting for now.
        # We will use HTML for rich text formatting inside the button.
        
        content = f"""
        <div style='text-align: center;'>
            <p style='font-size: 40px; margin-bottom: 10px;'>{icon_text}</p>
            <p style='font-size: 16px; font-weight: bold; margin-bottom: 5px;'>{title_text}</p>
            <p style='font-size: 12px; color: #666;'>{desc_text}</p>
        </div>
        """
        button.setText(title_text) # Fallback / Accessibility
        
        # To actually render complex HTML nicely in a standard QPushButton is tricky on some styles.
        # Let's use a QToolButton or just styling.
        # For simplicity and reliability, let's just create a custom widget that ACTS like a button 
        # or just use text with newlines.
        
        button.setText(f"{icon_text}\n\n{title_text}\n\n{desc_text}")
        
        # Style Sheet for Card look
        button.setStyleSheet("""
            QPushButton {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 12px;
                padding: 10px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #f9f9f9;
                border: 1px solid #bbb;
            }
            QPushButton:pressed {
                background-color: #eee;
            }
        """)
        return button

    def _setup_connections(self):
        self.checkerButton.clicked.connect(lambda: self.navigate_to.emit('checker'))
        self.textComparatorButton.clicked.connect(lambda: self.navigate_to.emit('text_comparator'))

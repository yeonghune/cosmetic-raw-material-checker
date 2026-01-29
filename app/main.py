# app/main.py
import sys
from PyQt5 import QtWidgets

# Pages
from app.ui.pages.landing_page import LandingPage
from app.ui.pages.checker_page import CheckerPage
from app.ui.pages.text_comparator_page import TextComparatorPage

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cosmetic Raw Material Checker")
        self.resize(1000, 800) # Slightly larger for comfortable view

        # Central Widget: Stacked Widget (Page Container)
        self.stacked_widget = QtWidgets.QStackedWidget()
        self.setCentralWidget(self.stacked_widget)

        # Page 1: Landing Page
        self.landing_page = LandingPage()
        self.stacked_widget.addWidget(self.landing_page)

        # Page 2: Checker Page
        self.checker_page = CheckerPage()
        self.stacked_widget.addWidget(self.checker_page)

        # Page 3: Text Comparator Page
        self.text_comparator_page = TextComparatorPage()
        self.stacked_widget.addWidget(self.text_comparator_page)

        # Signal Connections
        self.landing_page.navigate_to.connect(self.on_navigate_to)
        self.checker_page.navigate_home.connect(self.go_to_home)
        self.text_comparator_page.navigate_home.connect(self.go_to_home)

    def on_navigate_to(self, page_name: str):
        """랜딩 페이지에서의 내비게이션 요청 처리"""
        if page_name == 'checker':
            self.stacked_widget.setCurrentWidget(self.checker_page)
        elif page_name == 'text_comparator':
            self.stacked_widget.setCurrentWidget(self.text_comparator_page)
        elif page_name == 'new_feature':
            QtWidgets.QMessageBox.information(
                self, 
                "준비 중", 
                "이 기능은 아직 개발 중입니다.\n(Coming Soon)"
            )
        else:
            print(f"Unknown page: {page_name}")

    def go_to_home(self):
        """홈(랜딩 페이지)으로 복귀"""
        self.stacked_widget.setCurrentWidget(self.landing_page)


def main():
    app = QtWidgets.QApplication(sys.argv)
    
    # Global Font Setting (Optional polish)
    font = app.font()
    font.setFamily("Arial") 
    app.setFont(font)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

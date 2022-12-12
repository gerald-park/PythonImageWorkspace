from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5 import uic
import sys

class UI(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('mybutton.ui', self)

        # fine our widget
        button = self.findChild(QPushButton,'pushButton')
        button.clicked.connect(self.clicked_btn)

    def clicked_btn(self):
        print("button Clicked")


app = QApplication([])
window = UI()
window.show()
app.exec_()
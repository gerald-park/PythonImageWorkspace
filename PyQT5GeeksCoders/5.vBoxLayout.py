from PyQt5.QtWidgets import QApplication, QWidget, QPushButton
from PyQt5 import uic
import sys

class UI(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('5.vBoxLayout.ui', self)


app = QApplication([])
window = UI()
window.show()
app.exec_()
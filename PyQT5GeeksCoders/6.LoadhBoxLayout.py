from PyQt5.QtWidgets import QApplication, QWidget, QHBoxLayout,QPushButton
from PyQt5 import uic
import sys
from PyQt5.QtGui import QIcon

class Window(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('6.HBoxLayout.ui', self)

app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec_())
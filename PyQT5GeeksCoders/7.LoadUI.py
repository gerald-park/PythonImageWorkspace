from PyQt5.QtWidgets import QApplication,QWidget, QGridLayout, QPushButton
import sys
from PyQt5.QtGui import QIcon
from PyQt5 import uic

class Window(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi('7.gridlayout.ui', self)

app = QApplication(sys.argv)

window = Window()
window.show()
sys.exit(app.exec_())

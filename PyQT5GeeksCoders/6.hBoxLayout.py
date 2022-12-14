from PyQt5.QtWidgets import QApplication, QWidget, QHBoxLayout,QPushButton
from PyQt5 import uic
import sys
from PyQt5.QtGui import QIcon

class Window(QWidget):
    def __init__(self):
        super().__init__()
        # uic.loadUi('5.vBoxLayout.ui', self)

        self.setGeometry(200, 200, 400, 400)
        self.setWindowTitle("PtQt5 HBoxLayout")
        self.setWindowIcon(QIcon('python.png'))

        hbox = QHBoxLayout()
        btn1 = QPushButton("Click one")
        btn2 = QPushButton("Click Two")
        btn3 = QPushButton("Click Three")
        btn4 = QPushButton("Click Four")

        hbox.addWidget(btn1)
        hbox.addWidget(btn2)
        hbox.addWidget(btn3)
        hbox.addWidget(btn4)

        self.setLayout(hbox)


app = QApplication(sys.argv)
window = Window()
window.show()
sys.exit(app.exec_())
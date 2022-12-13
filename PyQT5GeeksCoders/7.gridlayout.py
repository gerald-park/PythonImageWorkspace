from PyQt5.QtWidgets import QApplication,QWidget, QGridLayout, QPushButton
import sys
from PyQt5.QtGui import QIcon

class Window(QWidget):
    def __init__(self):
        super().__init__()
        
        self.setGeometry(200,200,400,400)
        self.setWindowTitle('7.GridLayout')

        grid = QGridLayout()

        btn1 = QPushButton("Click one")
        btn2 = QPushButton("Click Two")
        btn3 = QPushButton("Click Three")
        btn4 = QPushButton("Click Four")
        btn5 = QPushButton("Click Five")
        btn6 = QPushButton("Click Six")
        btn7 = QPushButton("Click Seven")
        btn8 = QPushButton("Click Eight")

        grid.addWidget(btn1, 0 , 0)
        grid.addWidget(btn2, 0 , 1)
        grid.addWidget(btn3, 0 , 2)
        grid.addWidget(btn4, 0 , 3)
        grid.addWidget(btn5, 1 , 0)
        grid.addWidget(btn6, 1 , 1)
        grid.addWidget(btn7, 1 , 2)
        grid.addWidget(btn8, 1 , 3)

        self.setLayout(grid)

app = QApplication(sys.argv)

window = Window()
window.show()
sys.exit(app.exec_())





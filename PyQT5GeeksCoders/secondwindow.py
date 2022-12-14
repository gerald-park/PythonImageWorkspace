import sys
from PyQt5.QtWidgets import QApplication, QWidget, QDialog, QMainWindow,QPushButton
from PyQt5.QtGui import QIcon

class WindowExample(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(200, 200, 400, 300)
        self.setWindowTitle("GeeksCoders.com")
        self.setFixedHeight(300)
        self.setFixedWidth(400)
        self.setWindowIcon((QIcon('python.png')))
        # self.setStyleSheet('background-color:green')
        self.create_buttons()
        self.show()

    def create_buttons(self):
        btn1 = QPushButton("Click me", self)
        btn1.setGeometry(100, 100, 100, 100)
        # btn1.setText('our second text')
        btn1.setIcon(QIcon("python.png"))
        btn1.setStyleSheet('color:red')
        btn1.clicked.connect(self.clicked_btn)

        btn2 = QPushButton("click two", self)
        btn2.setGeometry(200, 100, 100, 100)
        btn2.setStyleSheet('color:white;background-color:purple')
        btn2.clicked.connect(self.second_clicked_btn)

    def clicked_btn(self):
        print("button one clicked!")

    def second_clicked_btn(self):
        print("second button one clicked!")

app = QApplication(sys.argv)
window = WindowExample()
sys.exit(app.exec_())
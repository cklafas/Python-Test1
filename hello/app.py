import sys
from PyQt5.QtWidgets import QApplication, Qwidget

def app():
    my_app = QApplication(sys.argv)
    w = Qwidget()
    w.setWindowTitle("Test Window")
    w.show()
    sys.exit(my_app.exec_())

app()



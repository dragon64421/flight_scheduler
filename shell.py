from MainWindow import MainWindow
import sys
import PyQt5.QtWidgets as qtwid

app = qtwid.QApplication(sys.argv)
mw = MainWindow()
mw.show()
app.exec_()
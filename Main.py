import logging
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLabel, QLineEdit, QGridLayout
from QPlainTextEditLogger import QPlainTextEditLogger
from AutoSendMSG import AutoSendMSG
from PyQt5.QtCore import QThread

class SendMsgApp(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Send_MSG')
        
        grid = QGridLayout()
        grid.setSpacing(10)

        self.Frash_sp_label = QLabel('새로고침 속도(분)', self)
        self.Frash_sp_input = QLineEdit(self)
        grid.addWidget(self.Frash_sp_label, 1, 0)
        grid.addWidget(self.Frash_sp_input, 1, 1)
        
        self.button = QPushButton('실행', self)
        self.button.clicked.connect(self.on_click)

        layout = QVBoxLayout()
        layout.addLayout(grid)
        layout.addWidget(self.button)
        
        self.logBrowser = QPlainTextEditLogger(self)
        logging.getLogger().addHandler(self.logBrowser)
        logging.getLogger().setLevel(logging.INFO)

        layout.addWidget(self.logBrowser.widget)

        self.setLayout(layout)
        self.setGeometry(100, 100, 300, 200)
        self.isdebug = False
        self.thread = QThread()
        logging.info(f'프로그램이 정상적으로 실행되었습니다.')

    def on_click(self):
        logging.info(f'프로그램 개시')
        Frash_sp_input = self.Frash_sp_input.text()
        self.crawling_thread = AutoSendMSG(self.isdebug,
                                           Frash_sp_input
                                                )
        self.crawling_thread.moveToThread(self.thread)
        self.crawling_thread.ReturnError.connect(self.ShowError)
        self.crawling_thread.returnWarning.connect(self.showlog)
        self.thread.started.connect(self.crawling_thread.run)
        self.thread.start()

    def showlog(self,str):
        logging.warning(f'{str}')

    def ShowError(self, str):
        logging.warning(f'PGM ERROR - {str}')
        self.thread.quit()
        self.thread.wait()
        self.Conversion_thread = None
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = SendMsgApp()
    ex.show()
    sys.exit(app.exec_())

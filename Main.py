import logging
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout
from QPlainTextEditLogger import QPlainTextEditLogger
from AutoSendMSG import AutoSendMSG
from PyQt5.QtCore import QThread

class SendMsgApp(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Send_MSG')

        self.button = QPushButton('실행', self)
        self.button.clicked.connect(self.on_click)

        layout = QVBoxLayout()
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
        self.crawling_thread = AutoSendMSG(self.isdebug
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

import pygame
import os
import sys
import parsers
# import easygui
import monitor
import time
# Импортируем наш интерфейс из файла
from mainwindow import *
import time
# import win32com.client
import pythoncom
import pywintypes

from PyQt5.QtCore import QObject, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QApplication, QPushButton, QTextEdit, QVBoxLayout, QWidget, QMainWindow,QInputDialog, QLineEdit
from PyQt5.QtMultimedia import QSound

from monitor import do_scenario


def trap_exc_during_debug(*args):
    # when app raises uncaught exception, print info
    print(args)


# install exception hook: without this, uncaught exception would cause application to exit
sys.excepthook = trap_exc_during_debug


class MyWin(QMainWindow):
    def __init__(self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.file_log = "log.txt"
        self.file_last_docs = 'last_docs.txt'
        self.file_config = "config.conf"
        self.sleep_time = 5
        self.mail_password = ''
        # self.notify = Notifier
        self._qSound_alert = QSound("SilentHill.wav", self)


        # mailTo = 'grafinia@veb.ru;zagreba@veb.ru'
        # self.mail_password = easygui.passwordbox('Введите пароль от Lotus Notes: ')
        # Здесь прописываем событие нажатия на кнопку
        self.ui.pushButton.clicked.connect(self.start_function)
        QThread.currentThread().setObjectName('main')  # threads can be named, useful for log output
        self.__workers_done = None
        self.__threads = []
        self.refresh_log_text()

    def get_password(self):
        text, ok = QInputDialog.getText(None, "Attention", "Password?", QLineEdit.Password)
        if ok and text:
            self.mail_password = text

    def start_function(self):
        # Функция, которая выполняется по кнопке "Старт"
        self.ui.pushButton.setDisabled(True)
        probes = parsers.get_config(self.file_config)  # List of the config file contains server, database, view
        print(probes)
        self.get_password()


        # monitor.do_scenario(probes)
        # time.sleep(self.sleep_time)
        self.__workers_done = 0
        worker = Worker(1, probes, self.file_log, self.file_last_docs, self.sleep_time, self.mail_password)
        thread = QThread()
        thread.setObjectName('thread_' + str(1))
        self.__threads.append((thread, worker))  # need to store worker too otherwise will be gc'd
        worker.moveToThread(thread)
        # get progress messages from worker:
        # worker.sig_step.connect(self.on_worker_step)
        worker.sig_done.connect(self.on_worker_done)
        worker.sig_msg.connect(self.worker_message)
        worker.sig_alert.connect(self.new_alert)

        # control worker:
        # self.sig_abort_workers.connect(worker.abort)

        # get read to start worker:
        # self.sig_start.connect(worker.work)  # needed due to PyCharm debugger bug (!); comment out next line
        thread.started.connect(worker.work)
        thread.start()  # this will emit 'started' and start thread's event loop

    # @pyqtSlot(int, str)
    # def on_worker_step(self, worker_id: int, data: str):
    #     self.log.append('Worker #{}: {}'.format(worker_id, data))
    #     self.progress.append('{}: {}'.format(worker_id, data))
    #
    @pyqtSlot(str)
    def worker_message(self, message):
        print('функция worker_message, что-то не так' + message)
        # self.ui.pushButton.setDisabled(False)


    @pyqtSlot(int)
    def on_worker_done(self, worker_id):
        print('функция on_worker_done. Всё ок,')
        # print(str(worker_id))
        # self.ui.pushButton.setDisabled(False)
        # self.refresh_log_text()


    @pyqtSlot(int)
    def new_alert(self, worker_id):
        print('функция new_alert, обновим лог')
        self.refresh_log_text()
        self.ui.pushButton.setDisabled(False)
        self.play_sound()
        # pygame.mixer.init()
        # pygame.mixer.music.load("SilentHill.mp3")
        # pygame.mixer.music.play()
        # while pygame.mixer.music.get_busy():
        #     pygame.time.Clock().tick(1)
        #     pygame.time.wait(1)
        #     pygame.mixer.music.stop()

    def play_sound(self):
        print('играем?')
        self._qSound_alert.play()

    #
    # @pyqtSlot()
    # def abort_workers(self):
    #     self.sig_abort_workers.emit()
    #     self.log.append('Asking each worker to abort')
    #     for thread, worker in self.__threads:  # note nice unpacking by Python, avoids indexing
    #         thread.quit()  # this will quit **as soon as thread event loop unblocks**
    #         thread.wait()  # <- so you need to wait for it to *actually* quit
    #
    #     # even though threads have exited, there may still be messages on the main thread's
    #     # queue (messages that threads emitted before the abort):
    #     self.log.append('All threads exited')

    def refresh_log_text(self):
        # Парсит лог-файл и выводит на экран в TextEdit поле
        # file_log = "log.txt"
        self.ui.textEdit.setText(parsers.log_parser(self.file_log))


class Worker(QObject):
    """
    Must derive from QObject in order to emit signals, connect slots to other signals, and operate in a QThread.
    """

    sig_step = pyqtSignal(int, str)  # worker id, step description: emitted every step through work() loop
    sig_done = pyqtSignal(int)  # worker id: emitted at end of work()
    sig_alert = pyqtSignal(int)  # worker id: emitted at end of work()
    sig_msg = pyqtSignal(str)  # message to be shown to user

    def __init__(self, id: int, probes, file_log, file_last_docs, sleep_time, mail_password):
        super().__init__()
        self.__id = id
        self.probes = probes
        self.file_log = file_log
        self.file_last_docs = file_last_docs
        self.sleep_time = sleep_time
        self.__abort = False
        self.mail_password = mail_password


    @pyqtSlot()
    def work(self):
        """
        Pretend this worker method does work that takes a long time. During this time, the thread's
        event loop is blocked, except if the application's processEvents() is called: this gives every
        thread (incl. main) a chance to process events, which in this sample means processing signals
        received from GUI (such as abort).
        """
        thread_name = QThread.currentThread().objectName()
        thread_id = int(QThread.currentThreadId())  # cast to int() is necessary
        # self.sig_msg.emit('Running worker #{} from thread "{}" (#{})'.format(self.__id, thread_name, thread_id))
        # print(self.probes)
        pythoncom.CoInitialize()
        # def do_scenario(probes, file_last_docs, file_log, mail_password, sleep_time):
        while True:
            # notesSession = win32com.client.Dispatch('Lotus.NotesSession')

            # notesSession.Initialize(mail_password)
            alarm_returned = monitor.do_scenario(self.probes, self.file_last_docs, self.file_log,
                                                 self.mail_password, self.sleep_time)
            print(alarm_returned)

            if alarm_returned is False:
                print("alarm returned False")
                self.sig_done.emit(self.__id)
            elif alarm_returned is True:
                print("alarm returned True")
                self.sig_alert.emit(self.__id)
            else:
                print("alarm returned something else")
                self.sig_msg.emit('что-то пошло не так =(')

            time.sleep(self.sleep_time)


        # for step in range(100):
        #     time.sleep(0.1)
        #     self.sig_step.emit(self.__id, 'step ' + str(step))
        #
        #     # check if we need to abort the loop; need to process events to receive signals;
        #     app.processEvents()  # this could cause change to self.__abort
        #     if self.__abort:
        #         # note that "step" value will not necessarily be same for every thread
        #         self.sig_msg.emit('Worker #{} aborting work at step {}'.format(self.__id, step))
        #         break
        # self.sig_msg.emit('Hello from thread')
        # self.sig_done.emit(self.__id)

        # def abort(self):
        #     self.sig_msg.emit('Worker #{} notified to abort'.format(self.__id))
        #     self.__abort = True


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())

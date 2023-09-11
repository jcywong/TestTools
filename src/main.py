import sys
import threading

from PySide6.QtCore import QFile, QIODevice, QObject, Signal
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QMessageBox, QFileDialog, QLineEdit, QComboBox, \
    QProgressBar, QCheckBox
from PySide6.QtUiTools import QUiLoader

from comm import *


class SignalStore(QObject):
    # 定义信号
    progress_update = Signal(int)
    download_state = Signal(bool)
    show_message = Signal(str, str)


so = SignalStore()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        so.progress_update.connect(self.setProgress)
        so.download_state.connect(self.update_download_state)
        so.show_message.connect(self.show_MessageBox)
        self.filename = []

        self.filePath = None
        self.ui_file_name = "../ui/main.ui"
        self.ui_file = QFile(self.ui_file_name)
        if not self.ui_file.open(QIODevice.ReadOnly):
            print(f"Cannot open {self.ui_file_name}: {self.ui_file.errorString()}")
            sys.exit(-1)
        self.window = QUiLoader().load(self.ui_file)
        self.ui_file.close()
        if not self.window:
            print(QUiLoader().errorString())
            sys.exit(-1)

        # 将UI中的控件添加到主窗口
        self.setCentralWidget(self.window)

        # 选择文件地址
        self.btn_choseFilePath = self.window.findChild(QPushButton, "btn_choseFilePath")
        self.lineEdit_save_path = self.window.findChild(QLineEdit, "lineEdit_save_path")
        self.btn_choseFilePath.clicked.connect(self.chose_file_path)

        # 选择ics/icc
        self.comboBox_Edition = self.window.findChild(QComboBox, "comboBox_Edition")
        self.comboBox_icc_model = self.window.findChild(QComboBox, "comboBox_icc_model")
        self.comboBox_ver = self.window.findChild(QComboBox, "comboBox_ver")
        self.comboBox_icc_model.addItems(['LITE', 'PRO', 'TURBO'])
        self.comboBox_ver.addItems([" ", 'v1.2', 'v1.3'])
        self.comboBox_Edition.addItems(['Debug', 'Release'])
        self.checkBox_ics = self.window.findChild(QCheckBox, "checkBox_ics")
        self.checkBox_icc = self.window.findChild(QCheckBox, "checkBox_icc")
        self.comboBox_Edition.currentIndexChanged.connect(self.SelectionChange_comboBox_Edition)

        # 下载
        self.btn_download = self.window.findChild(QPushButton, "btn_download")
        self.btn_download.clicked.connect(self.download_soft)
        self.progressBar_download = self.window.findChild(QProgressBar, "progressBar_download")
        self.progressBar_download.setRange(0, 2)
        self.downloading = False

        # 运行ICS Studio
        self.btn_run = self.window.findChild(QPushButton, "btn_run")
        self.btn_run.clicked.connect(self.run_ICSStudio)

    def chose_file_path(self):
        self.filePath = QFileDialog.getExistingDirectory(self.window, "选择存储路径")
        if isinstance(self.lineEdit_save_path, QLineEdit):
            self.lineEdit_save_path.setText(self.filePath)

    def SelectionChange_comboBox_Edition(self):
        if isinstance(self.comboBox_Edition, QComboBox):
            if self.comboBox_Edition.currentText() == "Release":
                self.comboBox_ver.setEnabled(True)
            else:
                self.comboBox_ver.setCurrentText(" ")
                self.comboBox_ver.setEnabled(False)

    def update_download_state(self):
        if self.downloading:
            self.comboBox_icc_model.setEnabled(False)
            self.comboBox_Edition.setEnabled(False)
            self.btn_choseFilePath.setEnabled(False)
            if self.comboBox_Edition.currentText() == "Release":
                self.comboBox_ver.setEnabled(False)
        else:
            self.comboBox_icc_model.setEnabled(True)
            self.comboBox_Edition.setEnabled(True)
            self.btn_choseFilePath.setEnabled(True)
            if self.comboBox_Edition.currentText() == "Release":
                self.comboBox_ver.setEnabled(True)


    def show_MessageBox(self, message, message_type):
        if message_type == "warning":
            return QMessageBox.warning(self.window, "警告", message)
        elif message_type == "question":
            return QMessageBox.question(self.window, '确认', message)

    def download_soft(self):
        def workerThreadFunc():
            self.downloading = True
            model = self.comboBox_icc_model.currentText()
            edition = self.comboBox_Edition.currentText()
            ver = self.comboBox_ver.currentText()
            icc_isChecked = self.checkBox_icc.isChecked()
            ics_isChecked = self.checkBox_ics.isChecked()
            file_save_path = self.filePath

            if not file_save_path:
                so.show_message.emit("请设置文件保存地址", "warning")
                self.downloading = False
                so.download_state.emit(self.downloading)
                return
            else:
                so.download_state.emit(self.downloading)

            self.filename.clear()
            if ics_isChecked and not icc_isChecked and edition == "Debug":
                self.filename.append(get_latest_filename())
            elif ics_isChecked and icc_isChecked and edition == "Debug":
                self.filename.append(get_latest_filename())
                self.filename.append(get_latest_filename(soft_type='ICC', model=model))
            elif icc_isChecked and not ics_isChecked and edition == "Debug":
                self.filename.append(get_latest_filename(soft_type='ICC', model=model))
            elif ics_isChecked and not icc_isChecked and edition == "Release":
                self.filename.append(get_latest_filename(edition="Release", ver=ver))
            elif ics_isChecked and icc_isChecked and edition == "Release":
                self.filename.append(get_latest_filename(edition="Release", ver=ver))
                self.filename.append(get_latest_filename(soft_type='ICC', model=model, edition="Release", ver=ver))
            elif icc_isChecked and not ics_isChecked and edition == "Release":
                self.filename.append(get_latest_filename(soft_type='ICC', model=model, edition="Release", ver=ver))
            elif not ics_isChecked and not icc_isChecked :
                so.show_message.emit("请勾选下载软件", "warning")
                self.downloading = False
                so.download_state.emit(self.downloading)
                return

            so.progress_update.emit(1)

            for name in self.filename:
                download_file(name, file_save_path, soft_type=name[:3], edition=edition)
                unzip_file(self.filePath, name)

            self.downloading = False

            so.progress_update.emit(2)

            so.download_state.emit(self.downloading)

        if self.downloading:
            QMessageBox.warning(
                self.window,
                '警告', '任务进行中，请等待完成')
            return

        worker = threading.Thread(target=workerThreadFunc)
        worker.start()

    def run_ICSStudio(self):
        if self.filename:
            for file in self.filename:
                if file[:3] == "ICS" and file[-9:-4] == "debug":
                    open_ics(self.filePath + "/" + file[:-4] + "/Debug")
                elif file[:3] == "ICS" and file[-11:-4] == "Release":
                    open_ics(self.filePath + "/" + file[:-4] + "/Release")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")

    def setProgress(self, value):
        self.progressBar_download.setValue(value)

    def execute(self):
        print("------------------------------\n"
              "上传下载测试工具\n"
              "注意：\n"
              "1.需搭配ICS Studio使用,ics可放置soft目录下\n"
              "2.工程文件必需要放置json目录下\n"
              "3.工程需提前配置好对应的PLC型号\n"
              "4.需要打开window拓展名\n"
              "------------------------------\n")
        test_type = input("请选择测试类型download / upload\n").replace(' ', "")
        path = input("请输入ICS Studio 目录地址（eg:D:\TestTools\soft\Release）\n").replace(' ', "")
        file_name = input("请输入工程文件名称(eg:test.json)\n").replace(' ', "")
        plc_ip = input("请输入plcIP地址(eg:192.168.1.211)\n").replace(' ', "")
        test_times = input("请输入上传下载次数\n").replace(' ', "")
        time_gap = input("请输入上传下载的间隔，单位秒\n").replace(' ', "")

        print("------------------------------\n")
        print("开始测试环境初始化")
        # 打开软件
        open_ics(path)
        set_language(model=2)
        # 打开工程
        open_json_file(file_name)
        # 连接plc
        net_path(plc_ip, model=2)
        # 登出
        logout()
        print("------------------------------\n")
        print("开始测试")

        for i in range(int(test_times)):
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:开始第{i + 1}次{test_type}')
            if test_type == "download":
                try:
                    download()
                    print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成第{i + 1}次{test_type}')
                    logout()
                except LookupError:
                    print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:第{i + 1}次{test_type}失败')
                    break
            elif test_type == "upload":
                try:
                    upload()
                    print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成第{i + 1}次{test_type}')
                    logout()
                except LookupError:
                    print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:第{i + 1}次{test_type}失败')
                    break
            else:
                print("测试类型错误")
            time.sleep(int(time_gap))
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成第等待间隔{int(time_gap)}秒')

        print("结束测试")
        print("------------------------------\n")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

    # print("按任意键退出...")
    # input()  # 等待用户输入
    # print("程序已退出")

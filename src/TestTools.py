import json
import sys
import threading

from PySide6.QtCore import QFile, QIODevice, QObject, Signal, QRegularExpression, Qt, QThread
from PySide6.QtGui import QRegularExpressionValidator, QIcon, QAction
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QMessageBox, QFileDialog, QLineEdit, QComboBox, \
    QProgressBar, QCheckBox, QRadioButton, QStatusBar, QTabWidget, QDialog, QLabel, QVBoxLayout, QMenuBar, QMenu

from comm import *

# 定义版本号
VERSION = "1.3.7"


class SignalStore(QObject):
    # 定义信号
    progress_update = Signal(int)
    download_state = Signal(bool)
    execute_state = Signal(bool)
    show_message = Signal(str, str)
    show_status = Signal(str)


# 信号类
so = SignalStore()

# 配置文件名
config_file = 'config.json'


class MemoryThread(QThread):
    memory_info_updated = Signal(float, float, float)  # 信号，用于传递内存信息

    def run(self):
        while not self.isInterruptionRequested():
            memory_info = psutil.virtual_memory()
            ICS_memory = monitor_process_memory()
            used_memory = memory_info.used / (1024 ** 2)  # 转换为MB
            total_memory = memory_info.total / (1024 ** 2)  # 转换为MB

            self.memory_info_updated.emit(ICS_memory, used_memory, total_memory)
            self.sleep(1)  # 每隔1秒更新一次内存信息

    def stop(self):
        self.requestInterruption()
        self.quit()
        self.wait()


class MemoryMonitorDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Memory Monitor')
        self.resize(200, 80)

        self.label = QLabel(self)

        layout = QVBoxLayout(self)
        layout.addWidget(self.label)
        self.setLayout(layout)

        self.label.setAlignment(Qt.AlignCenter)
        self.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint | Qt.WindowStaysOnTopHint)

        self.memory_thread = MemoryThread(self)
        self.memory_thread.memory_info_updated.connect(self.update_memory_info)
        self.memory_thread.start()

    def closeEvent(self, event):
        self.memory_thread.stop()
        event.accept()

    # def update_memory_info(self):
    #     memory_info = psutil.virtual_memory()
    #     ICS_memory = monitor_process_memory()
    #     used_memory = memory_info.used / (1024 ** 2)  # 转换为MB
    #     total_memory = memory_info.total / (1024 ** 2)  # 转换为MB
    #
    #     if ICS_memory:
    #         self.label.setText(
    #             f'ICS Studio: {ICS_memory:.2f} MB   {(ICS_memory / total_memory) * 100:.2f}%\nUsed: {used_memory:.2f} MB    {(used_memory / total_memory) * 100:.2f}%\n')
    #         self.start_timer()
    #     else:
    #         self.label.setText(
    #             f'ICS Studio:not found\nUsed: {used_memory:.2f} MB {(used_memory / total_memory) * 100:.2f}%\n')
    #         self.start_timer(60000)

    def update_memory_info(self, ICS_memory, used_memory, total_memory):
        if ICS_memory:
            self.label.setText(
                f'ICS Studio: {ICS_memory:.2f} MB   {(ICS_memory / total_memory) * 100:.2f}%\nUsed: {used_memory:.2f} MB    {(used_memory / total_memory) * 100:.2f}%\n')
        else:
            self.label.setText(
                f'ICS Studio:not found\nUsed: {used_memory:.2f} MB {(used_memory / total_memory) * 100:.2f}%\n')


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        so.progress_update.connect(self.setProgress)
        so.download_state.connect(self.update_download_state)
        so.show_message.connect(self.show_MessageBox)
        so.execute_state.connect(self.update_execute_state)
        so.show_status.connect(self.update_status)

        self.filename = {}  # jcywong 修改为字典 增加 icm 2023/11/13
        self.network = "LAN"

        self.filePath = None
        self.ui_file_name = "main.ui"
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

        # 添加menu jcywong add 2024/1/4
        self.menubar = self.window.findChild(QMenuBar, "menuBar")

        self.tool_menu = self.window.findChild(QMenu, "tool")
        self.memory_action = self.window.findChild(QAction, "action_memory")
        self.memory_action.triggered.connect(self.open_memory_monitor)
        # self.memory_monitor_dialog = MemoryMonitorDialog()  # 将对话框作为成员变量
        self.memory_monitor_dialog = None

        self.help_menu = self.window.findChild(QMenu, "help")
        self.ver_action = self.window.findChild(QAction, "action_ver")
        self.ver_action.triggered.connect(self.show_version)

        # tab选项  jcywong add 2023/11/13
        self.tab_tabMenu = self.window.findChild(QTabWidget, "tab")
        # self.tab_general = self.window.findChild(QWidget, "tab_general")
        # self.tab_upload_download = self.window.findChild(QWidget, "tab_upload_download")

        self.tab_tabMenu.removeTab(1)

        # 选择文件地址
        self.lineEdit_save_path = self.window.findChild(QLineEdit, "lineEdit_save_path")
        self.btn_choseFilePath = self.window.findChild(QPushButton, "btn_choseFilePath")
        self.btn_choseFilePath.clicked.connect(self.chose_file_path)
        self.btn_open_path = self.window.findChild(QPushButton, "btn_open_path")
        self.btn_open_path.clicked.connect(self.open_path)

        # 选择内外网
        self.radioButton_internet = self.window.findChild(QRadioButton, "radioButton_internet")
        self.radioButton_lan = self.window.findChild(QRadioButton, "radioButton_lan")
        self.radioButton_internet.toggled.connect(self.on_network_toggled)
        self.radioButton_lan.toggled.connect(self.on_network_toggled)

        # 选择ics/icc
        self.comboBox_Edition = self.window.findChild(QComboBox, "comboBox_Edition")
        self.comboBox_icc_model = self.window.findChild(QComboBox, "comboBox_icc_model")
        self.comboBox_ver = self.window.findChild(QComboBox, "comboBox_ver")
        self.comboBox_icc_model.addItems(['LITE', 'PRO', "PRO.B", 'TURBO', 'EVO'])  # 增加"PRO.B"  2024/1/31
        self.comboBox_ver.addItems([" ", 'v1.2', 'v1.3', 'v1.4', "v1.5"])
        self.comboBox_Edition.addItems(['Debug', 'Release'])
        self.checkBox_ics = self.window.findChild(QCheckBox, "checkBox_ics")
        self.checkBox_icc = self.window.findChild(QCheckBox, "checkBox_icc")
        self.checkBox_icm = self.window.findChild(QCheckBox, "checkBox_icm")  # jcywong 增加icm 2023/11/13
        self.checkBox_ics.toggled.connect(self.selection_change_checkbox_soft_type)  # jcywong 增加 2023/11/13
        self.checkBox_icc.toggled.connect(self.selection_change_checkbox_soft_type)  # jcywong 增加 2023/11/13
        self.checkBox_icm.toggled.connect(self.selection_change_checkbox_soft_type)  # jcywong 增加 2023/11/13
        self.comboBox_Edition.currentIndexChanged.connect(self.selection_change_comboBox_edition)

        # 下载
        self.btn_download = self.window.findChild(QPushButton, "btn_download")
        self.btn_download.clicked.connect(self.download_soft)
        self.progressBar_download = self.window.findChild(QProgressBar, "progressBar_download")
        self.progressBar_download.setRange(0, 2)
        self.downloading = False
        self.executing = False

        # 运行ICS Studio
        self.btn_run_ics = self.window.findChild(QPushButton, "btn_run_ics")
        self.btn_run_ics.clicked.connect(self.run_ICSStudio)

        self.btn_run_gateway = self.window.findChild(QPushButton, "btn_run_gateway")
        self.btn_run_gateway.clicked.connect(self.run_gateway)

        self.btn_run_update = self.window.findChild(QPushButton, "btn_run_update")
        self.btn_run_update.clicked.connect(self.run_update)

        # 复制ics版本号 jcywong add 2023/11/13
        self.btn_copy_ics_ver = self.window.findChild(QPushButton, "btn_copy_ver")
        self.btn_copy_ics_ver.clicked.connect(self.copy_ics_ver)

        # 打开ics 目录
        self.btn_ics_path = self.window.findChild(QPushButton, "btn_ics_path")
        self.btn_ics_path.clicked.connect(self.open_ics_path)

        # 控制PLC
        self.lineEdit_ip1 = self.window.findChild(QLineEdit, "lineEdit_ip1")
        self.lineEdit_ip2 = self.window.findChild(QLineEdit, "lineEdit_ip2")
        self.lineEdit_ip3 = self.window.findChild(QLineEdit, "lineEdit_ip3")
        self.lineEdit_ip4 = self.window.findChild(QLineEdit, "lineEdit_ip4")
        self.ip_parts = [self.lineEdit_ip1, self.lineEdit_ip2, self.lineEdit_ip3, self.lineEdit_ip4]
        # 创建一个用于验证 IP 地址部分的正则表达式
        ip_regex = QRegularExpression(
            r"^(25[0-5]\.?|2[0-4][0-9]\.?|[0-1]?[0-9][0-9]?\.?)$")  # jcywong 增加"."跳转到下一个  2024/2/24

        # 创建 QRegularExpressionValidator 并设置正则表达式
        ip_validator = QRegularExpressionValidator(ip_regex)
        for ip_part in self.ip_parts:
            ip_part.textChanged.connect(self.on_ip_part_changed)  # jcywong 增加"."跳转到下一个  2024/2/24
            ip_part.setValidator(ip_validator)

        self.comboBox_model_2 = self.window.findChild(QComboBox, "comboBox_model_2")
        self.comboBox_model_2.addItems(
            ['LITE', 'PRO', "PRO.B", 'TURBO', 'EVO', 'ICM-D1', 'ICM-D3', 'ICM-D5', 'ICM-D7'])  # 增加"PRO.B"  2024/1/31
        self.comboBox_command = self.window.findChild(QComboBox, "comboBox_command")
        self.comboBox_command.addItems([" ", '重启', "获取日志"])

        self.pushButton_execute = self.window.findChild(QPushButton, "pushButton_execute")
        self.pushButton_execute.clicked.connect(self.execute_command)

        self.statusbar = self.window.findChild(QStatusBar, "statusbar")

        # 加载配置
        self.load_config()

    def on_ip_part_changed(self, text):
        """
        当IP地址输入‘.’则跳转下一个
        :param text:
        :return:
        """
        current_line_edit = self.sender()
        if text.endswith('.'):
            current_line_edit.setText(text[:-1])
            index = self.ip_parts.index(current_line_edit)
            if index < len(self.ip_parts) - 1:
                next_line_edit = self.ip_parts[index + 1]
                next_line_edit.setFocus()

    def open_memory_monitor(self):
        self.memory_monitor_dialog = MemoryMonitorDialog()
        self.memory_monitor_dialog.show()

        screen_geometry = QApplication.primaryScreen().geometry()
        widget_rect = self.memory_monitor_dialog.geometry()

        # 计算右下角坐标
        x = screen_geometry.width() - widget_rect.width()
        y = screen_geometry.height() - widget_rect.height()

        self.memory_monitor_dialog.move(x, y - 80)

    def show_version(self):
        # 显示版本信息对话框
        QMessageBox.information(self, 'About', f'Version: {VERSION}\nAuthor: jcywong')

    def update_status(self, status: str):
        """
        更新状态栏
        :param status:
        :return:
        """
        self.statusbar.showMessage(status)

    def on_network_toggled(self):
        """
        勾选网络内网还是外网
        :return:
        """
        if self.radioButton_lan.isChecked():
            self.network = "LAN"
        elif self.radioButton_internet.isChecked():
            self.network = "Internet"

    def execute_command(self):
        def is_ip_address_empty(ip_parts):
            # 检查四个部分的文本是否都为空
            for ip_part in ip_parts:
                if not ip_part.text():
                    return True
            return False

        def worker_thread_func():
            self.executing = True

            command = self.comboBox_command.currentText()
            model = self.comboBox_model_2.currentText()

            if command == " ":
                so.show_message.emit("请选择执行的命令", "warning")
                self.executing = False
                so.execute_state.emit(self.executing)
                return
            elif is_ip_address_empty(self.ip_parts):
                so.show_message.emit("请输入IP地址", "warning")
                self.executing = False
                so.execute_state.emit(self.executing)
                return
            else:
                so.execute_state.emit(self.executing)
                so.show_status.emit("正在执行命令")

            ip_address = ".".join(list(map(lambda ip_part: ip_part.text(), self.ip_parts)))

            # if model == "LITE" or model == "PRO" or model == "PRO.B" or model == "EVO":  # jcywong 增加PRO.B\EVO 2024/1/31
            #     if command == "重启":
            #         if not telnet_to_device(ip_address, command="reboot"):
            #             self.executing = False
            #             so.execute_state.emit(self.executing)
            #             so.show_status.emit("命令执行失败")
            #             return
            #     elif command == "获取日志":
            #         # 判断保存地址 jcywong 解决未设置保存地址问题 2024/2/23
            #         if not self.filePath:
            #             self.executing = False
            #             so.execute_state.emit(self.executing)
            #             so.show_message.emit("请设置保存地址", "warning")
            #             so.show_status.emit("命令执行失败")
            #             return
            #
            #         if not get_icc_logs(model, self.filePath, ip_address):
            #             self.executing = False
            #             so.execute_state.emit(self.executing)
            #             so.show_status.emit("命令执行失败")
            #             return
            #         else:
            #             # 打开文件夹
            #             os.startfile(f"{self.filePath}/logs")
            # elif model == "TURBO":
            #     if command == "重启":
            #         if not ssh_to_device(ip_address, command="reboot"):
            #             self.executing = False
            #             so.execute_state.emit(self.executing)
            #             so.show_status.emit("命令执行失败")
            #             return
            #     elif command == "获取日志":
            #         if not get_icc_logs(model, self.filePath, ip_address):
            #             self.executing = False
            #             so.execute_state.emit(self.executing)
            #             so.show_status.emit("命令执行失败")
            #             return
            #         else:
            #             # 打开文件夹
            #             os.startfile(f"{self.filePath}/logs")
            # else:
            #     pass

            if command == "重启":
                if not reboot_device(device_model=model, ip=ip_address):
                    self.executing = False
                    so.execute_state.emit(self.executing)
                    so.show_status.emit("命令执行失败")
                    return
            elif command == "获取日志":
                # 判断保存地址 jcywong 解决未设置保存地址问题 2024/2/23
                if not self.filePath:
                    self.executing = False
                    so.execute_state.emit(self.executing)
                    so.show_message.emit("请设置保存地址", "warning")
                    so.show_status.emit("命令执行失败")
                    return

                if not get_device_logs(model, self.filePath, ip_address):
                    self.executing = False
                    so.execute_state.emit(self.executing)
                    so.show_status.emit("命令执行失败")
                    return
                else:
                    # 打开文件夹
                    os.startfile(f"{self.filePath}/logs")

            self.executing = False
            so.execute_state.emit(self.executing)
            so.show_status.emit("命令执行完成")

        if self.executing:
            QMessageBox.warning(
                self.window,
                '警告', '任务进行中，请等待完成')
            return

        worker = threading.Thread(target=worker_thread_func)
        worker.start()

    def chose_file_path(self):
        self.filePath = QFileDialog.getExistingDirectory(self.window, "选择存储路径")
        if isinstance(self.lineEdit_save_path, QLineEdit):
            self.lineEdit_save_path.setText(self.filePath)

    def open_path(self):
        """打开存储路径
        jcywong add 2023/11/13
        :return:
        """
        if self.filePath:
            os.startfile(self.filePath)
            so.show_status.emit("打开保存路径成功")
        else:
            QMessageBox.warning(
                self.window,
                '警告', '未选择保存路径')
            so.show_status.emit("打开保存路径失败")

    def selection_change_comboBox_edition(self):
        if isinstance(self.comboBox_Edition, QComboBox):
            if self.comboBox_Edition.currentText() == "Release":
                self.comboBox_ver.setEnabled(True)
            else:
                self.comboBox_ver.setCurrentText(" ")
                self.comboBox_ver.setEnabled(False)

    def selection_change_checkbox_soft_type(self):
        """
        勾选更新状态
        jcywong add 2023/11/13
        :return:
        """
        # # 只勾选ics
        # if self.checkBox_ics.isChecked() and not self.checkBox_icc.isChecked() and not self.checkBox_icm.isChecked():
        #     self.comboBox_icc_model.setEnabled(False)
        #     self.comboBox_Edition.setEnabled(True)
        #     if self.comboBox_Edition.currentText() == "Release":
        #         self.comboBox_ver.setEnabled(True)
        # # 只勾选icc
        # elif self.checkBox_icc.isChecked() and not self.checkBox_ics.isChecked() and not self.checkBox_icm.isChecked():
        #     self.comboBox_icc_model.setEnabled(True)
        #     self.comboBox_Edition.setEnabled(True)
        #     if self.comboBox_Edition.currentText() == "Release":
        #         self.comboBox_ver.setEnabled(True)
        # # 只勾选icm
        # elif self.checkBox_icm.isChecked() and not self.checkBox_ics.isChecked() and not self.checkBox_icc.isChecked():
        #     self.comboBox_icc_model.setEnabled(False)
        #     self.comboBox_Edition.setEnabled(False)
        #     self.comboBox_ver.setEnabled(False)
        # # 都不勾选
        # elif not self.checkBox_icm.isChecked() and not self.checkBox_ics.isChecked() and not self.checkBox_icc.isChecked():
        #     self.comboBox_icc_model.setEnabled(False)
        #     self.comboBox_Edition.setEnabled(False)
        #     self.comboBox_ver.setEnabled(False)
        # # 勾选ics和icc
        # elif self.checkBox_ics.isChecked() and self.checkBox_icc.isChecked() and not self.checkBox_icm.isChecked():
        #     self.comboBox_icc_model.setEnabled(True)
        #     self.comboBox_Edition.setEnabled(True)
        #     if self.comboBox_Edition.currentText() == "Release":
        #         self.comboBox_ver.setEnabled(True)
        # # 勾选ics和icm
        # elif self.checkBox_ics.isChecked() and self.checkBox_icm.isChecked() and not self.checkBox_icc.isChecked():
        #     self.comboBox_icc_model.setEnabled(False)
        #     self.comboBox_Edition.setEnabled(True)
        #     if self.comboBox_Edition.currentText() == "Release":
        #         self.comboBox_ver.setEnabled(True)
        # # 勾选icc和icm
        # elif self.checkBox_icc.isChecked() and self.checkBox_icm.isChecked() and not self.checkBox_ics.isChecked():
        #     self.comboBox_icc_model.setEnabled(True)
        #     self.comboBox_Edition.setEnabled(True)
        #     if self.comboBox_Edition.currentText() == "Release":
        #         self.comboBox_ver.setEnabled(True)
        # # 都勾选
        # elif not self.checkBox_icm.isChecked() and not self.checkBox_ics.isChecked() and not self.checkBox_icc.isChecked():
        #     self.comboBox_icc_model.setEnabled(False)
        #     self.comboBox_Edition.setEnabled(False)
        #     self.comboBox_ver.setEnabled(False)
        ics_checked = self.checkBox_ics.isChecked()
        icc_checked = self.checkBox_icc.isChecked()
        icm_checked = self.checkBox_icm.isChecked()

        self.comboBox_icc_model.setEnabled(icc_checked)
        self.comboBox_Edition.setEnabled(ics_checked or icc_checked)

        if self.comboBox_Edition.currentText() == "Release":
            self.comboBox_ver.setEnabled(ics_checked or icc_checked)

    def update_download_state(self):
        if self.downloading:
            self.radioButton_internet.setEnabled(False)  # jcywong /2023/11/13 增加radiobutton的禁用
            self.radioButton_lan.setEnabled(False)  # jcywong /2023/11/13 增加radiobutton的禁用
            self.checkBox_ics.setEnabled(False)  # jcywong /2023/11/13 增加checkbox的禁用
            self.checkBox_icc.setEnabled(False)  # jcywong /2023/11/13 增加checkbox的禁用
            self.checkBox_icm.setEnabled(False)  # jcywong /2023/11/13 增加checkbox的禁用
            self.comboBox_icc_model.setEnabled(False)
            self.comboBox_Edition.setEnabled(False)
            self.btn_choseFilePath.setEnabled(False)
            if self.comboBox_Edition.currentText() == "Release":
                self.comboBox_ver.setEnabled(False)
        else:
            self.radioButton_internet.setEnabled(True)  # jcywong /2023/11/13 增加radiobutton的禁用
            self.radioButton_lan.setEnabled(True)  # jcywong /2023/11/13 增加radiobutton的禁用
            self.checkBox_ics.setEnabled(True)  # jcywong /2023/11/13 增加checkbox的
            self.checkBox_icc.setEnabled(True)  # jcywong /2023/11/13 增加checkbox的
            self.checkBox_icm.setEnabled(True)  # jcywong /2023/11/13 增加checkbox的
            self.comboBox_icc_model.setEnabled(True)
            self.comboBox_Edition.setEnabled(True)
            self.btn_choseFilePath.setEnabled(True)
            if self.comboBox_Edition.currentText() == "Release":
                self.comboBox_ver.setEnabled(True)

    def update_execute_state(self):
        if self.executing:
            for ip_part in self.ip_parts:
                ip_part.setEnabled(False)
            self.comboBox_model_2.setEnabled(False)
            self.comboBox_command.setEnabled(False)
        else:
            for ip_part in self.ip_parts:
                ip_part.setEnabled(True)
            self.comboBox_model_2.setEnabled(True)
            self.comboBox_command.setEnabled(True)

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
            icm_isChecked = self.checkBox_icm.isChecked()  # jcywong 增加icm 2023/11/13
            file_save_path = self.filePath

            if not file_save_path:
                so.show_message.emit("请设置文件保存地址", "warning")
                self.downloading = False
                so.download_state.emit(self.downloading)
                return
            elif edition == "Release" and ver == " ":
                so.show_message.emit("请设置Release版本", "warning")
                self.downloading = False
                so.download_state.emit(self.downloading)
                return
            else:
                so.download_state.emit(self.downloading)
                so.show_status.emit("正在下载中")

            # self.filename.clear() jcywong  2023/11/13  修改保存配置为 ics+icc+icm
            if ics_isChecked and not icc_isChecked and not icm_isChecked and edition == "Debug":
                # debug ics only
                self.filename.clear()
                self.filename["ICS"] = get_latest_filename(soft_type="ICS", edition=edition, network=self.network)
            elif ics_isChecked and icc_isChecked and not icm_isChecked and edition == "Debug":
                # debug ics + icc
                self.filename.clear()
                self.filename["ICS"] = get_latest_filename(soft_type="ICS", edition=edition, network=self.network)
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', edition=edition, model=model,
                                                           network=self.network)
            elif icc_isChecked and not ics_isChecked and edition == "Debug" and not icm_isChecked:
                # debug icc only
                if "ICM" in self.filename:
                    self.filename.pop("ICM")
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', edition=edition, model=model,
                                                           network=self.network)
            elif ics_isChecked and not icc_isChecked and edition == "Release" and not icm_isChecked:
                # release ics only
                self.filename.clear()
                self.filename["ICS"] = get_latest_filename(edition="Release", network=self.network, ver=ver)
            elif ics_isChecked and icc_isChecked and edition == "Release" and not icm_isChecked:
                # release ics + icc
                self.filename.clear()
                self.filename["ICS"] = get_latest_filename(edition="Release", network=self.network, ver=ver)
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', model=model, edition="Release", ver=ver,
                                                           network=self.network)
            elif icc_isChecked and not ics_isChecked and edition == "Release" and not icm_isChecked:
                # release icc only
                if "ICM" in self.filename:
                    self.filename.pop("ICM")
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', model=model, edition="Release", ver=ver,
                                                           network=self.network)
            elif icm_isChecked and not ics_isChecked and not icc_isChecked:
                # icm only
                if "ICC" in self.filename:
                    self.filename.pop("ICC")
                self.comboBox_Edition.setEnabled(False)
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
            elif icm_isChecked and ics_isChecked and not icc_isChecked and edition == "Release":
                # icm + ics release
                self.filename.clear()
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
                self.filename["ICS"] = get_latest_filename(edition="Release", network=self.network, ver=ver)
            elif icm_isChecked and not ics_isChecked and icc_isChecked and edition == "Release":
                # icm + icc release
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', model=model, edition="Release", ver=ver,
                                                           network=self.network)
            elif icm_isChecked and ics_isChecked and icc_isChecked and edition == "Release":
                # icm + icc + ics release
                self.filename.clear()
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
                self.filename["ICS"] = get_latest_filename(edition="Release", network=self.network, ver=ver)
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', model=model, edition="Release", ver=ver,
                                                           network=self.network)
            elif icm_isChecked and ics_isChecked and not icc_isChecked and edition == "Debug":
                # icm + ics debug
                self.filename.clear()
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
                self.filename["ICS"] = get_latest_filename(edition="Debug", network=self.network)
            elif icm_isChecked and not ics_isChecked and icc_isChecked and edition == "Debug":
                # icm + icc debug
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', model=model, edition="Debug",
                                                           network=self.network)
            elif icm_isChecked and ics_isChecked and icc_isChecked and edition == "Debug":
                # icm + icc + ics debug
                self.filename.clear()
                self.filename["ICM"] = get_latest_filename(soft_type='ICM', network=self.network)
                self.filename["ICS"] = get_latest_filename(edition="Debug", network=self.network)
                self.filename["ICC"] = get_latest_filename(soft_type='ICC', model=model, edition="Debug",
                                                           network=self.network)
            elif not ics_isChecked and not icc_isChecked and not icm_isChecked:
                # none
                so.show_message.emit("请勾选下载软件", "warning")
                self.downloading = False
                so.download_state.emit(self.downloading)
                so.show_status.emit("")
                return

            so.progress_update.emit(1)
            try:
                for soft_type, name in self.filename.items():
                    download_file(file_name=name, file_save_path=self.filePath, soft_type=soft_type, edition=edition,
                                  network=self.network)
                    unzip_file(self.filePath, name)
            except TypeError:
                print("网络错误")
                self.downloading = False

                so.progress_update.emit(0)
                so.show_status.emit(f"网络错误/文件不存在，下载失败！")
                so.download_state.emit(self.downloading)
                so.show_message.emit("网络错误/文件不存在，下载失败！", "warning")
                return

            self.downloading = False

            so.progress_update.emit(2)
            so.show_status.emit(f"{self.filename}下载完成")
            so.download_state.emit(self.downloading)

        if self.downloading:
            QMessageBox.warning(
                self.window,
                '警告', '任务进行中，请等待完成')
            return

        worker = threading.Thread(target=workerThreadFunc)
        worker.start()

    def run_ICSStudio(self):
        """运行ics studio"""
        # jcywong 2023/11/13 修改filename列表为字典
        ics = self.filename["ICS"]
        if ics:
            # jcywong modify 2023/12/14
            # if ics[:3] == "ICS" and ics[-9:-4] == "debug":
            #     so.show_status.emit(f"打开ICS Studio版本：{ics}")
            #     # open_ics(self.filePath + "/" + ics[:-4] + "/Debug")
            #     open_ics(self.filePath + "/" + ics[:-4])  # jcywong modify 2023/12/11
            # elif ics[:3] == "ICS" and ics[-11:-4] == "Release":
            #     so.show_status.emit(f"打开ICS Studio版本：{ics}")
            #     # open_ics(self.filePath + "/" + ics[:-4] + "/Release")
            open_ics(self.filePath + "/" + ics[:-4])  # jcywong modify 2023/12/14
            so.show_status.emit("打开ICS Studio成功")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("打开ICS Studio失败")

    def run_gateway(self):
        """运行ics gateway"""
        ics = self.filename["ICS"]
        if ics:
            gateway_path = self.filePath + "/" + ics[:-4] + "/Extensions/ICSGateway/ICSGateway.exe"
            subprocess.Popen(gateway_path)
            so.show_status.emit("打开ICS Gateway成功")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("打开ICS Gateway失败")

    def run_update(self):
        """运行ics update plus"""
        ics = self.filename["ICS"]
        if ics:
            update_path = self.filePath + "/" + ics[:-4] + "/Extensions/IconUpdater/IconUpdater.exe"
            subprocess.Popen(update_path)
            so.show_status.emit("打开Icon Update Plus成功")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("打开Icon Update Plus失败")

    def copy_ics_ver(self):  # jcywong 2023/11/13
        """复制ics版本号到剪切板"""
        filename = self.filename["ICS"][:-4]
        if filename:
            QApplication.clipboard().setText(filename)
            so.show_status.emit(f"版本号已复制到剪贴板：{filename}")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("复制版本号失败")

    def open_ics_path(self):
        """打开ics路径"""
        if self.filename["ICS"]:
            ics_path = self.filePath + "/" + self.filename["ICS"][:-4]
            os.startfile(ics_path)
            so.show_status.emit("打开ICS Studio路径成功")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("打开ICS Studio路径失败")

    def setProgress(self, value):
        self.progressBar_download.setValue(value)

    def save_config(self):
        """保存配置"""

        # 创建配置字典
        config_data = {
            'save_path': self.filePath,
            'network': self.network,
            'filename': self.filename
        }

        # 将配置信息保存到配置文件
        with open(config_file, 'w') as file:
            json.dump(config_data, file)

    def load_config(self):
        try:
            with open(config_file, 'r') as file:
                config_data = json.load(file)

                # 从配置中获取信息
                saved_address = config_data.get('save_path', '')
                self.filePath = saved_address
                self.lineEdit_save_path.setText(saved_address)

                network = config_data.get('network', '')
                self.network = network
                if network == "LAN":
                    self.radioButton_lan.setChecked(True)
                elif network == "Internet":
                    self.radioButton_internet.setChecked(True)

                filename = config_data.get('filename', '')
                self.filename = filename
        except FileNotFoundError:
            # 如果配置文件不存在，不做任何操作
            pass

    def closeEvent(self, event):
        # 在窗口关闭时自动保存配置
        if self.memory_monitor_dialog:
            self.memory_monitor_dialog.close()
        self.save_config()
        event.accept()

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
    window.setWindowIcon(QIcon('tool.png'))
    window.setWindowTitle("TestTools")
    window.show()
    sys.exit(app.exec())

    # print("按任意键退出...")
    # input()  # 等待用户输入
    # print("程序已退出")

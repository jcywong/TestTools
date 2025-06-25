import json
import sys
import threading

import requests.exceptions
from PySide6.QtCore import QFile, QIODevice, QObject, Signal, QRegularExpression, Qt, QThread
from PySide6.QtGui import QRegularExpressionValidator, QIcon, QAction
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QMessageBox, QFileDialog, QLineEdit, QComboBox, \
    QProgressBar, QRadioButton, QStatusBar, QTabWidget, QDialog, QLabel, QVBoxLayout, QMenuBar, QMenu, QHBoxLayout

from comm import *

# 定义版本号
VERSION = "1.4.2"


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

configs = {}

tab_map = {
    0: "ics",
    1: "icc",
    2: "icm",
    3: "icp",
    4: "vp",
}


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


class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setFixedSize(400, 100)

        # 布局
        v_layout = QVBoxLayout(self)
        h_layout_button = QHBoxLayout(self)
        h_layout_network = QHBoxLayout(self)

        # 显示文件路径的标签
        self.label = QLabel("保存路径：")
        h_layout_button.addWidget(self.label)

        self.save_path_label = QLabel()
        if "save_path" in configs and configs["save_path"]:
            self.save_path_label.setText(configs["save_path"])
        else:
            self.save_path_label.setText("未设置保存路径")
        h_layout_button.addWidget(self.save_path_label)

        # 选择文件的按钮
        self.select_button = QPushButton("选择文件夹")
        self.select_button.clicked.connect(self.select_save_path)
        h_layout_button.addWidget(self.select_button)

        v_layout.addLayout(h_layout_button)

        # 网络选择
        self.label_network = QLabel("网络选择：")
        h_layout_network.addWidget(self.label_network)

        self.radioButton_lan = QRadioButton("局域网")
        self.radioButton_internet = QRadioButton("互联网")
        if "network" in configs and configs["network"]:
            if configs["network"] == "LAN":
                self.radioButton_lan.setChecked(True)
            else:
                self.radioButton_internet.setChecked(True)
        else:
            self.radioButton_lan.setChecked(True)

        h_layout_network.addWidget(self.radioButton_lan)
        h_layout_network.addWidget(self.radioButton_internet)

        v_layout.addLayout(h_layout_network)

        # 确认取消按钮
        h_layout_button = QHBoxLayout(self)
        self.ok_button = QPushButton("确定")
        self.ok_button.clicked.connect(self.accept)
        h_layout_button.addWidget(self.ok_button)

        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.reject)
        h_layout_button.addWidget(self.cancel_button)

        v_layout.addLayout(h_layout_button)

        self.setLayout(v_layout)

    def select_save_path(self):
        save_path = QFileDialog.getExistingDirectory(self, "选择存储路径")
        if save_path:
            self.save_path_label.setText(save_path)
    def accept(self):
        # 保存配置
        text = self.save_path_label.text()
        configs["save_path"] = text if text !="未设置保存路径" else None
        configs["network"] = "LAN" if self.radioButton_lan.isChecked() else "Internet"
        super().accept()


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
        self.settings_action = self.window.findChild(QAction, "action_settings")
        self.ver_action.triggered.connect(self.show_version)
        self.settings_action.triggered.connect(self.open_settings)

        # tab选项  jcywong add 2023/11/13
        self.tab_tabMenu = self.window.findChild(QTabWidget, "tab")
        # self.tab_general = self.window.findChild(QWidget, "tab_general")
        # self.tab_upload_download = self.window.findChild(QWidget, "tab_upload_download")

        self.tab_tabMenu.removeTab(5)


        # 初始化tab
        self._init_tab_ics()
        self._init_tab_icc()
        self._init_tab_icm()
        self._init_tab_visual_pro()
        self._init_tab_icp()

        self.statusbar = self.window.findChild(QStatusBar, "statusbar")

        self.downloading = False
        self.executing = False

        # 加载配置
        self.load_config()

        # 判断版本
        self.check_version()

    def _init_tab_ics(self):

        # 选择ics
        self.ics_comboBox_Edition = self.window.findChild(QComboBox, "ics_comboBox_Edition")
        self.ics_comboBox_ver = self.window.findChild(QComboBox, "ics_comboBox_ver")
        self.ics_comboBox_ver.addItems([" ", 'v1.2', 'v1.3', 'v1.4', "v1.5", "v1.6"])
        self.ics_comboBox_Edition.addItems(['Debug', 'Release'])
        self.ics_comboBox_Edition.currentIndexChanged.connect(lambda: self.selection_change_comboBox_edition("ics"))

        # 下载
        self.ics_btn_download = self.window.findChild(QPushButton, "ics_btn_download")
        self.ics_btn_download.clicked.connect(self.download_soft)
        self.ics_progressBar_download = self.window.findChild(QProgressBar, "ics_progressBar_download")
        self.ics_progressBar_download.setRange(0, 2)

        # 运行ICS Studio
        self.btn_run_ics = self.window.findChild(QPushButton, "btn_run_ics")
        self.btn_run_ics.clicked.connect(self.run_soft)

        self.btn_run_gateway = self.window.findChild(QPushButton, "btn_run_gateway")
        self.btn_run_gateway.clicked.connect(self.run_gateway)

        self.btn_run_update = self.window.findChild(QPushButton, "btn_run_update")
        self.btn_run_update.clicked.connect(self.run_update)

        # 复制ics版本号 jcywong add 2023/11/13
        self.btn_copy_ics_ver = self.window.findChild(QPushButton, "btn_copy_ver")
        self.btn_copy_ics_ver.clicked.connect(self.copy_ver)

        # 打开ics 目录
        self.btn_ics_path = self.window.findChild(QPushButton, "btn_ics_path")
        self.btn_ics_path.clicked.connect(self.open_soft_path)

    def _init_tab_icc(self):

        self.comboBox_icc_model = self.window.findChild(QComboBox, "comboBox_icc_model")
        self.comboBox_icc_model.addItems(['LITE', 'LITE.B', 'PRO', 'PRO.B', 'TURBO', 'EVO'])

        # 选择icc
        self.icc_comboBox_Edition = self.window.findChild(QComboBox, "icc_comboBox_Edition")
        self.icc_comboBox_ver = self.window.findChild(QComboBox, "icc_comboBox_ver")
        self.icc_comboBox_ver.addItems([" ", 'v1.2', 'v1.3', 'v1.4', "v1.5", "v1.6"])
        self.icc_comboBox_Edition.addItems(['Debug', 'Release'])
        self.icc_comboBox_Edition.currentIndexChanged.connect(lambda: self.selection_change_comboBox_edition("icc"))

        # 下载
        self.icc_btn_download = self.window.findChild(QPushButton, "icc_btn_download")
        self.icc_btn_download.clicked.connect(self.download_soft)
        self.icc_progressBar_download = self.window.findChild(QProgressBar, "icc_progressBar_download")
        self.icc_progressBar_download.setRange(0, 2)

        # 运行
        self.icc_btn_run_update = self.window.findChild(QPushButton, "icc_btn_run_update")
        self.icc_btn_run_update.clicked.connect(self.run_update)

        # 复制ics版本号 jcywong add 2023/11/13
        self.icc_btn_copy_ver = self.window.findChild(QPushButton, "icc_btn_copy_ver")
        self.icc_btn_copy_ver.clicked.connect(self.copy_ver)

        # 打开ics 目录
        self.icc_btn_path = self.window.findChild(QPushButton, "icc_btn_path")
        self.icc_btn_path.clicked.connect(self.open_soft_path)

        # 控制PLC
        self.icc_lineEdit_ip1 = self.window.findChild(QLineEdit, "icc_lineEdit_ip1")
        self.icc_lineEdit_ip2 = self.window.findChild(QLineEdit, "icc_lineEdit_ip2")
        self.icc_lineEdit_ip3 = self.window.findChild(QLineEdit, "icc_lineEdit_ip3")
        self.icc_lineEdit_ip4 = self.window.findChild(QLineEdit, "icc_lineEdit_ip4")
        self.icc_ip_parts = [self.icc_lineEdit_ip1, self.icc_lineEdit_ip2, self.icc_lineEdit_ip3, self.icc_lineEdit_ip4]
        # 创建一个用于验证 IP 地址部分的正则表达式
        ip_regex = QRegularExpression(
            r"^(25[0-5]\.?|2[0-4][0-9]\.?|[0-1]?[0-9][0-9]?\.?)$")  # jcywong 增加"."跳转到下一个  2024/2/24

        # 创建 QRegularExpressionValidator 并设置正则表达式
        ip_validator = QRegularExpressionValidator(ip_regex)
        for ip_part in self.icc_ip_parts:
            ip_part.textChanged.connect(self.on_ip_part_changed)
            ip_part.setValidator(ip_validator)

        self.icc_comboBox_model = self.window.findChild(QComboBox, "icc_comboBox_model")
        self.icc_comboBox_model.addItems(
            ['LITE', 'LITE.B', 'PRO', "PRO.B", 'TURBO', 'EVO'])  # 增加"PRO.B"  2024/1/31
        self.icc_comboBox_command = self.window.findChild(QComboBox, "icc_comboBox_command")
        self.icc_comboBox_command.addItems([" ", '重启', "获取日志"])

        self.icc_pushButton_execute = self.window.findChild(QPushButton, "icc_pushButton_execute")
        self.icc_pushButton_execute.clicked.connect(self.execute_command)

    def _init_tab_icm(self):

        self.comboBox_icm_model = self.window.findChild(QComboBox, "comboBox_icm_model")
        self.comboBox_icm_model.addItems(["Release",'D1', 'D3', "D5", 'D7'])

        # 选择icm
        self.icm_comboBox_Edition = self.window.findChild(QComboBox, "icm_comboBox_Edition")
        self.icm_comboBox_ver = self.window.findChild(QComboBox, "icm_comboBox_ver")
        self.icm_comboBox_ver.addItems([" ", 'v1.2', 'v1.3', 'v1.4', "v1.5", "v1.6"])
        self.icm_comboBox_Edition.addItems([' ', 'Debug',"Release"])
        self.icm_comboBox_Edition.currentIndexChanged.connect(lambda: self.selection_change_comboBox_edition("icm"))

        # 下载
        self.icm_btn_download = self.window.findChild(QPushButton, "icm_btn_download")
        self.icm_btn_download.clicked.connect(self.download_soft)
        self.icm_progressBar_download = self.window.findChild(QProgressBar, "icm_progressBar_download")
        self.icm_progressBar_download.setRange(0, 2)

        # 运行
        self.icm_btn_run_update = self.window.findChild(QPushButton, "icm_btn_run_update")
        self.icm_btn_run_update.clicked.connect(self.run_update)

        # 复制ics版本号 jcywong add 2023/11/13
        self.icm_btn_copy_ver = self.window.findChild(QPushButton, "icm_btn_copy_ver")
        self.icm_btn_copy_ver.clicked.connect(self.copy_ver)

        # 打开ics 目录
        self.icm_btn_path = self.window.findChild(QPushButton, "icm_btn_path")
        self.icm_btn_path.clicked.connect(self.open_soft_path)

        # 控制PLC
        self.icm_lineEdit_ip1 = self.window.findChild(QLineEdit, "icm_lineEdit_ip1")
        self.icm_lineEdit_ip2 = self.window.findChild(QLineEdit, "icm_lineEdit_ip2")
        self.icm_lineEdit_ip3 = self.window.findChild(QLineEdit, "icm_lineEdit_ip3")
        self.icm_lineEdit_ip4 = self.window.findChild(QLineEdit, "icm_lineEdit_ip4")
        self.icm_ip_parts = [self.icm_lineEdit_ip1, self.icm_lineEdit_ip2, self.icm_lineEdit_ip3, self.icm_lineEdit_ip4]
        # 创建一个用于验证 IP 地址部分的正则表达式
        ip_regex = QRegularExpression(
            r"^(25[0-5]\.?|2[0-4][0-9]\.?|[0-1]?[0-9][0-9]?\.?)$")  # jcywong 增加"."跳转到下一个  2024/2/24

        # 创建 QRegularExpressionValidator 并设置正则表达式
        ip_validator = QRegularExpressionValidator(ip_regex)
        for ip_part in self.icm_ip_parts:
            ip_part.textChanged.connect(self.on_ip_part_changed)  # jcywong 增加"."跳转到下一个  2024/2/24
            ip_part.setValidator(ip_validator)

        self.icm_comboBox_model = self.window.findChild(QComboBox, "icm_comboBox_model")
        self.icm_comboBox_model.addItems(['ICM-D1', 'ICM-D3', "ICM-D5", 'ICM-D7'])
        self.icm_comboBox_command = self.window.findChild(QComboBox, "icm_comboBox_command")
        self.icm_comboBox_command.addItems([" ", '重启', "获取日志"])

        self.icm_pushButton_execute = self.window.findChild(QPushButton, "icm_pushButton_execute")
        self.icm_pushButton_execute.clicked.connect(self.execute_command)

    def _init_tab_visual_pro(self):

        # 选择vp
        self.vp_comboBox_Edition = self.window.findChild(QComboBox, "vp_comboBox_Edition")
        self.vp_comboBox_ver = self.window.findChild(QComboBox, "vp_comboBox_ver")
        # self.vp_comboBox_ver.addItems([" ", 'v1.2', 'v1.3', 'v1.4', "v1.5", "v1.6"])
        self.vp_comboBox_Edition.addItems(['Debug', 'Release'])
        self.vp_comboBox_Edition.currentIndexChanged.connect(lambda: self.selection_change_comboBox_edition("vp"))

        # 下载
        self.vp_btn_download = self.window.findChild(QPushButton, "vp_btn_download")
        self.vp_btn_download.clicked.connect(self.download_soft)
        self.vp_progressBar_download = self.window.findChild(QProgressBar, "vp_progressBar_download")
        self.vp_progressBar_download.setRange(0, 2)

        # 运行ICS Studio
        self.vp_btn_run = self.window.findChild(QPushButton, "vp_btn_run")
        self.vp_btn_run.clicked.connect(self.run_soft)

        # 复制ics版本号 jcywong add 2023/11/13
        self.vp_btn_copy_ver = self.window.findChild(QPushButton, "vp_btn_copy_ver")
        self.vp_btn_copy_ver.clicked.connect(self.copy_ver)

        # 打开ics 目录
        self.vp_btn_path = self.window.findChild(QPushButton, "vp_btn_path")
        self.vp_btn_path.clicked.connect(self.open_soft_path)

    def _init_tab_icp(self):

        # 选择icp
        self.icp_comboBox_Edition = self.window.findChild(QComboBox, "icp_comboBox_Edition")
        self.icp_comboBox_ver = self.window.findChild(QComboBox, "icp_comboBox_ver")
        # self.icp_comboBox_ver.addItems()
        self.icp_comboBox_Edition.addItems(['Debug', 'Release'])
        self.icp_comboBox_Edition.currentIndexChanged.connect(lambda: self.selection_change_comboBox_edition("icp"))

        # 下载
        self.icp_btn_download = self.window.findChild(QPushButton, "icp_btn_download")
        self.icp_btn_download.clicked.connect(self.download_soft)
        self.icp_progressBar_download = self.window.findChild(QProgressBar, "icp_progressBar_download")
        self.icp_progressBar_download.setRange(0, 2)

        # 运行
        self.icp_btn_run_update = self.window.findChild(QPushButton, "icp_btn_run_update")
        self.icp_btn_run_update.clicked.connect(self.run_update)

        # 复制ics版本号 jcywong add 2023/11/13
        self.icp_btn_copy_ver = self.window.findChild(QPushButton, "icp_btn_copy_ver")
        self.icp_btn_copy_ver.clicked.connect(self.copy_ver)

        # 打开ics 目录
        self.icp_btn_path = self.window.findChild(QPushButton, "icp_btn_path")
        self.icp_btn_path.clicked.connect(self.open_soft_path)

        # 控制PLC
        self.icp_lineEdit_ip1 = self.window.findChild(QLineEdit, "icp_lineEdit_ip1")
        self.icp_lineEdit_ip2 = self.window.findChild(QLineEdit, "icp_lineEdit_ip2")
        self.icp_lineEdit_ip3 = self.window.findChild(QLineEdit, "icp_lineEdit_ip3")
        self.icp_lineEdit_ip4 = self.window.findChild(QLineEdit, "icp_lineEdit_ip4")
        self.icp_ip_parts = [self.icp_lineEdit_ip1, self.icp_lineEdit_ip2, self.icp_lineEdit_ip3, self.icp_lineEdit_ip4]
        # 创建一个用于验证 IP 地址部分的正则表达式
        ip_regex = QRegularExpression(
            r"^(25[0-5]\.?|2[0-4][0-9]\.?|[0-1]?[0-9][0-9]?\.?)$")  # jcywong 增加"."跳转到下一个  2024/2/24

        # 创建 QRegularExpressionValidator 并设置正则表达式
        ip_validator = QRegularExpressionValidator(ip_regex)
        for ip_part in self.icp_ip_parts:
            ip_part.textChanged.connect(self.on_ip_part_changed)  # jcywong 增加"."跳转到下一个  2024/2/24
            ip_part.setValidator(ip_validator)

        self.icp_comboBox_model = self.window.findChild(QComboBox, "icp_comboBox_model")
        self.icp_comboBox_model.addItems(['ICP'])
        self.icp_comboBox_command = self.window.findChild(QComboBox, "icp_comboBox_command")
        self.icp_comboBox_command.addItems([" ", '重启'])

        self.icp_pushButton_execute = self.window.findChild(QPushButton, "icp_pushButton_execute")
        self.icp_pushButton_execute.clicked.connect(self.execute_command)

    def on_ip_part_changed(self, text):
        """
        当IP地址输入‘.’则跳转下一个
        :param text:
        :return:
        """
        current_line_edit = self.sender()
        cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()]
        if text.endswith('.'):
            if cur_tab_name == "icc":
                current_line_edit.setText(text[:-1])
                index = self.icc_ip_parts.index(current_line_edit)
                if index < len(self.icc_ip_parts) - 1:
                    next_line_edit = self.icc_ip_parts[index + 1]
                    next_line_edit.setFocus()
            elif cur_tab_name == "icp":
                current_line_edit.setText(text[:-1])
                index = self.icp_ip_parts.index(current_line_edit)
                if index < len(self.icp_ip_parts) - 1:
                    next_line_edit = self.icp_ip_parts[index + 1]
                    next_line_edit.setFocus()
            elif cur_tab_name == "icm":
                current_line_edit.setText(text[:-1])
                index = self.icm_ip_parts.index(current_line_edit)
                if index < len(self.icm_ip_parts) - 1:
                    next_line_edit = self.icm_ip_parts[index + 1]
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

    def open_settings(self):
        # 打开设置对话框
        settings_dialog = SettingsDialog()
        if settings_dialog.exec():
            # 在对话框关闭时保存配置到全局字典
            self.filePath = configs["save_path"] if "save_path" in configs else None
            self.network = configs["network"] if "network" in configs else None

    def show_version(self):
        # 显示版本信息对话框
        ret = get_test_tools_last_version()
        last_version = ret["last-version"] if ret else None
        message = f'''
            <p>Author: Jcy Wong</p>
            <p>Version: {VERSION}</p>
            <p>Last Version: {last_version}</p>
            <p>
            <a href="https://jcywong.notion.site/TestTools-0d910d4d26ad44d5b813119b59f8dae7">User Manual</a>
            </p>
        '''

        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setTextFormat(Qt.RichText)  # 允许使用富文本
        msg.setText(message)
        msg.setWindowTitle("信息")
        msg.exec()

    def update_status(self, status: str):
        """
        更新状态栏
        :param status:
        :return:
        """
        self.statusbar.showMessage(status)

    def execute_command(self):
        def is_ip_address_empty(ip_parts):
            # 检查四个部分的文本是否都为空
            for ip_part in ip_parts:
                if not ip_part.text():
                    return True
            return False

        def worker_thread_func():
            self.executing = True
            cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()]

            command = eval(f"self.{cur_tab_name}_comboBox_command.currentText()")
            model = eval(f"self.{cur_tab_name}_comboBox_model.currentText()")
            ip_parts = eval(f"self.{cur_tab_name}_ip_parts")

            if command == " ":
                so.show_message.emit("请选择执行的命令", "warning")
                self.executing = False
                so.execute_state.emit(self.executing)
                return
            elif is_ip_address_empty(ip_parts):
                so.show_message.emit("请输入IP地址", "warning")
                self.executing = False
                so.execute_state.emit(self.executing)
                return
            else:
                so.execute_state.emit(self.executing)
                so.show_status.emit("正在执行命令")

            ip_address = ".".join(list(map(lambda ip_part: ip_part.text(), ip_parts)))


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

    def selection_change_comboBox_edition(self, type):
        if isinstance(self.ics_comboBox_Edition, QComboBox):
            if type == "ics":
                if self.ics_comboBox_Edition.currentText() == "Release":
                    self.ics_comboBox_ver.setEnabled(True)
                else:
                    self.ics_comboBox_ver.setCurrentText(" ")
                    self.ics_comboBox_ver.setEnabled(False)
            elif type == "icc":
                if self.icc_comboBox_Edition.currentText() == "Release":
                    self.icc_comboBox_ver.setEnabled(True)
                else:
                    self.icc_comboBox_ver.setCurrentText(" ")
                    self.icc_comboBox_ver.setEnabled(False)
            elif type == "icm":
                if self.icm_comboBox_Edition.currentText() == "Release":
                    self.icm_comboBox_ver.setEnabled(True)
                else:
                    self.icm_comboBox_ver.setCurrentText(" ")
                    self.icm_comboBox_ver.setEnabled(False)
            elif type == "icp":
                if self.icp_comboBox_Edition.currentText() == "Release":
                    self.icp_comboBox_ver.setEnabled(False)
                else:
                    self.icp_comboBox_ver.setCurrentText(" ")
                    self.icp_comboBox_ver.setEnabled(False)
            elif type == "vp":
                if self.vp_comboBox_Edition.currentText() == "Release":
                    self.vp_comboBox_ver.setEnabled(False)
                else:
                    self.vp_comboBox_ver.setCurrentText(" ")
                    self.vp_comboBox_ver.setEnabled(False)


    def update_download_state(self, state):
        current_index = self.tab_tabMenu.currentIndex()
        cur_tab_name = tab_map[current_index]

        if state:
            for i in range(self.tab_tabMenu.count()):
                if i != current_index:
                    self.tab_tabMenu.setTabEnabled(i, False)


            if cur_tab_name == "ics":
                self.ics_comboBox_Edition.setEnabled(False)
                if self.ics_comboBox_Edition.currentText() == "Release":
                    self.ics_comboBox_ver.setEnabled(False)
            elif cur_tab_name == "icc":
                self.comboBox_icc_model.setEnabled(False)
                self.icc_comboBox_Edition.setEnabled(False)
                if self.icc_comboBox_Edition.currentText() == "Release":
                    self.icc_comboBox_ver.setEnabled(False)
            # elif cur_tab_name == "icm":
            #     self.icm_comboBox_Edition.setEnabled(False)
            #     if self.icm_comboBox_Edition.currentText() == "Release":
            #         self.icm_comboBox_ver.setEnabled(False)
            elif cur_tab_name == "icp":
                self.icp_comboBox_Edition.setEnabled(False)
                if self.icp_comboBox_Edition.currentText() == "Release":
                    self.icp_comboBox_ver.setEnabled(False)
            # elif cur_tab_name == "vp":
            #     self.vp_comboBox_Edition.setEnabled(False)
            #     if self.vp_comboBox_Edition.currentText() == "Release":
            #         self.vp_comboBox_ver.setEnabled(False)
        else:
            for i in range(self.tab_tabMenu.count()):
                self.tab_tabMenu.setTabEnabled(i, True)

            if cur_tab_name == "ics":
                self.ics_comboBox_Edition.setEnabled(True)
                if self.ics_comboBox_Edition.currentText() == "Release":
                    self.ics_comboBox_ver.setEnabled(True)
            elif cur_tab_name == "icc":
                self.comboBox_icc_model.setEnabled(True)
                self.icc_comboBox_Edition.setEnabled(True)
                if self.icc_comboBox_Edition.currentText() == "Release":
                    self.icc_comboBox_ver.setEnabled(True)
            # elif cur_tab_name == "icm":
            #     self.icm_comboBox_Edition.setEnabled(False)
            #     if self.icm_comboBox_Edition.currentText() == "Release":
            #         self.icm_comboBox_ver.setEnabled(True)
            elif cur_tab_name == "icp":
                self.icp_comboBox_Edition.setEnabled(True)
                if self.icp_comboBox_Edition.currentText() == "Release":
                    self.icp_comboBox_ver.setEnabled(False)
            # elif cur_tab_name == "vp":
            #     self.vp_comboBox_Edition.setEnabled(True)
            #     if self.vp_comboBox_Edition.currentText() == "Release":
            #         self.vp_comboBox_ver.setEnabled(False)

    def update_execute_state(self):
        if self.executing:
            for ip_part in self.icc_ip_parts and self.icm_ip_parts and self.icp_ip_parts:
                ip_part.setEnabled(False)
            self.icc_comboBox_model.setEnabled(False)
            self.icm_comboBox_model.setEnabled(False)
            self.icp_comboBox_model.setEnabled(False)
            self.icc_comboBox_command.setEnabled(False)
            self.icm_comboBox_command.setEnabled(False)
            self.icp_comboBox_command.setEnabled(False)
        else:
            for ip_part in self.icc_ip_parts and self.icm_ip_parts and self.icp_ip_parts:
                ip_part.setEnabled(True)
            self.icc_comboBox_model.setEnabled(True)
            self.icm_comboBox_model.setEnabled(True)
            self.icp_comboBox_model.setEnabled(True)
            self.icc_comboBox_command.setEnabled(True)
            self.icm_comboBox_command.setEnabled(True)
            self.icp_comboBox_command.setEnabled(True)

    def show_MessageBox(self, message, message_type):
        if message_type == "warning":
            return QMessageBox.warning(self.window, "警告", message)
        elif message_type == "question":
            return QMessageBox.question(self.window, '确认', message)
        elif message_type == "information":
            return QMessageBox.information(self.window, '提示', message)

    def download_soft(self):
        def workerThreadFunc():
            self.downloading = True

            cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()]

            model = None

            if cur_tab_name == "icc":
                model = self.comboBox_icc_model.currentText()

            edition = eval(f"self.{cur_tab_name}_comboBox_Edition.currentText()")
            ver = eval(f"self.{cur_tab_name}_comboBox_ver.currentText()")

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

            try:
                if self.filename:
                    for soft_type, infos in self.filename.items():
                        infos["is_checked"] = False
                else:
                    self.filename = {}

                self.filename[cur_tab_name.upper()] = {
                    "name": get_latest_filename(soft_type=cur_tab_name.upper(), edition=edition, model=model, ver=ver,
                                                network=self.network),
                    "is_checked": True
                }
            except requests.exceptions.ConnectionError:

                self.downloading = False

                so.progress_update.emit(0)
                so.show_status.emit(f"网络错误，下载失败！")
                so.download_state.emit(self.downloading)
                so.show_message.emit("网络错误，下载失败！", "warning")
                return

            so.progress_update.emit(1)
            try:
                for soft_type, infos in self.filename.items():
                    if infos["is_checked"]:
                        name = infos["name"]
                    else:
                        continue

                    try:
                        download_file(file_name=name, file_save_path=self.filePath, soft_type=soft_type,
                                      edition=edition,
                                      network=self.network)

                        unzip_file(self.filePath, name)

                    except FileExistsError:
                        so.show_message.emit(f"{name}:文件已经存在！", "information")
                        so.show_status.emit(f"{name}:文件已经存在！")

                    except Exception as e:
                        print(f"{e}")
                        raise e

            except Exception as e:
                print(f"{e}")

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

    def run_soft(self):
        """运行ics visual pro"""
        try:
            cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()].upper()
            soft = self.filename[cur_tab_name]["name"]
            path = self.filePath + "/" + soft[:-4]
            if cur_tab_name == "ICS":
                open_ics(path)
            elif cur_tab_name == "VP":
                subprocess.Popen(path + "/VisualPro.exe")
            so.show_status.emit(f"打开{cur_tab_name}成功")
        except KeyError:
            QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
            so.show_status.emit(f"打开{cur_tab_name}失败")

    def run_gateway(self):
        """运行ics gateway"""
        ics = self.filename["ICS"]["name"]
        if ics:
            gateway_path = self.filePath + "/" + ics[:-4] + "/Extensions/ICSGateway/ICSGateway.exe"
            subprocess.Popen(gateway_path)
            so.show_status.emit("打开ICS Gateway成功")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("打开ICS Gateway失败")

    def run_update(self):
        """运行ics update plus"""
        ics = self.filename["ICS"]["name"]
        if ics:
            update_path = self.filePath + "/" + ics[:-4] + "/Extensions/IconUpdater/IconUpdater.exe"
            subprocess.Popen(update_path)
            so.show_status.emit("打开Icon Update Plus成功")
        else:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            so.show_status.emit("打开Icon Update Plus失败")

    def copy_ver(self):  # jcywong 2023/11/13
        """复制ics版本号到剪切板"""
        cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()].upper()

        try:
            filename = self.filename[cur_tab_name]["name"]
            QApplication.clipboard().setText(filename)
            so.show_status.emit(f"{cur_tab_name} 版本号已复制到剪贴板：{filename}")
        except KeyError:
            QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
            so.show_status.emit("复制版本号失败")


    def open_soft_path(self):
        """打开路径"""
        cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()].upper()
        try:
            filename = self.filename[cur_tab_name]["name"]
            cur_soft_path = self.filePath + "/" + filename[:-4]
            os.startfile(cur_soft_path)
            so.show_status.emit(f"打开{cur_tab_name}路径成功")
        except KeyError:
            QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
            so.show_status.emit(f"打开{cur_tab_name}路径失败")

    def setProgress(self, value):
        cur_tab_name = tab_map[self.tab_tabMenu.currentIndex()]
        if cur_tab_name == "ics":
            self.ics_progressBar_download.setValue(value)
        elif cur_tab_name == "icc":
            self.icc_progressBar_download.setValue(value)
        elif cur_tab_name == "icm":
            self.icm_progressBar_download.setValue(value)
        elif cur_tab_name == "icp":
            self.icp_progressBar_download.setValue(value)
        elif cur_tab_name == "vp":
            self.vp_progressBar_download.setValue(value)

    def save_config(self):
        """保存配置"""
        # 将配置信息保存到配置文件
        with open(config_file, 'w') as file:
            json.dump(configs, file)

    def load_config(self):
        try:
            with open(config_file, 'r') as file:
                config_data = json.load(file)

                # 从配置中获取信息
                saved_address = config_data.get('save_path', '')
                self.filePath = saved_address
                configs["save_path"] = saved_address

                network = config_data.get('network', '')
                self.network = network
                configs["network"] = network

                filename = config_data.get('filename', '')
                configs["filename"] = filename
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

    def check_version(self):

        def compare_versions(version1, version2):
            # 将版本号转换为整数列表
            parts1 = [int(part) for part in version1.split('.')]
            parts2 = [int(part) for part in version2.split('.')]
            # 比较版本号的每个部分
            for i in range(max(len(parts1), len(parts2))):
                part1 = parts1[i] if i < len(parts1) else 0
                part2 = parts2[i] if i < len(parts2) else 0
                if part1 > part2:
                    return 1
                elif part1 < part2:
                    return -1
            return 0

        ret = get_test_tools_last_version()
        if ret:
            last_version = ret['last-version']

            if compare_versions(last_version, VERSION) == 1:
                self.show_MessageBox(f"Last Version is {last_version}, please update！", "information")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.setWindowIcon(QIcon('tool.png'))
    window.setWindowTitle("TestTools")
    window.show()
    sys.exit(app.exec())

import json
import sys
import threading
import logging
import os
import subprocess
from typing import Dict, List, Optional, Any
from dataclasses import dataclass

import requests.exceptions

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('testtools.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
from PySide6.QtCore import QFile, QIODevice, QObject, Signal, QRegularExpression, Qt, QThread
from PySide6.QtGui import QRegularExpressionValidator, QIcon, QAction
from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QMessageBox, QFileDialog, QLineEdit, QComboBox, \
    QProgressBar, QRadioButton, QStatusBar, QTabWidget, QDialog, QLabel, QVBoxLayout, QMenuBar, QMenu, QHBoxLayout

from comm import *

# 定义版本号
VERSION = "1.4.2"

# 配置类
@dataclass
class AppConfig:
    """应用程序配置类"""
    VERSION: str = VERSION
    CONFIG_FILE: str = 'config.json'
    UI_FILE_NAME: str = "main.ui"
    IP_REGEX_PATTERN: str = r"^(25[0-5]\.?|2[0-4][0-9]\.?|[0-1]?[0-9][0-9]?\.?)$"
    
    # Tab映射
    TAB_MAP: Dict[int, str] = None
    
    def __post_init__(self):
        if self.TAB_MAP is None:
            self.TAB_MAP = {
                0: "ics",
                1: "icc", 
                2: "icm",
                3: "icf",
                4: "icp",
                5: "vp",
            }

# 全局配置实例
config = AppConfig()

class SignalStore(QObject):
    """信号存储类，统一管理所有信号"""
    progress_update = Signal(int)
    download_state = Signal(bool)
    execute_state = Signal(bool)
    show_message = Signal(str, str)
    show_status = Signal(str)

# 全局信号实例
signal_store = SignalStore()

# 全局配置字典
configs = {}

class Constants:
    """常量类"""
    # 消息类型
    MESSAGE_WARNING = "warning"
    MESSAGE_INFORMATION = "information"
    MESSAGE_QUESTION = "question"
    
    # 网络类型
    NETWORK_LAN = "LAN"
    NETWORK_INTERNET = "Internet"
    
    # 版本类型
    EDITION_DEBUG = "Debug"
    EDITION_RELEASE = "Release"
    
    # 命令类型
    COMMAND_REBOOT = "重启"
    COMMAND_GET_LOGS = "获取日志"
    
    # 默认值
    DEFAULT_VERSION = " "
    DEFAULT_NETWORK = "LAN"

class MemoryThread(QThread):
    """内存监控线程"""
    memory_info_updated = Signal(float, float, float)

    def run(self):
        while not self.isInterruptionRequested():
            try:
                memory_info = psutil.virtual_memory()
                ICS_memory = monitor_process_memory()
                used_memory = memory_info.used / (1024 ** 2)  # 转换为MB
                total_memory = memory_info.total / (1024 ** 2)  # 转换为MB

                self.memory_info_updated.emit(ICS_memory, used_memory, total_memory)
                self.sleep(1)  # 每隔1秒更新一次内存信息
            except Exception as e:
                logger.error(f"内存监控错误: {e}")
                self.sleep(5)  # 出错时延长等待时间

    def stop(self):
        """停止线程"""
        self.requestInterruption()
        self.quit()
        self.wait()

class MemoryMonitorDialog(QDialog):
    """内存监控对话框"""
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
        """关闭事件处理"""
        self.memory_thread.stop()
        event.accept()

    def update_memory_info(self, ICS_memory: float, used_memory: float, total_memory: float):
        """更新内存信息显示"""
        if ICS_memory:
            self.label.setText(
                f'ICS Studio: {ICS_memory:.2f} MB   {(ICS_memory / total_memory) * 100:.2f}%\n'
                f'Used: {used_memory:.2f} MB    {(used_memory / total_memory) * 100:.2f}%\n')
        else:
            self.label.setText(
                f'ICS Studio:not found\n'
                f'Used: {used_memory:.2f} MB {(used_memory / total_memory) * 100:.2f}%\n')

class SettingsDialog(QDialog):
    """设置对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setFixedSize(400, 100)
        self._init_ui()

    def _init_ui(self):
        """初始化UI"""
        v_layout = QVBoxLayout(self)
        
        # 保存路径设置
        h_layout_button = QHBoxLayout()
        self.label = QLabel("保存路径：")
        h_layout_button.addWidget(self.label)

        self.save_path_label = QLabel()
        if "save_path" in configs and configs["save_path"]:
            self.save_path_label.setText(configs["save_path"])
        else:
            self.save_path_label.setText("未设置保存路径")
        h_layout_button.addWidget(self.save_path_label)

        self.select_button = QPushButton("选择文件夹")
        self.select_button.clicked.connect(self.select_save_path)
        h_layout_button.addWidget(self.select_button)
        v_layout.addLayout(h_layout_button)

        # 网络选择
        h_layout_network = QHBoxLayout()
        self.label_network = QLabel("网络选择：")
        h_layout_network.addWidget(self.label_network)

        self.radioButton_lan = QRadioButton("局域网")
        self.radioButton_internet = QRadioButton("互联网")
        
        # 设置默认网络选择
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
        h_layout_button = QHBoxLayout()
        self.ok_button = QPushButton("确定")
        self.ok_button.clicked.connect(self.accept)
        h_layout_button.addWidget(self.ok_button)

        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.reject)
        h_layout_button.addWidget(self.cancel_button)
        v_layout.addLayout(h_layout_button)

        self.setLayout(v_layout)

    def select_save_path(self):
        """选择保存路径"""
        save_path = QFileDialog.getExistingDirectory(self, "选择存储路径")
        if save_path:
            self.save_path_label.setText(save_path)

    def accept(self):
        """确认设置"""
        text = self.save_path_label.text()
        configs["save_path"] = text if text != "未设置保存路径" else None
        configs["network"] = "LAN" if self.radioButton_lan.isChecked() else "Internet"
        super().accept()

class TabInitializer:
    """Tab初始化器，用于减少重复代码"""
    
    @staticmethod
    def init_combo_boxes(window, tab_name: str, ver_items: List[str] = None, edition_items: List[str] = ['Debug', 'Release']):
        """初始化下拉框"""
        edition_combo = window.findChild(QComboBox, f"{tab_name}_comboBox_Edition")
        ver_combo = window.findChild(QComboBox, f"{tab_name}_comboBox_ver")
        
        if ver_combo:
            ver_combo.addItems(ver_items)
        
        if edition_combo:
            edition_combo.addItems(edition_items)
            edition_combo.currentIndexChanged.connect(
                lambda: window.selection_change_comboBox_edition(tab_name)
            )
        
        return edition_combo, ver_combo

    @staticmethod
    def init_download_controls(window, tab_name: str):
        """初始化下载控件"""
        download_btn = window.findChild(QPushButton, f"{tab_name}_btn_download")
        progress_bar = window.findChild(QProgressBar, f"{tab_name}_progressBar_download")
        
        if download_btn:
            download_btn.clicked.connect(window.download_soft)
        if progress_bar:
            progress_bar.setRange(0, 2)
        
        return download_btn, progress_bar

    @staticmethod
    def init_action_buttons(window, tab_name: str):
        """初始化操作按钮"""
        buttons = {}

        
        # 复制版本号按钮
        copy_btn = window.findChild(QPushButton, f"{tab_name}_btn_copy_ver")
        if copy_btn:
            copy_btn.clicked.connect(window.copy_ver)
            buttons['copy_ver'] = copy_btn
        
        # 打开路径按钮
        path_btn = window.findChild(QPushButton, f"{tab_name}_btn_path")
        if path_btn:
            path_btn.clicked.connect(window.open_soft_path)
            buttons['path'] = path_btn
        
        return buttons

    @staticmethod
    def init_ip_controls(window, tab_name: str):
        """初始化IP地址控件"""
        ip_parts = []
        for i in range(1, 5):
            ip_edit = window.findChild(QLineEdit, f"{tab_name}_lineEdit_ip{i}")
            if ip_edit:
                ip_parts.append(ip_edit)
        
        # 设置IP验证器
        ip_regex = QRegularExpression(config.IP_REGEX_PATTERN)
        ip_validator = QRegularExpressionValidator(ip_regex)
        
        for ip_part in ip_parts:
            ip_part.textChanged.connect(window.on_ip_part_changed)
            ip_part.setValidator(ip_validator)
        
        return ip_parts

class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        super().__init__()
        self._init_signals()
        self._init_variables()
        self._init_ui()
        self._init_tabs()
        self._init_menu()
        self._load_config()
        self._check_version()

    def _init_signals(self):
        """初始化信号连接"""
        signal_store.progress_update.connect(self.setProgress)
        signal_store.download_state.connect(self.update_download_state)
        signal_store.show_message.connect(self.show_MessageBox)
        signal_store.execute_state.connect(self.update_execute_state)
        signal_store.show_status.connect(self.update_status)

    def _init_variables(self):
        """初始化变量"""
        self.filename: Dict[str, Dict[str, Any]] = {}
        self.network: str = "LAN"
        self.filePath: Optional[str] = None
        self.downloading: bool = False
        self.executing: bool = False
        self.memory_monitor_dialog: Optional[MemoryMonitorDialog] = None

    def _init_ui(self):
        """初始化UI"""
        # 获取当前脚本所在目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ui_file_path = os.path.join(current_dir, config.UI_FILE_NAME)
        
        # 检查UI文件是否存在
        if not os.path.exists(ui_file_path):
            logger.error(f"UI file not found: {ui_file_path}")
            print(f"UI file not found: {ui_file_path}")
            print(f"Current directory: {current_dir}")
            print(f"Available files in current directory: {os.listdir(current_dir)}")
            sys.exit(-1)
        
        logger.info(f"Loading UI file: {ui_file_path}")
        self.ui_file = QFile(ui_file_path)
        if not self.ui_file.open(QIODevice.ReadOnly):
            logger.error(f"Cannot open {ui_file_path}: {self.ui_file.errorString()}")
            print(f"Cannot open {ui_file_path}: {self.ui_file.errorString()}")
            sys.exit(-1)
        
        self.window = QUiLoader().load(self.ui_file)
        self.ui_file.close()
        
        if not self.window:
            error_msg = QUiLoader().errorString()
            logger.error(f"Failed to load UI: {error_msg}")
            print(f"Failed to load UI: {error_msg}")
            sys.exit(-1)

        logger.info("UI loaded successfully")
        self.setCentralWidget(self.window)
        self.statusbar = self.window.findChild(QStatusBar, "statusbar")
        self.tab_tabMenu = self.window.findChild(QTabWidget, "tab")

    def _init_tabs(self):
        """初始化所有标签页"""
        self._init_tab_ics()
        self._init_tab_icc()
        self._init_tab_icm()
        self._init_tab_icf()
        self._init_tab_icp()
        self._init_tab_visual_pro()

    def _init_menu(self):
        """初始化菜单"""
        self.menubar = self.window.findChild(QMenuBar, "menuBar")
        
        # 工具菜单
        self.tool_menu = self.window.findChild(QMenu, "tool")
        self.memory_action = self.window.findChild(QAction, "action_memory")
        self.memory_action.triggered.connect(self.open_memory_monitor)
        
        # 帮助菜单
        self.help_menu = self.window.findChild(QMenu, "help")
        self.ver_action = self.window.findChild(QAction, "action_ver")
        self.settings_action = self.window.findChild(QAction, "action_settings")
        self.ver_action.triggered.connect(self.show_version)
        self.settings_action.triggered.connect(self.open_settings)

    def _init_tab_ics(self):
        """初始化ICS标签页"""
        # 初始化下拉框
        edition_combo, ver_combo = TabInitializer.init_combo_boxes(self, "ics", [' ', 'v1.2', 'v1.3', 'v1.4', 'v1.5', 'v1.6'])
        self.ics_comboBox_Edition = edition_combo
        self.ics_comboBox_ver = ver_combo
        
        # 初始化下载控件
        download_btn, progress_bar = TabInitializer.init_download_controls(self, "ics")
        self.ics_btn_download = download_btn
        self.ics_progressBar_download = progress_bar
        
        # 初始化操作按钮
        buttons = TabInitializer.init_action_buttons(self, "ics")
        self.btn_copy_ics_ver = buttons.get('copy_ver')
        self.btn_ics_path = buttons.get('path')
        
        # 特殊按钮
        self.btn_run_ics = self.window.findChild(QPushButton, "ics_btn_run_ics")
        self.btn_run_gateway = self.window.findChild(QPushButton, "ics_btn_run_gateway")
        self.btn_run_update = self.window.findChild(QPushButton, "ics_btn_run_update")
        
        if self.btn_run_ics:
            self.btn_run_ics.clicked.connect(self.run_soft)
        if self.btn_run_gateway:
            self.btn_run_gateway.clicked.connect(self.run_gateway)
        if self.btn_run_update:
            self.btn_run_update.clicked.connect(self.run_update)

    def _init_tab_icc(self):
        """初始化ICC标签页"""
        # 初始化下拉框
        edition_combo, ver_combo = TabInitializer.init_combo_boxes(self, "icc", [' ', 'v1.2', 'v1.3', 'v1.4', 'v1.5', 'v1.6'])
        self.icc_comboBox_Edition = edition_combo
        self.icc_comboBox_ver = ver_combo

        # 特殊下拉框
        self.icc_comboBox_model_1 = self.window.findChild(QComboBox, "icc_comboBox_model_1")
        if self.icc_comboBox_model_1:
            self.icc_comboBox_model_1.addItems(['LITE', 'LITE.B', 'PRO', "PRO.B", 'TURBO', 'EVO'])
        
        # 初始化下载控件
        download_btn, progress_bar = TabInitializer.init_download_controls(self, "icc")
        self.icc_btn_download = download_btn
        self.icc_progressBar_download = progress_bar
        
        # 初始化操作按钮
        buttons = TabInitializer.init_action_buttons(self, "icc")
        self.icc_btn_copy_ver = buttons.get('copy_ver')
        self.icc_btn_path = buttons.get('path')

        # 特殊按钮
        self.icc_btn_run_update = self.window.findChild(QPushButton, "icc_btn_run_update")
        if self.icc_btn_run_update:
            self.icc_btn_run_update.clicked.connect(self.run_update)
        
        # 初始化IP控件
        self.icc_ip_parts = TabInitializer.init_ip_controls(self, "icc")
        
        # 初始化控制控件
        self.icc_comboBox_model_2 = self.window.findChild(QComboBox, "icc_comboBox_model_2")
        self.icc_comboBox_command = self.window.findChild(QComboBox, "icc_comboBox_command")
        self.icc_pushButton_execute = self.window.findChild(QPushButton, "icc_pushButton_execute")
        
        if self.icc_comboBox_model_2:
            self.icc_comboBox_model_2.addItems(['LITE', 'LITE.B', 'PRO', "PRO.B", 'TURBO', 'EVO'])
        if self.icc_comboBox_command:
            self.icc_comboBox_command.addItems([" ", '重启', "获取日志"])
        if self.icc_pushButton_execute:
            self.icc_pushButton_execute.clicked.connect(self.execute_command)

    def _init_tab_icm(self):
        """初始化ICM标签页"""
        # 初始化下拉框
        edition_combo, ver_combo = TabInitializer.init_combo_boxes(self, "icm", [" "], [' '])
        self.icm_comboBox_Edition = edition_combo
        self.icm_comboBox_ver = ver_combo

        # 特殊下拉框
        self.icm_comboBox_model_1 = self.window.findChild(QComboBox, "icm_comboBox_model_1")
        if self.icm_comboBox_model_1:
            self.icm_comboBox_model_1.addItems(['Release', 'ICM-D1', 'ICM-D3', "ICM-D5", 'ICM-D7'])
        
        
        # 初始化下载控件
        download_btn, progress_bar = TabInitializer.init_download_controls(self, "icm")
        self.icm_btn_download = download_btn
        self.icm_progressBar_download = progress_bar
        
        # 初始化操作按钮
        buttons = TabInitializer.init_action_buttons(self, "icm")
        self.icm_btn_copy_ver = buttons.get('copy_ver')
        self.icm_btn_path = buttons.get('path')

        # 特殊按钮
        self.icm_btn_run_update = self.window.findChild(QPushButton, "icm_btn_run_update")
        if self.icm_btn_run_update:
            self.icm_btn_run_update.clicked.connect(self.run_update)
        
        # 初始化IP控件
        self.icm_ip_parts = TabInitializer.init_ip_controls(self, "icm")
        
        # 初始化控制控件
        self.icm_comboBox_model_2 = self.window.findChild(QComboBox, "icm_comboBox_model_2")
        self.icm_comboBox_command = self.window.findChild(QComboBox, "icm_comboBox_command")
        self.icm_pushButton_execute = self.window.findChild(QPushButton, "icm_pushButton_execute")
        
        if self.icm_comboBox_model_2:
            self.icm_comboBox_model_2.addItems(['ICM-D1', 'ICM-D3', "ICM-D5", 'ICM-D7'])
        if self.icm_comboBox_command:
            self.icm_comboBox_command.addItems([" ", '重启', "获取日志"])
        if self.icm_pushButton_execute:
            self.icm_pushButton_execute.clicked.connect(self.execute_command)

    def _init_tab_icf(self):
        """初始化ICF标签页"""
        # 初始化下拉框
        edition_combo, ver_combo = TabInitializer.init_combo_boxes(self, "icf", [''])
        self.icf_comboBox_Edition = edition_combo
        self.icf_comboBox_ver = ver_combo

        # 特殊下拉框
        self.icf_comboBox_model_1 = self.window.findChild(QComboBox, "icf_comboBox_model_1")
        if self.icf_comboBox_model_1:
            self.icf_comboBox_model_1.addItems(['C1S4T013B', 'C1S2S7R5B', 'C1S2S2R8N', 'C1S2S4R6N'])

        # 初始化下载控件
        download_btn, progress_bar = TabInitializer.init_download_controls(self, "icf")
        self.icf_btn_download = download_btn
        self.icf_progressBar_download = progress_bar

        # 初始化操作按钮
        buttons = TabInitializer.init_action_buttons(self, "icf")
        self.icf_btn_copy_ver = buttons.get('copy_ver')
        self.icf_btn_path = buttons.get('path')

        # 特殊按钮
        self.icf_btn_run_update = self.window.findChild(QPushButton, "icf_btn_run_update")
        if self.icf_btn_run_update:
            self.icf_btn_run_update.clicked.connect(self.run_update)

        # 初始化IP控件
        self.icf_ip_parts = TabInitializer.init_ip_controls(self, "icf")

        # 初始化控制控件
        self.icf_comboBox_model_2 = self.window.findChild(QComboBox, "icf_comboBox_model_2")
        self.icf_comboBox_command = self.window.findChild(QComboBox, "icf_comboBox_command")
        self.icf_pushButton_execute = self.window.findChild(QPushButton, "icf_pushButton_execute")
        
        if self.icf_comboBox_model_2:
            self.icf_comboBox_model_2.addItems(['ICF-C'])
        if self.icf_comboBox_command:
            self.icf_comboBox_command.addItems([" ", '重启', "获取日志"])
        if self.icf_pushButton_execute:
            self.icf_pushButton_execute.clicked.connect(self.execute_command)

    def _init_tab_visual_pro(self):
        """初始化Visual Pro标签页"""
        # 初始化下拉框
        edition_combo, ver_combo = TabInitializer.init_combo_boxes(self, "vp", [' '])
        self.vp_comboBox_Edition = edition_combo
        self.vp_comboBox_ver = ver_combo
        
        # 初始化下载控件
        download_btn, progress_bar = TabInitializer.init_download_controls(self, "vp")
        self.vp_btn_download = download_btn
        self.vp_progressBar_download = progress_bar
        
        # 初始化操作按钮
        buttons = TabInitializer.init_action_buttons(self, "vp")
        self.vp_btn_copy_ver = buttons.get('copy_ver')
        self.vp_btn_path = buttons.get('path')

        # 特殊按钮
        self.vp_btn_run = self.window.findChild(QPushButton, "vp_btn_run")
        if self.vp_btn_run:
            self.vp_btn_run.clicked.connect(self.run_soft)

    def _init_tab_icp(self):
        """初始化ICP标签页"""
        # 初始化下拉框
        edition_combo, ver_combo = TabInitializer.init_combo_boxes(self, "icp", [' '])
        self.icp_comboBox_Edition = edition_combo
        self.icp_comboBox_ver = ver_combo
        
        # 初始化下载控件
        download_btn, progress_bar = TabInitializer.init_download_controls(self, "icp")
        self.icp_btn_download = download_btn
        self.icp_progressBar_download = progress_bar
        
        # 初始化操作按钮
        buttons = TabInitializer.init_action_buttons(self, "icp")
        self.icp_btn_copy_ver = buttons.get('copy_ver')
        self.icp_btn_path = buttons.get('path')

        # 特殊按钮
        self.icp_btn_run_update = self.window.findChild(QPushButton, "icp_btn_run_update")
        if self.icp_btn_run_update:
            self.icp_btn_run_update.clicked.connect(self.run_update)
        
        # 初始化IP控件
        self.icp_ip_parts = TabInitializer.init_ip_controls(self, "icp")
        
        # 初始化控制控件
        self.icp_comboBox_model = self.window.findChild(QComboBox, "icp_comboBox_model")
        self.icp_comboBox_command = self.window.findChild(QComboBox, "icp_comboBox_command")
        self.icp_pushButton_execute = self.window.findChild(QPushButton, "icp_pushButton_execute")
        
        if self.icp_comboBox_model:
            self.icp_comboBox_model.addItems(['ICP'])
        if self.icp_comboBox_command:
            self.icp_comboBox_command.addItems([" ", '重启'])
        if self.icp_pushButton_execute:
            self.icp_pushButton_execute.clicked.connect(self.execute_command)

    def on_ip_part_changed(self, text: str):
        """当IP地址输入'.'则跳转下一个"""
        current_line_edit = self.sender()
        cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()]
        
        if text.endswith('.'):
            current_line_edit.setText(text[:-1])
            
            # 根据当前标签页获取对应的IP部件列表
            ip_parts_map = {
                "icc": self.icc_ip_parts,
                "icp": self.icp_ip_parts,
                "icm": self.icm_ip_parts
            }
            
            if cur_tab_name in ip_parts_map:
                ip_parts = ip_parts_map[cur_tab_name]
                index = ip_parts.index(current_line_edit)
                if index < len(ip_parts) - 1:
                    next_line_edit = ip_parts[index + 1]
                    next_line_edit.setFocus()

    def open_memory_monitor(self):
        """打开内存监控对话框"""
        self.memory_monitor_dialog = MemoryMonitorDialog()
        self.memory_monitor_dialog.show()

        screen_geometry = QApplication.primaryScreen().geometry()
        widget_rect = self.memory_monitor_dialog.geometry()

        # 计算右下角坐标
        x = screen_geometry.width() - widget_rect.width()
        y = screen_geometry.height() - widget_rect.height()

        self.memory_monitor_dialog.move(x, y - 80)

    def open_settings(self):
        """打开设置对话框"""
        settings_dialog = SettingsDialog()
        if settings_dialog.exec():
            self.filePath = configs.get("save_path")
            self.network = configs.get("network")

    def show_version(self):
        """显示版本信息对话框"""
        try:
            ret = get_test_tools_last_version()
            last_version = ret.get("last-version") if ret else None
            message = f'''
                <p>Author: Jcy Wong</p>
                <p>Version: {config.VERSION}</p>
                <p>Last Version: {last_version}</p>
                <p>
                <a href="https://jcywong.notion.site/TestTools-0d910d4d26ad44d5b813119b59f8dae7">User Manual</a>
                </p>
            '''

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setTextFormat(Qt.RichText)
            msg.setText(message)
            msg.setWindowTitle("信息")
            msg.exec()
        except Exception as e:
            print(f"显示版本信息时出错: {e}")

    def update_status(self, status: str):
        """更新状态栏"""
        self.statusbar.showMessage(status)

    def execute_command(self):
        """执行命令"""
        def is_ip_address_empty(ip_parts: List[QLineEdit]) -> bool:
            """检查IP地址是否为空"""
            return any(not ip_part.text() for ip_part in ip_parts)

        def worker_thread_func():
            self.executing = True
            cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()]

            try:
                # 使用getattr安全地获取属性，避免使用eval
                command_combo = getattr(self, f"{cur_tab_name}_comboBox_command", None)
                
                # 特殊处理 ICM 标签页的 model 控件
                if cur_tab_name in ["icm", "icc", "icf"]:
                    model_combo = getattr(self, f"{cur_tab_name}_comboBox_model_2", None)
                else:
                    model_combo = getattr(self, f"{cur_tab_name}_comboBox_model", None)
                    
                ip_parts = getattr(self, f"{cur_tab_name}_ip_parts", None)
                
                if not all([command_combo, model_combo, ip_parts]):
                    signal_store.show_message.emit("控件初始化失败", "warning")
                    self.executing = False
                    signal_store.execute_state.emit(self.executing)
                    return
                
                command = command_combo.currentText()
                model = model_combo.currentText()

                if command == Constants.DEFAULT_VERSION:
                    signal_store.show_message.emit("请选择执行的命令", Constants.MESSAGE_WARNING)
                    self.executing = False
                    signal_store.execute_state.emit(self.executing)
                    return
                elif is_ip_address_empty(ip_parts):
                    signal_store.show_message.emit("请输入IP地址", Constants.MESSAGE_WARNING)
                    self.executing = False
                    signal_store.execute_state.emit(self.executing)
                    return
                else:
                    signal_store.execute_state.emit(self.executing)
                    signal_store.show_status.emit("正在执行命令")

                ip_address = ".".join(ip_part.text() for ip_part in ip_parts)

                if command == Constants.COMMAND_REBOOT:
                    if not reboot_device(device_model=model, ip=ip_address):
                        self.executing = False
                        signal_store.execute_state.emit(self.executing)
                        signal_store.show_status.emit("命令执行失败")
                        return
                elif command == Constants.COMMAND_GET_LOGS:
                    if not self.filePath:
                        self.executing = False
                        signal_store.execute_state.emit(self.executing)
                        signal_store.show_message.emit("请设置保存地址", Constants.MESSAGE_WARNING)
                        signal_store.show_status.emit("命令执行失败")
                        return

                    if not get_device_logs(model, self.filePath, ip_address):
                        self.executing = False
                        signal_store.execute_state.emit(self.executing)
                        signal_store.show_status.emit("命令执行失败")
                        return
                    else:
                        # 打开文件夹
                        normalized_path = os.path.normpath(f"{self.filePath}/logs")
                        os.startfile(normalized_path)

                self.executing = False
                signal_store.execute_state.emit(self.executing)
                signal_store.show_status.emit("命令执行完成")
                
            except Exception as e:
                logger.error(f"执行命令时出错: {e}")
                self.executing = False
                signal_store.execute_state.emit(self.executing)
                signal_store.show_status.emit("命令执行失败")

        if self.executing:
            QMessageBox.warning(self.window, '警告', '任务进行中，请等待完成')
            return

        worker = threading.Thread(target=worker_thread_func)
        worker.start()

    def open_path(self):
        """打开存储路径"""
        if self.filePath and os.path.exists(self.filePath):
            # 使用 os.path.normpath 标准化路径
            normalized_path = os.path.normpath(self.filePath)
            os.startfile(normalized_path)
            signal_store.show_status.emit("打开保存路径成功")
        else:
            QMessageBox.warning(self.window, '警告', '未选择保存路径或路径不存在')
            signal_store.show_status.emit("打开保存路径失败")

    def selection_change_comboBox_edition(self, type_name: str):
        """处理版本选择变化"""
        if not isinstance(self.ics_comboBox_Edition, QComboBox):
            return
            
        combo_map = {
            "ics": (self.ics_comboBox_Edition, self.ics_comboBox_ver),
            "icc": (self.icc_comboBox_Edition, self.icc_comboBox_ver),
            "icm": (self.icm_comboBox_Edition, self.icm_comboBox_ver),
            "icp": (self.icp_comboBox_Edition, self.icp_comboBox_ver),
            "vp": (self.vp_comboBox_Edition, self.vp_comboBox_ver),
        }
        
        if type_name in combo_map:
            edition_combo, ver_combo = combo_map[type_name]
            if edition_combo and ver_combo:
                if edition_combo.currentText() == "Release":
                    ver_combo.setEnabled(True)
                else:
                    ver_combo.setCurrentText(" ")
                    ver_combo.setEnabled(False)

    def update_download_state(self, state: bool):
        """更新下载状态"""
        current_index = self.tab_tabMenu.currentIndex()
        cur_tab_name = config.TAB_MAP[current_index]

        # 禁用/启用其他标签页
        for i in range(self.tab_tabMenu.count()):
            self.tab_tabMenu.setTabEnabled(i, not state or i == current_index)

        # 根据当前标签页更新控件状态
        self._update_tab_controls_state(cur_tab_name, state)

    def _update_tab_controls_state(self, tab_name: str, disabled: bool):
        """更新标签页控件状态"""
        control_map = {
            "ics": [
                (self.ics_comboBox_Edition, True),
                (self.ics_comboBox_ver, self.ics_comboBox_Edition.currentText() == "Release" if not disabled else False)
            ],
            "icc": [
                (self.icc_comboBox_model_1, True),
                (self.icc_comboBox_Edition, True),
                (self.icc_comboBox_ver, self.icc_comboBox_Edition.currentText() == "Release" if not disabled else False)
            ],
            "icp": [
                (self.icp_comboBox_Edition, False),
                (self.icp_comboBox_ver, False)  # ICP的版本选择总是禁用
            ],
            "icm": [
                (self.icm_comboBox_model_1, False),
                (self.icm_comboBox_Edition, False),
                (self.icm_comboBox_ver, self.icm_comboBox_Edition.currentText() == "Release" if not disabled else False)
            ],
            "vp": [
                (self.vp_comboBox_Edition, False),
                (self.vp_comboBox_ver, self.vp_comboBox_Edition.currentText() == "Release" if not disabled else False)
            ],
        }
        
        if tab_name in control_map:
            for control, enabled in control_map[tab_name]:
                if control:
                    control.setEnabled(not disabled and enabled)

    def update_execute_state(self, state: bool):
        """更新执行状态"""
        self.executing = state
        
        # 更新所有IP输入框和控制控件
        all_ip_parts = []
        if hasattr(self, 'icc_ip_parts'):
            all_ip_parts.extend(self.icc_ip_parts)
        if hasattr(self, 'icm_ip_parts'):
            all_ip_parts.extend(self.icm_ip_parts)
        if hasattr(self, 'icp_ip_parts'):
            all_ip_parts.extend(self.icp_ip_parts)
            
        for ip_part in all_ip_parts:
            ip_part.setEnabled(not state)
            
        # 更新控制控件
        control_combos = [
            self.icc_comboBox_model_2, self.icm_comboBox_model_2, self.icp_comboBox_model,
            self.icc_comboBox_command, self.icm_comboBox_command, self.icp_comboBox_command
        ]
        
        for combo in control_combos:
            if combo:
                combo.setEnabled(not state)

    def show_MessageBox(self, message: str, message_type: str):
        """显示消息框"""
        message_map = {
            "warning": QMessageBox.warning,
            "question": QMessageBox.question,
            "information": QMessageBox.information
        }
        
        if message_type in message_map:
            return message_map[message_type](self.window, "提示", message)
        return None

    def download_soft(self):
        """下载软件"""
        def worker_thread_func():
            self.downloading = True
            cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()]

            try:
                model = None
                if cur_tab_name in ["icc", "icf"]:
                    model = getattr(self, f"{cur_tab_name}_comboBox_model_1", None)
                    if model:
                        model = model.currentText()

                # 使用getattr安全地获取属性，避免使用eval
                edition_combo = getattr(self, f"{cur_tab_name}_comboBox_Edition", None)
                ver_combo = getattr(self, f"{cur_tab_name}_comboBox_ver", None)
                
                if not edition_combo or not ver_combo:
                    signal_store.show_message.emit("控件初始化失败", "warning")
                    self.downloading = False
                    signal_store.download_state.emit(self.downloading)
                    return
                
                edition = edition_combo.currentText()
                ver = ver_combo.currentText()

                if not self.filePath:
                    signal_store.show_message.emit("请设置文件保存地址", "warning")
                    self.downloading = False
                    signal_store.download_state.emit(self.downloading)
                    return
                elif edition == "Release" and ver == " ":
                    signal_store.show_message.emit("请设置Release版本", "warning")
                    self.downloading = False
                    signal_store.download_state.emit(self.downloading)
                    return
                else:
                    signal_store.download_state.emit(self.downloading)
                    signal_store.show_status.emit("正在下载中")

                # 重置文件名状态
                if self.filename:
                    for soft_type, infos in self.filename.items():
                        infos["is_checked"] = False
                else:
                    self.filename = {}

                # 获取最新文件名
                self.filename[cur_tab_name.upper()] = {
                    "name": get_latest_filename(
                        soft_type=cur_tab_name.upper(), 
                        edition=edition, 
                        model=model, 
                        ver=ver,
                        network=self.network
                    ),
                    "is_checked": True
                }

                signal_store.progress_update.emit(1)

                # 下载和解压文件
                for soft_type, infos in self.filename.items():
                    if not infos["is_checked"]:
                        continue

                    name = infos["name"]
                    file_path = os.path.join(self.filePath, name)
                    
                    # 下载文件
                    try:
                        download_file(
                            file_name=name, 
                            file_save_path=self.filePath, 
                            soft_type=soft_type,
                            edition=edition,
                            network=self.network
                        )
                        signal_store.show_status.emit(f"{name}:文件下载成功")
                    except FileExistsError:
                        signal_store.show_message.emit(f"{name}:文件已经存在！", "information")
                        signal_store.show_status.emit(f"{name}:文件已经存在！")
                    except Exception as e:
                        logger.error(f"下载文件时出错: {e}")
                        raise e
                    
                    # 解压文件（无论是否已下载）
                    try:
                        unzip_file(self.filePath, name)
                        signal_store.show_status.emit(f"{name}:文件解压成功")
                    except FileExistsError:
                        signal_store.show_message.emit(f"{name}:文件已经解压！", "information")
                        signal_store.show_status.emit(f"{name}:文件已经解压！")
                    except Exception as e:
                        logger.error(f"解压文件时出错: {e}")
                        # 解压失败不影响下载完成状态
                        signal_store.show_status.emit(f"{name}:文件解压失败")

                self.downloading = False
                signal_store.progress_update.emit(2)
                signal_store.show_status.emit(f"{self.filename}下载完成")
                signal_store.download_state.emit(self.downloading)
                
                # 保存 filename 到配置中
                configs["filename"] = self.filename

            except requests.exceptions.ConnectionError:
                self.downloading = False
                signal_store.progress_update.emit(0)
                signal_store.show_status.emit("网络错误，下载失败！")
                signal_store.download_state.emit(self.downloading)
                signal_store.show_message.emit("网络错误，下载失败！", "warning")
            except Exception as e:
                logger.error(f"下载过程中出错: {e}")
                self.downloading = False
                signal_store.progress_update.emit(0)
                signal_store.show_status.emit("网络错误/文件不存在，下载失败！")
                signal_store.download_state.emit(self.downloading)
                signal_store.show_message.emit("网络错误/文件不存在，下载失败！", "warning")

        if self.downloading:
            QMessageBox.warning(self.window, '警告', '任务进行中，请等待完成')
            return

        worker = threading.Thread(target=worker_thread_func)
        worker.start()

    def run_soft(self):
        """运行软件"""
        try:
            cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()].upper()
            filename = self.filename[cur_tab_name]["name"]
            if filename and len(filename) > 4:
                # 使用 os.path.normpath 标准化路径
                path = os.path.normpath(os.path.join(self.filePath, filename[:-4]))
                if cur_tab_name == "ICS":
                    if os.path.exists(path):
                        open_ics(path)
                        signal_store.show_status.emit(f"打开{cur_tab_name}成功")
                    else:
                        QMessageBox.warning(self.window, "警告", f"路径不存在: {path}")
                        signal_store.show_status.emit(f"打开{cur_tab_name}失败")
                elif cur_tab_name == "VP":
                    exe_path = os.path.normpath(os.path.join(path, "VisualPro.exe"))
                    if os.path.exists(exe_path):
                        subprocess.Popen(exe_path)
                        signal_store.show_status.emit(f"打开{cur_tab_name}成功")
                    else:
                        QMessageBox.warning(self.window, "警告", f"文件不存在: {exe_path}")
                        signal_store.show_status.emit(f"打开{cur_tab_name}失败")
            else:
                QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
                signal_store.show_status.emit(f"打开{cur_tab_name}失败")
        except KeyError:
            QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
            signal_store.show_status.emit(f"打开{cur_tab_name}失败")

    def run_gateway(self):
        """运行ICS Gateway"""
        try:
            ics = self.filename["ICS"]["name"]
            if ics and len(ics) > 4:
                # 使用 os.path.normpath 标准化路径
                gateway_path = os.path.normpath(os.path.join(self.filePath, ics[:-4], "Extensions", "ICSGateway", "ICSGateway.exe"))
                if os.path.exists(gateway_path):
                    subprocess.Popen(gateway_path)
                    signal_store.show_status.emit("打开ICS Gateway成功")
                else:
                    QMessageBox.warning(self.window, "警告", f"文件不存在: {gateway_path}")
                    signal_store.show_status.emit("打开ICS Gateway失败")
            else:
                QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
                signal_store.show_status.emit("打开ICS Gateway失败")
        except KeyError:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            signal_store.show_status.emit("打开ICS Gateway失败")

    def run_update(self):
        """运行ICS Update Plus"""
        try:
            ics = self.filename["ICS"]["name"]
            if ics and len(ics) > 4:
                # 使用 os.path.normpath 标准化路径
                update_path = os.path.normpath(os.path.join(self.filePath, ics[:-4], "Extensions", "IconUpdater", "IconUpdater.exe"))
                if os.path.exists(update_path):
                    subprocess.Popen(update_path)
                    signal_store.show_status.emit("打开Icon Update Plus成功")
                else:
                    QMessageBox.warning(self.window, "警告", f"文件不存在: {update_path}")
                    signal_store.show_status.emit("打开Icon Update Plus失败")
            else:
                QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
                signal_store.show_status.emit("打开Icon Update Plus失败")
        except KeyError:
            QMessageBox.information(self.window, "提示", "最近未下载最新ICS Studio")
            signal_store.show_status.emit("打开Icon Update Plus失败")

    def copy_ver(self):
        """复制版本号到剪贴板"""
        cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()].upper()

        try:
            filename = self.filename[cur_tab_name]["name"]
            QApplication.clipboard().setText(filename)
            signal_store.show_status.emit(f"{cur_tab_name} 版本号已复制到剪贴板：{filename}")
        except KeyError:
            QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
            signal_store.show_status.emit("复制版本号失败")

    def open_soft_path(self):
        """打开软件路径"""
        cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()].upper()
        try:
            filename = self.filename[cur_tab_name]["name"]
            if filename and len(filename) > 4:
                # 使用 os.path.normpath 标准化路径，避免混合正斜杠和反斜杠
                cur_soft_path = os.path.normpath(os.path.join(self.filePath, filename[:-4]))
                if os.path.exists(cur_soft_path):
                    os.startfile(cur_soft_path)
                    signal_store.show_status.emit(f"打开{cur_tab_name}路径成功")
                else:
                    QMessageBox.warning(self.window, "警告", f"路径不存在: {cur_soft_path}")
                    signal_store.show_status.emit(f"打开{cur_tab_name}路径失败")
            else:
                QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
                signal_store.show_status.emit(f"打开{cur_tab_name}路径失败")
        except KeyError:
            QMessageBox.information(self.window, "提示", f"最近未下载最新{cur_tab_name}")
            signal_store.show_status.emit(f"打开{cur_tab_name}路径失败")

    def setProgress(self, value: int):
        """设置进度条"""
        cur_tab_name = config.TAB_MAP[self.tab_tabMenu.currentIndex()]
        progress_bars = {
            "ics": self.ics_progressBar_download,
            "icc": self.icc_progressBar_download,
            "icm": self.icm_progressBar_download,
            "icp": self.icp_progressBar_download,
            "vp": self.vp_progressBar_download,
        }
        
        if cur_tab_name in progress_bars and progress_bars[cur_tab_name]:
            progress_bars[cur_tab_name].setValue(value)

    def save_config(self):
        """保存配置"""
        try:
            with open(config.CONFIG_FILE, 'w', encoding='utf-8') as file:
                json.dump(configs, file, ensure_ascii=False, indent=2)
            logger.info("配置保存成功")
        except Exception as e:
            logger.error(f"保存配置时出错: {e}")

    def _load_config(self):
        """加载配置"""
        try:
            with open(config.CONFIG_FILE, 'r', encoding='utf-8') as file:
                config_data = json.load(file)

                # 从配置中获取信息
                self.filePath = config_data.get('save_path', '')
                configs["save_path"] = self.filePath

                self.network = config_data.get('network', '')
                configs["network"] = self.network

                # 正确处理 filename 数据
                filename_data = config_data.get('filename', {})
                configs["filename"] = filename_data
                # 确保 self.filename 是字典类型
                if isinstance(filename_data, dict):
                    self.filename = filename_data
                else:
                    self.filename = {}
                
            logger.info("配置加载成功")
        except FileNotFoundError:
            # 如果配置文件不存在，不做任何操作
            logger.info("配置文件不存在，使用默认配置")
        except Exception as e:
            logger.error(f"加载配置时出错: {e}")

    def closeEvent(self, event):
        """窗口关闭事件"""
        if self.memory_monitor_dialog:
            self.memory_monitor_dialog.close()
        self.save_config()
        event.accept()

    def _check_version(self):
        """检查版本更新"""
        def compare_versions(version1: str, version2: str) -> int:
            """比较版本号"""
            try:
                parts1 = [int(part) for part in version1.split('.')]
                parts2 = [int(part) for part in version2.split('.')]
                
                for i in range(max(len(parts1), len(parts2))):
                    part1 = parts1[i] if i < len(parts1) else 0
                    part2 = parts2[i] if i < len(parts2) else 0
                    if part1 > part2:
                        return 1
                    elif part1 < part2:
                        return -1
                return 0
            except Exception:
                return 0

        try:
            ret = get_test_tools_last_version()
            if ret:
                last_version = ret.get('last-version')
                if last_version and compare_versions(last_version, config.VERSION) == 1:
                    self.show_MessageBox(f"Last Version is {last_version}, please update！", "information")
        except Exception as e:
            print(f"检查版本时出错: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    
    # 设置图标，尝试多个路径
    icon_paths = [
        'tool.png',  # 当前目录
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tool.png'),  # src目录
        os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'img', 'tool.png'),  # img目录
    ]
    
    for icon_path in icon_paths:
        if os.path.exists(icon_path):
            window.setWindowIcon(QIcon(icon_path))
            break
    
    window.setWindowTitle("TestTools")
    window.show()
    sys.exit(app.exec())

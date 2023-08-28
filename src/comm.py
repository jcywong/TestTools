import datetime
import os
import subprocess
import time

import uiautomation as auto

# 设置全局搜索时间
auto.uiautomation.SetGlobalSearchTimeout(15)

# 目录信息

workspace = os.getcwd()

soft_path = workspace + '\\soft\\'

# json_path = os.path.dirname(workspace) + '\\json\\'
json_path = workspace + "\\json\\"

# ics窗口
ics_window = auto.WindowControl(SubName='ICS Studio', ClassName='Window', AutomationId='VisualStudioMainWindow')


def open_ics(path):
    # 打开ics

    ics_path = f'{path}\\ICSStudio.exe'
    if ics_window.Exists():
        close_window(ics_window)
        time.sleep(5)
        subprocess.Popen(ics_path)
    else:
        subprocess.Popen(ics_path)
    time.sleep(10)
    ics_window.SetTopmost(True)
    print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成打开ICS并置顶')

    # 最大化ics
    if not ics_window.IsMaximize():
        ics_window.ButtonControl(AutomationId='Maximize', Name='Maximize').Click()
    print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成ICS最大化')


def close_window(window):
    # 关闭ics
    if window.Exists():
        exe = None
        name = window.Name
        try:
            # 使用 taskkill 命令终止进程
            if window == ics_window:
                exe = "ICSStudio.exe"
            taskkill_cmd = f"taskkill /F /IM {exe}"
            subprocess.check_output(taskkill_cmd, shell=True)
            print(f"{name}已成功关闭。")

        except subprocess.CalledProcessError:
            print(f"{name}进程不存在或无法终止。")
    else:
        print(f"进程不存在.")


def set_language(language=2, model=1):
    # 设置语言
    # language 1 为英语 2 为中文（默认）
    # model 1 为快捷键操作，2为点击操作
    help_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
        MenuItemControl(Name='Help', ClassName='MenuItem')
    if model == 1:
        # 方式一： 快捷键
        ics_window.SendKeys('{Alt}h')
    elif model == 2:
        help_menu.Click()
    language_menu = help_menu.MenuItemControl(Name='Language')
    if language == 1:
        if not language_menu.Exists():
            help_menu.MenuItemControl(Name='语言').Click()
            help_menu.MenuItemControl(Name='语言').MenuItemControl(Name='English').Click()
        else:
            ics_window.Click()
    elif language == 2:
        if language_menu.Exists():
            language_menu.Click()
            language_menu.MenuItemControl(Name='简体中文').Click()
        else:
            ics_window.Click()


def open_json_file(json_filename='test.json', model=1):
    """
    打开json文件
    :param json_filename: json名称
    :param model: 1 为快捷方式测试；2 为页面点击方式测试
    :return:
    """

    # 方式一： 快捷键
    if model == 1:
        ics_window.SendKeys('{Ctrl}o')
    elif model == 2:
        # 方式二： 点击选择信息
        file_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
            MenuItemControl(Name='File', ClassName='MenuItem')
        file_menu.Click()
        file_menu.MenuItemControl(Name='打开...').Click()

    # 处理文件对话框
    import_dig = auto.WindowControl(Name='Import file')
    time.sleep(1)
    address_bar = import_dig.Control(AutomationId='1001', ClassName='ToolbarWindow32')
    up_button = import_dig.Control(ClassName='UpBand')

    for i in range(len(address_bar.GetChildren())-1):
        up_button.Click()
    address_bar.Click()
    auto.SendKeys("{Ctrl}a")  # 模拟按下 Ctrl + A 快捷键，全选地址栏内容
    auto.SendKeys(json_path)  # 输入新的地址路径
    auto.SendKeys('{Enter}')
    time.sleep(1)

    # 选择文件test
    file_item = import_dig.ListControl(Name='项目视图', ClassName='UIItemsView').ListItemControl(Name=json_filename)
    if file_item.Exists():
        file_item.Click()

    import_dig.ButtonControl(Name='打开(O)').Click()

    # 程序加载中 todo 需等确认窗口再修改
    while True:
        loading_dig = ics_window.WindowControl(Name='Downloading', ClassName='Window')
        if not loading_dig.Exists():
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:程序加载完成')
            break


def net_path(plc_ip='192.168.1.211', model=1):
    """
    网络路径中的操作
    :param plc_ip: ip地址
    :param model: 1 为快捷方式测试 2 为页面点击方式测试
    :return:
    """
    communications_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
        MenuItemControl(Name='Communications', ClassName='MenuItem')

    if model == 1:
        # 方式一： 快捷键
        ics_window.SendKeys('{Alt}c')
    elif model == 2:
        # 方式二： 点击选择信息
        communications_menu.Click()

    # 网络路径设置
    communications_menu.MenuItemControl(Name='网络路径').Click()
    net_dlg = ics_window.WindowControl(Name='网络路径')

    time.sleep(5)
    tree = net_dlg.TreeControl(ClassName='TreeView', AutomaitonID='DeviceTreeView')

    # 遍历树找到对应IP
    is_find = False
    for item, depth in auto.WalkControl(tree, includeTop=True):
        if isinstance(item, auto.TreeItemControl):  # or item.ControlType == auto.ControlType.TreeItemControl
            item.GetSelectionItemPattern().Select(waitTime=0.05)
        if plc_ip == item.Name.split(',')[0]:
            item.Click()
            is_find = True
            break

    if is_find:
        net_dlg.ButtonControl(Name='下载').Click()
    else:
        print("未找到对应IP设备")

    # 程序验证进度条
    # todo 增加加载失败判断
    while True:
        build_dlg = ics_window.WindowControl(Name='Downloading')
        if not build_dlg.Exists():
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:程序验证完成')
            break

    # 下载对话框中点击下载
    time.sleep(1)
    ics_window.WindowControl(Name='下载').ButtonControl(Name='下载').Click()

    # 程序编译下载进度条
    # todo 增加加载失败判断
    while True:
        build_dlg = ics_window.WindowControl(Name='下载中')
        if not build_dlg.Exists():
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:程序编译和下载完成')
            break


def login(model=2):
    """
    登入
    :param model: 1 为快捷方式测试 2 为页面点击方式测试
    :return:
    """
    communications_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
        MenuItemControl(Name='Communications', ClassName='MenuItem')
    if model == 1:
        # 方式一： 快捷键
        ics_window.SendKeys('{Alt}c')
    elif model == 2:
        # 方式二： 点击选择信息
        communications_menu.Click()

    # 网络路径设置
    try:
        communications_menu.MenuItemControl(Name='登入').Click()
    except LookupError:
        print("现在已经是登入状态")


def logout(model=2):
    """
    登出
    :param model: 1 为快捷方式测试 2 为页面点击方式测试
    :return:
    """
    communications_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
        MenuItemControl(Name='Communications', ClassName='MenuItem')
    if model == 1:
        # 方式一： 快捷键
        ics_window.SendKeys('{Alt}c')
    elif model == 2:
        # 方式二： 点击选择信息
        communications_menu.Click()

    # 网络路径设置
    try:
        communications_menu.MenuItemControl(Name='登出').Click()
    except LookupError:
        print("现在已经是登出状态")


def download(model=2):
    """
    下载
    :param model: 1 为快捷方式测试 2 为页面点击方式测试
    :return:
    """
    communications_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
        MenuItemControl(Name='Communications', ClassName='MenuItem')
    if model == 1:
        # 方式一： 快捷键
        ics_window.SendKeys('{Alt}c')
    elif model == 2:
        # 方式二： 点击选择信息
        communications_menu.Click()

    # 下载点击
    communications_menu.MenuItemControl(Name='下载').Click()

    # 程序验证进度条 todo 需等确认窗口再修改
    # todo 增加加载失败判断
    while True:
        build_dlg = ics_window.WindowControl(Name='Downloading')
        if not build_dlg.Exists():
            break

    # 下载对话框中点击下载
    time.sleep(1)
    ics_window.WindowControl(Name='下载').ButtonControl(Name='下载').Click()

    # 程序编译下载进度条
    # todo 增加加载失败判断
    #  todo 需等确认窗口再修改

    while True:
        build_dlg = ics_window.WindowControl(Name='下载中')
        if not build_dlg.Exists():
            break


def upload(model=2):
    """
    上传
    :param model: 1 为快捷方式测试 2 为页面点击方式测试
    :return:
    """
    communications_menu = ics_window.MenuBarControl(ClassName='Menu', AutomationID='MenuBar'). \
        MenuItemControl(Name='Communications', ClassName='MenuItem')
    if model == 1:
        # 方式一： 快捷键
        ics_window.SendKeys('{Alt}c')
    elif model == 2:
        # 方式二： 点击选择信息
        communications_menu.Click()

    # 上传点击
    communications_menu.MenuItemControl(Name='上传').Click()

    # 保存文件
    save_file_dlg = ics_window.WindowControl(Name="保存文件")
    time.sleep(1)
    save_file_dlg.ButtonControl(Name="保存(S)").Click()

    # 名称已经存在对话框
    if ics_window.WindowControl(Name="确认另存为").Exists():
        ics_window.WindowControl(Name="确认另存为").ButtonControl(Name="是(Y)").Click()

    # 程序验证进度条 todo 需等确认窗口再修改
    # todo 增加加载失败判断
    while True:
        build_dlg = ics_window.WindowControl(Name='Downloading')
        if not build_dlg.Exists():
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:程序上传完成')
            break

    while True:
        build_dlg = ics_window.WindowControl(Name='Downloading')
        if not build_dlg.Exists():
            print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:程序验证完成')
            break

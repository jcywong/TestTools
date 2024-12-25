import datetime
import os
import re
import subprocess
import telnetlib
import time
import zipfile
from ftplib import FTP

import paramiko
import psutil
import rarfile
import requests
import uiautomation as auto
from bs4 import BeautifulSoup

# 目录信息
workspace = os.getcwd()

soft_path = workspace + '\\soft\\'

# json_path = os.path.dirname(workspace) + '\\json\\'
json_path = workspace + "\\json\\"

# ics窗口
ics_window = auto.WindowControl(SubName='ICS Studio', ClassName='Window', AutomationId='VisualStudioMainWindow')


def get_server_url(soft_type='ICS', edition="Debug", network="LAN"):
    """
    获得网址
    :param soft_type: 软件类型 ICS/ICC/ICM
    :param edition: 软件版本
    :param network: 内网LAN 外网Internet
    :return:
    """
    if soft_type == 'ICS' and edition == "Debug" and network == "LAN":
        return 'http://172.16.2.240/autobuild/icsstudio/'
    elif soft_type == 'ICC' and edition == "Debug" and network == "LAN":
        return 'http://172.16.2.240/autobuild/firmwares/'
    elif soft_type == 'ICS' and edition == "Release" and network == "LAN":
        return "http://172.16.2.240/autobuild/release/"
    elif soft_type == 'ICC' and edition == "Release" and network == "LAN":
        return "http://172.16.2.240/autobuild/release/"
    elif soft_type == 'ICM' and network == "LAN":  # jcywong add 2023/11/13
        return "http://172.16.2.240/autobuild/Driver/"
    elif soft_type == 'ICS' and edition == "Debug" and network == "Internet":
        return 'http://hub.i-con.cn:32208/autobuild/icsstudio/'
    elif soft_type == 'ICC' and edition == "Debug" and network == "Internet":
        return 'http://hub.i-con.cn:32208/autobuild/firmwares/'
    elif soft_type == 'ICS' and edition == "Release" and network == "Internet":
        return "http://hub.i-con.cn:32208/autobuild/release/"
    elif soft_type == 'ICC' and edition == "Release" and network == "Internet":
        return "http://hub.i-con.cn:32208/autobuild/release/"
    elif soft_type == 'ICM' and network == "Internet":  # jcywong add 2023/11/13
        return "http://hub.i-con.cn:32208/autobuild/Driver/"


def download_file(file_name, file_save_path, soft_type='ICS', edition="Debug", network="LAN"):
    """
    下载文件
    :param network: 内网LAN 外网Internet
    :param file_name: 文件名
    :param file_save_path: 保存路径
    :param soft_type: 软件类型 ICC / ICS /ICM
    :param edition: 软件版本 Debug / Release
    :return:
    """

    file_server = get_server_url(soft_type, edition, network)
    if not file_server:
        return False

    try:
        response = requests.get(file_server + file_name)
        file_path = os.path.join(file_save_path, file_name)
        if response.status_code == 200:
            if not os.path.isfile(file_path):  # 判断目录下是有同样文件
                with open(file_path, 'wb') as file:
                    file.write(response.content)
                print(f"{file_name}:File downloaded successfully.")
                return True
            else:
                print(f"{file_name}:已存在该文件，不进行下载")
                return False
        else:
            print(f"{file_name}:Failed to download file.")
            return False

    except Exception as e:
        print(f"{file_name}: An error occurred: {e}")
        return False


def unzip_file(zip_file_path, zip_file_name, extract_dir=None):
    """
    解压文件 可解压zip和rar格式
    :param zip_file_path: 解压文件地址
    :param zip_file_name: 解压文件名字包含拓展名
    :param extract_dir:  解压地址
    :return:
    """
    if extract_dir is None:
        extract_dir = zip_file_path
    path = os.path.join(extract_dir, zip_file_name[:-4])
    if not os.path.isdir(path):
        if zip_file_name[-3:] == 'zip':
            with zipfile.ZipFile(zip_file_path + "/" + zip_file_name, 'r') as zip_ref:
                zip_ref.extractall(str(path))
        elif zip_file_name[-3:] == 'rar':
            with rarfile.RarFile(zip_file_path + "/" + zip_file_name, 'r') as rar_file:
                rar_file.extractall(path)
        print(f"{zip_file_name[:-4]}:RAR file extracted successfully.")
        return True
    else:
        print(f"{zip_file_name[:-4]}:已解压，不进行再次解压")
        return False


def get_latest_filename(soft_type='ICS', edition="Debug", network="LAN", model=None, ver=None, ):
    """
    得到最新的文件名
    :param network: 内网LAN 外网Internet
    :param soft_type: 软件类型ICS/ICC/ICM
    :param edition: 软件版本 "Debug" 、"Release"
    :param model: ICC型号：LITE、PRO、PRO.B、TURBO、EVO
    :param ver: release 版本号
    :return:
    """
    if soft_type == 'ICS' and edition == "Debug":
        # 发送 HTTP 请求获取页面内容
        file_server = get_server_url(soft_type, edition, network)

        try:
            response = requests.get(file_server)
            # 解析 HTML 内容
            soup = BeautifulSoup(response.text, 'html.parser')

            tbody = soup.select('tbody')
            second_tr = tbody[0].select('tr')[1]
            first_td = second_tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            return filename
        except requests.exceptions.ConnectTimeout:
            print("网络错误")
            return False

    elif soft_type == 'ICC' and edition == "Debug":
        file_server = get_server_url(soft_type, edition, network)
        try:
            response = requests.get(file_server)

            # 解析 HTML 内容
            soup = BeautifulSoup(response.text, 'html.parser')

            tbody = soup.select('tbody')
            tr_list = tbody[0].select('tr')
            for tr in tr_list:
                first_td = tr.select('td')[0]
                a_tag = first_td.find('a')
                filename = a_tag.get("href")
                if filename[-9:-4] == "debug":  # jcywong add 2023/11/13  解决固件firmwares下载debug中包含release和debug问题
                    if model == 'LITE':
                        if model == filename[4:8]:
                            return filename
                    elif model == 'TURBO':
                        if model == filename[4:9]:
                            return filename
                    elif model == 'PRO':
                        if model == filename[4:7] and filename[4:9] != 'PRO.B':
                            return filename
                    elif model == 'PRO.B':  # jcywong add 2024/1/31
                        if model == filename[4:9]:
                            return filename
                    elif model == 'EVO':  # jcywong add 2023/12/12
                        if model == filename[4:7]:
                            return filename
        except requests.exceptions.ConnectTimeout:
            print("网络错误")
            return False
    elif soft_type == 'ICS' and edition == "Release":
        # 发送 HTTP 请求获取页面内容
        file_server = get_server_url(soft_type, edition, network)
        try:
            response = requests.get(file_server)
            # 解析 HTML 内容
            soup = BeautifulSoup(response.text, 'html.parser')

            tbody = soup.select('tbody')
            tr_list = tbody[0].select('tr')
            for tr in tr_list:
                first_td = tr.select('td')[0]
                a_tag = first_td.find('a')
                filename = a_tag.get("href")
                if filename[:9] == "ICSStudio" and filename[10:14] == ver:
                    return filename
        except requests.exceptions.ConnectTimeout:
            print("网络错误")
            return False
    elif soft_type == 'ICC' and edition == "Release":
        file_server = get_server_url(soft_type, edition, network)
        try:
            response = requests.get(file_server)
            # 解析 HTML 内容
            soup = BeautifulSoup(response.text, 'html.parser')

            tbody = soup.select('tbody')
            tr_list = tbody[0].select('tr')
            for tr in tr_list:
                first_td = tr.select('td')[0]
                a_tag = first_td.find('a')
                filename = a_tag.get("href")
                if model == 'LITE':
                    if model == filename[4:8] and filename[-16:-12] == ver:
                        return filename
                elif model == 'TURBO':
                    if model == filename[4:9] and filename[-16:-12] == ver:
                        return filename
                elif model == 'PRO':
                    if model == filename[4:7] and filename[-16:-12] == ver:
                        return filename
                elif model == 'PRO.B':  # jcywong add 2024/4/2
                    if model == filename[4:9] and filename[-16:-12] == ver:
                        return filename
                elif model == 'EVO':  # jcywong add 2024/4/2
                    if model == filename[4:7] and filename[-16:-12] == ver:
                        return filename
        except requests.exceptions.ConnectTimeout:
            print("网络错误")
            return False
    elif soft_type == 'ICM':
        file_server = get_server_url(soft_type, network)
        try:
            response = requests.get(file_server)
            # 解析 HTML 内容
            soup = BeautifulSoup(response.text, 'html.parser')

            tbody = soup.select('tbody')
            second_tr = tbody[0].select('tr')[1]
            first_td = second_tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            return filename
        except requests.exceptions.ConnectTimeout:
            print("网络错误")
            return False
    else:
        return False


def open_ics(path):
    # 打开ics
    ics_path = f'{path}/ICSStudio.exe'
    # if ics_window.Exists():
    #     close_window(ics_window)
    #     time.sleep(5)
    #     subprocess.Popen(ics_path)
    # else:
    #     subprocess.Popen(ics_path)
    # subprocess.Popen(ics_path)
    subprocess.Popen([ics_path], cwd=ics_path[:-13])  # 2023/11/13 jcywong 解决工具外挂3个程序无法打开
    # time.sleep(10)
    # ics_window.SetTopmost(True)
    # print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成打开ICS并置顶')
    #
    # # 最大化ics
    # if not ics_window.IsMaximize():
    #     ics_window.ButtonControl(AutomationId='Maximize', Name='Maximize').Click()
    # print(f'{datetime.datetime.now().strftime("%H:%M:%S")}:完成ICS最大化')


def monitor_process_memory(process_name="ICSStudio.exe"):
    """
    monitor process memory
    :param process_name:default ICSStudio.exe
    :return: MB
    """
    while True:
        try:
            # 获取进程列表
            for process in psutil.process_iter(['pid', 'name', 'memory_info']):
                if process.name() == process_name:
                    pid = process.pid
                    memory_info = process.memory_info().rss / (1024 ** 2)
                    print(f"{process_name}, Process ID: {pid}, Memory used: {memory_info:.2f} MB")
                    return memory_info
            else:
                print(f"Process '{process_name}' not found.")
                return
        except Exception as e:
            print(f"Error: {e}")

        time.sleep(1)


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


def ssh_to_device(ip="192.168.1.211", device_model="TURBO", command="reboot"):
    """
    针对turbo进行 远程ssh指令
    :param device_model:
    :param ip:
    :param command:
    :return:
    """

    # 设置SSH连接参数
    hostname = ip
    port = 22
    # 从环境变量获取SSH凭据
    username = os.getenv('SSH_USERNAME', default='icon' if device_model == "TURBO" else 'root')
    password = os.getenv('SSH_PASSWORD', default='Icon!@#123')

    # 创建SSH客户端
    ssh_client = paramiko.SSHClient()

    # 自动添加远程主机的SSH密钥（可选）
    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    try:
        # 连接远程服务器
        ssh_client.connect(hostname, port, username, password)

        if command == "reboot":
            # 创建一个交互式shell
            ssh_shell = ssh_client.invoke_shell()

            ssh_shell.send(b"clear\n")
            # 等待命令执行完毕
            while not ssh_shell.recv_ready():
                time.sleep(1)

            print(f"clear:{ssh_shell.recv(1024).decode('utf-8')}")

            # # 执行提权命令（su -）
            # ssh_shell.send(b"su -\n")
            #
            # # 等待命令执行完毕
            # while not ssh_shell.recv_ready():
            #     time.sleep(1)
            #
            # print(f"su:{ssh_shell.recv(1024).decode('utf-8')}")

            # # 输入提权密码
            # ssh_shell.send(b"Icon!@#123\n")
            #
            # # 等待命令执行完毕
            # while not ssh_shell.recv_ready():
            #     time.sleep(2)

            if device_model == "TURBO":
                # 执行提权命令（su -）
                ssh_shell.send(b"su -\n")

                # 等待命令执行完毕
                while not ssh_shell.recv_ready():
                    time.sleep(1)

                print(f"su:{ssh_shell.recv(1024).decode('utf-8')}")

                # 输入提权密码
                ssh_shell.send(b"Icon!@#123\n")

                # 等待命令执行完毕
                while not ssh_shell.recv_ready():
                    time.sleep(1)

                # 读取stdout输出，检查是否切换到root用户
                su_output = ssh_shell.recv(1024).decode('utf-8')
                if "root@" in su_output or "#" in su_output:
                    print("提权成功")
                    # 在此处执行需要root权限的操作

                    # 执行重启命令

                    ssh_shell.send(b"reboot\n")

                    while not ssh_shell.recv_ready():
                        time.sleep(1)

                    print(f"reboot:{ssh_shell.recv(1024).decode('utf-8')}")
                    return True
                else:
                    print("提权失败")
                    return False
            elif device_model == "ICM" or device_model == "ANTER" or device_model == "BANTER":
                ssh_shell.send(b"reboot\n")
                while not ssh_shell.recv_ready():
                    time.sleep(1)

                print(f"reboot:{ssh_shell.recv(1024).decode('utf-8')}")
                return True

    except paramiko.AuthenticationException:
        print("认证失败，请检查用户名和密码或SSH密钥。")
        return False
    except paramiko.SSHException as e:
        print("SSH连接或执行命令时发生错误:", str(e))
        return False
    except Exception as e:
        print("发生错误:", str(e))
        return False
    finally:
        # 确保在异常情况下关闭SSH连接
        ssh_client.close()


def telnet_to_device(ip="192.168.1.211", command="reboot"):
    """
    针对B/PRO进行 远程telnet指令
    :param command: 执行的命令
    :param ip:PLC的IP
    :return:
    """

    HOST = ip  # 设备的 IP 地址
    PORT = 23  # Telnet 端口号

    try:
        # 连接到设备
        tn = telnetlib.Telnet(HOST, PORT, 3)

        # 登录设备
        icon_login = tn.read_until(b"Icon login: ", 5)
        if not icon_login:
            raise OSError
        tn.write(b"root\r\n")
        tn.read_until(b"Password: ", 5)
        tn.write(b"Icon!@#123\r\n")

        # 执行命令
        tn.read_until(b"# ", 5)

        if command == "free":
            tn.write(b"free -h\n")
            output = tn.read_until(b"# ", 5).decode('ascii')

            # 解析命令输出，获取内存使用情况
            lines = output.splitlines()
            men = {}
            for line in lines:
                if "Mem:" in line:
                    _, total, used, free, *_ = re.split(r'\s+', line)
                    men["总内存"] = total
                    men["已使用内存"] = used
                    men["可用内存"] = free
                    break
            # 关闭 Telnet 连接
            tn.close()
            print(f"{command}命令执行成功！")
            return men
        elif command == "ls":
            # 执行命令获取文件大小信息
            cfg_path = "/mnt/data0/config/project.cfg"
            tn.write(f"ls -l {cfg_path} \n".encode('ascii'))
            output = tn.read_until(b"# ", 5).decode('ascii')

            # 解析命令输出，获取文件大小
            lines = output.splitlines()
            men = {}
            for line in lines:
                if line.startswith("-"):
                    _, _, _, _, size, *_ = line.split()
                    men["size"] = size
                    break
            tn.close()
            print(f"{command}命令执行成功！")
            return men
        elif command == "reboot":
            tn.write(command.encode('ascii') + b"\r\n")
            time.sleep(1)
            # 关闭 Telnet 连接
            tn.close()
            print(f"{command}命令执行成功！")
            return True
    except TimeoutError:
        print("连接超时")
        return False
    except OSError:
        print("网络错误")
        return False
    except Exception as e:
        print("发生错误:", str(e))
        raise e


def reboot_device(device_model, ip):
    """
    重启设备
    :param device_model: 设备型号：eg ICC-B010ERM
    :param ip:
    :return:
    """
    print(f"重启设备:设备型号：{device_model},设备IP：{ip}")
    if device_model in ['LITE', 'PRO', 'PRO.B', 'EVO', 'ICM-D3', 'ICM-D5']:
        return telnet_to_device(ip)
    elif device_model in ['ICM-D1', 'ICM-D7']:
        return ssh_to_device(ip, device_model="ICM")
    elif device_model == 'ICD-ANTER':
        return ssh_to_device(ip, device_model="ANTER")
    elif device_model == 'ICC-BANTER':
        return ssh_to_device(ip, device_model="BANTER")
    elif device_model == "TURBO":
        return ssh_to_device(ip, device_model="TURBO")
    else:
        return True


def is_directory(connection, item):
    try:
        if isinstance(connection, FTP):
            connection.cwd(item)
        elif isinstance(connection, paramiko.sftp_client.SFTPClient):
            connection.chdir(item)
        return True
    except Exception as e:
        print(e)
        return False


def get_files_By_FTP(icc_model, remote_path, local_path, ip="192.168.1.211"):
    os.makedirs(local_path, exist_ok=True)
    password, username = get_username(icc_model)
    ftp = FTP(ip)
    try:
        ftp.login(username, password)
        ftp.cwd(remote_path)
        files = ftp.nlst()
        for filename in files:
            local_file_path = f"{local_path}/{filename}"
            remote_file_path = f"{remote_path}/{filename}"

            # 如果是目录，则递归下载
            if is_directory(ftp, remote_file_path):
                get_files_By_FTP(ftp, remote_file_path, local_file_path, ip)
            # 如果是文件，则下载
            else:
                with open(local_file_path, 'wb') as local_file:
                    ftp.retrbinary('RETR ' + remote_file_path, local_file.write)
        ftp.quit()
    except Exception as e:
        ftp.quit()
        raise e


def get_files_By_SFTP(icc_model, remote_path, local_path, ip="192.168.1.211"):
    """

    :param icc_model:
    :param remote_path:
    :param local_path:
    :param ip:
    :return:
    """
    os.makedirs(local_path, exist_ok=True)

    password, username = get_username(icc_model)
    ssh = paramiko.SSHClient()
    try:
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh.connect(ip, 22, username, password)

        sftp = ssh.open_sftp()

        sftp.chdir(remote_path)
        filenames = sftp.listdir()

        for filename in filenames:
            local_file_path = f"{local_path}/{filename}"
            remote_file_path = f"{remote_path}/{filename}"

            # 如果是目录，则递归下载
            if is_directory(sftp, remote_file_path):
                # 判断有无权限
                try:
                    get_files_By_SFTP(icc_model, remote_file_path, local_file_path, ip)
                except PermissionError:
                    print(f"{remote_file_path}：没有权限")
                    continue
                except Exception as e:
                    raise e
            # 如果是文件，则下载
            else:
                sftp.get(remote_file_path, local_file_path)
    except FileNotFoundError:
        print(f"目录不存在: {remote_path}")
    except PermissionError:
        print(f"没有权限访问目录: {remote_path}")
    except Exception as e:
        print(f"发生错误: {str(e)}")
        raise e
    finally:
        ssh.close()


def get_device_logs(device_model, local_path, ip="192.168.1.211"):
    """
    获取日志
    :param local_path:
    :param device_model:
    :param ip:
    :return:
    """
    try:
        # 生成文件夹名称为日期
        local_directory = f"{local_path}/logs/{datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"

        if device_model in ["PRO", "LITE", "PRO.B", "EVO"]:
            get_files_By_FTP(device_model, "/mnt/data0/config", local_directory + "/config", ip)
            get_files_By_FTP(device_model, "/tmp", local_directory + "/tmp", ip)
        elif device_model in ["TURBO"]:
            get_files_By_SFTP(device_model, "/mnt/data0/config", local_directory + "/config", ip)
            get_files_By_SFTP(device_model, "/tmp", local_directory + "/tmp", ip)
        elif device_model in ["ICM-D1", "ICM-D7"]:
            get_files_By_SFTP(device_model, "/mnt/mmc/", local_directory + "/mmc", ip)
        elif device_model in ["ICM-D3", "ICM-D5"]:
            get_files_By_FTP(device_model, "/mnt/mmc/", local_directory + "/mmc", ip)

        # 压缩整个目录下的所有文件
        zip_files(local_directory, local_directory + ".zip")
        return True
    except Exception as e:
        print(e)
        return False


def zip_files(source_file_path, zip_file_path):
    try:
        with zipfile.ZipFile(zip_file_path, 'w') as zip_file:
            # 压缩目标文件夹内的所有文件和子目录
            for root, dirs, files in os.walk(source_file_path):
                for file in files:
                    file_path = str(os.path.join(root, file))
                    zip_file.write(file_path, str(os.path.relpath(file_path, source_file_path)))

                # 创建空文件夹
                for empty_dir in dirs:
                    empty_dir_path = str(os.path.join(root, empty_dir))
                    zip_file.write(empty_dir_path, os.path.relpath(empty_dir_path, source_file_path))

        print(f"Compression successful. Zip file: {zip_file_path}")

    except Exception as e:
        raise e


def get_username(icc_model):
    """
    get password, username
    :param icc_model:
    :return:
    """
    if icc_model in ["LITE", "PRO", "PRO.B", "EVO", "ICM-D1", "ICM-D3", "ICM-D5", "ICM-D7"]:
        username = "root"
        password = "Icon!@#123"
        return password, username
    elif icc_model in ["TURBO"]:
        username = "icon"
        password = "Icon!@#123"
        return password, username


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

    for i in range(len(address_bar.GetChildren()) - 1):
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

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
from bs4 import BeautifulSoup

# 目录信息
workspace = os.getcwd()

soft_path = workspace + '\\soft\\'

# json_path = os.path.dirname(workspace) + '\\json\\'
json_path = workspace + "\\json\\"


def get_server_url(soft_type='ICS', edition="Debug", network="LAN"):
    """
    获得网址
    :param soft_type: 软件类型 ICS/ICC/ICM
    :param edition: 软件版本
    :param network: 内网LAN 外网Internet
    :return:
    """
    lan_url = "http://172.16.2.240/"
    internet_url = "http://hub.i-con.cn:32208/"

    url_mapping = {
        ("ICS", "Debug", "LAN"): f'{lan_url}autobuild/icsstudio/',
        ("ICC", "Debug", "LAN"): f'{lan_url}autobuild/firmwares/',
        (soft_type, "Release", "LAN"): f'{lan_url}autobuild/release/',
        ("ICM", edition, "LAN"): f'{lan_url}autobuild/ICON/driver/',
        ("AENTR", edition, "LAN"): f'{lan_url}autobuild/ICD_AENTR/',
        ("BAENTR", edition, "LAN"): f'{lan_url}autobuild/ICC_BAENTR/',
        ("VP", edition, "LAN"): f'{lan_url}autobuild/vstudio/',
        ("ICP", edition, "LAN"): f'{lan_url}autobuild/hmi/',
        ("ICF", edition, "LAN"): f'{lan_url}autobuild/driver/',
        ("ICS", "Debug", "Internet"): f'{internet_url}autobuild/icsstudio/',
        ("ICC", "Debug", "Internet"): f'{internet_url}autobuild/firmwares/',
        (soft_type, "Release", "Internet"): f'{internet_url}autobuild/release/',
        ("ICM", edition, "Internet"): f'{internet_url}autobuild/ICON/driver/',
        ("AENTR", edition, "Internet"): f'{internet_url}autobuild/ICD_AENTR/',
        ("BAENTR", edition, "Internet"): f'{internet_url}autobuild/ICC_BAENTR/',
        ("VP", edition, "Internet"): f'{internet_url}autobuild/vstudio/',
        ("ICP", edition, "Internet"): f'{internet_url}autobuild/hmi/',
        ("ICF", edition, "Internet"): f'{internet_url}autobuild/driver/',
    }

    # 参数校验
    valid_soft_types = ['ICS', 'ICC', 'ICM', 'AENTR', 'BAENTR', 'VP', 'ICP', 'ICF']
    valid_editions = [' ', 'Debug', 'Release']
    valid_networks = ['LAN', 'Internet']

    if soft_type not in valid_soft_types or edition not in valid_editions or network not in valid_networks:
        raise ValueError("Invalid input parameters")

    try:
        return url_mapping[(soft_type, edition, network)]
    except KeyError:
        raise ValueError("No URL mapping found for the provided parameters")


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
                raise FileExistsError
        else:
            print(f"{file_name}:Failed to download file.")
            return False

    except Exception as e:
        print(f"{file_name}: An error occurred: {e}")
        raise e


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
        raise FileExistsError


def get_latest_filename(soft_type='ICS', edition="Debug", network="LAN", model=None, ver=None, ):
    """
    得到最新的文件名
    :param network: 内网LAN 外网Internet
    :param soft_type: 软件类型ICS/ICC/ICM
    :param edition: 软件版本 "Debug" 、"Release"
    :param model: ICC型号：LITE、PRO、PRO.B、TURBO、EVO、,LITE.B
    :param ver: release 版本号
    :return:
    """
    # 发送 HTTP 请求获取页面内容
    file_server = get_server_url(soft_type, edition, network)

    try:
        response = requests.get(file_server)
        # 解析 HTML 内容
        soup = BeautifulSoup(response.text, 'html.parser')
        tbody = soup.select('tbody')
    except requests.exceptions.ConnectTimeout:
        print("网络错误")
        return False

    if soft_type == 'ICS' and edition == "Debug":
        tr_list = tbody[0].select('tr')
        for tr in tr_list:
            first_td = tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            if "ICSStudio" not in filename:
                continue
            if  "refs" in filename:
                continue
            if not filename.startswith("ICSStudio"):
                continue
            return filename
    elif soft_type == 'ICC' and edition == "Debug":
        tr_list = tbody[0].select('tr')
        for tr in tr_list:
            first_td = tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            if filename[-9:-4] == "debug":  # jcywong add 2023/11/13  解决固件firmwares下载debug中包含release和debug问题
                if model == 'LITE' and filename[4:10] != 'LITE.B':
                    if model == filename[4:8]:
                        return filename
                elif model == 'LITE.B':
                    if filename[4:10] == model:
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
    elif soft_type == 'ICS' and edition == "Release":
        tr_list = tbody[0].select('tr')
        for tr in tr_list:
            first_td = tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            if filename[:9] == "ICSStudio" and filename[10:14] == ver:
                return filename
    elif soft_type == 'ICC' and edition == "Release":
        tr_list = tbody[0].select('tr')
        for tr in tr_list:
            first_td = tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            if model == 'LITE':
                if model == filename[4:8] and filename[4:10] != 'LITE.B' and filename[-16:-12] == ver:
                    return filename
            elif model == 'LITE.B':
                if model == filename[4:10] and filename[-16:-12] == ver:
                    return filename
            elif model == 'TURBO':
                if model == filename[4:9] and filename[-16:-12] == ver:
                    return filename
            elif model == 'PRO':
                if model == filename[4:7] and filename[4:9] != 'PRO.B' and filename[-16:-12] == ver:
                    return filename
            elif model == 'PRO.B':  # jcywong add 2024/4/2
                if model == filename[4:9] and filename[-16:-12] == ver:
                    return filename
            elif model == 'EVO':  # jcywong add 2024/4/2
                if model == filename[4:7] and filename[-16:-12] == ver:
                    return filename
    elif soft_type == 'ICM':
        tbody = soup.select('tbody')
        second_tr = tbody[0].select('tr')[1]
        first_td = second_tr.select('td')[0]
        a_tag = first_td.find('a')
        filename = a_tag.get("href")
        return filename
    elif soft_type == 'VP':
        tbody = soup.select('tbody')
        second_tr = tbody[0].select('tr')[2]
        first_td = second_tr.select('td')[0]
        a_tag = first_td.find('a')
        filename = a_tag.get("href")
        return filename
    elif soft_type == 'ICP':
        tr_list = tbody[0].select('tr')
        for tr in tr_list:
            first_td = tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            parts = filename.split('.')
            if len(parts) < 3:
                continue
            master_part = parts[-3]
            edition_part = parts[-2]
            if master_part != "master":
                continue
            if edition_part == edition.lower():
                return filename
    elif soft_type == 'ICF':
        tr_list = tbody[0].select('tr')
        for tr in tr_list:
            first_td = tr.select('td')[0]
            a_tag = first_td.find('a')
            filename = a_tag.get("href")
            if not filename.startswith('ICF'):
                continue
            parts = filename.split('.')
            if len(parts) < 3:
                continue
            if parts[0].split('-')[1] == model:
                return filename
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
    password, username = get_username(device_model)

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

            print(f"clear:{ssh_shell.recv(5048).decode('utf-8')}")
            if device_model in ["TURBO", "ICP"]:
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
    if device_model in ['LITE', "LITE.B", 'PRO', 'PRO.B', 'EVO', 'ICM-D3', 'ICM-D5', 'ICF-C', 'ICM-D7',"KCU"]:
        return telnet_to_device(ip)
    elif device_model in ['ICM-D1']:
        return ssh_to_device(ip, device_model="ICM")
    elif device_model == 'ICD-ANTER':
        return ssh_to_device(ip, device_model="ANTER")
    elif device_model == 'ICC-BANTER':
        return ssh_to_device(ip, device_model="BANTER")
    elif device_model in ["TURBO", "ICP"]:
        return ssh_to_device(ip, device_model=device_model)
    else:
        return


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

        if device_model in ["PRO", "LITE", "LITE.B", "PRO.B", "EVO"]:
            get_files_By_FTP(device_model, "/mnt/data0/config", local_directory + "/config", ip)
            get_files_By_FTP(device_model, "/tmp", local_directory + "/tmp", ip)
        elif device_model in ["TURBO"]:
            get_files_By_SFTP(device_model, "/mnt/data0/config", local_directory + "/config", ip)
            get_files_By_SFTP(device_model, "/tmp", local_directory + "/tmp", ip)
        elif device_model in ["ICM-D1"]:
            get_files_By_SFTP(device_model, "/mnt/mmc/", local_directory + "/mmc", ip)
        elif device_model in ["ICM-D3", "ICM-D5", "ICF", "ICM-D7", "KCU"]:
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


def get_test_tools_last_version():
    """
    获取测试工具最新版本号
    :return:
    """
    try:
        response = requests.get("https://testtools-version.pages.dev/version.json")
        if response.status_code == 200:
            return response.json()
        else:
            return None

    except Exception as e:
        print(e)
        return None


def get_username(model):
    """
    get password, username
    :param model:
    :return:
    """
    # if model in ["LITE", "LITE.B", "PRO", "PRO.B", "EVO", "ICM-D1", "ICM-D3", "ICM-D5", "ICM-D7","KCU"]:
    #     username = "root"
    #     password = "Icon!@#123"
    #     return password, username
    # elif model in ["TURBO"]:
    #     username = "icon"
    #     password = "Icon!@#123"
    #     return password, username
    # elif model == "ICP":
    #     username = "cat"
    #     password = "Icon!@#123"
    if model in ["TURBO"]:
        username = "icon"
        password = "Icon!@#123"
    else:
        username = "root"
        password = "Icon!@#123"

    return password, username

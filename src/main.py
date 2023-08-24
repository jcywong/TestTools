from comm import *


def main():
    print("------------------------------\n"
          "上传下载测试工具\n"
          "注意：\n"
          "需搭配ICS Studio使用,可放置soft目录下\n"
          "工程文件必需要放置json目录下\n"
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
    main()

    print("按任意键退出...")
    input()  # 等待用户输入
    print("程序已退出")

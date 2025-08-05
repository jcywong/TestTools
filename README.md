# TestTools

## 项目简介

TestTools 是一个基于 PySide6 的多功能桌面工具，集成了固件下载、设备管理、日志获取、在线升级等多项功能，适用于嵌入式设备开发、测试和维护场景。

## 主要功能

- **固件下载与解压**：支持多种设备型号（ICS、ICC、ICM、ICF、ICP、VP）的固件下载、自动解压和路径管理。
- **设备管理**：支持设备重启、日志获取等常用操作。
- **在线升级**：自动检测新版本，支持一键下载并自我更新。
- **配置管理**：支持保存和加载用户配置（如下载路径、网络类型等）。
- **多标签页界面**：每种设备类型独立标签页，操作互不干扰。
- **日志记录**：自动记录操作日志，便于问题追踪。

## 目录结构

```
TestTools/
├── env/                  # 虚拟环境（可选）
├── img/                  # 图标资源
│   ├── tool.ico
│   └── tool.png
├── src/                  # 源码目录
│   ├── comm.py           # 通用通信与下载解压逻辑
│   ├── config.json       # 配置文件
│   ├── main.ui           # Qt Designer 设计的主界面
│   └── TestTools.py      # 主程序入口
├── pyinstaller.txt       # 打包命令模板
└── README.md             # 项目说明文档
```

## 依赖环境

- Python 3.8+

安装依赖：
```bash
pip install -r requirements.txt
```
> 注：如无 requirements.txt，请根据实际 import 手动安装。

## 使用说明

### 运行开发版

```bash
cd src
python TestTools.py
```

### 打包发布版

1. 确保已安装 PyInstaller：
   ```bash
   pip install pyinstaller
   ```
2. 使用 `pyinstaller.txt` 中的命令打包（推荐在项目根目录下执行）：
   ```bash
   pyinstaller -F src/TestTools.py --icon="img/tool.ico" --noconsole --add-data "src/main.ui;." --add-data "img/tool.png;img"
   ```
3. 打包完成后，`dist/` 目录下会生成 `TestTools.exe`，可直接运行。

### 常见问题

- **UI 文件找不到**  
  打包时需用 `--add-data "src/main.ui;."` 参数，确保 UI 文件被正确打包。
- **图标不显示**  
  同理，需用 `--add-data "img/tool.png;img"` 参数。


## 配置文件说明

`src/config.json` 用于保存用户的个性化设置，如下载路径、网络类型、自动检查更新等。

## 日志

所有操作日志会记录在 `testtools.log` 文件中，便于排查问题。

## 贡献与反馈

如有建议或发现 bug，欢迎提交 issue 或 pull request。
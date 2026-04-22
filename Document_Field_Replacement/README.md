# 文档字段批量替换工具

**[🇺🇸 English Version](README_EN.md)** | **[🇨🇳 中文版](README.md)**

**v0.1 | by keloder | Python 3.8.20**

基于 Web 的文档批量字段替换工具，支持 Word (.docx/.doc) 和 WPS 格式，提供正向替换与反向还原功能。

## 功能特性

- **正向替换** — 按规则批量替换文档中的字段内容
- **反向还原** — 将已替换的内容还原为原始字段
- **多文件处理** — 支持同时上传多个 .docx / .doc / .wps 文件
- **规则管理** — 支持添加、编辑、删除、排序、清空替换规则
- **导入/导出** — 支持 JSON 和 TXT 格式的规则导入导出
- **批量下载** — 处理完成后支持单文件下载或打包下载全部结果
- **无痕模式** — 支持无痕模式，不保留任何数据
- **主题切换** — 支持明暗主题切换
- **多语言** — 支持中文/英文界面切换
- **便携模式** — 配置文件保存在软件目录，便于携带

## 项目结构

```
word_test/
├── web_server.py          # Flask Web 服务主程序
├── replacer.py            # 核心替换逻辑（正向/反向）
├── document_handler.py    # 文档读写处理
├── build_exe.py           # PyInstaller 打包脚本
├── build.bat              # 一键打包批处理（支持 onedir/onefile）
├── templates/
│   └── index.html         # 前端页面
├── static/
│   └── favicon.png        # 网页图标
├── icon/
│   └── TH.png             # 应用图标
└── requirements.txt       # 依赖列表
```

## 快速开始

### 开发模式运行

```bash
python web_server.py
```

访问 http://127.0.0.1:5000

### 打包为 EXE

双击 `build.bat`，选择打包模式：

- **1. onedir** — 目录模式，生成多个文件，启动速度快
- **2. onefile** — 单文件模式，生成单个 exe，便于分发

等待几分钟后在 `dist/` 目录生成输出文件。

### 依赖安装

```bash
pip install flask python-docx lxml pyinstaller pyinstaller-hooks-contrib psutil
```

## 使用说明

1. 在左侧面板添加替换规则（原文本 → 替换文本）
2. 上传需要处理的文档文件（支持拖拽）
3. 选择「正向替换」或「反向还原」模式
4. 点击开始处理，等待完成
5. 下载处理后的文件

## 数据目录

- **默认数据目录**：软件所在目录
- **配置文件位置**：软件目录下的 `config.json`
- **便携模式**：所有配置和数据保存在软件目录，便于携带
- **自定义目录**：可在设置中修改数据目录路径

## 无痕模式

- 点击无痕模式按钮启用
- 该模式下不会保留任何上传文件和处理结果
- 再次点击可退出无痕模式

## 技术栈

- 后端：Python 3.8 + Flask
- 前端：原生 HTML/CSS/JavaScript
- 文档处理：python-docx + lxml
- 打包：PyInstaller 5.13.2

## 依赖

| 库 | 用途 |
|---|---|
| flask | Web 框架 |
| python-docx | Word 文档处理 |
| lxml | XML 解析 |
| pyinstaller | 打包工具 |
| psutil | 进程管理 |

## 作者

**keloder**

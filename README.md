# Word转PDF工具

一个带图形界面的工具，用于批量将Word文档转换为PDF文件，支持Microsoft Word和WPS Office。

## 功能特性

- ✅ **批量转换**：支持批量转换文件夹中的所有Word文档
- ✅ **多Office支持**：支持Microsoft Word和WPS Office
- ✅ **智能检测**：自动检测系统中可用的Office软件
- ✅ **三种模式**：
  - 自动检测（推荐）：自动选择可用的Office软件
  - Microsoft Word：使用Word应用程序转换
  - WPS Office：使用WPS应用程序转换
- ✅ **实时进度**：显示转换进度和当前文件状态
- ✅ **详细日志**：记录转换过程中的详细信息
- ✅ **错误处理**：自动处理常见错误并提供解决建议

## 系统要求

- Windows操作系统
- 已安装Microsoft Word或WPS Office任一软件
- Python 3.6+（如果运行源代码）

## 安装

### 方式1：使用预编译的exe文件（推荐）

1. 从[Releases](https://github.com/longjueyuyu/word-to-pdf-converter/releases)下载最新版本
2. 解压后双击运行`Word转PDF工具.exe`

### 方式2：从源代码运行

1. 克隆仓库：
   ```bash
   git clone https://github.com/longjueyuyu/word-to-pdf-converter.git
   ```

2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

3. 运行程序：
   ```bash
   python word_to_pdf_converter.py
   ```

## 使用方法

1. 启动程序
2. 点击"📁 选择目录"按钮，选择包含Word文档的文件夹
3. 程序会自动扫描并显示找到的Word文件数量
4. 选择转换方式（自动检测/Word/WPS）
5. 点击"🔄 开始批量转换"按钮开始转换
6. 查看详细日志了解转换进度
7. 转换完成后，PDF文件将保存在原Word文件的同目录下

## 常见问题

**Q: 提示"未检测到可用的Office应用程序"？**
A: 请确保已安装Microsoft Word或WPS Office任一软件。

**Q: 转换失败怎么办？**
A: 查看详细日志中的错误信息，通常失败原因包括：
- Word文档损坏
- 文档使用了特殊字体
- 文件被其他程序占用
- Office软件未正确安装

## 技术栈

- Python 3.12
- tkinter（GUI界面）
- win32com（Office自动化）
- PyInstaller（打包工具）

## 许可证

MIT License
"""
打包Word转PDF工具为可执行文件
使用PyInstaller将程序打包成独立的exe文件
"""
import os
import sys
import subprocess

print("="*70)
print("Word转PDF工具 - 打包脚本")
print("="*70)

# 检查是否安装了PyInstaller
try:
    import PyInstaller
    print("\n✓ PyInstaller已安装")
except ImportError:
    print("\n✗ 未安装PyInstaller")
    print("正在安装PyInstaller...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    print("✓ PyInstaller安装完成")

# 打包命令
print("\n开始打包...")
print("-"*70)

# PyInstaller参数说明:
# --onefile: 打包成单个exe文件
# --windowed: 不显示控制台窗口(GUI程序)
# --name: 指定exe文件名
# --icon: 指定图标(如果有)
# --add-data: 添加额外文件
# --hidden-import: 添加隐藏导入的模块

cmd = [
    "pyinstaller",
    "--onefile",                    # 单文件模式
    "--windowed",                   # GUI模式,不显示控制台
    "--name=Word转PDF工具",          # 程序名称
    "--clean",                      # 清理临时文件
    # 添加隐藏导入
    "--hidden-import=win32com.client",
    "--hidden-import=pythoncom",
    "--hidden-import=pywintypes",
    "--hidden-import=win32com.gen_py",
    # 排除不需要的模块以减小体积
    "--exclude-module=matplotlib",
    "--exclude-module=numpy",
    "--exclude-module=pandas",
    # 主程序文件
    "word_to_pdf_converter.py"
]

print(f"执行命令: {' '.join(cmd)}")
print()

try:
    # 执行打包命令
    result = subprocess.run(cmd, check=True, capture_output=True, text=True)
    
    print("\n" + "="*70)
    print("✓✓✓ 打包成功！")
    print("="*70)
    print(f"\n生成的exe文件位置: dist\\Word转PDF工具.exe")
    print(f"\n文件大小: ", end="")
    
    exe_path = os.path.join("dist", "Word转PDF工具.exe")
    if os.path.exists(exe_path):
        size = os.path.getsize(exe_path)
        if size > 1024 * 1024:
            print(f"{size / (1024*1024):.2f} MB")
        else:
            print(f"{size / 1024:.2f} KB")
    
    print("\n使用说明:")
    print("  1. 将 dist\\Word转PDF工具.exe 复制到任意位置")
    print("  2. 双击运行即可,无需安装Python")
    print("  3. 确保电脑已安装Microsoft Word")
    print("\n注意事项:")
    print("  - 首次运行可能需要几秒钟启动")
    print("  - 某些杀毒软件可能误报,添加信任即可")
    print("="*70)
    
except subprocess.CalledProcessError as e:
    print("\n" + "="*70)
    print("✗✗✗ 打包失败！")
    print("="*70)
    print(f"\n错误信息: {e}")
    if e.output:
        print(f"\n详细输出:\n{e.output}")
    print("\n请检查错误信息并重试")
    sys.exit(1)

except Exception as e:
    print(f"\n✗ 发生错误: {str(e)}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

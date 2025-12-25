"""
自动化GUI测试 - 测试Word转PDF转换器的Word应用程序转换功能
"""
import os
import time
import sys

# 设置测试目录
test_dir = r"D:\my_code\wtp\test_docs"

print("="*60)
print("自动化GUI测试 - Word应用程序转换")
print("="*60)
print(f"\n测试目录: {test_dir}\n")

# 导入转换器类
sys.path.insert(0, r"D:\my_code\wtp")
from word_to_pdf_converter import WordToPdfConverter
import tkinter as tk

# 创建GUI实例
root = tk.Tk()
root.withdraw()  # 隐藏主窗口，只在后台运行

app = WordToPdfConverter(root)

# 模拟用户操作
print("1. 设置文件夹路径...")
app.selected_folder = test_dir
app.folder_path_var.set(test_dir)

print("2. 扫描Word文件...")
app.scan_word_files()

print(f"   找到 {len(app.word_files)} 个文件")
for i, f in enumerate(app.word_files, 1):
    print(f"   {i}. {os.path.basename(f)}")

# 清理之前的PDF
print("\n3. 清理旧PDF文件...")
for f in os.listdir(test_dir):
    if f.endswith('.pdf'):
        os.remove(os.path.join(test_dir, f))
        print(f"   删除: {f}")

# 测试Word应用程序方式
print("\n4. 设置转换方式为: Word应用程序")
app.use_word_app.set(True)

print("\n5. 开始转换...")
print("-"*60)

# 启动转换
app.start_conversion()

# 等待转换完成
print("\n等待转换完成...")
max_wait = 60  # 最多等待60秒
wait_time = 0
while app.is_converting and wait_time < max_wait:
    time.sleep(1)
    root.update()
    wait_time += 1
    if wait_time % 5 == 0:
        print(f"  已等待 {wait_time} 秒...")

if wait_time >= max_wait:
    print("警告: 转换超时!")
else:
    print(f"  转换完成, 耗时 {wait_time} 秒")

print("\n" + "="*60)
print("转换结果:")
print("="*60)

# 检查结果
pdf_files = []
for f in os.listdir(test_dir):
    if f.endswith('.pdf'):
        pdf_files.append(f)
        size = os.path.getsize(os.path.join(test_dir, f))
        print(f"✓ {f} ({size} 字节)")

print(f"\n总计: 成功生成 {len(pdf_files)} 个PDF文件")
print("="*60)

# 关闭GUI
root.destroy()

print("\n测试完成！")


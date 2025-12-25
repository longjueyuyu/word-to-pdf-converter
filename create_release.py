"""
创建发布包
将exe和说明文档打包成zip文件,方便分发
"""
import os
import shutil
import zipfile
from datetime import datetime

print("="*70)
print("创建发布包")
print("="*70)

# 检查exe是否存在
exe_path = r"dist\Word转PDF工具.exe"
if not os.path.exists(exe_path):
    print("\n✗ 错误: 未找到exe文件")
    print("请先运行 python build_exe.py 进行打包")
    exit(1)

print(f"\n✓ 找到exe文件: {exe_path}")
exe_size = os.path.getsize(exe_path) / (1024*1024)
print(f"  文件大小: {exe_size:.2f} MB")

# 创建发布目录
release_dir = "release"
if os.path.exists(release_dir):
    shutil.rmtree(release_dir)
os.makedirs(release_dir)

print(f"\n✓ 创建发布目录: {release_dir}")

# 复制文件
print("\n正在复制文件...")
files_to_copy = [
    ("dist\\Word转PDF工具.exe", "Word转PDF工具.exe"),
    ("dist\\使用说明.txt", "使用说明.txt"),
]

for src, dst in files_to_copy:
    if os.path.exists(src):
        dst_path = os.path.join(release_dir, dst)
        shutil.copy2(src, dst_path)
        print(f"  ✓ {dst}")
    else:
        print(f"  ⚠ 跳过: {dst} (源文件不存在)")

# 创建zip压缩包
version = datetime.now().strftime("%Y%m%d")
zip_name = f"Word转PDF工具_v1.0_{version}.zip"
zip_path = zip_name

print(f"\n正在创建压缩包: {zip_name}")

with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for root, dirs, files in os.walk(release_dir):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, release_dir)
            zipf.write(file_path, arcname)
            print(f"  添加: {arcname}")

zip_size = os.path.getsize(zip_path) / (1024*1024)

print("\n" + "="*70)
print("✓✓✓ 发布包创建成功！")
print("="*70)
print(f"\n压缩包位置: {zip_path}")
print(f"压缩包大小: {zip_size:.2f} MB")
print(f"\n包含文件:")
print("  - Word转PDF工具.exe")
print("  - 使用说明.txt")
print("\n分发说明:")
print("  1. 将zip文件发送给用户")
print("  2. 用户解压后双击exe即可使用")
print("  3. 确保用户电脑已安装Microsoft Word")
print("="*70)

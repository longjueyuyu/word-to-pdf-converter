"""
简单命令行测试 - 直接测试Word转换功能
"""
import os
import time

print("="*60)
print("Word转PDF 简单测试")
print("="*60)

# 测试目录
test_dir = r"D:\my_code\wtp\test_docs"
print(f"\n测试目录: {test_dir}")

# 清理旧PDF
print("\n1. 清理旧PDF文件...")
for f in os.listdir(test_dir):
    if f.endswith('.pdf'):
        os.remove(os.path.join(test_dir, f))
        print(f"   删除: {f}")

# 获取Word文件
word_files = []
for f in os.listdir(test_dir):
    if f.endswith('.docx'):
        word_files.append(os.path.join(test_dir, f))

print(f"\n2. 找到 {len(word_files)} 个Word文件:")
for i, f in enumerate(word_files, 1):
    print(f"   {i}. {os.path.basename(f)}")

# 使用Word应用程序转换
print("\n3. 使用Word应用程序转换...")
print("-"*60)

try:
    import win32com.client
    import pythoncom
    
    success = 0
    failed = 0
    total_time = 0
    
    for i, word_file in enumerate(word_files, 1):
        filename = os.path.basename(word_file)
        pdf_file = word_file.replace('.docx', '.pdf')
        
        print(f"\n[{i}/{len(word_files)}] 正在转换: {filename}")
        
        word = None
        doc = None
        try:
            start = time.time()
            
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            
            doc = word.Documents.Open(os.path.abspath(word_file))
            doc.SaveAs(os.path.abspath(pdf_file), FileFormat=17)
            doc.Close(False)
            
            elapsed = time.time() - start
            total_time += elapsed
            
            if os.path.exists(pdf_file):
                size = os.path.getsize(pdf_file)
                print(f"  ✓ 成功: {os.path.basename(pdf_file)}")
                print(f"     大小: {size} 字节")
                print(f"     耗时: {elapsed:.2f} 秒")
                success += 1
            else:
                print(f"  ✗ 失败: PDF文件未生成")
                failed += 1
                
        except Exception as e:
            print(f"  ✗ 错误: {str(e)}")
            failed += 1
            import traceback
            traceback.print_exc()
        finally:
            try:
                if doc:
                    doc.Close(False)
            except:
                pass
            try:
                if word:
                    word.Quit()
            except:
                pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    print("\n" + "="*60)
    print("转换结果:")
    print("="*60)
    print(f"成功: {success} 个")
    print(f"失败: {failed} 个")
    if success > 0:
        print(f"平均耗时: {total_time/success:.2f} 秒/文件")
    print("="*60)
    
    # 列出生成的PDF
    print("\n生成的PDF文件:")
    for f in os.listdir(test_dir):
        if f.endswith('.pdf'):
            size = os.path.getsize(os.path.join(test_dir, f))
            print(f"  ✓ {f} ({size} 字节)")
    
except Exception as e:
    print(f"\n错误: {str(e)}")
    import traceback
    traceback.print_exc()

print("\n测试完成！")


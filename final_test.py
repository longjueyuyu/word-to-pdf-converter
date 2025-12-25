"""
最终验证测试 - Word转PDF转换器
测试目录: D:\my_code\wtp\test_docs
"""
import os
import time

print("="*70)
print("Word转PDF转换器 - 最终验证测试")
print("="*70)

# 测试目录
test_dir = r"D:\my_code\wtp\test_docs"
print(f"\n测试目录: {test_dir}")

# 清理旧PDF文件
print("\n步骤1: 清理旧PDF文件")
print("-"*70)
pdf_count = 0
for f in os.listdir(test_dir):
    if f.endswith('.pdf'):
        os.remove(os.path.join(test_dir, f))
        print(f"  删除: {f}")
        pdf_count += 1
if pdf_count == 0:
    print("  无需清理")
else:
    print(f"  已清理 {pdf_count} 个PDF文件")

# 扫描Word文件
print("\n步骤2: 扫描Word文件")
print("-"*70)
word_files = []
for f in os.listdir(test_dir):
    if f.endswith('.docx') or f.endswith('.doc'):
        word_files.append(os.path.join(test_dir, f))

print(f"找到 {len(word_files)} 个Word文件:")
for i, f in enumerate(word_files, 1):
    size = os.path.getsize(f)
    print(f"  {i}. {os.path.basename(f)} ({size} 字节)")

# 使用Word应用程序转换
print("\n步骤3: 使用Word应用程序转换 (默认方式)")
print("-"*70)

try:
    import win32com.client
    import pythoncom
    
    success = 0
    failed = 0
    total_time = 0
    errors = []
    
    for i, word_file in enumerate(word_files, 1):
        filename = os.path.basename(word_file)
        pdf_file = word_file.replace('.docx', '.pdf').replace('.doc', '.pdf')
        
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
                print(f"  ✓ 成功")
                print(f"     PDF文件: {os.path.basename(pdf_file)}")
                print(f"     文件大小: {size:,} 字节")
                print(f"     转换耗时: {elapsed:.2f} 秒")
                success += 1
            else:
                print(f"  ✗ 失败: PDF文件未生成")
                failed += 1
                errors.append(f"{filename}: PDF文件未生成")
                
        except Exception as e:
            print(f"  ✗ 错误: {str(e)}")
            failed += 1
            errors.append(f"{filename}: {str(e)}")
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
    
    # 显示结果统计
    print("\n" + "="*70)
    print("转换结果统计")
    print("="*70)
    print(f"总文件数: {len(word_files)}")
    print(f"成功转换: {success} 个 ({success*100//len(word_files) if word_files else 0}%)")
    print(f"转换失败: {failed} 个")
    if success > 0:
        print(f"平均耗时: {total_time/success:.2f} 秒/文件")
        print(f"总耗时: {total_time:.2f} 秒")
    
    if errors:
        print("\n错误详情:")
        for err in errors:
            print(f"  - {err}")
    
    # 列出生成的PDF文件
    print("\n" + "="*70)
    print("生成的PDF文件列表")
    print("="*70)
    pdf_files = []
    for f in os.listdir(test_dir):
        if f.endswith('.pdf'):
            pdf_files.append(f)
            size = os.path.getsize(os.path.join(test_dir, f))
            print(f"  ✓ {f}")
            print(f"     大小: {size:,} 字节")
    
    if not pdf_files:
        print("  (无PDF文件)")
    
    # 最终结论
    print("\n" + "="*70)
    if success == len(word_files) and len(word_files) > 0:
        print("✓✓✓ 测试通过！所有文件转换成功！")
    elif success > 0:
        print(f"⚠ 部分成功：{success}/{len(word_files)} 个文件转换成功")
    else:
        print("✗✗✗ 测试失败：没有文件转换成功")
    print("="*70)
    
except Exception as e:
    print(f"\n✗ 程序错误: {str(e)}")
    import traceback
    traceback.print_exc()

print("\n测试完成！\n")


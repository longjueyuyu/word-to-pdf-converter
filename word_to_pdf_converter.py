"""
Word转PDF工具
一个带图形界面的工具，用于批量将Word文件转换为PDF文件
使用Microsoft Word应用程序进行转换
"""
import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread
import time

# 导入win32com用于Word/WPS应用程序转换
try:
    import win32com.client
    import pythoncom
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

# 检测可用的Office应用程序
def detect_office_apps():
    """检测系统中可用的Office应用程序"""
    available_apps = []
    
    if not HAS_WIN32COM:
        return available_apps
    
    # 检测Microsoft Word
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Quit()
        pythoncom.CoUninitialize()
        available_apps.append("Word")
    except:
        pass
    
    # 检测WPS Office (使用KWPS.Application)
    try:
        pythoncom.CoInitialize()
        wps = win32com.client.DispatchEx("KWPS.Application")  # 金山WPS文字
        wps.Quit()
        pythoncom.CoUninitialize()
        available_apps.append("WPS")
    except:
        pass
    
    return available_apps


class WordToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word转PDF工具")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 设置图标(如果有)
        try:
            self.root.iconbitmap(default='default')
        except:
            pass
        
        self.selected_folder = None
        self.word_files = []
        self.is_converting = False
        self.stop_conversion = False
        self.office_app = tk.StringVar(value="auto")  # 转换方式：auto/word/wps
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置用户界面"""
        # 设置主框架的padding
        main_frame = tk.Frame(self.root, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 目录选择部分
        folder_frame = tk.LabelFrame(main_frame, text="目录选择", font=("微软雅黑", 10, "bold"), padx=10, pady=10)
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 文件夹路径显示
        path_frame = tk.Frame(folder_frame)
        path_frame.pack(fill=tk.X, pady=5)
        
        self.folder_path_var = tk.StringVar(value="请选择包含Word文档的文件夹...")
        folder_label = tk.Label(path_frame, textvariable=self.folder_path_var, 
                               bg="white", relief=tk.SUNKEN, anchor="w", padx=5, pady=8)
        folder_label.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        select_btn = tk.Button(path_frame, text="📁 选择目录", command=self.select_folder,
                              font=("微软雅黑", 9), padx=15, pady=5,
                              bg="#2196F3", fg="white", cursor="hand2")
        select_btn.pack(side=tk.RIGHT)
        
        # 批量任务状态部分
        status_frame = tk.LabelFrame(main_frame, text="批量任务状态", font=("微软雅黑", 10, "bold"), padx=10, pady=10)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 状态信息
        status_info_frame = tk.Frame(status_frame)
        status_info_frame.pack(fill=tk.X, pady=5)
        
        self.status_text_var = tk.StringVar(value="等待选择目录...")
        status_text = tk.Label(status_info_frame, textvariable=self.status_text_var,
                              font=("微软雅黑", 9), anchor="w")
        status_text.pack(fill=tk.X, pady=2)
        
        self.file_count_var = tk.StringVar(value="当前文件: -")
        file_count_label = tk.Label(status_info_frame, textvariable=self.file_count_var,
                                   font=("微软雅黑", 9), anchor="w")
        file_count_label.pack(fill=tk.X, pady=2)
        
        # 转换控制部分
        control_frame = tk.LabelFrame(main_frame, text="转换控制", font=("微软雅黑", 10, "bold"), padx=10, pady=10)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 转换方式选择
        method_frame = tk.LabelFrame(control_frame, text="转换方式", font=("微软雅黑", 9))
        method_frame.pack(fill=tk.X, pady=5)
        
        radio_frame = tk.Frame(method_frame)
        radio_frame.pack(fill=tk.X, pady=5)
        
        tk.Radiobutton(radio_frame, text="自动检测（推荐）", 
                      variable=self.office_app, value="auto",
                      font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(radio_frame, text="使用Microsoft Word", 
                      variable=self.office_app, value="word",
                      font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=10)
        
        tk.Radiobutton(radio_frame, text="使用WPS Office", 
                      variable=self.office_app, value="wps",
                      font=("微软雅黑", 9)).pack(side=tk.LEFT, padx=10)
        
        info_label = tk.Label(method_frame, 
                             text="💡 需要已安装Microsoft Word或WPS Office",
                             font=("微软雅黑", 8), foreground="blue", anchor="w")
        info_label.pack(fill=tk.X, pady=2)
        
        # 开始按钮
        button_frame = tk.Frame(control_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        self.start_btn = tk.Button(button_frame, text="🔄 开始批量转换", 
                                   command=self.start_conversion,
                                   font=("微软雅黑", 10), state=tk.DISABLED,
                                   bg="#4CAF50", fg="white", padx=20, pady=10,
                                   cursor="hand2", relief=tk.RAISED)
        self.start_btn.pack(pady=5)
        
        # 停止按钮
        self.stop_btn = tk.Button(button_frame, text="⏸ 停止转换",
                                  command=self.stop_conversion_process,
                                  font=("微软雅黑", 9), state=tk.DISABLED,
                                  bg="#f44336", fg="white", padx=15, pady=8,
                                  cursor="hand2")
        self.stop_btn.pack(pady=5)
        
        # 总进度
        progress_frame = tk.Frame(control_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.total_progress_var = tk.StringVar(value="总进度: 0%")
        total_progress_label = tk.Label(progress_frame, textvariable=self.total_progress_var,
                                       font=("微软雅黑", 9), anchor="w")
        total_progress_label.pack(fill=tk.X, pady=2)
        
        self.total_progress_bar = ttk.Progressbar(progress_frame, length=400, mode='determinate')
        self.total_progress_bar.pack(fill=tk.X, pady=5)
        
        # 当前文件进度
        current_frame = tk.Frame(control_frame)
        current_frame.pack(fill=tk.X, pady=5)
        
        self.current_progress_var = tk.StringVar(value="当前文件进度: 0%")
        current_progress_label = tk.Label(current_frame, textvariable=self.current_progress_var,
                                         font=("微软雅黑", 9), anchor="w")
        current_progress_label.pack(fill=tk.X, pady=2)
        
        self.current_progress_bar = ttk.Progressbar(current_frame, length=400, mode='determinate')
        self.current_progress_bar.pack(fill=tk.X, pady=5)
        
        # 详细日志部分
        log_frame = tk.LabelFrame(main_frame, text="详细日志", font=("微软雅黑", 10, "bold"), padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建文本框和滚动条
        log_scroll = tk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD, 
                               yscrollcommand=log_scroll.set, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)
        
    def select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含Word文档的文件夹")
        if folder:
            self.selected_folder = folder
            self.folder_path_var.set(folder)
            self.scan_word_files()
            
    def scan_word_files(self):
        """扫描文件夹中的Word文件"""
        if not self.selected_folder:
            return
        
        self.word_files = []
        extensions = ['.doc', '.docx']
        
        for root, dirs, files in os.walk(self.selected_folder):
            for file in files:
                if any(file.lower().endswith(ext) for ext in extensions):
                    full_path = os.path.join(root, file)
                    self.word_files.append(full_path)
        
        count = len(self.word_files)
        self.file_count_var.set(f"当前文件: {count}")
        
        if count > 0:
            self.status_text_var.set(f"找到 {count} 个Word文件，点击开始转换")
            self.start_btn.config(state=tk.NORMAL, bg="#4CAF50", fg="white")
            self.log_message(f"✓ 扫描完成，找到 {count} 个Word文件")
            
            # 显示文件列表
            for i, file in enumerate(self.word_files[:5], 1):  # 只显示前5个
                self.log_message(f"  {i}. {os.path.basename(file)}")
            if count > 5:
                self.log_message(f"  ... 还有 {count - 5} 个文件")
        else:
            self.status_text_var.set("未找到Word文件")
            self.start_btn.config(state=tk.DISABLED, bg="#cccccc", fg="#666666")
            self.log_message("⚠ 该文件夹中没有找到Word文件")
            
    def log_message(self, message):
        """添加日志消息"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def start_conversion(self):
        """开始转换"""
        if self.is_converting:
            return
        
        if not self.word_files:
            messagebox.showwarning("警告", "没有找到Word文件")
            return
        
        # 检查Office应用程序是否可用
        self.log_message("\n检测转换环境...")
        
        if not HAS_WIN32COM:
            error_msg = "错误: 未安装pywin32库"
            self.log_message(f"\n✗ {error_msg}")
            self.log_message("解决方法: pip install pywin32")
            messagebox.showerror("错误", 
                               "需要安装pywin32库：\n\npip install pywin32")
            return
        
        # 检测可用的应用程序
        self.log_message("正在检测Office应用程序...")
        self.log_message("  - pywin32库: ✓ 已安装")
        
        available_apps = detect_office_apps()
        
        if "Word" in available_apps:
            self.log_message("  - Microsoft Word: ✓ 已安装")
        else:
            self.log_message("  - Microsoft Word: ✗ 未检测到")
        
        if "WPS" in available_apps:
            self.log_message("  - WPS Office: ✓ 已安装")
        else:
            self.log_message("  - WPS Office: ✗ 未检测到")
        
        if not available_apps:
            error_msg = "未检测到可用的Office应用程序"
            self.log_message(f"\n✗ {error_msg}")
            self.log_message("\n请安装以下任一软件:")
            self.log_message("  1. Microsoft Word")
            self.log_message("  2. WPS Office")
            messagebox.showerror("错误",
                               f"未检测到可用的Office应用程序！\n\n" +
                               f"请安装Microsoft Word或WPS Office")
            return
        
        # 根据用户选择确定使用哪个应用
        selected = self.office_app.get()
        if selected == "auto":
            # 自动模式：优先Word，其次WPS
            if "Word" in available_apps:
                self.log_message("\n转换方式: Microsoft Word（自动检测）")
            elif "WPS" in available_apps:
                self.log_message("\n转换方式: WPS Office（自动检测）")
        elif selected == "word":
            if "Word" not in available_apps:
                error_msg = "未检测到Microsoft Word"
                self.log_message(f"\n✗ {error_msg}")
                messagebox.showerror("错误", "未检测到Microsoft Word！\n\n请安装Word或选择其他转换方式")
                return
            self.log_message("\n转换方式: Microsoft Word")
        elif selected == "wps":
            if "WPS" not in available_apps:
                error_msg = "未检测到WPS Office"
                self.log_message(f"\n✗ {error_msg}")
                messagebox.showerror("错误", "未检测到WPS Office！\n\n请安装WPS或选择其他转换方式")
                return
            self.log_message("\n转换方式: WPS Office")
        
        self.log_message("✓ 环境检测通过\n")
        
        self.is_converting = True
        self.stop_conversion = False
        self.start_btn.config(state=tk.DISABLED, bg="#cccccc", fg="#666666")
        self.stop_btn.config(state=tk.NORMAL)
        
        # 在新线程中执行转换
        thread = Thread(target=self.convert_files, daemon=True)
        thread.start()
        
    def stop_conversion_process(self):
        """停止转换过程"""
        if self.is_converting:
            self.stop_conversion = True
            self.log_message("\n⚠ 用户请求停止转换...")
            self.status_text_var.set("正在停止转换...")
    
    def convert_files(self):
        """转换文件（在后台线程中运行）"""
        total_files = len(self.word_files)
        converted_count = 0
        failed_count = 0
        failed_files = []  # 记录失败的文件
        
        self.log_message("\n" + "="*60)
        self.log_message("开始批量转换...")
        self.log_message("="*60 + "\n")
        
        for i, word_file in enumerate(self.word_files, 1):
            # 检查是否需要停止
            if self.stop_conversion:
                self.log_message("\n⚠ 转换已被用户停止")
                break
                
            try:
                filename = os.path.basename(word_file)
                self.status_text_var.set(f"正在转换文件: {filename}")
                self.log_message(f"[{i}/{total_files}] 正在转换: {filename}")
                
                # 更新当前文件进度
                self.current_progress_var.set(f"当前文件进度: 0%")
                self.current_progress_bar['value'] = 0
                
                # 生成PDF文件路径
                pdf_file = os.path.splitext(word_file)[0] + '.pdf'
                
                # 执行转换
                success = self.convert_word_to_pdf(word_file, pdf_file)
                
                if success:
                    converted_count += 1
                    self.log_message(f"  ✓ 转换成功: {os.path.basename(pdf_file)}")
                else:
                    failed_count += 1
                    failed_files.append(filename)
                    self.log_message(f"  ✗ 转换失败: {filename} (详见错误信息)")
                
                # 更新当前文件进度为100%
                self.current_progress_var.set(f"当前文件进度: 100%")
                self.current_progress_bar['value'] = 100
                
                # 更新总进度
                total_progress = int((i / total_files) * 100)
                self.total_progress_var.set(f"总进度: {total_progress}%")
                self.total_progress_bar['value'] = total_progress
                
            except Exception as e:
                failed_count += 1
                failed_files.append(filename)
                self.log_message(f"  ✗ 转换异常: {filename}")
                self.log_message(f"     错误详情: {str(e)}")
        
        # 转换完成
        self.log_message("\n" + "="*60)
        if self.stop_conversion:
            self.log_message(f"转换已停止！")
            self.log_message(f"已处理: {converted_count + failed_count}/{total_files} 个")
        else:
            self.log_message(f"转换完成！")
        self.log_message(f"成功: {converted_count} 个，失败: {failed_count} 个")
        
        # 显示失败文件列表
        if failed_files:
            self.log_message("\n失败文件列表:")
            for i, file in enumerate(failed_files, 1):
                self.log_message(f"  {i}. {file}")
            self.log_message("\n建议: 请手动用Word打开上述文件检查是否有错误")
        
        self.log_message("="*60)
        
        status_msg = f"转换{'(已停止)' if self.stop_conversion else '完成'}，成功 {converted_count} 个，失败 {failed_count} 个"
        self.status_text_var.set(status_msg)
        self.is_converting = False
        self.stop_conversion = False
        self.start_btn.config(state=tk.NORMAL, bg="#4CAF50", fg="white")
        self.stop_btn.config(state=tk.DISABLED)
        
        messagebox.showinfo("完成", 
                          f"转换{'(已停止)' if self.stop_conversion else '完成'}！\n\n成功: {converted_count} 个\n失败: {failed_count} 个")
    
    def convert_word_to_pdf(self, word_path, pdf_path):
        """转换Word文档为PDF"""
        # 根据用户选择确定使用哪个应用
        selected = self.office_app.get()
        available_apps = detect_office_apps()
        
        # 确定实际使用的应用
        use_app = None
        if selected == "auto":
            # 自动模式：优先Word，其次WPS
            if "Word" in available_apps:
                use_app = "word"
            elif "WPS" in available_apps:
                use_app = "wps"
        elif selected == "word" and "Word" in available_apps:
            use_app = "word"
        elif selected == "wps" and "WPS" in available_apps:
            use_app = "wps"
        
        if use_app == "word":
            return self.convert_with_word(word_path, pdf_path)
        elif use_app == "wps":
            return self.convert_with_wps(word_path, pdf_path)
        else:
            self.log_message("     ✗ 未找到可用的转换应用")
            return False
    
    def convert_with_word(self, word_path, pdf_path):
        """使用Microsoft Word转换"""
        word = None
        doc = None
        try:
            if not HAS_WIN32COM:
                raise Exception("未安装pywin32库")
            
            pythoncom.CoInitialize()  # 初始化COM
            
            word = win32com.client.DispatchEx("Word.Application")  # 使用DispatchEx创建新实例
            word.Visible = False
            word.DisplayAlerts = 0  # 禁用警告对话框
            
            # 打开文档，忽略缺失字体警告
            doc = word.Documents.Open(
                os.path.abspath(word_path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Revert=False
            )
            
            # 另存为PDF - 使用最简单的参数以兼容所有Word版本
            try:
                # 尝试使用标准参数
                doc.SaveAs(
                    os.path.abspath(pdf_path),
                    FileFormat=17  # wdFormatPDF
                )
            except Exception as e:
                # 如果失败，使用最基本的参数
                doc.SaveAs(os.path.abspath(pdf_path), 17)
            
            doc.Close(False)  # 关闭文档不保存
            
            return True
            
        except Exception as e:
            error_str = str(e)
            
            # 分析常见错误原因
            if '此命令无效' in error_str or 'Command failed' in error_str:
                self.log_message(f"     ⚠ Word文档问题: 该文档可能包含:")
                self.log_message(f"        - 缺失的字体或特殊字体")
                self.log_message(f"        - 受保护的内容")
                self.log_message(f"        - 损坏的格式")
                self.log_message(f"     建议: 手动用Word打开文档，更换字体后再试")
            elif '没有注册类' in error_str or 'Class not registered' in error_str:
                self.log_message(f"     ⚠ Word未正确安装或注册")
            elif '访拒绝' in error_str or 'Access denied' in error_str:
                self.log_message(f"     ⚠ 文件权限问题或文件被占用")
            else:
                self.log_message(f"     Word转换错误: {error_str}")
            
            return False
        finally:
            # 确保Word进程被正确关闭
            try:
                if doc is not None:
                    doc.Close(False)
            except:
                pass
            try:
                if word is not None:
                    word.Quit()
            except:
                pass
            try:
                pythoncom.CoUninitialize()  # 清理COM
            except:
                pass
    
    def convert_with_wps(self, word_path, pdf_path):
        """使用WPS Office转换"""
        wps = None
        doc = None
        try:
            if not HAS_WIN32COM:
                raise Exception("未安装pywin32库")
            
            pythoncom.CoInitialize()  # 初始化COM
            
            wps = win32com.client.DispatchEx("KWPS.Application")  # 金山WPS文字应用程序
            wps.Visible = False
            wps.DisplayAlerts = 0  # 禁用警告对话框
            
            # 打开文档
            doc = wps.Documents.Open(
                os.path.abspath(word_path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False
            )
            
            # 另存为PDF (WPS使用与Word相同的格式代码)
            try:
                # 尝试使用标准参数
                doc.SaveAs(
                    os.path.abspath(pdf_path),
                    FileFormat=17  # wdFormatPDF
                )
            except Exception as e:
                # 如果失败，使用最基本的参数
                doc.SaveAs(os.path.abspath(pdf_path), 17)
            
            doc.Close(False)  # 关闭文档不保存
            
            return True
            
        except Exception as e:
            error_str = str(e)
            
            # 分析常见错误原因
            if '此命令无效' in error_str or 'Command failed' in error_str:
                self.log_message(f"     ⚠ 文档问题: 该文档可能包含:")
                self.log_message(f"        - 缺失的字体或特殊字体")
                self.log_message(f"        - 受保护的内容")
                self.log_message(f"        - 损坏的格式")
                self.log_message(f"     建议: 手动用WPS打开文档检查")
            elif '没有注册类' in error_str or 'Class not registered' in error_str:
                self.log_message(f"     ⚠ WPS未正确安装或注册")
            elif '访拒绝' in error_str or 'Access denied' in error_str:
                self.log_message(f"     ⚠ 文件权限问题或文件被占用")
            else:
                self.log_message(f"     WPS转换错误: {error_str}")
            
            return False
        finally:
            # 确保WPS进程被正确关闭
            try:
                if doc is not None:
                    doc.Close(False)
            except:
                pass
            try:
                if wps is not None:
                    wps.Quit()
            except:
                pass
            try:
                pythoncom.CoUninitialize()  # 清理COM
            except:
                pass


def main():
    """主函数"""
    root = tk.Tk()
    app = WordToPdfConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()

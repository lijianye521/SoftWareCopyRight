import tkinter as tk
from tkinter import ttk, font, Toplevel, messagebox, filedialog
import os

class FontSelector:
    @staticmethod
    def choose_font(parent, font_var):
        """打开字体选择对话框"""
        font_dialog = Toplevel(parent)
        font_dialog.title("选择字体")
        font_dialog.geometry("400x300")
        font_dialog.resizable(False, False)
        font_dialog.transient(parent)
        
        # 获取系统字体列表
        font_list = list(font.families())
        font_list.sort()
        
        # 创建字体列表框
        font_listbox = tk.Listbox(font_dialog, font=('楷体', 12), height=10)
        font_listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 添加字体到列表框
        for f in font_list:
            font_listbox.insert(tk.END, f)
        
        # 如果当前字体在列表中，选中它
        current_font = font_var.get()
        if current_font in font_list:
            index = font_list.index(current_font)
            font_listbox.see(index)
            font_listbox.selection_set(index)
        
        # 确定按钮回调
        def on_select():
            selection = font_listbox.curselection()
            if selection:
                selected_font = font_list[selection[0]]
                font_var.set(selected_font)
            font_dialog.destroy()
        
        # 确定按钮
        select_btn = ttk.Button(font_dialog, text="确定", command=on_select)
        select_btn.pack(pady=10)

class CustomModeDialog:
    @staticmethod
    def show_dialog(parent, config_manager, mode_selector_callback):
        """显示自定义模式保存对话框"""
        mode_dialog = Toplevel(parent)
        mode_dialog.title("保存自定义模式")
        mode_dialog.geometry("400x150")
        mode_dialog.resizable(False, False)
        mode_dialog.transient(parent)
        
        ttk.Label(mode_dialog, text="请输入模式名称:").pack(pady=(20, 10))
        
        mode_name_var = tk.StringVar()
        mode_name_entry = ttk.Entry(mode_dialog, textvariable=mode_name_var, width=30, font=('微软雅黑', 12))
        mode_name_entry.pack(padx=20, fill=tk.X)
        
        def on_save():
            mode_name = mode_name_var.get().strip()
            if not mode_name:
                messagebox.showerror("错误", "模式名称不能为空")
                return
                
            if mode_name in ["前端模式", "后端模式", "混合模式", "其他模式"]:
                messagebox.showerror("错误", "不能使用预设模式名称")
                return
            
            # 这里需要从外部传入扩展名列表
            # 在实际使用时会通过回调函数处理
            if hasattr(on_save, 'extensions'):
                config_manager.add_custom_mode(mode_name, on_save.extensions)
                mode_selector_callback(mode_name)
                messagebox.showinfo("成功", f"已保存模式 '{mode_name}'")
                mode_dialog.destroy()
        
        # 按钮框架
        btn_frame = ttk.Frame(mode_dialog)
        btn_frame.pack(pady=20, fill=tk.X)
        
        ttk.Button(btn_frame, text="保存", command=on_save).pack(side=tk.RIGHT, padx=10)
        ttk.Button(btn_frame, text="取消", command=mode_dialog.destroy).pack(side=tk.RIGHT, padx=10)
        
        # 聚焦输入框
        mode_name_entry.focus_set()
        
        return on_save  # 返回回调函数以便设置扩展名

class ProgressWindow:
    def __init__(self, parent, total_files):
        self.window = Toplevel(parent)
        self.window.title("正在生成")
        self.window.geometry("400x120")
        self.window.resizable(False, False)
        self.window.transient(parent)  # 保持在主窗口之上
        
        ttk.Label(self.window, text="正在处理文件...", font=('微软雅黑', 12)).pack(pady=10)
        
        self.progress_bar = ttk.Progressbar(self.window, orient="horizontal", length=350, mode="determinate")
        self.progress_bar.pack(pady=5)
        self.progress_bar["maximum"] = total_files
        
        self.progress_label = ttk.Label(self.window, text="", font=('微软雅黑', 9))
        self.progress_label.pack(pady=5)
        
        parent.update_idletasks()
    
    def update_progress(self, current, filename):
        """更新进度"""
        self.progress_bar['value'] = current
        self.progress_label['text'] = f"正在处理: {filename}"
        self.window.update_idletasks()
    
    def destroy(self):
        """销毁窗口"""
        if self.window:
            self.window.destroy()
            self.window = None

class SimilarityAnalysisFrame(ttk.Frame):
    def __init__(self, parent, analyzer):
        super().__init__(parent, padding=20)
        self.analyzer = analyzer
        self.file_paths = []
        
        # 创建文件列表框架
        self.create_file_list_frame()
        
        # 创建按钮框架
        self.create_button_frame()
        
        # 创建结果框架
        self.create_result_frame()
    
    def create_file_list_frame(self):
        """创建文件列表框架"""
        file_list_frame = ttk.LabelFrame(self, text="文件列表", padding=15)
        file_list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 创建文件列表
        self.file_listbox = tk.Listbox(file_list_frame, height=10, font=('微软雅黑', 12))
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(file_list_frame, orient="vertical", command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)
    
    def create_button_frame(self):
        """创建按钮框架"""
        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 添加文件按钮
        add_btn = ttk.Button(button_frame, text="添加文件", command=self.add_file)
        add_btn.pack(side=tk.LEFT, padx=5)
        
        # 删除选中文件按钮
        remove_btn = ttk.Button(button_frame, text="删除选中", command=self.remove_selected_file)
        remove_btn.pack(side=tk.LEFT, padx=5)
        
        # 清空列表按钮
        clear_btn = ttk.Button(button_frame, text="清空列表", command=self.clear_files)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 分析按钮
        analyze_btn = ttk.Button(button_frame, text="开始分析", command=self.analyze_files, style='Generate.TButton')
        analyze_btn.pack(side=tk.RIGHT, padx=5)
    
    def create_result_frame(self):
        """创建结果框架"""
        result_frame = ttk.LabelFrame(self, text="分析结果", padding=15)
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建结果文本框
        self.result_text = tk.Text(result_frame, wrap=tk.WORD, font=('微软雅黑', 12), height=10)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        self.result_text.config(state=tk.DISABLED)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
    
    def add_file(self):
        """添加文件"""
        file_paths = filedialog.askopenfilenames(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        
        if file_paths:
            for path in file_paths:
                if path not in self.file_paths:
                    self.file_paths.append(path)
                    self.file_listbox.insert(tk.END, os.path.basename(path))
    
    def remove_selected_file(self):
        """删除选中的文件"""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.file_listbox.delete(index)
            self.file_paths.pop(index)
    
    def clear_files(self):
        """清空文件列表"""
        self.file_listbox.delete(0, tk.END)
        self.file_paths.clear()
    
    def analyze_files(self):
        """分析文件相似度"""
        if len(self.file_paths) < 2:
            messagebox.showwarning("警告", "请至少添加两个文件进行比较")
            return
        
        try:
            # 显示进度窗口
            progress_window = ProgressWindow(self, len(self.file_paths))
            self.update_idletasks()
            
            # 获取相似度报告
            report = self.analyzer.get_similarity_report(self.file_paths)
            
            # 关闭进度窗口
            progress_window.destroy()
            
            # 显示结果
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, report)
            self.result_text.config(state=tk.DISABLED)
            
        except Exception as e:
            messagebox.showerror("错误", f"分析过程中发生错误: {str(e)}")
            if 'progress_window' in locals():
                progress_window.destroy() 
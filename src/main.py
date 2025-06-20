import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pathlib
import datetime as dt
from PIL import Image, ImageTk, ImageDraw  # 用于美化界面图标
import sv_ttk  # 用于现代化Tkinter主题

# 导入自定义模块
from config_manager import ConfigManager
from file_processor import FileProcessor
from document_generator import DocumentGenerator
from similarity_analyzer import SimilarityAnalyzer
from ui_components import FontSelector, CustomModeDialog, ProgressWindow, SimilarityAnalysisFrame

class CodeToWordApp:
    # 添加创建圆角图像的方法
    def create_rounded_button_images(self):
        """创建圆角按钮图像"""
        # 这个方法会在初始化时被调用，用于创建圆角按钮图像
        self.rounded_images = {}
        
    def create_rounded_rectangle(self, width, height, radius, fill_color):
        """创建圆角矩形图像"""
        image = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle([(0, 0), (width, height)], radius, fill=fill_color)
        return ImageTk.PhotoImage(image)
    
    def setup_theme(self):
        """延迟设置主题以提高启动速度"""
        try:
            # 应用Sun Valley主题 - 一个现代化的Tkinter主题，提供圆角和阴影效果
            sv_ttk.set_theme("light")
        except Exception as e:
            print(f"主题加载失败，使用默认主题: {e}")
        
    def __init__(self, root):
        self.root = root
        self.root.title("软件著作权源代码文档生成器")
        self.root.geometry("900x950") # 增大窗口宽度以适应新设计
        
        # 延迟加载主题以提高启动速度
        self.root.after_idle(self.setup_theme)
        
        # 设置应用主题和样式
        self.style = ttk.Style()
        
        # 定义统一的楷体字体和大小
        label_font = ('楷体', 14)
        entry_font = ('楷体', 12)  # 输入框字体改为12号
        button_font = ('楷体', 14)
        
        # 自定义样式
        self.style.configure('TLabel', font=label_font)
        self.style.configure('TEntry', font=entry_font)
        self.style.configure('TButton', font=button_font)
        self.style.configure('TLabelframe.Label', font=('楷体', 18, 'bold'))
        self.style.configure('TCheckbutton', font=label_font)
        self.style.configure('TCombobox', font=entry_font)
        
        self.style.configure('Small.TButton', font=button_font)
        self.style.configure('Add.TButton', font=button_font, padding=(10, 5))
        self.style.configure('Generate.TButton', font=('楷体', 20, 'bold'), padding=(30, 15))
        self.style.configure('Title.TLabel', font=('楷体', 32, 'bold'))
        self.style.configure('Subtitle.TLabel', font=('楷体', 24))
        self.style.configure('Info.TLabel', font=('楷体', 14))
        
        # 创建圆角按钮图像
        self.create_rounded_button_images()
        
        # 设置背景颜色 - 使用更柔和的颜色
        self.root.configure(bg="#f8f9fa")
        
        # 初始化模块化组件
        self.config_manager = ConfigManager()
        self.file_processor = FileProcessor()
        self.document_generator = DocumentGenerator()
        self.similarity_analyzer = SimilarityAnalyzer()
        
        # 延迟加载用户自定义模式以提高启动速度
        self.root.after_idle(self.config_manager.load_config)
        
        # 项目路径列表
        self.project_paths = []
        
        # 创建主框架，使用滚动条
        outer_frame = ttk.Frame(root)
        outer_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        outer_frame.columnconfigure(0, weight=1)
        outer_frame.rowconfigure(0, weight=1)
        
        # 创建标签页
        self.notebook = ttk.Notebook(outer_frame)
        self.notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 第一个标签页 - 源代码导出
        self.code_export_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.code_export_tab, text="源代码导出")
        
        # 第二个标签页 - 文件相似度分析
        self.similarity_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.similarity_tab, text="文件相似度分析")
        
        # 在源代码导出标签页中创建滚动区域
        self.create_code_export_tab()
        
        # 在文件相似度分析标签页中创建内容
        self.create_similarity_tab()
        
        # 使窗口可调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def create_code_export_tab(self):
        """创建源代码导出标签页内容"""
        # 创建Canvas和滚动条
        canvas = tk.Canvas(self.code_export_tab, bg="#f8f9fa", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.code_export_tab, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        
        # 主内容框架
        main_frame = ttk.Frame(canvas, padding="30")
        main_frame.columnconfigure(1, weight=1)
        
        # 创建窗口用于放置main_frame
        canvas_window = canvas.create_window((0, 0), window=main_frame, anchor="nw")
        
        # 配置Canvas滚动区域
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # 设置canvas宽度与main_frame相同
            canvas.itemconfigure(canvas_window, width=event.width)
        
        main_frame.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(canvas_window, width=e.width))
        
        # 绑定鼠标滚轮事件到Canvas
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 标题和图标
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=3, pady=(0, 30))
        
        # 创建标题 - 使用tk.Label而不是ttk.Label以直接控制字体
        title_label = tk.Label(
            header_frame, 
            text="软件源代码导出工具", 
            font=('楷体', 40, 'bold'),
            fg="#4361ee",
            bg="#f8f9fa"  # 与背景色保持一致
        )
        title_label.pack(pady=15)
        
        # 副标题 - 同样使用tk.Label
        subtitle_label = tk.Label(
            header_frame, 
            text="用于软件著作权申请的源代码文档生成器", 
            font=('楷体', 24),
            bg="#f8f9fa"  # 与背景色保持一致
        )
        subtitle_label.pack(pady=5)
        
        # 主根目录设置框架
        root_dir_frame = ttk.LabelFrame(main_frame, text="主根目录设置", padding=15)
        root_dir_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15, padx=5)
        root_dir_frame.columnconfigure(1, weight=1)
        
        ttk.Label(root_dir_frame, text="主根目录:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.root_dir_var = tk.StringVar()
        root_dir_entry = ttk.Entry(root_dir_frame, textvariable=self.root_dir_var, width=50)
        root_dir_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        root_dir_browse_btn = ttk.Button(root_dir_frame, text="浏览...", command=self.choose_root_directory)
        root_dir_browse_btn.grid(row=0, column=2, padx=10, pady=5)
        
        # 添加说明文本
        root_dir_info = ttk.Label(
            root_dir_frame, 
            text="设置主根目录后，所有文件路径将从此目录开始计算。如不设置，则使用各项目路径作为起点。", 
            style='Info.TLabel',
            wraplength=750
        )
        root_dir_info.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # 项目路径选择框架 - 可动态添加多个路径
        paths_frame = ttk.LabelFrame(main_frame, text="项目路径", padding=15)
        paths_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15, padx=5)
        paths_frame.columnconfigure(0, weight=1)
        
        # 保存paths_frame的引用，供add_path_row使用
        self.paths_frame = paths_frame
        
        # 初始化第一个路径选择行
        self.path_vars = []
        self.path_frames = []
        self.add_path_row()
        
        # 添加路径按钮
        add_path_btn = ttk.Button(
            paths_frame, 
            text="+ 添加路径", 
            command=self.add_path_row,
            style='Add.TButton'
        )
        add_path_btn.grid(row=100, column=0, sticky=tk.W, pady=15)
        
        # Gitignore 文件选择框架
        gitignore_frame = ttk.LabelFrame(main_frame, text="Gitignore 设置", padding=15)
        gitignore_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15, padx=5)
        gitignore_frame.columnconfigure(1, weight=1)
        
        ttk.Label(gitignore_frame, text="Gitignore 文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.gitignore_path_var = tk.StringVar()
        gitignore_entry = ttk.Entry(gitignore_frame, textvariable=self.gitignore_path_var, width=50)
        gitignore_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)
        gitignore_browse_btn = ttk.Button(gitignore_frame, text="浏览...", command=self.choose_gitignore_file)
        gitignore_browse_btn.grid(row=0, column=2, padx=10, pady=5)
        
        # 文件类型模式选择框架
        file_types_frame = ttk.LabelFrame(main_frame, text="文件类型设置", padding=15)
        file_types_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15, padx=5)
        file_types_frame.columnconfigure(1, weight=1)
        
        # 获取所有模式名称
        mode_names = list(self.config_manager.get_file_type_modes().keys()) + ["其他模式"]
        self.file_mode_var = tk.StringVar(value=mode_names[0])
        
        # 模式选择标签和下拉框
        ttk.Label(file_types_frame, text="选择模式:").grid(row=0, column=0, sticky=tk.W, pady=10)
        
        # 模式选择下拉框
        mode_selector = ttk.Combobox(
            file_types_frame, 
            textvariable=self.file_mode_var, 
            values=mode_names, 
            state="readonly", 
            width=20
        )
        mode_selector.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        mode_selector.bind("<<ComboboxSelected>>", self.on_mode_selected)
        
        # 文件类型标签和输入框
        ttk.Label(file_types_frame, text="文件类型:").grid(row=1, column=0, sticky=tk.W, pady=10)
        
        # 自定义文件类型输入框
        self.extension_var = tk.StringVar()
        self.extension_entry = ttk.Entry(file_types_frame, textvariable=self.extension_var)
        self.extension_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        
        # 添加自定义模式按钮
        save_mode_btn = ttk.Button(
            file_types_frame, 
            text="保存为自定义模式", 
            command=self.save_custom_mode
        )
        save_mode_btn.grid(row=2, column=1, sticky=tk.E, pady=(10, 0), padx=5)
        
        # 初始化文件类型
        self.update_extension_field()
        
        # 文档设置框架
        doc_settings_frame = ttk.LabelFrame(main_frame, text="文档设置", padding=15)
        doc_settings_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15, padx=5)
        doc_settings_frame.columnconfigure(1, weight=1)
        
        # 字体选择
        ttk.Label(doc_settings_frame, text="字体:").grid(row=0, column=0, sticky=tk.W, pady=10)
        self.font_var = tk.StringVar(value="微软雅黑")
        
        font_display = ttk.Entry(doc_settings_frame, textvariable=self.font_var, state='readonly')
        font_display.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        
        font_btn = ttk.Button(doc_settings_frame, text="选择字体", command=lambda: FontSelector.choose_font(self.root, self.font_var))
        font_btn.grid(row=0, column=2, padx=5, pady=10)
        
        # 每页行数设置
        ttk.Label(doc_settings_frame, text="每页行数:").grid(row=1, column=0, sticky=tk.W, pady=10)
        self.lines_per_page_var = tk.StringVar(value="55")
        lines_per_page_entry = ttk.Entry(doc_settings_frame, textvariable=self.lines_per_page_var)
        lines_per_page_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        
        # 应用名称和版本号设置
        ttk.Label(doc_settings_frame, text="应用名称:").grid(row=2, column=0, sticky=tk.W, pady=10)
        self.app_name_var = tk.StringVar(value="AI简报")
        self.app_name_font_var = tk.StringVar(value="微软雅黑")
        
        app_name_entry = ttk.Entry(doc_settings_frame, textvariable=self.app_name_var)
        app_name_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        
        app_font_btn = ttk.Button(doc_settings_frame, text="字体", command=lambda: FontSelector.choose_font(self.root, self.app_name_font_var), width=5)
        app_font_btn.grid(row=2, column=2, padx=5, pady=10)
        
        # 版本号和字体选择按钮
        ttk.Label(doc_settings_frame, text="版本号:").grid(row=3, column=0, sticky=tk.W, pady=10)
        self.version_var = tk.StringVar(value="V1.0")
        self.version_font_var = tk.StringVar(value="微软雅黑")
        
        version_entry = ttk.Entry(doc_settings_frame, textvariable=self.version_var)
        version_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        
        version_font_btn = ttk.Button(doc_settings_frame, text="字体", command=lambda: FontSelector.choose_font(self.root, self.version_font_var), width=5)
        version_font_btn.grid(row=3, column=2, padx=5, pady=10)
        
        # 输出设置框架
        output_frame = ttk.LabelFrame(main_frame, text="输出设置", padding=15)
        output_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=15, padx=5)
        output_frame.columnconfigure(1, weight=1)
        
        # 输出文档路径
        ttk.Label(output_frame, text="输出路径:").grid(row=0, column=0, sticky=tk.W, pady=10)
        self.output_path_var = tk.StringVar()
        
        output_path_entry = ttk.Entry(output_frame, textvariable=self.output_path_var)
        output_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=10, padx=5)
        
        output_path_btn = ttk.Button(output_frame, text="浏览...", command=self.choose_output_directory)
        output_path_btn.grid(row=0, column=2, padx=5, pady=10)
        
        # 选项框架
        options_frame = ttk.LabelFrame(output_frame, text="文档选项", padding=10)
        options_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(15, 0))
        options_frame.columnconfigure(0, weight=1)
        options_frame.columnconfigure(1, weight=1)
        
        # 生成目录选项
        self.generate_toc_var = tk.BooleanVar(value=False)
        toc_check = ttk.Checkbutton(options_frame, text="生成目录", variable=self.generate_toc_var)
        toc_check.grid(row=0, column=0, sticky=tk.W, padx=20, pady=10)
        
        # 显示行号选项
        self.show_line_numbers_var = tk.BooleanVar(value=False)
        line_numbers_check = ttk.Checkbutton(options_frame, text="显示行号", variable=self.show_line_numbers_var)
        line_numbers_check.grid(row=0, column=1, sticky=tk.W, padx=20, pady=10)
        
        # 生成按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=30)
        
        # 创建一个带有阴影效果的生成按钮
        generate_btn = ttk.Button(
            button_frame, 
            text="生成文档", 
            command=self.generate_document,
            style='Generate.TButton'
        )
        generate_btn.pack(padx=10)
        
        # 底部信息栏
        footer_frame = ttk.Frame(main_frame)
        footer_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        footer_frame.columnconfigure(1, weight=1)
        
        # 版本信息
        version_label = ttk.Label(footer_frame, text="v1.2.0")
        version_label.grid(row=0, column=0, sticky=tk.W)
        
        # 作者信息
        author_label = ttk.Label(footer_frame, text="作者: lijianye")
        author_label.grid(row=0, column=1, sticky=tk.E)
        
        # 使滚动区域可调整大小
        self.code_export_tab.columnconfigure(0, weight=1)
        self.code_export_tab.rowconfigure(0, weight=1)
    
    def create_similarity_tab(self):
        """创建文件相似度分析标签页内容"""
        # 创建标题框架
        header_frame = ttk.Frame(self.similarity_tab, padding=20)
        header_frame.pack(fill=tk.X)
        
        # 创建标题
        title_label = tk.Label(
            header_frame, 
            text="文件相似度分析", 
            font=('楷体', 40, 'bold'),
            fg="#4361ee",
            bg="#f8f9fa"  # 与背景色保持一致
        )
        title_label.pack(pady=15)
        
        # 副标题
        subtitle_label = tk.Label(
            header_frame, 
            text="分析多个Word文档之间的相似度", 
            font=('楷体', 24),
            bg="#f8f9fa"  # 与背景色保持一致
        )
        subtitle_label.pack(pady=5)
        
        # 创建相似度分析框架
        similarity_frame = SimilarityAnalysisFrame(self.similarity_tab, self.similarity_analyzer)
        similarity_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
    
    def add_path_row(self):
        """添加一个新的项目路径选择行"""
        # 创建新的路径变量
        path_var = tk.StringVar()
        self.path_vars.append(path_var)
        
        row_idx = len(self.path_frames)
        path_frame = ttk.Frame(self.paths_frame)
        path_frame.grid(row=row_idx, column=0, sticky=(tk.W, tk.E), pady=5)
        path_frame.columnconfigure(0, weight=1)
        
        self.path_frames.append(path_frame)
        
        # 路径输入框 - 优化长路径显示性能
        path_entry = ttk.Entry(path_frame, textvariable=path_var, width=50, font=('微软雅黑', 12))
        path_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # 优化长路径的显示 - 当路径过长时自动滚动到末尾
        def on_path_focus_out(event):
            if len(path_var.get()) > 50:
                path_entry.xview_moveto(1.0)  # 滚动到末尾显示文件夹名
        
        path_entry.bind('<FocusOut>', on_path_focus_out)
        path_entry.bind('<Button-1>', lambda e: path_entry.xview_moveto(1.0) if len(path_var.get()) > 50 else None)
        
        # 浏览按钮
        browse_btn = ttk.Button(
            path_frame, 
            text="浏览...", 
            command=lambda v=path_var: self.choose_directory(v),
            style='TButton'
        )
        browse_btn.grid(row=0, column=1, padx=5)
        
        # 删除按钮（第一行不显示）
        if row_idx > 0:
            delete_btn = ttk.Button(
                path_frame,
                text="删除",
                command=lambda idx=row_idx: self.delete_path_row(idx),
                style='Small.TButton'
            )
            delete_btn.grid(row=0, column=2)
    
    def delete_path_row(self, idx):
        """删除指定索引的路径行"""
        if idx < len(self.path_frames) and idx > 0:  # 不允许删除第一行
            # 销毁UI元素
            self.path_frames[idx].destroy()
            # 从列表中移除
            self.path_frames.pop(idx)
            self.path_vars.pop(idx)
            
            # 重新排列剩余行
            for i in range(idx, len(self.path_frames)):
                self.path_frames[i].grid(row=i, column=0)
                # 更新删除按钮的命令
                delete_btn = self.path_frames[i].winfo_children()[2]
                delete_btn.configure(command=lambda idx=i: self.delete_path_row(idx))
    
    def choose_directory(self, path_var):
        """选择目录或文件并设置到指定的StringVar中"""
        # 创建选择框架
        choice_frame = tk.Toplevel(self.root)
        choice_frame.title("选择")
        choice_frame.geometry("300x150")
        choice_frame.resizable(False, False)
        
        # 居中显示
        choice_frame.update_idletasks()
        width = choice_frame.winfo_width()
        height = choice_frame.winfo_height()
        x = (choice_frame.winfo_screenwidth() // 2) - (width // 2)
        y = (choice_frame.winfo_screenheight() // 2) - (height // 2)
        choice_frame.geometry(f"{width}x{height}+{x}+{y}")
        
        tk.Label(choice_frame, text="请选择要添加的类型:", font=("楷体", 14)).pack(pady=10)
        
        def select_directory():
            directory = filedialog.askdirectory()
            if directory:
                path_var.set(directory)
            choice_frame.destroy()
        
        def select_file():
            file_path = filedialog.askopenfilename(filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")])
            if file_path:
                path_var.set(file_path)
            choice_frame.destroy()
        
        button_frame = tk.Frame(choice_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="选择目录", command=select_directory).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="选择文件", command=select_file).pack(side=tk.LEFT, padx=10)
    
    def choose_output_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_path_var.set(directory)
    
    def choose_gitignore_file(self):
        """打开文件对话框选择.gitignore文件"""
        filepath = filedialog.askopenfilename(
            title="选择 .gitignore 文件",
            filetypes=[("Gitignore files", ".gitignore"), ("All files", "*.*")]
        )
        if filepath:
            self.gitignore_path_var.set(filepath)
            
    def choose_root_directory(self):
        """选择主根目录"""
        directory = filedialog.askdirectory()
        if directory:
            self.root_dir_var.set(directory)
    
    def on_mode_selected(self, event=None):
        """当文件类型模式被选择时更新扩展名输入框"""
        self.update_extension_field()
    
    def update_extension_field(self):
        """根据选择的模式更新扩展名输入框"""
        mode = self.file_mode_var.get()
        
        if mode == "其他模式":
            # 启用输入框
            self.extension_entry.config(state='normal')
            # 清空输入框
            self.extension_var.set("")
        else:
            # 显示选定模式的扩展名
            extensions = self.config_manager.get_file_type_modes().get(mode, [])
            self.extension_var.set(','.join(extensions))
            # 对于默认模式，禁用编辑；对于自定义模式，允许编辑
            if mode == "默认模式":
                self.extension_entry.config(state='readonly')
            else:
                self.extension_entry.config(state='normal')
    
    def save_custom_mode(self):
        """保存当前扩展名设置为自定义模式"""
        extensions = [ext.strip() for ext in self.extension_var.get().split(',') if ext.strip()]
        if not extensions:
            messagebox.showerror("错误", "请输入至少一个文件类型")
            return
        
        def update_mode_selector(mode_name):
            """更新模式选择器"""
            mode_names = list(self.config_manager.get_file_type_modes().keys()) + ["其他模式"]
            # 这里需要找到并更新下拉框，简化实现
            self.file_mode_var.set(mode_name)
        
        # 使用模块化的对话框
        save_callback = CustomModeDialog.show_dialog(self.root, self.config_manager, update_mode_selector)
        save_callback.extensions = extensions
    
    def generate_document(self):
        # 收集所有项目路径
        paths = [var.get() for var in self.path_vars if var.get()]
        if not paths:
            messagebox.showerror("错误", "请至少选择一个项目路径")
            return
            
        # 获取文件类型
        mode = self.file_mode_var.get()
        if mode == "其他模式":
            extensions = [ext.strip() for ext in self.extension_var.get().split(',')]
        else:
            extensions = self.config_manager.get_file_type_modes().get(mode, [])
            
        if not extensions:
            messagebox.showerror("错误", "请选择或输入文件类型")
            return
        
        font_name = self.font_var.get()
        if not font_name:
            font_name = "微软雅黑"
            
        app_name = self.app_name_var.get()
        version = self.version_var.get()
        app_name_font = self.app_name_font_var.get() or "微软雅黑"
        version_font = self.version_font_var.get() or "微软雅黑"
        
        # 获取主根目录
        root_dir = self.root_dir_var.get()
        
        # 获取并验证每页行数
        try:
            lines_per_page = int(self.lines_per_page_var.get())
            if lines_per_page <= 0:
                raise ValueError("每页行数必须大于0")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的每页行数")
            return
            
        # 根据每页行数计算适当的字体大小（简单算法，可能需要调整）
        line_height_cm = 26.7 / lines_per_page
        font_size = int(line_height_cm * 28)
        font_size = max(8, min(12, font_size))
        
        try:
            # --- Stage 1: File Discovery & .gitignore Filtering ---
            gitignore_path_str = self.gitignore_path_var.get()
            files_to_process, ignored_files = self.file_processor.collect_files(
                paths, extensions, gitignore_path_str if gitignore_path_str else None
            )

            # --- Stage 1.5: Setup Progress Bar ---
            progress_window = ProgressWindow(self.root, len(files_to_process))

            # --- Stage 2: Collect all lines from all files ---
            def progress_callback(current, filename):
                progress_window.update_progress(current, filename)
            
            all_styled_lines, lines_by_ext = self.file_processor.process_files(
                files_to_process, paths, root_dir, progress_callback
            )

            file_count = len(files_to_process)
            progress_window.destroy()

            # --- Stage 3: 使用所有代码行 ---
            lines_to_print = all_styled_lines

            # --- Stage 4: Generate the document from the prepared line list ---
            doc = self.document_generator.create_document(
                lines_to_print, font_name, font_size, app_name, version,
                app_name_font, version_font, self.generate_toc_var.get(), 
                self.show_line_numbers_var.get()
            )

            # --- Stage 5: Save and open ---
            # 计算真实的总代码行数
            total_code_lines = sum(lines_by_ext.values())
            
            # 准备文件名
            filename = f"{app_name}_{version}_源代码_{total_code_lines}行.docx"
            
            save_path = ""
            output_dir = self.output_path_var.get()
            
            if output_dir:
                # 如果指定了输出目录，直接在该目录下保存文件
                save_path = os.path.join(output_dir, filename)
            else:
                # 否则弹出保存对话框
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word documents", "*.docx")],
                    initialfile=filename
                )
            
            if save_path:
                doc.save(save_path)
                
                total_code_lines = sum(lines_by_ext.values())
                
                # 构建详细的成功信息
                stats_details = "\n\n代码行数统计:\n"
                for ext, count in lines_by_ext.items():
                    stats_details += f"  - {ext} 文件: {count} 行\n"
                stats_details += f"总代码行数: {total_code_lines} 行"

                ignored_details = ""
                if ignored_files:
                    ignored_details = "\n\n根据 .gitignore 忽略了以下文件:\n"
                    for f in ignored_files[:10]: # 最多显示10个
                        ignored_details += f"  - {f}\n"
                    if len(ignored_files) > 10:
                        ignored_details += f"  ...等共 {len(ignored_files)} 个文件\n"

                actual_pages = len(lines_to_print) / lines_per_page
                success_message = f"文档已生成完成！\n共处理 {file_count} 个文件。\n共 {actual_pages:.1f} 页"
                
                success_message += stats_details
                success_message += ignored_details

                # 创建日志文件
                log_filename = f"{app_name}_{version}_导出记录.txt"
                log_path = os.path.join(os.path.dirname(save_path), log_filename)
                
                with open(log_path, 'w', encoding='utf-8') as log_file:
                    log_file.write(f"文档名称: {os.path.basename(save_path)}\n")
                    log_file.write(f"生成时间: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                    log_file.write("处理的路径:\n")
                    
                    # 获取根目录
                    root_dir = self.root_dir_var.get()
                    root_path = pathlib.Path(root_dir) if root_dir else None
                    
                    for path in paths:
                        path_obj = pathlib.Path(path)
                        if root_path and path_obj.is_relative_to(root_path):
                            rel_path = path_obj.relative_to(root_path)
                            log_file.write(f"  - {rel_path} (相对路径)\n")
                        else:
                            log_file.write(f"  - {path}\n")
                    
                    log_file.write(f"\n总文件数: {file_count}\n")
                    log_file.write(f"总代码行数: {total_code_lines}\n")
                    log_file.write(f"总页数: {actual_pages:.1f}\n")
                    
                    # 写入文件类型统计
                    log_file.write("\n文件类型统计:\n")
                    for ext, count in lines_by_ext.items():
                        log_file.write(f"  - {ext}: {count} 行\n")

                messagebox.showinfo("成功", success_message + f"\n\n已在同目录下创建导出记录文件: {log_filename}")
                try:
                    os.startfile(save_path)
                except Exception as e:
                    messagebox.showwarning("提示", f"无法自动打开文件：{e}")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成文档时发生错误：{str(e)}")

def main():
    root = tk.Tk()
    app = CodeToWordApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
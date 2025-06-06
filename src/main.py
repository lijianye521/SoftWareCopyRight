import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font
from docx import Document
from docx.shared import RGBColor, Pt, Inches, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
import os
import pathlib
import chardet

class CodeToWordApp:
    def __init__(self, root):
        self.root = root
        self.root.title("软件著作权源代码文档生成器")
        self.root.geometry("800x850")  # 增大窗口尺寸以容纳新字段
        
        # 设置应用主题和样式
        self.style = ttk.Style()
        self.style.configure('TLabel', font=('微软雅黑', 12))
        self.style.configure('TEntry', font=('微软雅黑', 12))
        self.style.configure('TButton', font=('微软雅黑', 12))
        
        # 设置背景颜色
        self.root.configure(bg="#f0f0f0")
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_font = font.Font(family="微软雅黑", size=18, weight="bold")
        title_label = tk.Label(main_frame, text="软件源代码导出工具", font=title_font, bg="#f0f0f0")
        title_label.grid(row=0, column=0, columnspan=3, pady=20)
        
        # 项目路径选择
        ttk.Label(main_frame, text="项目路径:").grid(row=1, column=0, sticky=tk.W, pady=15)
        self.path_var = tk.StringVar()
        path_entry = ttk.Entry(main_frame, textvariable=self.path_var, width=50, font=('微软雅黑', 12))
        path_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        browse_btn = ttk.Button(main_frame, text="浏览...", command=self.choose_directory, style='TButton')
        browse_btn.grid(row=1, column=2, padx=10, pady=15)
        
        # 文件后缀输入
        ttk.Label(main_frame, text="文件后缀:").grid(row=2, column=0, sticky=tk.W, pady=15)
        self.extension_var = tk.StringVar(value="py")
        extension_entry = ttk.Entry(main_frame, textvariable=self.extension_var, width=50, font=('微软雅黑', 12))
        extension_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 字体选择
        ttk.Label(main_frame, text="字体:").grid(row=3, column=0, sticky=tk.W, pady=15)
        self.font_var = tk.StringVar(value="微软雅黑")
        font_entry = ttk.Entry(main_frame, textvariable=self.font_var, width=50, font=('微软雅黑', 12))
        font_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 应用名称
        ttk.Label(main_frame, text="应用名称:").grid(row=4, column=0, sticky=tk.W, pady=15)
        self.app_name_var = tk.StringVar(value="软件")
        app_name_entry = ttk.Entry(main_frame, textvariable=self.app_name_var, width=50, font=('微软雅黑', 12))
        app_name_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 应用名称字体
        ttk.Label(main_frame, text="应用名称字体:").grid(row=5, column=0, sticky=tk.W, pady=15)
        self.app_name_font_var = tk.StringVar(value="微软雅黑")
        app_name_font_entry = ttk.Entry(main_frame, textvariable=self.app_name_font_var, width=50, font=('微软雅黑', 12))
        app_name_font_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 版本号
        ttk.Label(main_frame, text="版本号:").grid(row=6, column=0, sticky=tk.W, pady=15)
        self.version_var = tk.StringVar(value="V1.0")
        version_entry = ttk.Entry(main_frame, textvariable=self.version_var, width=50, font=('微软雅黑', 12))
        version_entry.grid(row=6, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 版本号字体
        ttk.Label(main_frame, text="版本号字体:").grid(row=7, column=0, sticky=tk.W, pady=15)
        self.version_font_var = tk.StringVar(value="微软雅黑")
        version_font_entry = ttk.Entry(main_frame, textvariable=self.version_font_var, width=50, font=('微软雅黑', 12))
        version_font_entry.grid(row=7, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 每页行数设置
        ttk.Label(main_frame, text="每页行数:").grid(row=8, column=0, sticky=tk.W, pady=15)
        self.lines_per_page_var = tk.StringVar(value="55")
        lines_per_page_entry = ttk.Entry(main_frame, textvariable=self.lines_per_page_var, width=50, font=('微软雅黑', 12))
        lines_per_page_entry.grid(row=8, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 生成模式
        ttk.Label(main_frame, text="生成模式:").grid(row=9, column=0, sticky=tk.W, pady=15)
        self.mode_var = tk.StringVar(value="默认模式")
        mode_selector = ttk.Combobox(main_frame, textvariable=self.mode_var, values=["默认模式", "标准模式"], state="readonly", width=47, font=('微软雅黑', 12))
        mode_selector.grid(row=9, column=1, sticky=(tk.W, tk.E), pady=15, padx=5)
        
        # 模式说明
        standard_mode_info = ttk.Label(
            main_frame, 
            text="标准模式：用于软件著作权申请，将生成源代码文件的前30页和后30页。\n要求总代码量超过60页，否则将按默认模式生成。", 
            font=('微软雅黑', 10),
            wraplength=450, 
            justify=tk.LEFT,
            background="#f0f0f0"
        )
        standard_mode_info.grid(row=10, column=1, sticky=tk.W, pady=5, padx=5)
        
        # 生成按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=11, column=0, columnspan=3, pady=30)
        
        self.style.configure('Generate.TButton', font=('微软雅黑', 14, 'bold'))
        generate_btn = ttk.Button(
            button_frame, 
            text="生成文档", 
            command=self.generate_document,
            style='Generate.TButton',
            padding=(20, 10)
        )
        generate_btn.pack(padx=10)
        
        # 作者信息
        author_label = ttk.Label(main_frame, text="作者: lijianye", font=('微软雅黑', 10))
        author_label.grid(row=12, column=0, columnspan=3, sticky=tk.E, pady=10)
        
        # 状态变量（不显示状态框）
        self.status_var = tk.StringVar()
        
        # 使窗口可调整大小
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
    def choose_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.path_var.set(directory)
    
    def detect_encoding(self, file_path):
        """检测文件编码"""
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read())
        return result['encoding'] or 'utf-8'

    def read_file_lines(self, file_path, encoding):
        """读取文件内容，将软回车转为硬回车，移除空行，并返回行列表"""
        with open(file_path, 'r', encoding=encoding) as f:
            content = f.read()
        # 将软回车（\v 或 ^l）替换为硬回车（\n 或 ^p）
        content = content.replace('\v', '\n')
        lines = content.splitlines()
        # 只保留非空行
        return [line for line in lines if line.strip()]
    
    def apply_style_to_paragraph(self, paragraph, font_name="微软雅黑", font_size=10):
        """应用统一样式到段落"""
        # 设置段落行距为精确值
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        paragraph.paragraph_format.line_spacing = Pt(font_size * 1.05)  # 进一步减小行距，设为字体大小的1.05倍
        # 设置段落间距为0
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            # 显式设置东亚字体
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    
    def add_header(self, doc, app_name, version, app_name_font, version_font):
        """添加页眉"""
        # 获取第一节
        section = doc.sections[0]
        
        # 创建页眉
        header = section.header
        header_para = header.paragraphs[0]
        
        # 添加应用名称
        app_run = header_para.add_run(app_name)
        app_run.font.name = app_name_font
        app_run.font.size = Pt(12)
        app_run.bold = True
        # 显式设置东亚字体
        r_app = app_run._element
        r_app.rPr.rFonts.set(qn('w:eastAsia'), app_name_font)
        
        # 添加空格
        header_para.add_run(" ")
        
        # 添加版本号
        ver_run = header_para.add_run(version)
        ver_run.font.name = version_font
        ver_run.font.size = Pt(12)
        # 显式设置东亚字体
        r_ver = ver_run._element
        r_ver.rPr.rFonts.set(qn('w:eastAsia'), version_font)
        
        # 设置页眉段落格式
        header_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
            
    def generate_document(self):
        path = self.path_var.get()
        if not path:
            messagebox.showerror("错误", "请选择项目路径")
            return
            
        extensions = [ext.strip() for ext in self.extension_var.get().split(',')]
        if not extensions:
            messagebox.showerror("错误", "请输入文件后缀")
            return
        
        font_name = self.font_var.get()
        if not font_name:
            font_name = "微软雅黑"
            
        app_name = self.app_name_var.get()
        version = self.version_var.get()
        app_name_font = self.app_name_font_var.get() or "微软雅黑"
        version_font = self.version_font_var.get() or "微软雅黑"
        
        mode = self.mode_var.get()
        
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
            # --- Stage 1: Collect all lines from all files ---
            all_styled_lines = []
            base_path = pathlib.Path(path)
            lines_by_ext = {}
            
            for extension in extensions:
                extension = extension.strip('.')
                lines_by_ext[f".{extension}"] = 0
                for file_path in base_path.rglob(f"*.{extension}"):
                    relative_path = str(file_path.relative_to(base_path))
                    all_styled_lines.append((relative_path, 'path'))
                    
                    try:
                        encoding = self.detect_encoding(file_path)
                        lines = self.read_file_lines(file_path, encoding)
                        lines_by_ext[f".{extension}"] += len(lines)
                        for line in lines:
                            all_styled_lines.append((line, 'code'))
                    except Exception as e:
                        all_styled_lines.append((f"无法读取文件: {relative_path} ({e})", 'error'))

            file_count = len([1 for text, style in all_styled_lines if style == 'path'])

            # --- Stage 2: Decide which lines to print based on mode ---
            lines_to_print = []
            total_lines = len(all_styled_lines)
            total_pages = total_lines / lines_per_page
            
            final_mode = mode
            if mode == "标准模式":
                if total_pages <= 60:
                    messagebox.showwarning("模式切换", f"总代码约 {total_pages:.1f} 页，不足60页，无法按标准模式生成。\n将自动切换到默认模式。")
                    final_mode = "默认模式"
                else:
                    lines_for_30_pages = int(30 * lines_per_page)
                    first_30 = all_styled_lines[:lines_for_30_pages]
                    last_30 = all_styled_lines[-lines_for_30_pages:]
                    separator_text = f"... (中间部分代码省略) ..."
                    
                    lines_to_print.extend(first_30)
                    lines_to_print.append((separator_text, 'separator'))
                    lines_to_print.extend(last_30)
            
            if final_mode == "默认模式":
                lines_to_print = all_styled_lines

            # --- Stage 3: Generate the document from the prepared line list ---
            doc = Document()
            # 设置文档页面大小和边距
            for section in doc.sections:
                section.page_height = Cm(29.7)  # A4高度
                section.page_width = Cm(21.0)   # A4宽度
                # 设置更小的页边距以容纳更多内容
                section.top_margin = Cm(0.8)    # 进一步减小上边距
                section.bottom_margin = Cm(0.8) # 进一步减小下边距
                section.left_margin = Cm(1.5)   # 减小左边距
                section.right_margin = Cm(1.5)  # 减小右边距
                # 设置页眉和页脚距离
                section.header_distance = Cm(0.3)
                section.footer_distance = Cm(0.3)
            
            # 设置文档默认字体和段落格式
            style = doc.styles['Normal']
            style.font.name = font_name
            style.font.size = Pt(font_size)
            rpr = style.element.rPr
            rpr.rFonts.set(qn('w:eastAsia'), font_name)
            style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            style.paragraph_format.line_spacing = Pt(font_size * 1.05)
            style.paragraph_format.space_before = Pt(0)
            style.paragraph_format.space_after = Pt(0)
            
            self.add_header(doc, app_name, version, app_name_font, version_font)
            
            for text, style_type in lines_to_print:
                if style_type == 'path':
                    path_para = doc.add_paragraph()
                    path_run = path_para.add_run(text)
                    path_run.bold = True
                    path_run.font.name = font_name
                    path_run.font.size = Pt(font_size + 1)
                    r_path = path_run._element
                    r_path.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                    path_para.paragraph_format.space_before = Pt(0)
                    path_para.paragraph_format.space_after = Pt(0)
                    path_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    path_para.paragraph_format.line_spacing = Pt((font_size + 1) * 1.1)
                elif style_type == 'code' or style_type == 'error':
                    content_para = doc.add_paragraph(text)
                    self.apply_style_to_paragraph(content_para, font_name, font_size)
                elif style_type == 'separator':
                    sep_para = doc.add_paragraph()
                    sep_para.add_run(text).italic = True
                    sep_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    self.apply_style_to_paragraph(sep_para, font_name, font_size)

            # --- Stage 4: Save and open ---
            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word documents", "*.docx")],
                initialfile="源代码导出.docx"
            )
            
            if save_path:
                doc.save(save_path)
                
                total_code_lines = sum(lines_by_ext.values())
                
                # 构建详细的成功信息
                stats_details = "\n\n代码行数统计:\n"
                for ext, count in lines_by_ext.items():
                    stats_details += f"  - {ext} 文件: {count} 行\n"
                stats_details += f"总代码行数: {total_code_lines} 行"

                success_message = f"文档已生成完成！\n共处理 {file_count} 个文件。"
                if final_mode == "标准模式":
                    success_message += "\n模式：标准模式 (前30页 + 后30页)"
                else:
                    success_message += f"\n模式：默认模式 (共 {total_pages:.1f} 页)"
                
                success_message += stats_details

                messagebox.showinfo("成功", success_message)
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
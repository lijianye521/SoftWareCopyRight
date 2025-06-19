import os
import re
import jieba
import hashlib
import numpy as np
import difflib
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
from simhash import Simhash
from collections import Counter

class SimilarityAnalyzer:
    def __init__(self):
        """初始化相似度分析器"""
        self.vectorizer = None
    
    def extract_text_from_docx(self, file_path):
        """从Word文档中提取文本内容，按行分割"""
        try:
            doc = Document(file_path)
            lines = []
            
            # 提取正文段落
            for para in doc.paragraphs:
                if para.text and para.text.strip():
                    lines.append(para.text.strip())
            
            # 提取表格内容
            for table in doc.tables:
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        if cell.text and cell.text.strip():
                            for para in cell.paragraphs:
                                if para.text and para.text.strip():
                                    row_text.append(para.text.strip())
                    if row_text:
                        lines.append(" | ".join(row_text))
            
            return lines
        except Exception as e:
            print(f"无法读取文件 {file_path}: {e}")
            return []
    
    def clean_text(self, text):
        """清理文本，保留中文标点和重要符号"""
        # 移除多余空格
        text = re.sub(r'\s+', ' ', text)
        # 移除无意义的特殊字符，但保留中文标点
        text = re.sub(r'[^\w\s\u4e00-\u9fff，。、；：""（）《》？！]', '', text)
        return text.strip()
    
    def clean_lines(self, lines):
        """清理文本行列表"""
        return [self.clean_text(line) for line in lines if line.strip()]
    
    def segment_chinese_text(self, text):
        """使用结巴分词对中文文本进行分词"""
        # 分词
        words = jieba.cut(text)
        return " ".join(words)
    
    def calculate_md5(self, text):
        """计算文本的MD5哈希值，用于精确比较"""
        return hashlib.md5(text.encode('utf-8')).hexdigest()
    
    def calculate_simhash(self, text):
        """计算文本的SimHash值，用于近似比较"""
        return Simhash(text.split()).value
    
    def find_identical_lines(self, lines1, lines2):
        """使用精确匹配找出两个文档中完全相同的行"""
        # 创建行内容到行号的映射
        lines1_dict = {}
        for i, line in enumerate(lines1):
            if line.strip():
                lines1_dict[line] = i
        
        # 找出相同的行
        identical_lines = []
        for i, line in enumerate(lines2):
            if line.strip() and line in lines1_dict:
                identical_lines.append((lines1_dict[line], i, line))
        
        return identical_lines
    
    def analyze_similarity(self, file_paths):
        """分析多个文件之间的相似度，专注于序列相似度和相同行"""
        if len(file_paths) < 2:
            return [], []
        
        # 提取文件名（不带路径和扩展名）
        file_names = [os.path.splitext(os.path.basename(path))[0] for path in file_paths]
        
        # 从文件中提取文本行
        all_lines = []
        for path in file_paths:
            # 提取原始文本行
            lines = self.extract_text_from_docx(path)
            # 清理文本行
            cleaned_lines = self.clean_lines(lines)
            all_lines.append(cleaned_lines)
        
        # 格式化结果
        results = []
        for i in range(len(file_paths)):
            for j in range(i+1, len(file_paths)):
                # 找出完全相同的行
                identical_lines = self.find_identical_lines(all_lines[i], all_lines[j])
                
                # 计算序列相似度
                total_lines = max(len(all_lines[i]), len(all_lines[j]))
                identical_count = len(identical_lines)
                seq_sim = (identical_count / total_lines * 100) if total_lines > 0 else 0
                
                # 使用difflib计算更精确的序列相似度
                diff_ratio = difflib.SequenceMatcher(None, 
                                                    "\n".join(all_lines[i]), 
                                                    "\n".join(all_lines[j])).ratio() * 100
                
                results.append({
                    'file1': file_names[i],
                    'file2': file_names[j],
                    'identical_lines': identical_count,
                    'identical_lines_details': identical_lines[:100],  # 最多显示100个相同行
                    'total_lines1': len(all_lines[i]),
                    'total_lines2': len(all_lines[j]),
                    'max_total_lines': total_lines,
                    'sequence_similarity': seq_sim,
                    'difflib_similarity': diff_ratio
                })
        
        # 按相同行数降序排序
        results.sort(key=lambda x: x['identical_lines'], reverse=True)
        
        return results, None
    
    def get_similarity_report(self, file_paths):
        """生成相似度报告，仅包含序列相似度和相同行信息"""
        if len(file_paths) < 2:
            return "需要至少两个文件才能进行相似度分析。"
        
        # 分析相似度
        results, _ = self.analyze_similarity(file_paths)
        
        # 生成报告
        report = "文件序列相似度分析报告:\n\n"
        
        # 添加详细的相似度结果
        for result in results:
            report += f"{result['file1']} 与 {result['file2']} 的序列相似度: {result['sequence_similarity']:.2f}%\n"
            report += f"  - 完全相同的代码行: {result['identical_lines']} 行\n"
            report += f"  - 文件1总行数: {result['total_lines1']} 行\n"
            report += f"  - 文件2总行数: {result['total_lines2']} 行\n"
            report += f"  - 较大文件总行数: {result['max_total_lines']} 行\n"
            
            # 如果有相同行，显示部分示例
            if result['identical_lines'] > 0:
                report += "  - 相同行示例 (最多显示10行):\n"
                for idx, (line1_idx, line2_idx, line_content) in enumerate(result['identical_lines_details'][:10]):
                    if len(line_content) > 50:
                        line_content = line_content[:50] + "..."
                    report += f"    {idx+1}. 行 {line1_idx+1}↔{line2_idx+1}: {line_content}\n"
                
                if result['identical_lines'] > 10:
                    report += f"    ... 还有 {result['identical_lines'] - 10} 行相同 ...\n"
            
            report += "\n"
        
        # 找出最相似的文件对
        if results:
            most_similar = results[0]
            report += f"\n最相似的文件对: {most_similar['file1']} 和 {most_similar['file2']}\n"
            report += f"完全相同的代码行: {most_similar['identical_lines']} 行 (占较大文件的 {most_similar['sequence_similarity']:.2f}%)"
        
        return report 
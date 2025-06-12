import os
import pathlib
import chardet
import pathspec

class FileProcessor:
    def __init__(self):
        pass
    
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
    
    def collect_files(self, paths, extensions, gitignore_path=None):
        """收集所有匹配的文件"""
        all_files = []
        ignored_files = []
        
        # 获取 .gitignore 规则
        spec = None
        if gitignore_path and os.path.exists(gitignore_path):
            gitignore_path = pathlib.Path(gitignore_path)
            with open(gitignore_path, 'r', encoding='utf-8') as f:
                spec = pathspec.PathSpec.from_lines('gitwildmatch', f)
        
        # 遍历所有项目路径收集文件
        for path in paths:
            base_path = pathlib.Path(path)
            
            # 为每个项目路径单独处理gitignore
            current_spec = spec
            if not current_spec and (base_path / '.gitignore').is_file():
                with open(base_path / '.gitignore', 'r', encoding='utf-8') as f:
                    current_spec = pathspec.PathSpec.from_lines('gitwildmatch', f)
            
            # 收集当前路径下的所有匹配文件
            current_files = []
            for extension in extensions:
                extension = extension.strip('.')
                current_files.extend(base_path.rglob(f"*.{extension}"))
            
            # 应用gitignore过滤
            if current_spec:
                for f in current_files:
                    try:
                        relative_path = f.relative_to(base_path)
                        if not current_spec.match_file(str(relative_path)):
                            all_files.append(f)
                        else:
                            ignored_files.append(str(relative_path))
                    except ValueError:
                        # 如果无法计算相对路径，直接添加
                        all_files.append(f)
            else:
                all_files.extend(current_files)
        
        return all_files, ignored_files
    
    def process_files(self, files_to_process, paths, progress_callback=None):
        """处理文件列表，返回样式化的行列表和统计信息"""
        all_styled_lines = []
        lines_by_ext = {}
        
        for i, file_path in enumerate(files_to_process):
            if progress_callback:
                progress_callback(i + 1, file_path.name)
            
            ext = file_path.suffix
            if ext not in lines_by_ext:
                lines_by_ext[ext] = 0
            
            # 找到文件所属的项目路径
            project_base = None
            for path in paths:
                try:
                    file_path.relative_to(pathlib.Path(path))
                    project_base = pathlib.Path(path)
                    break
                except ValueError:
                    continue
            
            if project_base:
                relative_path = str(file_path.relative_to(project_base))
            else:
                relative_path = str(file_path)
                
            all_styled_lines.append((relative_path, 'path'))
            
            try:
                encoding = self.detect_encoding(file_path)
                lines = self.read_file_lines(file_path, encoding)
                lines_by_ext[ext] += len(lines)
                for line in lines:
                    all_styled_lines.append((line, 'code'))
            except Exception as e:
                all_styled_lines.append((f"无法读取文件: {relative_path} ({e})", 'error'))
        
        return all_styled_lines, lines_by_ext 
import configparser
import os

class ConfigManager:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
        
        # 预定义文件类型模式
        self.file_type_modes = {
            "前端模式": ["json", "js", "jsx", "ts", "tsx", "html", "css", "scss", "vue"],
            "后端模式": ["java", "xml", "py", "php", "cs", "go", "rb"],
            "混合模式": ["json", "js", "jsx", "ts", "tsx", "html", "css", "scss", "vue", "java", "xml", "py", "php"]
        }
    
    def load_config(self):
        """加载配置文件"""
        try:
            if os.path.exists(self.config_file):
                self.config.read(self.config_file, encoding='utf-8')
                
                # 加载自定义模式
                if 'CustomModes' in self.config:
                    for mode_name, extensions in self.config['CustomModes'].items():
                        if mode_name not in self.file_type_modes:
                            self.file_type_modes[mode_name] = extensions.split(',')
        except Exception as e:
            print(f"加载配置文件失败: {e}")
    
    def save_config(self):
        """保存配置到文件"""
        try:
            # 确保存在CustomModes部分
            if 'CustomModes' not in self.config:
                self.config['CustomModes'] = {}
            
            # 保存自定义模式
            custom_modes = {k: ','.join(v) for k, v in self.file_type_modes.items() 
                           if k not in ["前端模式", "后端模式", "混合模式"]}
            
            self.config['CustomModes'] = custom_modes
            
            # 写入文件
            os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def get_file_type_modes(self):
        """获取文件类型模式"""
        return self.file_type_modes
    
    def add_custom_mode(self, mode_name, extensions):
        """添加自定义模式"""
        self.file_type_modes[mode_name] = extensions
        self.save_config() 
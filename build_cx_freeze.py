#!/usr/bin/env python
"""
使用cx_Freeze打包软件著作权源代码生成器
cx_Freeze是一个轻量级的打包工具，通常生成较小的可执行文件
"""
import os
import sys
import subprocess
import shutil
from pathlib import Path

def print_header(text):
    print("\n" + "=" * 60)
    print(text)
    print("=" * 60)

def check_cx_freeze():
    """检查cx_Freeze是否已安装"""
    try:
        import cx_Freeze
        print("✓ cx_Freeze已安装")
        return True
    except ImportError:
        print("✗ cx_Freeze未安装，正在安装...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "cx_Freeze"], check=True)
            print("✓ cx_Freeze安装成功")
            return True
        except Exception as e:
            print(f"✗ cx_Freeze安装失败: {e}")
            return False

def clean_build_files():
    """清理旧的构建文件"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print(f"✓ 已删除 {dir_name}/")
            except Exception as e:
                print(f"✗ 无法删除 {dir_name}: {e}")

def create_setup_file():
    """创建cx_Freeze的setup.py文件"""
    setup_content = """
import sys
from cx_Freeze import setup, Executable

# 依赖项
build_exe_options = {
    "packages": [
        "docx", "chardet", "pathspec", "sv_ttk", 
        "simhash", "jieba", "sklearn.feature_extraction.text", 
        "sklearn.metrics.pairwise"
    ],
    "excludes": [
        "matplotlib", "PyQt5", "PySide2", "wx", "pandas", 
        "scipy", "notebook", "jupyter", "IPython", "tornado", 
        "zmq", "sqlite3", "test", "unittest", "pydoc"
    ],
    "include_files": [
        ("assets", "assets")
    ],
    "optimize": 2,  # 最高级别的优化
    "include_msvcr": True,  # 包含MSVC运行时
}

# 可执行文件设置
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # 不显示控制台窗口

setup(
    name="SoftwareCopyrightGenerator",
    version="1.2.0",
    description="软件著作权源代码文档生成器",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "src/main.py", 
            base=base,
            target_name="SoftwareCopyrightGenerator.exe",
            icon="assets/image-20250612130814003.png"
        )
    ]
)
"""
    
    with open("setup_cx_freeze.py", "w", encoding="utf-8") as f:
        f.write(setup_content)
    
    print("✓ 创建了cx_Freeze配置文件: setup_cx_freeze.py")

def build_with_cx_freeze():
    """使用cx_Freeze构建应用程序"""
    # 确保assets目录存在
    if not os.path.exists('assets'):
        os.makedirs('assets')
        print("✓ 创建了assets目录")
    
    # 创建setup文件
    create_setup_file()
    
    # 运行cx_Freeze
    print("\n正在使用cx_Freeze构建应用程序...")
    print("这可能需要几分钟时间，请耐心等待...")
    
    cmd = [sys.executable, "setup_cx_freeze.py", "build"]
    result = subprocess.run(cmd)
    
    if result.returncode != 0:
        print("\n✗ 构建失败！")
        return False
    
    # 查找构建目录
    build_dirs = [d for d in os.listdir("build") if d.startswith("exe.")]
    if not build_dirs:
        print("\n✗ 构建失败！未找到构建目录。")
        return False
    
    build_dir = os.path.join("build", build_dirs[0])
    
    # 创建dist目录
    if not os.path.exists("dist"):
        os.makedirs("dist")
    
    # 复制可执行文件到dist目录
    exe_path = os.path.join(build_dir, "SoftwareCopyrightGenerator.exe")
    if not os.path.exists(exe_path):
        print("\n✗ 构建失败！未找到可执行文件。")
        return False
    
    dist_exe_path = "dist/SoftwareCopyrightGenerator.exe"
    shutil.copy2(exe_path, dist_exe_path)
    
    # 获取文件大小
    exe_path = Path(dist_exe_path)
    size_mb = exe_path.stat().st_size / (1024 * 1024)
    print(f"\n✓ 构建成功! 文件大小: {size_mb:.2f} MB")
    
    # 尝试使用UPX进一步压缩
    try:
        print("\n尝试使用UPX进一步压缩...")
        subprocess.run(["upx", "--best", "--lzma", str(exe_path)], check=False)
        
        # 获取最终大小
        final_size_mb = exe_path.stat().st_size / (1024 * 1024)
        reduction = (size_mb - final_size_mb)
        reduction_percent = (reduction / size_mb) * 100
        
        print(f"✓ UPX压缩成功")
        print(f"✓ 最终大小: {final_size_mb:.2f} MB")
        print(f"✓ 减少了: {reduction:.2f} MB ({reduction_percent:.1f}%)")
    except FileNotFoundError:
        print("✗ UPX未安装，跳过额外压缩")
    
    return True

def main():
    print_header("软件著作权源代码生成器 - cx_Freeze打包工具")
    
    print("\n[1/3] 检查环境...")
    if not check_cx_freeze():
        print("请先安装cx_Freeze: pip install cx_Freeze")
        return
    
    print("\n[2/3] 清理旧的构建文件...")
    clean_build_files()
    
    print("\n[3/3] 开始构建...")
    if build_with_cx_freeze():
        print_header("构建完成! 可执行文件位于: dist/SoftwareCopyrightGenerator.exe")
    else:
        print_header("构建失败，请检查错误信息")

if __name__ == "__main__":
    main() 
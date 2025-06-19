#!/usr/bin/env python
"""
使用Nuitka打包软件著作权源代码生成器
Nuitka通常比PyInstaller生成更小、更快的可执行文件
"""
import os
import sys
import subprocess
import shutil
import time
import ctypes
from pathlib import Path

def print_header(text):
    print("\n" + "=" * 60)
    print(text)
    print("=" * 60)

def check_nuitka():
    """检查Nuitka是否已安装"""
    try:
        import nuitka
        print("✓ Nuitka已安装")
        return True
    except ImportError:
        print("✗ Nuitka未安装，正在安装...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "nuitka"], check=True)
            print("✓ Nuitka安装成功")
            return True
        except Exception as e:
            print(f"✗ Nuitka安装失败: {e}")
            return False

def clean_build_files():
    """清理旧的构建文件"""
    dirs_to_clean = [
        'build', 
        'dist', 
        '__pycache__', 
        'SoftwareCopyrightGenerator.build',
        'SoftwareCopyrightGenerator.dist',
        'SoftwareCopyrightGenerator.onefile-build'
    ]
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print(f"✓ 已删除 {dir_name}/")
            except Exception as e:
                print(f"✗ 无法删除 {dir_name}: {e}")

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def main():
    # 设置构建选项
    build_options = [
        "--standalone",
        "--assume-yes-for-downloads",
        "--plugin-enable=tk-inter",
        "--show-progress",
        "--output-dir=dist",
        "--windows-disable-console",
        "--windows-icon-from-ico=assets/logo.ico" if os.path.exists("assets/logo.ico") else "",
        "--enable-plugin=numpy",
        "--include-package=simhash",
        "--include-package=jieba",
        "--include-package=docx",
        "--include-package=tkinter",
        "--include-package=tkinterdnd2",
        "--include-package=Pillow",
        "--include-package=PIL",
    ]
    
    # 添加UPX压缩支持（如果安装了UPX）
    upx_path = shutil.which("upx")
    if upx_path:
        build_options.append(f"--upx-binary={upx_path}")
        build_options.append("--compress")
    
    # 检查是否有Windows Defender实时保护
    has_defender = False
    try:
        result = subprocess.run(
            ["powershell", "-Command", "Get-MpPreference -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RealTimeProtectionEnabled"],
            capture_output=True, text=True
        )
        has_defender = "True" in result.stdout
    except Exception as e:
        print(f"无法检查Windows Defender状态: {e}")
    
    # 如果有Windows Defender并且不是管理员权限，给出警告
    if has_defender:
        if is_admin():
            print("====================================================================")
            print("检测到Windows Defender可能会干扰构建过程。")
            print("尝试临时禁用实时保护...")
            try:
                subprocess.run(
                    ["powershell", "-Command", "Set-MpPreference -DisableRealtimeMonitoring $true"],
                    check=True
                )
                print("已临时禁用Windows Defender实时保护")
                defender_disabled = True
            except Exception as e:
                print(f"无法禁用Windows Defender: {e}")
                defender_disabled = False
        else:
            print("====================================================================")
            print("警告: 检测到Windows Defender开启，可能会干扰构建过程")
            print("建议以管理员权限运行，或暂时关闭Windows Defender实时保护")
            print("====================================================================")
            input("按Enter键继续...")
    
    # 构建命令
    build_command = [sys.executable, "-m", "nuitka"]
    build_command.extend(build_options)
    build_command.append("src/main.py")
    
    print("====================================================================")
    print("开始构建...")
    print("构建命令:", " ".join(build_command))
    print("====================================================================")
    
    try:
        # 使用临时环境变量来避免资源嵌入问题
        env = os.environ.copy()
        env["NUITKA_RESOURCE_MODE"] = "copy"  # 使用复制模式而不是嵌入模式
        
        # 运行构建
        result = subprocess.run(build_command, env=env)
        
        if result.returncode == 0:
            print("====================================================================")
            print("构建成功！")
            print("输出位于: dist/main.dist/")
            print("====================================================================")
        else:
            print("====================================================================")
            print("构建失败！")
            print("====================================================================")
    except Exception as e:
        print(f"构建过程出错: {e}")
    finally:
        # 如果之前禁用了Defender，尝试重新启用
        if has_defender and is_admin() and defender_disabled:
            try:
                subprocess.run(
                    ["powershell", "-Command", "Set-MpPreference -DisableRealtimeMonitoring $false"],
                    check=True
                )
                print("已重新启用Windows Defender实时保护")
            except Exception as e:
                print(f"无法重新启用Windows Defender: {e}")

if __name__ == "__main__":
    main() 
#!/usr/bin/env python
"""
Minimal build script for SoftwareCopyrightGenerator
Uses direct PyInstaller commands for maximum size reduction
"""
import os
import sys
import shutil
import subprocess
from pathlib import Path

print("Starting minimal build process...")

# Clean previous builds
if os.path.exists("build"):
    shutil.rmtree("build")
if os.path.exists("dist"):
    shutil.rmtree("dist")

# Set optimization environment variables
os.environ["PYTHONOPTIMIZE"] = "2"
os.environ["PYTHONHASHSEED"] = "1"

# Create a minimal spec content with aggressive optimization
spec_content = """# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_submodules

# 确保distutils能被正确处理
distutils_hidden_imports = collect_submodules('distutils')

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets')],
    hiddenimports=[
        'docx', 'chardet', 'pathspec', 'sv_ttk', 
        'simhash', 'jieba',
        'sklearn.feature_extraction.text._stop_words',
        'distutils.version',  # 添加必要的distutils子模块
    ] + distutils_hidden_imports,  # 添加所有distutils子模块
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'PyQt5', 'PySide2', 'wx',
        'pandas', 'scipy', 'numpy.random', 'numpy.linalg',
        'notebook', 'jupyter', 'IPython', 'tornado', 'zmq',
        'sqlite3', 'email', 'html', 'http', 'pydoc', 'xml',
        'unittest', 'test', 'setuptools',
        # 移除 'distutils'，因为它可能被其他模块需要
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove unnecessary binaries
a.binaries = [x for x in a.binaries if not any(p in x[0] for p in [
    'scipy', 'sklearn/externals', 'tcl', 'tk', 'qt', 'PyQt', 
    'numpy/core/tests', 'numpy/lib/tests', 'numpy/testing',
    'unittest', 'test', 'cryptography', 'PIL',
])]

# Remove unnecessary data files
a.datas = [x for x in a.datas if not any(p in x[0] for p in [
    'test', 'tests', 'docs', 'examples', 'samples', 'demo',
])]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SoftwareCopyrightGenerator_min',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=True,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/image-20250612130814003.png',
)
"""

# Write the minimal spec file
with open("minimal.spec", "w") as f:
    f.write(spec_content)

print("Created minimal spec file...")

# 创建hook文件来处理simhash警告
with open("hook-simhash.py", "w") as f:
    f.write('''
"""
PyInstaller hook for simhash module
This hook suppresses the SyntaxWarning from simhash's invalid escape sequence
"""
from PyInstaller.utils.hooks import collect_all
# Collect all package data
datas, binaries, hiddenimports = collect_all('simhash')
# Fix for the SyntaxWarning in simhash module
import warnings
warnings.filterwarnings("ignore", category=SyntaxWarning, module="simhash")
''')

# 修改spec文件内容，添加钩子目录
spec_content = spec_content.replace(
    "hookspath=[],", 
    "hookspath=['.'],  # 使用当前目录中的hook文件"
)

# 重新写入修改后的spec文件
with open("minimal.spec", "w") as f:
    f.write(spec_content)

# Run PyInstaller with the minimal spec
cmd = [
    sys.executable,
    "-OO",
    "-m",
    "PyInstaller",
    "--clean",
    "--noconfirm",
    "--log-level=WARN",  # 减少日志输出
    "minimal.spec"
]

print("Running PyInstaller with minimal configuration...")
subprocess.run(cmd)

# Apply additional UPX compression if available
exe_path = Path("dist/SoftwareCopyrightGenerator_min.exe")
if exe_path.exists():
    size_mb = exe_path.stat().st_size / (1024 * 1024)
    print(f"Initial build size: {size_mb:.2f} MB")
    
    try:
        print("Applying additional UPX compression...")
        subprocess.run(["upx", "--best", "--lzma", str(exe_path)], 
                       check=False, stdout=subprocess.PIPE)
        
        final_size_mb = exe_path.stat().st_size / (1024 * 1024)
        print(f"Final size: {final_size_mb:.2f} MB")
        print(f"Size reduction: {(size_mb - final_size_mb):.2f} MB ({(1 - final_size_mb/size_mb) * 100:.1f}%)")
    except FileNotFoundError:
        print("UPX not found. Install UPX for additional compression.")
    
    print(f"\nBuild successful! Executable is at: {exe_path}")
else:
    print("Build failed. Check the error messages above.")
    sys.exit(1) 
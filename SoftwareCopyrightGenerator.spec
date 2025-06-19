# -*- mode: python ; coding: utf-8 -*-

import sys
sys.setrecursionlimit(5000)

block_cipher = None

a = Analysis(
    ['src/main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets')],  # Include assets folder
    hiddenimports=[
        'docx', 'chardet', 'pathspec', 'sv_ttk', 
        'simhash', 'jieba',
        # Minimal sklearn imports - only what's absolutely needed
        'sklearn.feature_extraction.text._stop_words',
        'sklearn.feature_extraction.text._document_frequency',
        'sklearn.metrics.pairwise_fast',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Graphics/UI libraries
        'matplotlib', 'PyQt5', 'PySide2', 'PySide6', 'PyQt6', 'wx', 'gi',
        'tkinter', 'Tkinter', '_tkinter', 'Tcl', 'Tk',
        
        # Data science libraries
        'notebook', 'jupyter', 'ipython', 'ipykernel', 'nbconvert', 'nbformat',
        'pandas', 'numba', 'h5py', 'skimage', 'plotly', 'bokeh', 'altair',
        'scipy', 'sympy', 'statsmodels', 'seaborn', 'dask', 'numexpr',
        
        # Web and networking
        'tornado', 'zmq', 'sqlite3', 'http', 'html', 'urllib', 'requests',
        'socketserver', 'pydoc', 'www',
        
        # Development tools
        'pip', 'test', 'distutils', 'setuptools', 'pytz', 'xml', 'curses',
        'unittest', 'pdb', 'doctest', 'pytest', 'nose', 'sphinx',
        
        # Other large libraries
        'cryptography', 'Crypto', 'PIL', 'numpy.random',
        'numpy.linalg', 'numpy.fft', 'numpy.testing',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Remove unnecessary binary files to reduce size
def filter_binaries(binaries):
    # Define patterns to exclude
    exclude_patterns = [
        # Large libraries
        'scipy', 'sklearn/externals', 'tcl', 'tk',
        # Numpy components we don't need
        'numpy/core/tests', 'numpy/lib/tests', 'numpy/linalg',
        'numpy/random', 'numpy/testing', 'numpy/f2py',
        # Testing and documentation
        'unittest', 'test', 'tests', 'testing',
        # Qt libraries
        'qt', 'Qt', 'PyQt5', 'PySide',
        # Other large components
        'lib2to3', 'multiprocessing/tests', 'mpl-data',
        'cryptography', 'PIL', 'Pillow', 'pandas',
        # File formats we don't use
        'sqlite3', 'bz2', 'lzma',
        # Development tools
        'distutils', 'ensurepip', 'pip', 'setuptools',
        # Unnecessary encodings
        'encodings/cp', 'encodings/mac', 'encodings/johab',
        # Locales we don't need
        'locale/de', 'locale/fr', 'locale/ja', 'locale/ko',
    ]
    
    # Keep only binaries that don't match any exclude pattern
    filtered = []
    for binary in binaries:
        if not any(pattern in binary[0] for pattern in exclude_patterns):
            filtered.append(binary)
    
    # Keep only essential encodings
    encodings_to_keep = ['utf_8', 'utf_16', 'ascii', 'cp936', 'gb2312', 'gbk']
    encoding_binaries = [b for b in binaries if 'encodings' in b[0]]
    for binary in encoding_binaries:
        if any(enc in binary[0] for enc in encodings_to_keep):
            if binary not in filtered:
                filtered.append(binary)
    
    print(f"Reduced binaries from {len(binaries)} to {len(filtered)} files")
    return filtered

# Apply the filter
a.binaries = filter_binaries(a.binaries)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Filter datas to remove unnecessary files
a.datas = [x for x in a.datas if not any(pattern in x[0] for pattern in [
    'test', 'tests', 'testing', 'examples', 'docs', 'demo',
    'samples', 'tutorial', 'changelog', 'README', 'LICENSE',
    '.pyc', '.pyo', '.pyd', '.git', '.github', '.gitignore',
    'numpy/doc', 'sklearn/datasets'
])]

# Create a minimal executable
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SoftwareCopyrightGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,           # Strip symbols to reduce size
    upx=True,             # Use UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,  # Store temporary files in memory
    console=False,
    disable_windowed_traceback=True,  # Disable traceback to save space
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/image-20250612130814003.png',  # Use the screenshot as icon
) 

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

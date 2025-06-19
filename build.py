#!/usr/bin/env python
"""
Build script for SoftwareCopyrightGenerator
This script automates the process of creating an optimized executable
"""

import os
import sys
import shutil
import subprocess
import time
from pathlib import Path

def print_status(message):
    """Print a status message with formatting"""
    print(f"\n[+] {message}")

def check_requirements():
    """Check if all required tools are installed"""
    print_status("Checking requirements...")
    
    try:
        import PyInstaller
        print("✓ PyInstaller is installed")
    except ImportError:
        print("✗ PyInstaller is not installed. Installing...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    try:
        # Check if UPX is available
        result = subprocess.run(["upx", "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            print("✓ UPX is installed")
        else:
            print("! UPX is not available. Installing UPX is recommended for smaller executables.")
            print("  Download from: https://github.com/upx/upx/releases")
    except FileNotFoundError:
        print("! UPX is not available. Installing UPX is recommended for smaller executables.")
        print("  Download from: https://github.com/upx/upx/releases")

def clean_build_directories():
    """Clean build and dist directories"""
    print_status("Cleaning build directories...")
    
    for directory in ['build', 'dist', '__pycache__']:
        if os.path.exists(directory):
            shutil.rmtree(directory)
            print(f"✓ Removed {directory}/")
    
    # Clean up Python cache files
    for root, dirs, files in os.walk('.'):
        for file in files:
            if file.endswith('.pyc') or file.endswith('.pyo'):
                os.remove(os.path.join(root, file))
                print(f"✓ Removed {os.path.join(root, file)}")
        
        for dir in dirs:
            if dir == '__pycache__':
                shutil.rmtree(os.path.join(root, dir))
                print(f"✓ Removed {os.path.join(root, dir)}")

def optimize_python_files():
    """Optimize Python files to reduce size"""
    print_status("Optimizing Python files...")
    
    # Compile Python files with optimization
    subprocess.run([sys.executable, "-OO", "-m", "compileall", "src"], check=True)
    print("✓ Compiled Python files with optimization")

def build_application():
    """Build the application using PyInstaller"""
    print_status("Building application...")
    
    # Set environment variables to optimize build
    os.environ['PYTHONOPTIMIZE'] = '2'  # Maximum optimization
    os.environ['PYTHONHASHSEED'] = '1'
    
    # Start timer
    start_time = time.time()
    
    # Run PyInstaller
    cmd = [
        sys.executable, 
        "-OO",  # Use maximum optimization
        "-m", 
        "PyInstaller",
        "--clean",
        "--noconfirm",
        "--log-level=WARN",  # Reduce log output
        "SoftwareCopyrightGenerator.spec"
    ]
    
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        universal_newlines=True
    )
    
    # Show output in real-time with a simple progress indicator
    print("\nBuilding: ", end="", flush=True)
    progress_chars = ["|", "/", "-", "\\"]
    i = 0
    
    while True:
        output = process.stdout.readline()
        if output == '' and process.poll() is not None:
            break
        if output:
            print(f"\rBuilding: {progress_chars[i % len(progress_chars)]}", end="", flush=True)
            i += 1
    
    # Get the return code
    return_code = process.poll()
    
    # Calculate elapsed time
    elapsed_time = time.time() - start_time
    
    print(f"\rBuild completed in {elapsed_time:.2f} seconds.")
    
    if return_code == 0:
        exe_path = Path("dist/SoftwareCopyrightGenerator.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ Successfully built: {exe_path} ({size_mb:.2f} MB)")
            
            # Apply additional UPX compression if available
            try:
                print("\nApplying additional UPX compression...")
                subprocess.run(["upx", "--best", "--lzma", str(exe_path)], 
                               check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                # Show final size after additional compression
                final_size_mb = exe_path.stat().st_size / (1024 * 1024)
                print(f"Final size after additional compression: {final_size_mb:.2f} MB")
                print(f"Size reduction: {(size_mb - final_size_mb):.2f} MB ({(1 - final_size_mb/size_mb) * 100:.1f}%)")
            except FileNotFoundError:
                pass  # UPX not available for additional compression
            
            return True
        else:
            print("\n✗ Build failed: Executable not found")
            return False
    else:
        print("\n✗ Build failed with error code:", return_code)
        return False

def main():
    """Main function"""
    print("\n" + "=" * 60)
    print("SoftwareCopyrightGenerator Build Script")
    print("=" * 60)
    
    check_requirements()
    clean_build_directories()
    optimize_python_files()
    success = build_application()
    
    if success:
        print("\n" + "=" * 60)
        print("Build successful! Executable is in the 'dist' folder.")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("Build failed. Please check the error messages above.")
        print("=" * 60)
        sys.exit(1)

if __name__ == "__main__":
    main() 
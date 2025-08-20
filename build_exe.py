#!/usr/bin/env python3
"""
Build script for creating Windows executable using PyInstaller
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_pyinstaller():
    """Install PyInstaller if not already installed"""
    try:
        import PyInstaller
        print("PyInstaller is already installed")
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Also install pillow for icon conversion
    try:
        import PIL
        print("Pillow is already installed")
    except ImportError:
        print("Installing Pillow for icon processing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pillow"])

def convert_icon():
    """Convert PNG icon to ICO format for Windows"""
    if not os.path.exists("icon.png"):
        print("No icon.png found, skipping icon conversion")
        return None
    
    try:
        from PIL import Image
        print("Converting icon.png to icon.ico...")
        
        # Load the PNG image
        img = Image.open("icon.png")
        
        # Convert to ICO format with multiple sizes
        icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        img.save("icon.ico", format='ICO', sizes=icon_sizes)
        print("Icon converted successfully!")
        return "icon.ico"
        
    except Exception as e:
        print(f"Failed to convert icon: {e}")
        print("Proceeding without icon...")
        return None

def clean_build_dirs():
    """Clean previous build directories"""
    dirs_to_clean = ["build", "dist", "__pycache__"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"Cleaning {dir_name}...")
            shutil.rmtree(dir_name)
    
    # Also clean any .spec files
    for spec_file in Path(".").glob("*.spec"):
        print(f"Removing {spec_file}...")
        spec_file.unlink()

def create_spec_file(icon_path=None):
    """Create PyInstaller spec file with proper configuration"""
    icon_line = f"    icon='{icon_path}'," if icon_path else "    icon=None,  # Add icon path here if you have one"
    
    spec_content = f'''
# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

block_cipher = None

# Get the absolute path to the project directory
project_path = Path(__file__).parent.absolute()

a = Analysis(
    ['main.py'],
    pathex=[str(project_path)],
    binaries=[],
    datas=[
        ('modules', 'modules'),
    ],
    hiddenimports=[
        'PyQt6.QtCore',
        'PyQt6.QtGui', 
        'PyQt6.QtWidgets',
        'matplotlib.backends.backend_qt5agg',
        'matplotlib.backends.backend_qtagg',
        'scipy.sparse.csgraph._validation',
        'scipy.sparse.linalg',
        'scipy.interpolate',
        'scipy.ndimage',
        'scipy.signal',
        'sklearn.utils._cython_blas',
        'sklearn.neighbors.typedefs',
        'sklearn.neighbors.quad_tree',
        'sklearn.tree._utils',
        'pandas._libs.tslibs.base',
        'olefile',
        'numpy',
        'csv',
        'json',
        'struct',
        'datetime',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'unittest',
        'test',
        'tests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='FTIR-Tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True for debugging
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
{icon_line}
)
'''
    
    with open('ftir_tools.spec', 'w') as f:
        f.write(spec_content.strip())
    print("Created ftir_tools.spec file")

def build_executable():
    """Build the executable using PyInstaller"""
    print("Building executable...")
    try:
        # Run PyInstaller with the spec file
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller", 
            "--clean",
            "ftir_tools.spec"
        ])
        print("Build completed successfully!")
        print("Executable location: dist/FTIR-Tools.exe")
        
    except subprocess.CalledProcessError as e:
        print(f"Build failed with error: {e}")
        return False
    return True

def main():
    """Main build process"""
    print("=== FTIR Tools Windows Build Script ===")
    
    # Check if we're in the right directory
    if not os.path.exists("main.py"):
        print("Error: main.py not found. Please run this script from the project root directory.")
        return
    
    print("Step 1: Installing PyInstaller...")
    install_pyinstaller()
    
    print("\nStep 2: Cleaning previous builds...")
    clean_build_dirs()
    
    print("\nStep 3: Converting icon...")
    icon_path = convert_icon()
    
    print("\nStep 4: Creating spec file...")
    create_spec_file(icon_path)
    
    print("\nStep 5: Building executable...")
    if build_executable():
        print("\n=== Build Summary ===")
        print("✅ Windows executable created successfully!")
        print("📁 Location: dist/FTIR-Tools.exe")
        print("📦 This is a standalone executable that includes all dependencies")
        print("\n💡 Tips for distribution:")
        print("- The executable is self-contained and doesn't require Python installation")
        print("- Test on a clean Windows machine to ensure it works properly")
        print("- Consider creating an installer using tools like NSIS or Inno Setup")
    else:
        print("\n❌ Build failed. Check the error messages above.")

if __name__ == "__main__":
    main()
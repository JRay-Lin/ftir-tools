# Building Windows Executable for FTIR Tools

This document explains how to build a standalone Windows executable for the FTIR Tools application.

## Prerequisites

1. **Python 3.12+** installed on your system
2. **All project dependencies** (will be installed automatically)
3. **Windows 10/11** (recommended for building Windows executables)

## Quick Build (Windows)

1. Open Command Prompt or PowerShell in the project directory
2. Run the batch file:
   ```
   build_windows.bat
   ```
3. The executable will be created in the `dist/` folder

## Manual Build Process

### Step 1: Install Dependencies

```bash
# Install build requirements
pip install -r requirements-build.txt

# Or install PyInstaller directly
pip install pyinstaller
```

### Step 2: Run Build Script

```bash
python build_exe.py
```

### Step 3: Find Your Executable

The executable will be located at: `dist/FTIR-Tools.exe`

## Build Configuration

The build process is configured through `ftir_tools.spec` (auto-generated). Key settings:

- **Name**: FTIR-Tools.exe
- **Mode**: Windowed application (no console)
- **Bundling**: All dependencies included in single file
- **Modules**: Includes all required Python packages
- **Data**: Includes the `modules/` directory

## Troubleshooting

### Common Issues

1. **Missing Qt platforms**
   - Solution: Ensure PyQt6 is properly installed
   - Try: `pip install --upgrade PyQt6`

2. **Matplotlib backend issues**
   - The spec file includes Qt backend configuration
   - If issues persist, try: `pip install --upgrade matplotlib`

3. **Large executable size**
   - This is normal for PyQt6 applications (50-150MB)
   - Use UPX compression (enabled by default) to reduce size

4. **Antivirus false positives**
   - Some antivirus software may flag PyInstaller executables
   - Add the `dist/` folder to antivirus exclusions during build

### Debug Mode

To enable console output for debugging:

1. Edit `ftir_tools.spec`
2. Change `console=False` to `console=True`
3. Rebuild: `pyinstaller --clean ftir_tools.spec`

## Distribution

### Single File Distribution

The built executable is self-contained and includes:
- Python runtime
- All required libraries (PyQt6, matplotlib, pandas, etc.)
- Application modules
- No external dependencies required

### Testing

1. Test on the build machine first
2. Test on a clean Windows machine without Python installed
3. Verify all features work (file loading, plotting, baseline creation, CSV export)

### Creating an Installer (Optional)

For professional distribution, consider creating an installer:

1. **NSIS** (Nullsoft Scriptable Install System)
2. **Inno Setup**
3. **Advanced Installer**
4. **WiX Toolset**

## File Structure After Build

```
project/
├── build/              # Temporary build files
├── dist/               # Final executable
│   └── FTIR-Tools.exe  # Your application
├── ftir_tools.spec     # PyInstaller configuration
└── build_exe.py        # Build script
```

## Performance Notes

- First startup may be slower (3-5 seconds) due to extraction
- Subsequent launches are faster
- Memory usage similar to running from Python
- File I/O performance identical to Python version

## Advanced Configuration

### Custom Icon

Add an icon file and modify the spec:
```python
exe = EXE(
    ...
    icon='icon.ico',  # Path to your icon file
    ...
)
```

### Version Information

Add version info to the executable:
```python
exe = EXE(
    ...
    version='version.txt',  # Create version info file
    ...
)
```

### Multiple Files

To create a directory distribution instead of single file:
```python
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='FTIR-Tools'
)
```

## Support

If you encounter issues:
1. Check the console output during build
2. Try building with debug mode enabled
3. Verify all dependencies are correctly installed
4. Test on a virtual machine if possible
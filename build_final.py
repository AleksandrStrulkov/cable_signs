import os
import sys
import subprocess

def build_app():
    base_python = sys.base_prefix
    tcl_root = os.path.join(base_python, 'tcl')
    dlls_path = os.path.join(base_python, 'DLLs')
    
    print("=== Building Application ===")
    print(f"Python: {base_python}")
    print(f"Tcl8.6: {os.path.join(tcl_root, 'tcl8.6')} - exists: {os.path.exists(os.path.join(tcl_root, 'tcl8.6'))}")
    print(f"Tk8.6: {os.path.join(tcl_root, 'tk8.6')} - exists: {os.path.exists(os.path.join(tcl_root, 'tk8.6'))}")
    print(f"DLLs: {dlls_path} - exists: {os.path.exists(dlls_path)}")
    
    cmd = [
        'pyinstaller',
        '--onefile',
        '--noconsole',
        '--add-data', f'{os.path.join(tcl_root, "tcl8.6")};tcl/tcl8.6',
        '--add-data', f'{os.path.join(tcl_root, "tk8.6")};tk/tk8.6',
        '--add-binary', f'{os.path.join(dlls_path, "tk86t.dll")};.',
        '--add-binary', f'{os.path.join(dlls_path, "_tkinter.pyd")};.',
        '--add-binary', f'{os.path.join(dlls_path, "tcl86t.dll")};.',
        'main.py'
    ]
    
    print(f"\nCommand: {' '.join(cmd)}")
    print("\nStarting build...")
    
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        print("\n✅ Build successful!")
        print(f"Executable: dist/main.exe")
    else:
        print("\n❌ Build failed!")
    
    return result.returncode

if __name__ == '__main__':
    build_app()

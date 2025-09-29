
import sys
import os
import subprocess

base_python = sys.base_prefix
tcl_path = os.path.join(base_python, 'tcl', 'tcl8.6')
dlls_path = os.path.join(base_python, 'DLLs')

print(f"Using base Python: {base_python}")
print(f"Tcl path: {tcl_path} - exists: {os.path.exists(tcl_path)}")
print(f"DLLs path: {dlls_path} - exists: {os.path.exists(dlls_path)}")

# Команда сборки с включением всех необходимых файлов
cmd = [
    'pyinstaller',
    '--onefile',
    '--noconsole',
    f'--add-data={tcl_path};tcl/tcl8.6',
    f'--add-data={os.path.join(dlls_path, "tk86t.dll")};.',
    f'--add-binary={os.path.join(dlls_path, "_tkinter.pyd")};.',
    '--hidden-import=tkinter',
    'main.py'
]

print("\nRunning:", ' '.join(cmd))
result = subprocess.run(cmd)

if result.returncode == 0:
    print("Build successful!")
else:
    print("Build failed!")
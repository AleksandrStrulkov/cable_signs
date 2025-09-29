import sys
import os
import subprocess
import tkinter

base_python = sys.base_prefix

# Получим путь к Tk
tk = tkinter.Tk()
tk_library = tk.eval('info library')
tk_root = os.path.dirname(os.path.dirname(tk_library))
tk.destroy()

tcl_path = os.path.join(base_python, 'tcl', 'tcl8.6')
dlls_path = os.path.join(base_python, 'DLLs')

print(f"Tcl path: {tcl_path} - exists: {os.path.exists(tcl_path)}")
print(f"Tk root: {tk_root} - exists: {os.path.exists(tk_root)}")
print(f"DLLs path: {dlls_path} - exists: {os.path.exists(dlls_path)}")

# Команда сборки
cmd = [
    'pyinstaller',
    '--onefile',
    '--noconsole',
    f'--add-data={tcl_path};tcl/tcl8.6',
    f'--add-data={tk_root};tk',  # Добавляем всю папку tk
    f'--add-binary={os.path.join(dlls_path, "tk86t.dll")};.',
    f'--add-binary={os.path.join(dlls_path, "_tkinter.pyd")};.',
    f'--add-binary={os.path.join(dlls_path, "tcl86t.dll")};.',
    '--hidden-import=tkinter',
    'main.py'
]

print("\nRunning:", ' '.join(cmd))
result = subprocess.run(cmd)

if result.returncode == 0:
    print("Build successful!")
else:
    print("Build failed!")

import sys
import os
import subprocess

# Находим пути автоматически
python_path = sys.prefix
tcl_path = os.path.join(python_path, 'tcl', 'tcl8.6')
tk_path = os.path.join(python_path, 'tk', 'tk8.6')

if not os.path.exists(tcl_path):
    print(f"Tcl path not found: {tcl_path}")
    sys.exit(1)
if not os.path.exists(tk_path):
    print(f"Tk path not found: {tk_path}")
    sys.exit(1)

print(f"Tcl path: {tcl_path}")
print(f"Tk path: {tk_path}")

# Команда сборки
cmd = [
    'pyinstaller',
    '--onefile',
    '--noconsole',
    f'--add-data={tcl_path};tcl/tcl8.6',
    f'--add-data={tk_path};tk/tk8.6',
    'main.py'
]

print("Running:", ' '.join(cmd))
subprocess.run(cmd)

import sys
import os

base_python = sys.base_prefix
print(f"Base Python: {base_python}")

# Проверим возможные расположения Tk
possible_tk_paths = [
    os.path.join(base_python, 'tk', 'tk8.6'),
    os.path.join(base_python, 'Lib', 'tkinter'),
    os.path.join(base_python, 'DLLs', 'tk86t.dll'),
    os.path.join(base_python, 'DLLs', '_tkinter.pyd'),
]

for path in possible_tk_paths:
    exists = os.path.exists(path)
    print(f"Path: {path} - exists: {exists}")
    if exists:
        if os.path.isdir(path):
            files = os.listdir(path)
            print(f"  Files: {files[:5]}...")  # Первые 5 файлов
        else:
            print(f"  File size: {os.path.getsize(path)} bytes")

# Также проверим корневую папку tk
tk_root = os.path.join(base_python, 'tk')
if os.path.exists(tk_root):
    print(f"\nContents of {tk_root}:")
    for item in os.listdir(tk_root):
        item_path = os.path.join(tk_root, item)
        print(f"  {item} - dir: {os.path.isdir(item_path)}")
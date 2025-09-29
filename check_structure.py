import os
import sys

# Проверим системный Python
base_python = r"C:\Users\aastrulkov\AppData\Local\Programs\Python\Python312"
print(f"System Python: {base_python}")

# Проверим Tcl/Tk в системном Python
tcl_path = os.path.join(base_python, 'tcl', 'tcl8.6')
tk_path = os.path.join(base_python, 'tcl', 'tk8.6')

print(f"Tcl path: {tcl_path}")
print(f"Tcl exists: {os.path.exists(tcl_path)}")
if os.path.exists(tcl_path):
    print(f"Files in tcl8.6: {len(os.listdir(tcl_path))}")

print(f"Tk path: {tk_path}")
print(f"Tk exists: {os.path.exists(tk_path)}")
if os.path.exists(tk_path):
    files = os.listdir(tk_path)
    print(f"Files in tk8.6: {len(files)}")
    if 'tk.tcl' in files:
        print("✅ tk.tcl FOUND!")
    else:
        print("❌ tk.tcl NOT found")
    print(f"First 10 files: {files[:10]}")
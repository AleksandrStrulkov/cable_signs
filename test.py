#import os
#dlls_path = r'C:\Users\developer\AppData\Local\Programs\Python\Python312\DLLs'
#tcl_dll = os.path.join(dlls_path, 'tcl86t.dll')
#print(f"tcl86t.dll exists: {os.path.exists(tcl_dll)}")

#import sys
#import os
#import tkinter

#base_python = sys.base_prefix
#print(f"Base Python: {base_python}")

# Получим путь к Tk от самого tkinter
#tk = tkinter.Tk()
#tk_library = tk.eval('info library')
#print(f"Tk library path: {tk_library}")

# Найдем корневую папку tk
#tk_root = os.path.dirname(os.path.dirname(tk_library))
#print(f"Tk root directory: {tk_root}")

# Проверим что там внутри
#if os.path.exists(tk_root):
#    print(f"Contents of {tk_root}:")
#    for item in os.listdir(tk_root):
#        item_path = os.path.join(tk_root, item)
#        if os.path.isdir(item_path):
#            print(f"  [DIR] {item}")
#            # Покажем содержимое папки tk8.6 если она есть
#            if item == 'tk8.6':
#                tk_files = os.listdir(item_path)
#                print(f"    Files in tk8.6: {tk_files[:10]}...")
#        else:
#            print(f"  [FILE] {item}")

#tk.destroy()

#import os
#base_python = r'C:\Users\developer\AppData\Local\Programs\Python\Python312'
#tk_path = os.path.join(base_python, 'tk')
#print(f"Tk path exists: {os.path.exists(tk_path)}")
#if os.path.exists(tk_path):
#    print("Tk folder contents:", os.listdir(tk_path))

# import os
# base_python = r'C:\Users\developer\AppData\Local\Programs\Python\Python312'
# tcl_path = os.path.join(base_python, 'tcl')

# print("Contents of tcl folder:")
# for item in os.listdir(tcl_path):
#     item_path = os.path.join(tcl_path, item)
#     print(f"  {item} - dir: {os.path.isdir(item_path)}")
#     if os.path.isdir(item_path):
#         files = os.listdir(item_path)[:3]  # первые 3 файла
#         print(f"    files: {files}...")

test_str1 = "Рассчитываем общий шрифт для ОБЕИХ строк"
test_str2 = "Рассчитываем общий шрифт"
length_before_slash = len(test_str1)
length_after_slash = len(test_str2)
flag_before = False
flag_after = False
if length_before_slash <= 15:
    track_font_size = 14.0
    flag_before = True
elif length_after_slash <= 15:
    track_font_size = 14.0
    flag_after = True
elif length_before_slash >= 19:
    track_font_size = 12.0
    flag_before = True
elif length_after_slash >= 19:
    track_font_size = 12.0
    flag_after = True

print(track_font_size)
print(flag_before)
print(flag_after)
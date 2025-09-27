import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- Попробуем загрузить Times New Roman Bold ---
try:
    pdfmetrics.registerFont(TTFont('Times-Bold', 'timesbd.ttf'))
except:
    pass  # Используем fallback

# --- Параметры ---
TRIANGLE_BASE = 60 * mm
TRIANGLE_HEIGHT = 49 * mm
PAGE_WIDTH, PAGE_HEIGHT = A4

MAX_COLS = 5
MAX_ROWS = 5

FONT_SYSTEM = 18
FONT_TRACK = 14
FONT_CABLE = 14
FONT_LENGTH = 14

MIN_FONT_SIZE = 10

PRINTER_OFFSET_X = 0.5 * mm  # Компенсация смещения принтера (только на обратной стороне)

# --- Функция: разбивка текста по символам ---
def wrap_text_simple(text, max_chars=26):
    """Разбивает текст на строки по количеству символов"""
    if not text or not text.strip():
        return [""]
    words = text.strip().split()
    lines = []
    line = ""
    for word in words:
        sep = " " if line else ""
        test = f"{line}{sep}{word}"
        if len(test) <= max_chars:
            line = test
        else:
            if line:
                lines.append(line)
            # Если слово длиннее max_chars — режем
            while len(word) > max_chars:
                lines.append(word[:max_chars])
                word = word[max_chars:]
            line = word
    if line:
        lines.append(line)
    return lines[:3]  # максимум 3 строки


class CableLabelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор бирок")
        self.root.geometry("500x300")

        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding=20) 
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Генератор обозначений для маркировки кабеля", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="Excel файл:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.input_file, width=40).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Обзор", command=self.browse_input).grid(row=1, column=2, padx=5)

        ttk.Label(frame, text="Папка сохранения:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=self.output_dir, width=40).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Обзор", command=self.browse_output).grid(row=2, column=2, padx=5)

        ttk.Button(frame, text="Создать PDF", command=self.generate).grid(row=3, column=0, columnspan=3, pady=20)

        self.progress = ttk.Progressbar(frame, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")

    def browse_input(self):
        file = filedialog.askopenfilename(title="Выберите Excel", filetypes=[("Excel", "*.xlsx")])
        if file:
            self.input_file.set(file)

    def browse_output(self):
        folder = filedialog.askdirectory(title="Выберите папку")
        if folder:
            self.output_dir.set(folder)

    def generate(self):
        input_path = self.input_file.get()
        output_dir = self.output_dir.get()

        if not input_path or not output_dir:
            messagebox.showerror("Ошибка", "Укажите файл и папку!")
            return

        try:
            wb = openpyxl.load_workbook(input_path)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]

            system_idx = headers.index("system")
            track_idx = headers.index("track")
            cable_idx = headers.index("cable")
            length_idx = headers.index("lenght")
            quantity_idx = headers.index("quantity")

            data = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                system = str(row[system_idx] or "").strip()
                track = str(row[track_idx] or "").strip()
                cable = str(row[cable_idx] or "").strip()
                length_val = str(row[length_idx] or "").strip()
                try:
                    qty = int(row[quantity_idx])
                except:
                    qty = 1
                for _ in range(qty):
                    data.append({
                        "system": system,
                        "track": track,
                        "cable": cable,
                        "length": length_val
                    })

            output_path = os.path.join(output_dir, "cable_labels.pdf")
            c = canvas.Canvas(output_path, pagesize=A4)
            c.setFont("Times-Bold", 12)

            index = 0
            total_sides = len(data) * 2
            self.progress["maximum"] = total_sides
            self.progress["value"] = 0

            while index < len(data):
                # Лицевая сторона
                self.draw_page(c, data, index, side='front')
                c.showPage()
                self.progress["value"] += 1

                # Обратная сторона
                self.draw_page(c, data, index, side='back')
                if index + MAX_COLS * MAX_ROWS < len(data):
                    c.showPage()
                self.progress["value"] += 1

                index += MAX_COLS * MAX_ROWS

            c.save()
            messagebox.showinfo("Готово", f"PDF сохранён:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка: {str(e)}")

    def draw_page(self, c, data, start_index, side):
        col_step = TRIANGLE_BASE / 2
        x_centers = [45*mm, 75*mm, 105*mm, 135*mm, 165*mm]
        Y_START = 76.5 * mm

        shift_x = PRINTER_OFFSET_X if side == 'back' else 0

        count = 0
        for i in range(start_index, min(start_index + MAX_COLS * MAX_ROWS, len(data))):
            item = data[i]
            col = count % MAX_COLS
            row = count // MAX_COLS

            if row >= MAX_ROWS:
                break

            center_x = x_centers[col] + shift_x
            y_base = Y_START + row * TRIANGLE_HEIGHT
            is_upside_down = col % 2 == 1

            if side == 'front':
                main_text = item["system"]
                sub_text = item["track"]
                main_font_size = FONT_SYSTEM
                sub_font_size = FONT_TRACK
                max_lines_sub = 2
            else:
                main_text = item["cable"]
                sub_text = item["length"]
                main_font_size = FONT_CABLE
                sub_font_size = FONT_LENGTH
                max_lines_sub = 3

            self.draw_triangle(c, center_x, y_base, is_upside_down, main_text, sub_text,
                               main_font_size, sub_font_size, max_lines_sub, side)

            count += 1

    def draw_triangle(self, c, center_x, y_base, upside_down, main_text, sub_text,
                      main_font_size, sub_font_size, max_lines_sub, side):
        base = TRIANGLE_BASE
        height = TRIANGLE_HEIGHT
        x_left = center_x - base/2
        x_right = center_x + base/2

        if upside_down:
            points = [(x_left, y_base), (x_right, y_base), (center_x, y_base - height)]
        else:
            points = [(x_left, y_base - height), (x_right, y_base - height), (center_x, y_base)]

        c.setLineWidth(1.8)
        c.setStrokeColorRGB(0, 0, 0)
        c.lines([
            (points[0][0], points[0][1], points[1][0], points[1][1]),
            (points[1][0], points[1][1], points[2][0], points[2][1]),
            (points[2][0], points[2][1], points[0][0], points[0][1])
        ])

        dy_main = height * 0.35
        dy_sub = height * 0.1

        c.saveState()

        if upside_down:
            c.translate(center_x, y_base)
            c.rotate(180)
            c.translate(-center_x, -y_base)
            y_main = y_base + dy_main
            y_sub = y_base + dy_sub
        else:
            base_y = y_base - height
            y_main = base_y + dy_main
            y_sub = base_y + dy_sub

        # --- Основной текст (system/cable) ---
        fs = main_font_size
        max_chars = 26 if side == 'back' else 20
        lines = wrap_text_simple(main_text, max_chars)

        while len(lines) > 3 and fs > MIN_FONT_SIZE:
            fs -= 1
            larger_max = max_chars + int((main_font_size - fs) * 2)
            lines = wrap_text_simple(main_text, larger_max)

        c.setFont("Times-Bold", fs)
        for j, line in enumerate(lines):
            tw = len(line) * fs * 0.6  # грубая ширина
            c.drawString(center_x - tw/2, y_main - j * (fs * 1.4), line)

        # --- Подзаголовок (track/length) ---
        lines_sub = wrap_text_simple(sub_text, 30)[:max_lines_sub]
        temp_fs = sub_font_size
        while len(lines_sub) > max_lines_sub and temp_fs > MIN_FONT_SIZE:
            temp_fs -= 1
            lines_sub = wrap_text_simple(sub_text, 30)[:max_lines_sub]

        c.setFont("Times-Bold", temp_fs)
        for j, line in enumerate(lines_sub):
            tw = len(line) * temp_fs * 0.6
            c.drawString(center_x - tw/2, y_sub - j * (temp_fs * 1.4), line)

        c.restoreState()


if __name__ == "__main__":
    root = tk.Tk()
    app = CableLabelApp(root)
    root.mainloop()
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image, ImageDraw, ImageFont
import textwrap

# Регистрация Times New Roman
try:
    pdfmetrics.registerFont(TTFont('Times-Bold', 'timesbd.ttf'))  # Windows
except:
    try:
        pdfmetrics.registerFont(TTFont('Times-Bold', '/usr/share/fonts/truetype/msttcorefonts/timesbd.ttf'))  # Linux
    except:
        pass  # Используем fallback

# Параметры
TRIANGLE_BASE = 60 * mm
TRIANGLE_HEIGHT = 49 * mm  # Оптимальная высота для 5 рядов на A4 49 мм
PAGE_WIDTH, PAGE_HEIGHT = A4

MAX_COLS = 5   # 5 треугольников в ряду
MAX_ROWS = 5   # 5 рядов

MIN_FONT_SIZE = 8
MAX_FONT_SYSTEM = 18
MAX_FONT_TRACK = 14
MAX_FONT_CABLE = 14



PRINTER_OFFSET_X_MM = 0.5  # Настройте под свой принтер
#PRINTER_OFFSET_Y_MM = 0.2 


# Укажем явный путь
font_path = os.path.join(os.path.dirname(__file__), "timesbd.ttf")

if os.path.exists(font_path):
    pdfmetrics.registerFont(TTFont('Times-Bold', font_path))
    print(f"✅ Шрифт загружен: {font_path}")
else:
    print(f"❌ Файл шрифта не найден: {font_path}")
    # fallback
    try:
        pdfmetrics.registerFont(TTFont('Times-Bold', 'timesbd.ttf'))
    except:
        pass


def generate(self):
    # Проверка шрифта
    from reportlab.pdfbase import _fonts
    print("Зарегистрированные шрифты:", list(_fonts.keys()))
    if 'Times-Bold' in _fonts:
        print("✅ Times-Bold успешно зарегистрирован")
    else:
        print("❌ Times-Bold НЕ зарегистрирован — будет использоваться fallback")
        messagebox.showwarning("Предупреждение", "Шрифт Times-Bold не найден. Текст может отображаться некорректно.")


def force_wrap(text, max_chars=25):
    """Разбивает текст на строки по количеству символов"""
    if not text or not text.strip():
        return [""]
    words = text.strip().split()
    lines = []
    current_line = ""
    
    for word in words:
        spacer = " " if current_line else ""
        test_line = current_line + spacer + word
        if len(test_line) <= max_chars:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            # Если слово слишком длинное — режем
            while len(word) > max_chars:
                lines.append(word[:max_chars])
                word = word[max_chars:]
            current_line = word
    
    if current_line:
        lines.append(current_line)
    
    return lines[:3]  # максимум 3 строки


# Кэшируем шрифт, чтобы не загружать каждый раз
_font_cache = {}

def get_pil_font(font_size):
    global _font_cache
    key = int(font_size)
    if key not in _font_cache:
        try:
            # Попробуем Times New Roman
            font = ImageFont.truetype("timesbd.ttf", key)
        except:
            try:
                font = ImageFont.truetype("Arial", key)
            except:
                font = ImageFont.load_default()
        _font_cache[key] = font
    return _font_cache[key]

""" def wrap_text(text, max_width_mm, font_size, max_lines=3):
    if not text.strip():
        return [""]
    
    # Экспериментально: при font_size=14, 1 мм ≈ 0.9 символов
    max_chars = int(max_width_mm * 0.9)
    if max_chars <= 0:
        max_chars = 1

    words = text.strip().split()
    lines = []
    line = ""

    for word in words:
        space = " " if line else ""
        test = line + space + word
        
        if len(test) <= max_chars:
            line = test
        else:
            if line:
                lines.append(line)
                if len(lines) >= max_lines:
                    break
            # Если слово длиннее max_chars — режем
            while len(word) > max_chars:
                lines.append(word[:max_chars])
                word = word[max_chars:]
                if len(lines) >= max_lines:
                    break
            if len(lines) < max_lines:
                line = word
            else:
                line = ""

    if line and len(lines) < max_lines:
        lines.append(line)

    print(f"✂️ Разбивка: '{text}' -> {lines}")   

    return lines[:max_lines] """


class CableLabelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор обозначений для для маркировки кабеля")
        self.root.geometry("600x400")

        self.input_file = tk.StringVar()
        self.output_dir = tk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="Генератор этикеток", font=("Helvetica", 16, "bold")).grid(
            row=0, column=0, columnspan=3, pady=10)

        ttk.Label(frame, text="Excel файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.input_file, width=40).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Обзор", command=self.browse_input).grid(row=1, column=2, padx=5)

        ttk.Label(frame, text="Папка сохранения:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.output_dir, width=40).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Обзор", command=self.browse_output).grid(row=2, column=2, padx=5)

        ttk.Button(frame, text="Сгенерировать PDF", command=self.generate).grid(
            row=3, column=0, columnspan=3, pady=20)

        self.progress = ttk.Progressbar(frame, orient="horizontal", length=400, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)

    def browse_input(self):
        file = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file:
            self.input_file.set(file)

    def browse_output(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder:
            self.output_dir.set(folder)

    def generate(self):
        input_path = self.input_file.get()
        output_folder = self.output_dir.get()

        if not input_path or not output_folder:
            messagebox.showerror("Ошибка", "Укажите файл и папку сохранения!")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Ошибка", "Файл не найден!")
            return

        try:
            wb = openpyxl.load_workbook(input_path)
            ws = wb.active

            data = []
            headers = [cell.value for cell in ws[1]]
            system_idx = headers.index("system") if "system" in headers else None
            track_idx = headers.index("track") if "track" in headers else None
            cable_idx = headers.index("cable") if "cable" in headers else None
            length_idx = headers.index("lenght") if "lenght" in headers else None
            quantity_idx = headers.index("quantity") if "quantity" in headers else None

            if None in (system_idx, track_idx, cable_idx, length_idx, quantity_idx):
                messagebox.showerror("Ошибка", "Не хватает столбцов: system, track, cable, lenght, quantity")
                return

            for row in ws.iter_rows(min_row=2, values_only=True):
                system = str(row[system_idx] or "").strip()
                track = str(row[track_idx] or "").strip()
                cable = str(row[cable_idx] or "").strip()
                length_val = str(row[length_idx] or "").strip()
                try:
                    quantity = int(row[quantity_idx])
                except:
                    quantity = 1

                for _ in range(quantity):
                    data.append({
                        "system": system,
                        "track": track,
                        "cable": cable,
                        "length": length_val
                    })

            output_pdf = os.path.join(output_folder, "cable_labels_double_sided.pdf")
            self.progress["maximum"] = len(data) * 2
            self.progress["value"] = 0

            c = canvas.Canvas(output_pdf, pagesize=A4)
            c.setFont("Times-Bold", 12)

            index = 0
            while index < len(data):
                # ЛИЦЕВАЯ СТОРОНА
                self.draw_page(c, data, index, side='front')
                c.showPage()
                self.progress["value"] += 1

                # ОБРАТНАЯ СТОРОНА
                self.draw_page(c, data, index, side='back')
                if index + MAX_COLS * MAX_ROWS < len(data):
                    c.showPage()
                self.progress["value"] += 1

                index += MAX_COLS * MAX_ROWS

            c.save()
            messagebox.showinfo("Успех", f"PDF создан:\n{output_pdf}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")

    def draw_page(self, c, data, start_index, side):
        col_step = TRIANGLE_BASE / 2  # 30 мм
        row_height = TRIANGLE_HEIGHT  # 49 мм

        # Центры треугольников по X: отступы по 15 мм
        x_centers = [45 * mm, 75 * mm, 105 * mm, 135 * mm, 165 * mm]

        # Центрирование по вертикали
        Y_START = 76.5 * mm

        # Компенсация смещения принтера: +1 мм только на обратной стороне
        # shift_x = 1 * mm if side == 'back' else 0
        shift_x = PRINTER_OFFSET_X_MM * mm if side == 'back' else 0 # для оси X
        # shift_y = PRINTER_OFFSET_Y_MM * mm if side == 'back' else 0 # для оси Y

        count = 0
        for i in range(start_index, min(start_index + MAX_COLS * MAX_ROWS, len(data))):
            item = data[i]
            col = count % MAX_COLS
            row = count // MAX_COLS

            if row >= 5:
                break

            center_x = x_centers[col] + shift_x  # Сдвигаем только обратную сторону
            y_base = Y_START + row * row_height
            is_upside_down = col % 2 == 1

            if side == 'front':
                main_text = item["system"]
                sub_text = item["track"]
                main_font = MAX_FONT_SYSTEM
                sub_font = MAX_FONT_TRACK
                max_sub_lines = 2
            else:
                main_text = item["cable"]
                sub_text = item["length"]
                main_font = MAX_FONT_CABLE
                sub_font = MAX_FONT_CABLE
                max_sub_lines = 3

            self.draw_triangle_aligned(c, center_x, y_base, is_upside_down, main_text, sub_text,
                                   main_font, sub_font, max_sub_lines, side)
            count += 1

    def draw_triangle_aligned(self, c, center_x, y_base, upside_down, main_text, sub_text,
                          main_font_size, sub_font_size, max_sub_lines, side):
        base = TRIANGLE_BASE
        height = TRIANGLE_HEIGHT

        x_left = center_x - base / 2
        x_right = center_x + base / 2

        if upside_down:
            points = [
                (x_left, y_base),
                (x_right, y_base),
                (center_x, y_base - height)
            ]
        else:
            points = [
                (x_left, y_base - height),
                (x_right, y_base - height),
                (center_x, y_base)
            ]

        # Контур
        c.setLineWidth(1.8)
        c.setStrokeColorRGB(0, 0, 0)
        c.lines([
            (points[0][0], points[0][1], points[1][0], points[1][1]),
            (points[1][0], points[1][1], points[2][0], points[2][1]),
            (points[2][0], points[2][1], points[0][0], points[0][1])
        ])

        # Позиции текста
        dy_system = height * 0.35
        dy_track = height * 0.1

        c.saveState()

        if upside_down:
            c.translate(center_x, y_base)
            c.rotate(180)
            c.translate(-center_x, -y_base)
            y_system = y_base + dy_system
            y_track = y_base + dy_track
        else:
            base_y = y_base - height
            y_system = base_y + dy_system
            y_track = base_y + dy_track

                # --- Основной текст (system или cable) ---
        # Начальный размер шрифта
        fs = main_font_size

        # Максимальная длина строки в символах
        if side == 'back':  # Это cable
            max_chars = 26
        else:  # Это system
            max_chars = 20

        # Разбиваем текст
        lines = force_wrap(main_text, max_chars)

        # Принудительно уменьшаем шрифт, пока строк больше 3
        while len(lines) > 3 and fs > 10:
            fs -= 0.5  # плавное уменьшение
            larger_max_chars = max_chars + int((main_font_size - fs) * 2.5)
            lines = force_wrap(main_text, larger_max_chars)

        # Рисуем все строки
        c.setFont("Times-Bold", fs)
        line_height = fs * 1.4  # расстояние между строками
        for j, line in enumerate(lines):
            # Грубая оценка ширины: символ ≈ 0.6 * размер_шрифта
            estimated_width = len(line) * fs * 0.6
            x_pos = center_x - estimated_width / 2
            y_pos = y_system - j * line_height
            c.drawString(x_pos, y_pos, line)

        # --- ПОДЗАГОЛОВОК ---
        wrapped_sub = force_wrap(sub_text, max_chars=30)[:2]
        temp_fs = sub_font_size

        while len(wrapped_sub) > 2 and temp_fs > 10:
            temp_fs -= 1
            wrapped_sub = force_wrap(sub_text, max_chars=30 + (14 - temp_fs))[:2]

        c.setFont("Times-Bold", temp_fs)
        for j, line in enumerate(wrapped_sub):
            try:
                tw = pdfmetrics.stringWidth(line, "Times-Bold", temp_fs)
            except:
                tw = len(line) * temp_fs * 0.6
            c.drawString(center_x - tw / 2, y_track - j * (temp_fs * 1.5), line)

        c.restoreState()


if __name__ == "__main__":
    root = tk.Tk()
    app = CableLabelApp(root)
    root.mainloop()
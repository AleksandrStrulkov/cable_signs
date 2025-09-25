import os
import math
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import simpleSplit
from reportlab.lib.colors import black
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class TriangleLabelGenerator:
    def __init__(self):
        self.base_width = 60 * mm
        self.height = 52 * mm  # Высота равнобедренного треугольника с основанием 60 мм
        self.page_size = A4
        self.rows_per_page = 4
        self.triangles_per_row = 5
        self.triangles_per_page = self.rows_per_page * self.triangles_per_row
        self.register_fonts()

    def register_fonts(self):
        """Регистрируем шрифты Times New Roman"""
        try:
            # Попробуем разные возможные пути к шрифту
            font_paths = [
                'timesbd.ttf',
                'TIMESBD.TTF',
                '/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman_Bold.ttf',
                '/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf',
                'C:/Windows/Fonts/timesbd.ttf'
            ]
            
            font_registered = False
            for font_path in font_paths:
                try:
                    pdfmetrics.registerFont(TTFont('Times-Bold', font_path))
                    font_registered = True
                    break
                except:
                    continue
            
            if not font_registered:
                print("Предупреждение: Шрифт Times New Roman не найден, используется стандартный шрифт")
                
        except Exception as e:
            print(f"Ошибка при загрузке шрифта: {e}")

    def normalize_column_names(self, df_columns):
        """Нормализуем названия столбцов для обработки различных вариантов"""
        column_mapping = {}
        
        # Приводим все к нижнему регистру и удаляем лишние пробелы
        normalized_columns = [col.strip().lower() for col in df_columns]
        
        # Сопоставляем возможные варианты названий
        possible_names = {
            'system': ['system', 'система', 'sys'],
            'track': ['track', 'трасса', 'trk'],
            'cable': ['cable', 'кабель', 'cab'],
            'length': ['lenght', 'length', 'длина', 'len'],
            'quantity': ['quantity', 'количество', 'qty', 'кол-во']
        }
        
        for standard_name, variants in possible_names.items():
            for variant in variants:
                if variant in normalized_columns:
                    idx = normalized_columns.index(variant)
                    column_mapping[standard_name] = df_columns[idx]
                    break
        
        return column_mapping

    def calculate_positions(self, page_width, page_height):
        """Вычисляем позиции треугольников на странице с учетом трапециевидного расположения"""
        positions = []
        
        # Вычисляем доступную ширину для трапеции (3 треугольника по 60 мм = 180 мм)
        trapezoid_width = 3 * self.base_width  # 180 мм
        start_x = (page_width - trapezoid_width) / 2  # Центрируем трапецию
        
        # Высота ряда - высота треугольника + отступ
        row_height = self.height + 10 * mm
        
        # Увеличиваем верхний отступ, чтобы вершины треугольников не обрезались
        top_margin = 50 * mm  # Увеличили отступ сверху
        
        for row in range(self.rows_per_page):
            # Y-координата основания для ряда (для нормальных треугольников)
            y_base = page_height - top_margin - row * row_height
            
            for col in range(self.triangles_per_row):
                # Определяем ориентацию треугольника
                is_inverted = (col % 2 == 1)  # Второй и четвертый перевернуты
                
                # Позиция треугольника в трапеции
                if col == 0:
                    x = start_x
                elif col == 1:
                    x = start_x + self.base_width / 2
                elif col == 2:
                    x = start_x + self.base_width
                elif col == 3:
                    x = start_x + self.base_width * 1.5
                else:  # col == 4
                    x = start_x + self.base_width * 2
                
                # Для перевернутых треугольников основание находится выше
                if is_inverted:
                    y = y_base - self.height  # Основание перевернутого треугольника на уровне вершин нормальных
                else:
                    y = y_base  # Основание нормального треугольника внизу
                
                positions.append({
                    'x': x,
                    'y': y,
                    'inverted': is_inverted,
                    'row': row,
                    'col': col
                })
                
        return positions

    def draw_triangle(self, c, x, y, inverted=False, row=0, col=0):
        """Рисуем равнобедренный треугольник с правильным примыканием"""
        # Толщина линий примыкания
        thick_line_width = 3
        normal_line_width = 1
        
        if inverted:
            # Перевернутый треугольник - острый угол внизу
            # Основание вверху, вершина внизу
            base_y = y + self.height
            points = [
                (x, base_y),  # Левый угол основания
                (x + self.base_width, base_y),  # Правый угол основания
                (x + self.base_width/2, y)  # Вершина внизу
            ]
        else:
            # Обычный треугольник - острый угол вверху
            # Основание внизу, вершина вверху
            base_y = y
            points = [
                (x, base_y),  # Левый угол основания
                (x + self.base_width, base_y),  # Правый угол основания
                (x + self.base_width/2, y + self.height)  # Вершина вверху
            ]
        
        # Рисуем контур треугольника
        c.setLineWidth(normal_line_width)
        path = c.beginPath()
        path.moveTo(points[0][0], points[0][1])
        path.lineTo(points[1][0], points[1][1])
        path.lineTo(points[2][0], points[2][1])
        path.close()
        c.drawPath(path, stroke=1, fill=0)
        
        # Рисуем жирные линии примыкания
        if col > 0:  # Не первый треугольник в ряду
            # Левая сторона - жирная
            c.setLineWidth(thick_line_width)
            c.line(points[0][0], points[0][1], points[2][0], points[2][1])
        
        if col < self.triangles_per_row - 1:  # Не последний треугольник в ряду
            # Правая сторона - жирная
            c.setLineWidth(thick_line_width)
            c.line(points[1][0], points[1][1], points[2][0], points[2][1])

    def draw_text_in_triangle(self, c, x, y, inverted, system_text, track_text, 
                            cable_text, length_text, is_front_side):
        """Вписываем текст в треугольник"""
        if is_front_side:
            self.draw_front_side_text(c, x, y, inverted, system_text, track_text)
        else:
            self.draw_back_side_text(c, x, y, inverted, cable_text, length_text)

    def draw_front_side_text(self, c, x, y, inverted, system_text, track_text):
        """Текст для лицевой стороны"""
        center_x = x + self.base_width / 2
        
        if inverted:
            # Для перевернутого треугольника
            triangle_bottom = y  # Вершина внизу
            triangle_top = y + self.height  # Основание вверху
            
            # System text - в центре треугольника
            text_y = triangle_bottom + self.height * 0.6
            
            # Track text - внизу треугольника с отступом 5 мм от вершины
            track_y = triangle_bottom + 5 * mm
        else:
            # Для обычного треугольника
            triangle_bottom = y  # Основание внизу
            triangle_top = y + self.height  # Вершина вверху
            
            # System text - в центре треугольника
            text_y = triangle_bottom + self.height * 0.4
            
            # Track text - внизу треугольника с отступом 5 мм от основания
            track_y = triangle_bottom + 5 * mm

        # System text - автоматическая подгонка размера
        self.draw_adaptive_text(c, system_text, center_x, text_y, 
                               self.base_width - 10*mm, self.height * 0.4, 
                               max_font_size=18, min_font_size=8, is_bold=True, inverted=inverted)

        # Track text - фиксированный размер, максимум 2 строки
        self.draw_fixed_text(c, track_text, center_x, track_y, 
                            self.base_width - 10*mm, self.height * 0.2, 
                            font_size=12, max_lines=2, is_bold=True, inverted=inverted)

    def draw_back_side_text(self, c, x, y, inverted, cable_text, length_text):
        """Текст для обратной стороны"""
        center_x = x + self.base_width / 2
        
        if inverted:
            # Для перевернутого треугольника
            triangle_bottom = y  # Вершина внизу
            triangle_top = y + self.height  # Основание вверху
            
            # Cable text - в центре треугольника
            text_y = triangle_bottom + self.height * 0.6
            
            # Length text - внизу треугольника с отступом 5 мм от вершины
            length_y = triangle_bottom + 5 * mm
        else:
            # Для обычного треугольника
            triangle_bottom = y  # Основание внизу
            triangle_top = y + self.height  # Вершина вверху
            
            # Cable text - в центре треугольника
            text_y = triangle_bottom + self.height * 0.4
            
            # Length text - внизу треугольника с отступом 5 мм от основания
            length_y = triangle_bottom + 5 * mm

        # Cable text - максимум 3 строки, центрируем по вертикали
        self.draw_fixed_text(c, cable_text, center_x, text_y, 
                            self.base_width - 10*mm, self.height * 0.5, 
                            font_size=12, max_lines=3, is_bold=True, inverted=inverted)

        # Length text
        self.draw_fixed_text(c, length_text, center_x, length_y, 
                            self.base_width - 10*mm, self.height * 0.2, 
                            font_size=12, max_lines=1, is_bold=True, inverted=inverted)

    def draw_adaptive_text(self, c, text, x, y, max_width, max_height, 
                          max_font_size, min_font_size, is_bold=False, inverted=False):
        """Текст с автоматической подгонкой размера"""
        # Определяем доступный шрифт
        font_name = self.get_available_font(is_bold)
        
        for font_size in range(max_font_size, min_font_size - 1, -1):
            try:
                c.setFont(font_name, font_size)
                lines = simpleSplit(str(text), font_name, font_size, max_width)
                
                total_height = len(lines) * font_size * 1.2
                if total_height <= max_height:
                    self.draw_text_lines(c, lines, x, y, font_size, font_name, inverted)
                    break
            except Exception as e:
                print(f"Ошибка при отрисовке текста: {e}")
                # Используем стандартный шрифт в случае ошибки
                font_name = 'Helvetica-Bold' if is_bold else 'Helvetica'
                c.setFont(font_name, font_size)
                lines = simpleSplit(str(text), font_name, font_size, max_width)
                
                total_height = len(lines) * font_size * 1.2
                if total_height <= max_height:
                    self.draw_text_lines(c, lines, x, y, font_size, font_name, inverted)
                    break

    def draw_fixed_text(self, c, text, x, y, max_width, max_height, 
                       font_size, max_lines, is_bold=False, inverted=False):
        """Текст с фиксированным размером шрифта"""
        font_name = self.get_available_font(is_bold)
        
        try:
            c.setFont(font_name, font_size)
            lines = simpleSplit(str(text), font_name, font_size, max_width)
            lines = lines[:max_lines]
            
            # Проверяем, не превышает ли текст максимальную высоту
            total_height = len(lines) * font_size * 1.2
            if total_height <= max_height:
                self.draw_text_lines(c, lines, x, y, font_size, font_name, inverted)
            else:
                # Если не помещается, уменьшаем количество строк
                lines = lines[:max_lines-1] if max_lines > 1 else lines[:1]
                self.draw_text_lines(c, lines, x, y, font_size, font_name, inverted)
        except Exception as e:
            print(f"Ошибка при отрисовке фиксированного текста: {e}")
            # Используем стандартный шрифт в случае ошибки
            font_name = 'Helvetica-Bold' if is_bold else 'Helvetica'
            c.setFont(font_name, font_size)
            lines = simpleSplit(str(text), font_name, font_size, max_width)
            lines = lines[:max_lines]
            
            total_height = len(lines) * font_size * 1.2
            if total_height <= max_height:
                self.draw_text_lines(c, lines, x, y, font_size, font_name, inverted)
            else:
                lines = lines[:max_lines-1] if max_lines > 1 else lines[:1]
                self.draw_text_lines(c, lines, x, y, font_size, font_name, inverted)

    def get_available_font(self, is_bold=False):
        """Получаем доступный шрифт"""
        if 'Times-Bold' in pdfmetrics.getRegisteredFontNames() and is_bold:
            return 'Times-Bold'
        elif 'Times-Roman' in pdfmetrics.getRegisteredFontNames() and not is_bold:
            return 'Times-Roman'
        else:
            return 'Helvetica-Bold' if is_bold else 'Helvetica'

    def draw_text_lines(self, c, lines, x, y, font_size, font_name, inverted=False):
        """Рисуем строки текста с центрированием"""
        line_height = font_size * 1.2
        total_height = len(lines) * line_height
        
        # Центрируем текст по вертикали
        triangle_center_y = y + self.height / 2
        start_y = triangle_center_y - total_height / 2 + line_height / 2
        
        # Если треугольник перевернутый, переворачиваем и текст
        if inverted:
            c.saveState()
            # Центр вращения - центр треугольника
            c.translate(x + self.base_width/2, y + self.height/2)
            c.rotate(180)
            c.translate(-x - self.base_width/2, -y - self.height/2)
            
            for i, line in enumerate(lines):
                text_width = c.stringWidth(line, font_name, font_size)
                # Центрируем по горизонтали
                text_x = x + self.base_width/2 - text_width / 2
                c.drawString(text_x, start_y - i * line_height, line)
            
            c.restoreState()
        else:
            for i, line in enumerate(lines):
                text_width = c.stringWidth(line, font_name, font_size)
                # Центрируем по горизонтали
                text_x = x + self.base_width/2 - text_width / 2
                c.drawString(text_x, start_y - i * line_height, line)

    def generate_pdf(self, excel_path, output_path):
        """Генерируем PDF файл"""
        try:
            # Читаем Excel файл
            df = pd.read_excel(excel_path)
            
            # Выводим информацию о столбцах для отладки
            print("Найденные столбцы в файле:", list(df.columns))
            
            # Нормализуем названия столбцов
            column_mapping = self.normalize_column_names(df.columns)
            print("Сопоставление столбцов:", column_mapping)
            
            # Проверяем, что все обязательные столбцы найдены
            required_columns = ['system', 'track', 'cable', 'length', 'quantity']
            missing_columns = [col for col in required_columns if col not in column_mapping]
            
            if missing_columns:
                raise ValueError(f"Отсутствуют колонки: {missing_columns}. Найдены: {list(df.columns)}")
            
            # Генерируем данные для треугольников
            triangle_data = []
            for _, row in df.iterrows():
                quantity_val = row[column_mapping['quantity']]
                # Обрабатываем возможные NaN значения
                if pd.isna(quantity_val):
                    quantity_val = 1
                quantity = int(quantity_val)
                
                for _ in range(quantity):
                    triangle_data.append({
                        'system': str(row[column_mapping['system']]),
                        'track': str(row[column_mapping['track']]),
                        'cable': str(row[column_mapping['cable']]),
                        'length': str(row[column_mapping['length']])
                    })
            
            # Разбиваем на страницы
            total_triangles = len(triangle_data)
            total_pages = math.ceil(total_triangles / self.triangles_per_page)
            
            # Создаем PDF
            c = canvas.Canvas(output_path, pagesize=self.page_size)
            page_width, page_height = self.page_size
            
            for page_num in range(total_pages):
                start_idx = page_num * self.triangles_per_page
                end_idx = min(start_idx + self.triangles_per_page, total_triangles)
                page_data = triangle_data[start_idx:end_idx]
                
                # Лицевая сторона
                self.draw_page(c, page_data, page_width, page_height, True)
                c.showPage()
                
                # Обратная сторона
                if page_data:  # Рисуем только если есть данные
                    self.draw_page(c, page_data, page_width, page_height, False)
                    if page_num < total_pages - 1:
                        c.showPage()
            
            c.save()
            return True
            
        except Exception as e:
            raise Exception(f"Ошибка при генерации PDF: {str(e)}")

    def draw_page(self, c, page_data, page_width, page_height, is_front_side):
        """Рисуем страницу с треугольниками"""
        positions = self.calculate_positions(page_width, page_height)
        
        for i, pos in enumerate(positions):
            if i < len(page_data):
                data = page_data[i]
                self.draw_triangle(c, pos['x'], pos['y'], pos['inverted'], pos['row'], pos['col'])
                self.draw_text_in_triangle(
                    c, pos['x'], pos['y'], pos['inverted'],
                    data['system'], data['track'], 
                    data['cable'], data['length'],
                    is_front_side
                )

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор треугольных бирок для кабелей")
        self.geometry("600x300")
        self.generator = TriangleLabelGenerator()
        
        self.create_widgets()

    def create_widgets(self):
        """Создаем элементы интерфейса"""
        # Выбор файла Excel
        tk.Label(self, text="Excel файл:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.excel_path = tk.StringVar()
        tk.Entry(self, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self, text="Обзор", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)
        
        # Выбор папки для сохранения
        tk.Label(self, text="Сохранить в:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        self.output_path = tk.StringVar()
        tk.Entry(self, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self, text="Обзор", command=self.browse_output).grid(row=1, column=2, padx=5, pady=5)
        
        # Информация о требуемых колонках
        info_text = "Требуемые колонки: system, track, cable, lenght (или length), quantity"
        tk.Label(self, text=info_text, fg="blue", wraplength=500).grid(row=2, column=0, columnspan=3, pady=5)
        
        # Информация о формате
        format_text = "Треугольники: 60x52 мм, трапециевидное расположение, 5 треугольников в ряду"
        tk.Label(self, text=format_text, fg="green", wraplength=500).grid(row=3, column=0, columnspan=3, pady=5)
        
        # Информация о размещении
        layout_text = "Размещение: треугольники примыкают друг к другу, жирные линии на стыках"
        tk.Label(self, text=layout_text, fg="purple", wraplength=500).grid(row=4, column=0, columnspan=3, pady=5)
        
        # Кнопка генерации
        tk.Button(self, text="Сгенерировать PDF", command=self.generate, 
                 bg="lightblue", font=("Arial", 12)).grid(row=5, column=1, pady=10)
        
        # Прогресс бар
        self.progress = ttk.Progressbar(self, mode='indeterminate')
        self.progress.grid(row=6, column=0, columnspan=3, sticky="we", padx=10, pady=5)

    def browse_excel(self):
        """Выбор Excel файла"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)

    def browse_output(self):
        """Выбор папки для сохранения"""
        filename = filedialog.asksaveasfilename(
            title="Сохранить PDF как",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)

    def generate(self):
        """Запуск генерации PDF"""
        if not self.excel_path.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл")
            return
            
        if not self.output_path.get():
            messagebox.showerror("Ошибка", "Укажите путь для сохранения")
            return
        
        try:
            self.progress.start()
            self.update_idletasks()
            
            success = self.generator.generate_pdf(
                self.excel_path.get(), 
                self.output_path.get()
            )
            
            self.progress.stop()
            
            if success:
                messagebox.showinfo("Успех", 
                    "PDF успешно сгенерирован!\n\n"
                    "Для двухсторонней печати:\n"
                    "1. Распечатайте все нечетные страницы (лицевая сторона)\n"
                    "2. Переверните бумагу\n"
                    "3. Распечатайте все четные страницы в обратном порядке (обратная сторона)\n"
                    "4. Убедитесь, что треугольники совпадают с обеих сторон")
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    app = Application()
    app.mainloop()
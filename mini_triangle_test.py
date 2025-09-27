from PIL import Image, ImageDraw, ImageFont

# Параметры
WIDTH = 600
HEIGHT = 550
BG_COLOR = "white"
TRIANGLE_COLOR = "black"
TEXT_COLOR = "black"
FONT_SIZE_MAIN = 40
FONT_SIZE_SUB = 30

# Тестовые данные
CABLE_TEXT = "ParLan 4x2x0,57"
LENGTH_TEXT = "L=120 м"

# Создаём основное изображение
img = Image.new("RGB", (WIDTH, HEIGHT), BG_COLOR)
draw = ImageDraw.Draw(img)

# Координаты треугольника (остриём вниз)
left = 50
right = WIDTH - 50
top_y = 50          # основание сверху
bottom_y = HEIGHT - 50
center_x = WIDTH // 2

points = [(left, top_y), (right, top_y), (center_x, bottom_y)]
draw.polygon(points, outline=TRIANGLE_COLOR, width=5)

# Загрузка шрифтов
try:
    font_main = ImageFont.truetype("timesbd.ttf", FONT_SIZE_MAIN)
except:
    font_main = ImageFont.load_default()

try:
    font_sub = ImageFont.truetype("timesbd.ttf", FONT_SIZE_SUB)
except:
    font_sub = ImageFont.load_default()

def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines = []
    line = ""
    for word in words:
        test = f"{line} {word}".strip()
        bbox = draw.textbbox((0, 0), test, font=font)
        if bbox[2] - bbox[0] <= max_width:
            line = test
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines[:3]

# --- Подготовка текста ---
max_width = right - left - 40

main_lines = wrap_text(draw, CABLE_TEXT, font_main, max_width)
sub_lines = wrap_text(draw, LENGTH_TEXT, font_sub, max_width)

# Размеры строк
line_height_main = FONT_SIZE_MAIN + 10
line_height_sub = FONT_SIZE_SUB + 8

# Общая высота блока текста
total_main_height = len(main_lines) * line_height_main
total_sub_height = len(sub_lines) * line_height_sub

# Базовая позиция: центр основания
base_center_x = center_x
base_center_y = top_y

# Создаём временный холст для текста (достаточно большой)
temp_size = (400, 300)
temp_img = Image.new("RGBA", temp_size, (0, 0, 0, 0))
temp_draw = ImageDraw.Draw(temp_img)

# Начальная позиция текста на временном холсте (от верха)
y_main_temp = 50
for i, line in enumerate(main_lines):
    bbox = temp_draw.textbbox((0, 0), line, font=font_main)
    w = bbox[2] - bbox[0]
    x_pos = (temp_size[0] - w) // 2
    temp_draw.text((x_pos, y_main_temp + i * line_height_main), line, font=font_main, fill=(0, 0, 0, 255))

y_sub_temp = y_main_temp + total_main_height + 20
for i, line in enumerate(sub_lines):
    bbox = temp_draw.textbbox((0, 0), line, font=font_sub)
    w = bbox[2] - bbox[0]
    x_pos = (temp_size[0] - w) // 2
    temp_draw.text((x_pos, y_sub_temp + i * line_height_sub), line, font=font_sub, fill=(0, 0, 0, 255))

# Поворачиваем на 180°
rotated_temp = temp_img.rotate(180, expand=False)

# Центр поворота: центр основания
pivot_x = base_center_x - temp_size[0] // 2
pivot_y = base_center_y - temp_size[1] // 2

# Накладываем повёрнутый текст
img.paste(rotated_temp, (pivot_x, pivot_y), rotated_temp)

# Сохраняем
img.save("test_label_fixed.png")
img.show()

print("✅ Готово! Текст развёрнут на 180°, по центру основания, не выходит за границы.")
print("Строки cable:", main_lines)
print("Строки length:", sub_lines)
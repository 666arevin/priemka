import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

# Настройка стилей
thin_side = openpyxl.styles.Side(border_style="thin", color="000000")
my_border = openpyxl.styles.Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
my_font = openpyxl.styles.Font(name="Times New Roman", size=10)
my_alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center", horizontal="center")

pattern: str = "Спортивная подготовка по виду спорта. Тренерско-преподавательская деятельность в образовании"

def insert_data_to_row(workshhet, row_out: list | tuple, row_id, styles: bool = False):
    """Функция вставляет строки в определенную строку

    Args:
        workshhet (WorkSheet): рабочий лист WoorkBook
        row_out (list | tuple): список ячеек, которые нужно вставить 
        row_id (int): id строки, куда надо вставить ячейки
        styles (bool): флаг, указывающий, нужно ли применять стили
    """
    
    row_in = workshhet[row_id]

    for cell_in, cell_out in zip(row_in, row_out):
        cell_in.value = cell_out.value
        if styles:
            cell_in.border = my_border
            cell_in.font = my_font
            cell_in.alignment = my_alignment
            
            # мержим некоторые колонки 
            workshhet.merge_cells(f"O{row_id}:P{row_id}")
            workshhet.merge_cells(f"R{row_id}:T{row_id}")
            workshhet.merge_cells(f"U{row_id}:V{row_id}")

            if cell_out.value in ["Заочная", "Очная"]:
                cell_in.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")
            elif cell_out.value and isinstance(cell_out.value, str) and cell_out.value.startswith(pattern):
                cell_in.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")
            
            # настраиваем высоту
            workshhet.row_dimensions[row_id].height = 160
    
    workshhet[row_id][0].value = int(workshhet[row_id - 1][0].value) + 1


def safe_insert_rows(ws, start_row, amount):
    """
    Безопасно вставляет строки, сдвигая координаты всех объединенных ячеек ниже.
    """
    # 1. Сначала сдвигаем границы всех объединенных диапазонов, которые находятся ниже
    for merged_range in list(ws.merged_cells.ranges):
        # Если объединенная ячейка начинается ниже или на уровне вставки
        if merged_range.min_row >= start_row:
            merged_range.shift(row_shift=amount, col_shift=0)
        # Если вставка происходит ВНУТРИ объединенного диапазона (расширяем его)
        elif merged_range.max_row >= start_row:
            merged_range.max_row += amount
        

    # 2. Теперь выполняем стандартную вставку строк
    ws.insert_rows(idx=start_row + 1, amount=amount)
    
    

def formatted_merged_cells(ws):
    for merged_range in ws.merged_cells.ranges:
        columns_count = merged_range.max_col - merged_range.min_col + 1

        if columns_count >= 18:
            row_idx = merged_range.min_row
            ws.row_dimensions[row_idx].height = 13

def height_formatted(ws: Worksheet):
    for row_idx, row in enumerate(ws.iter_rows(), 1):
        if (row[8].value and row[8].value.startswith(pattern)) or (row[7].value and row[7].value in ["Заочная", "Очная"]) or row[0].value == "№":
            ws.row_dimensions[row_idx].height = 160
            if row[8].value.startswith(pattern):
                ws.merge_cells(f"O{row_idx}:P{row_idx}")
                ws.merge_cells(f"R{row_idx}:T{row_idx}")
                ws.merge_cells(f"U{row_idx}:V{row_idx}")
        elif row[8].value is None and row[0].value is None:
            ws.row_dimensions[row_idx].height = 13
            
            
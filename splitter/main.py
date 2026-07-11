import openpyxl
import re
from tools import *
from pathlib import Path
import json
import os

BASE_DIR = Path(__file__).parent.resolve()
READY_DIR = BASE_DIR / "ready"
SOURCE_DIR = BASE_DIR / "source"


# ищем по регулярке файл
def find_file_by_regex(folder_path, pattern):
    # Компилируем регулярное выражение для скорости
    regex = re.compile(pattern)
    
    # Перебираем все файлы в указанной папке
    for filename in os.listdir(folder_path):
        # Проверяем, подходит ли имя файла под регулярку
        if regex.match(filename):
            # Возвращаем полный путь к файлу
            return os.path.join(folder_path, filename)
    
    raise "Файл не найден"

path_gosuslugi = find_file_by_regex(SOURCE_DIR, r"^Все_заявления_\(бак_спец_БВО\).*\.xlsx$")
path_admission = Path(find_file_by_regex(SOURCE_DIR, r"^Списки поступающих бакалавриат.*\.xlsx$"))

# wd = openpyxl.load_workbook(path_admission)
# ws = wd.active
# print(ws.max_row)


def gosuslugi_splitter():
    print("\nРаботаю с файлом с госуслуг.\n")
    wd = openpyxl.load_workbook(path_gosuslugi)
    ws = wd["Sheet1"]

    data: dict = {}
    count: int = 0
    current_row: int = 1
    # проходимся по всем строкам таблицы
    while current_row <= ws.max_row:
        row = ws[current_row]
        
        # ищем в какой строке в столбце 34 есть запятая
        if re.search(r"\s*,\s*", str(row[34].value)):
            # получаем номер строки
            id = row[34].row + 1
            # узнаем сколько нужно сделать строк
            sports = re.split(pattern=r"\s*,\s*", string=row[34].value)
            sports = list(set(sports))
            print(sports)
            amount_split = len(sports) - 1
            # сохраняем виды, которое нужно добавить, а также индентификатор
            data[row[1].value] = sports
            # первую изменяем сразу
            row[34].value = sports[0]

            if amount_split > 0:
                # вставляем пустую строку двигая другие строки
                ws.insert_rows(idx=id, amount=amount_split)
            # вставляем новые строки
            for sport in sports[1:]:
                insert_data_to_row(ws, row, id)
                ws[id][34].value = sport
                id += 1

        elif row[34].value and row[1].value not in data:
            data[row[1].value] = row[34].value
            
        current_row += 1

    print("\nфайл с госуслуг готов!\n")
    wd.save(BASE_DIR / "temp" / "gosuslugi.xlsx")
    return data

res = gosuslugi_splitter()
with open(BASE_DIR / "json.json", mode="w", encoding="utf-8") as f:
    json.dump(res, f, ensure_ascii=False, indent=4)


def admission_lists_splitter_add(data):
    print("\nРаботаю с файлом вступительных списков\n")

    wd_admission = openpyxl.load_workbook(path_admission)
    ws_adm = wd_admission.active

    pattern: str = "Спортивная подготовка по виду спорта. Тренерско-преподавательская деятельность в образовании"
    current_row: int = 700

    stopper = 0
    while current_row <= ws_adm.max_row:
        row = ws_adm[current_row]

        if row[8].value == pattern:
            sports = data.get(int(row[1].value))

            # производим корректную нумеровку
            row[0].value = 1 if ws_adm[row[0].row - 1][0].value == "№" else int(ws_adm[row[0].row - 1][0].value) + 1
            print(f"row: {current_row}, sports: {sports}")
            if isinstance(sports, str):
                # добавляем вид спорта и фомратируем
                row[8].value += "; " + sports
                ws_adm.row_dimensions[current_row].height = 160

               
            elif isinstance(sports, list):
                # изменяем текущию строку
                row[8].value += "; " + sports[0]

                # получаем количество строк которое нужно вставить
                amount_split = len(sports) - 1

                # важная проверка, если в списке только один обьект
                if amount_split <= 0:
                    continue
                id = current_row + 1

                safe_insert_rows(ws_adm, current_row, amount_split)         
                # заполняем новые строки
                for sport in sports[1:]:
                    # вставляем строки
                    insert_data_to_row(ws_adm, row, id, styles=True)
                    ws_adm[id][8].value = pattern+  "; " + sport
                    id += 1
        
                current_row += amount_split
            else:
                print(f"Тип данных -> {type(sports)}, ключ -> {row[1].value}, тип ключа -> {type(row[1].value)}")
                print(f"ДЛина словаря -> {len(data)}")
            

                
        stopper += 1
        current_row += 1

    # форматируем все обьединенные ячейки
    print("Форматирую высоту merge клеток")
    formatted_merged_cells(ws_adm)
    print("Форматирую высоту клеток и делаю merge")
    height_formatted(ws_adm)
    wd_admission.save(BASE_DIR / "ready" / "Списки_поступающих_обработанные.xlsx")

    print("\nЗакончил работу с файлом вступительных списков\n")


admission_lists_splitter_add(res)







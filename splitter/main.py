import openpyxl
import re
from tools import *
from pathlib import Path
import json

BASE_DIR = Path(__file__).parent.resolve()
READY_DIR = BASE_DIR / "ready"
path_gosuslugi= r"C:\Users\arevi\Yandex.Disk\Computer STO_ORMIK\Documents\Програмирование\python\projects\priemka\splitter\Все_заявления_(бак_спец_БВО)_09-07_11-07-10.xlsx"
path_admission = r"C:\Users\arevi\Yandex.Disk\Computer STO_ORMIK\Documents\Програмирование\python\projects\priemka\splitter\Списки поступающих бакалавриат.xlsx"




def gosuslugi_splitter():
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

    wd.save(BASE_DIR / "temp" / "gosuslugi.xlsx")
    return data

# res = gosuslugi_splitter()
# with open(BASE_DIR / "json.json", mode="w", encoding="utf-8") as f:
#     print(res)
#     json.dump(res, f, ensure_ascii=False, indent=4)

with open(BASE_DIR / "json.json", mode="r", encoding="utf-8") as f:
    data = json.load(f)

def admission_lists_splitter_add():
    wd_admission = openpyxl.load_workbook(path_admission)
    # wd_gosuslugi = openpyxl.load_workbook(BASE_DIR / "temp" / "gosuslugi.xlsx")
    ws_adm = wd_admission.active

    pattern: str = "Спортивная подготовка по виду спорта. Тренерско-преподавательская деятельность в образовании"
    current_row: int = 700

    stopper = 0
    while current_row <= ws_adm.max_row:
        row = ws_adm[current_row]

        if row[8].value == pattern:
            sports = data.get(row[1].value)

            # производим корректную нумеровку
            row[0].value = 1 if ws_adm[row[0].row - 1][0].value == "№" else int(ws_adm[row[0].row - 1][0].value) + 1
            print(f"row: {current_row}, sports: {sports}, max_row: {ws_adm.max_row}")
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

                
        stopper += 1
        current_row += 1

    # форматируем все обьединенные ячейки
    print("Форматирую высоту merge клеток")
    formatted_merged_cells(ws_adm)
    print("Форматирую высоту клеток и делаю merge")
    height_formatted(ws_adm)
    wd_admission.save(BASE_DIR / "ready" / "execel.xlsx")


admission_lists_splitter_add()







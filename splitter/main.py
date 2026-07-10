from openpyxl import *
import re
from tools import *
from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
READY_DIR = BASE_DIR / "ready"
path_from = r"C:\Users\arevi\Yandex.Disk\Computer STO_ORMIK\Documents\Програмирование\python\projects\priemka\splitter\Все_заявления_(бак_спец_БВО)_09-07_11-07-10.xlsx"

wd = load_workbook(path_from)
ws = wd["Sheet1"]

def gosuslugi_splitter():
    
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
            
        current_row += 1
    print(row[34].row)

    # заполняем строки данными
    # for id, value in data.items():
    #     for sport in sports[1:]:
    #         insert_data_to_row(ws, value, id)
    #         ws[id][34].value = sport
    #         id += 1


gosuslugi_splitter()
wd.save(READY_DIR / "test_new.xlsx")
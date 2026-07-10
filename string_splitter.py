import pandas as pd
import re
import os
from tkinter import messagebox


# настройка пандас
pd.set_option('display.max_columns', None)

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

reg_file_name = r"^Все_заявления_\(бак_спец_БВО\).*\.xlsx$"
folder = "."

# путь к файлу
path = find_file_by_regex(folder, reg_file_name)

def split_rows():
    target_column = "Выбранный вид спорта"
    splitter = r"\s*,\s"
    data = pd.read_excel(path, dtype={
        "СНИЛС": str, "Телефон": str, "Номер заявления на ЕПГУ": str,
        "Музыкальный инструмент": str, "Выбранный вид спорта": str,
        "Онлайн-договор на платное обучение": str})

    data[target_column] = data[target_column].str.split(splitter)
    df = data.explode('Выбранный вид спорта').reset_index(drop=True)

    df.to_excel(f"{path[2:-5]}_splitted.xlsx")


split_rows()

messagebox.showinfo("Программа разбиения", f"Все успешно, изменения сохранены в документе {path[2:-5]}_splitted.xlsx")

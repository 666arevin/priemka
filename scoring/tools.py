import json
from pathlib import Path
import os
from openpyxl.cell.cell import Cell

# определяем основные директории
BASE_DIR = Path(__file__).resolve().parent

JSON_PATH = BASE_DIR / "cache" / "json.json"



def render_student_major(sport: str, student_major: str) -> str:
    """Обрабатывает строку с направлением подготовки,
    если указан вид спорта добавляем его, если не указан ничего не деалем.
    Args:
        sport (str): вид спорта из таблицу excel.
        student_major (str): направление подготовки из таблицы excel.
    Returns:
        str: обработанное направление подготовки.
    """
    if sport:
        student_major = student_major + "; " + sport.lower()
    else:
        pass

    return student_major

def insert_data_to_dict(data: dict, student_fio: str, student_major: str, 
                        test_name: str, test_points: int) -> dict:
    """Функция заполняет словарик, проверяя правильность структуры словаря,
    если структура нарушена создаст ее.

    Args:
        data (dict): словарик в который нужно вставить данные.
        student_fio (str): ФИО студента которого нужно добавить в словарик.
        student_major (str): Направление из excel, которое нужно добавить в словарик.
        test_name (str): Название теста по которому нужно проставить баллы.
        test_points (int): Баллы за тест

    Returns:
        dict: Обработанный сдловарик.
    """
    # если нет такого студента создаем для него структуру
    if not data.get(student_fio):
        # заполняем словарик
        data[student_fio] = {}
        data[student_fio][student_major] = {}

    # проверяем, есть ли в словарике нужное направление
    elif not data[student_fio].get(student_major):
        data[student_fio][student_major] = {}

    # проверили что все структуры существуют теперь заполняем словарик
    data[student_fio][student_major][test_name] = test_points

    return data

def get_data_from_dict(data: dict, student_fio: str, student_major: str, test_name: str) -> str | int:
    """Проверяет, что в словарики есть нуобходимый ФИО, направление и название теста.

    Args:
        data (dict): словарик с данными о баллах за тесты.
        student_fio (str): ФИО студента, чьи баллы надо найти.
        student_major (str): Направление по которому нужно получить баллы.
        test_name (str): Название теста за который нужно получить баллы.

    Returns:
        dict: Возвращает баллы за нужный тест, если удалось найти.
        В противном случае вернет пустую строку.
    """
    # получаем конкретного человека из словарика
    input_data = data.get(student_fio)

    # первая проверка, что такой человек есть в словарике
    if not input_data:
        return ""

    input_data = input_data.get(student_major)

    # вторая проверка, если фио есть но направления нету
    if not input_data:
        return ""
    
    input_data = input_data.get(test_name)

    # третья проверка, если фио и направление есть, но баллов за тест нет
    if not input_data:
        return ""
    
    return input_data


def read_headers(row: tuple[Cell, ...], *args) ->  dict:
    """Функция динамически ищет нужные столбцы и возврщает их id.

    Args:
        row (tuple[Cell, ...]): кортеж ячеек строки.

    Returns:
        dict: возвращает словарик найденныз колонок
    """

    columns_id = dict()

    # проходим циклом по ячейкам в строчке
    for id, cell in enumerate(row):

        # находим нужные колонки и возвращаем их column_id
        if isinstance(cell.value, str) and cell.value.lower() in args:
            cell_text_lower = cell.value.lower()
            columns_id[cell_text_lower] = id

    return columns_id

def one_of_list_true(arr: list) -> bool:
    """Если в списке хотябы один элемент True либо содержит значение
    отличное от None или False, функцитя вернет True.

    Args:
        arr (list): Список который нужно проверить.

    Returns:
        bool: 
    """
    for i in arr:
        if i:
            return True
        
    return False
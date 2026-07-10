

def insert_data_to_row(workshhet, row_out: list | tuple, row_id):
    """Функция вставляет строки в определенную строку

    Args:
        workshhet (WorkSheet): рабочий лист WoorkBook
        row_out (list | tuple): список ячеек, которые нужно вставить 
        row_id (int): id строки, куда надо вставить ячейки
    """
    row_in = workshhet[row_id]
    for cell_in, cell_out in zip(row_in, row_out):
        cell_in.value = cell_out.value
import pandas as pd
from pdfminer.high_level import extract_text

def extract_first_page_text(pdf_path):
    text = extract_text(pdf_path, page_numbers=[0])
    return text

def get_departure_and_destination(text):
    parts_of_texts = text.split(" — ")
    departure = parts_of_texts[0]
    destination = parts_of_texts[1].split()[0]
    return departure, destination

new_cols = ["WAYPOINT", "AIRWAY", "HDG", "CRS", "ALT", "CMP", "DIR/SPD", "ISA", "TAS", "GS", "LEG", "REM", "USED", "REM", "ACT", "LEG", "REM", "ETE"]
def check_row(row, i, cols_coords):
    for j in range(len(row)):
        for tcol in new_cols:
            if type(row[j]) is str:
                col_words = row[j].strip().split()
                for col_word in col_words:
                    if col_word.strip() == tcol:
                        cols_coords[tcol] = (i, j)

def prepare_df(df):
    df = df.applymap(lambda x: x.replace('_x000D_', '') if isinstance(x, str) else x)
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    df = df.fillna('-')
    df.columns = df.columns.str.replace('_x000D_', '').str.strip()
    return df


def convert_to_main_table(table_path, departure, destination):
    df = pd.read_excel(table_path)
    df = prepare_df(df)
    
    table_start_row_index = 0
    table_end_row_index = len(df)

    for i, value in enumerate(df.iloc[:, 0]):
        if value == departure:
            table_start_row_index = i
        if value == destination:
            table_end_row_index = i
            
    if len(df.columns) >= 17:
        df = df.iloc[table_start_row_index:len(df), 0:18]
        df = df.iloc[:table_end_row_index, :]
        df.columns = new_cols
        cols_indices = {"WAYPOINT": 0, "ALT": 4, "HDG": 2, "CRS": 3, "DIST": 10, "TIME": 15, "EFOB": 13}
        shorten_df = df.iloc[:, [0, 4, 2, 3, 10, 15, 13]]
        shorten_df.columns = cols_indices.keys()
        return shorten_df
    else:
        return None


def convert_to_sub_table(sub_table_path):
    df = pd.read_excel(sub_table_path)
    df = prepare_df(df)

    if len(df.columns) == 10:
        new_cols = list(df.columns)
        new_cols[-1] = "LONGEST RWY LENGHT"
        new_cols[-2] = "LONGEST RWY ANGLE"
        new_cols[0] = "TYPE"

        df.columns = new_cols
        df['LONGEST RWY ANGLE'] = df['LONGEST RWY ANGLE'].astype(str).str.replace(r'[^0-9/rl]', '', regex=True)
        df = df.set_index(df.columns[0])
        return df
    else:
        return None
    




import openpyxl
from copy import copy

import pandas as pd
from tabulate import tabulate

def show_excel_table(file_path, sheet_name=0):
    """Красиво выводит Excel-таблицу в консоль"""
    # header=None — потому что у тебя первые строки — это заголовки/шапка, а не обычные колонки
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    print(f"\n=== Таблица из файла: {file_path} ===\n")
    print(tabulate(
        df, 
        headers='keys',        # нумерует столбцы как 0, 1, 2...
        tablefmt="grid",       # варианты: "grid", "pretty", "fancy_grid", "psql", "github"
        showindex=False,
        numalign="center",
        stralign="left"
    ))

def modify_excel(template_path, N, output_path):
    """
    Модифицирует Excel-шаблон: повторяет промежуточные строки (4 и 5) N раз,
    сохраняя стили, объединенные ячейки и форматирование.
    
    :param template_path: Путь к шаблону Excel.
    :param N: Общее количество повторений промежуточных строк (минимум 1).
    :param output_path: Путь для сохранения модифицированного файла.
    """
    if N < 1:
        raise ValueError("N должно быть не меньше 1")
    
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active  # Или wb['Sheet1'] если нужно указать имя листа
    
    # Промежуточные строки (исходные)
    intermediate_rows = [4, 5]
    
    # Текущее количество наборов промежуточных строк (1 в шаблоне)
    current_count = 1
    inserts_needed = N - current_count
    
    if inserts_needed <= 0:
        wb.save(output_path)
        return
    
    # Позиция для вставки новых наборов: сразу после существующих промежуточных (строка 6)
    insert_position = 6
    
    for _ in range(inserts_needed):
        # Вставляем 2 новые строки в позицию insert_position
        ws.insert_rows(insert_position, amount=2)
        
        # Копируем значения, стили и форматирование из исходных строк 4 и 5 в новые
        for offset in range(2):
            src_row = intermediate_rows[offset]
            dest_row = insert_position + offset
            
            # Копируем ячейки по столбцам
            for col in range(1, ws.max_column + 1):
                src_cell = ws.cell(row=src_row, column=col)
                dest_cell = ws.cell(row=dest_row, column=col)
                
                dest_cell.value = src_cell.value
                
                # Копируем стиль, если он есть
                if src_cell.has_style:
                    dest_cell.font = copy(src_cell.font)
                    dest_cell.border = copy(src_cell.border)
                    dest_cell.fill = copy(src_cell.fill)
                    dest_cell.number_format = copy(src_cell.number_format)
                    dest_cell.protection = copy(src_cell.protection)
                    dest_cell.alignment = copy(src_cell.alignment)
            
            # Копируем высоту строки
            if ws.row_dimensions[src_row].height is not None:
                ws.row_dimensions[dest_row].height = ws.row_dimensions[src_row].height
        
        # Копируем объединенные ячейки (merged cells) для новых строк
        for merge in list(ws.merged_cells.ranges):
            # Если объединение в исходных промежуточных строках
            if merge.min_row in intermediate_rows:
                # Рассчитываем новые координаты для вставленной копии
                row_offset = insert_position - 4  # Смещение относительно исходной строки 4
                new_min_row = merge.min_row + row_offset
                new_max_row = merge.max_row + row_offset
                new_merge = f"{openpyxl.utils.get_column_letter(merge.min_col)}{new_min_row}:{openpyxl.utils.get_column_letter(merge.max_col)}{new_max_row}"
                ws.merge_cells(new_merge)
        
        # Сдвигаем позицию вставки вниз на 2 строки для следующего набора
        insert_position += 2
    wb.save(output_path)
    return wb



def insert_values_into_template(table, ws):
    col_to_letter = {"ALT": "C", "HDG": "D", "CRS": "E", "DIST": "F", "TIME": "G", "EFOB": "H"}

    for i in range(1, len(table) * 2 + 1):
        if i % 2 == 0:
            ws[f"A{i}"].value = i / 2
            ws[f"B{i}"].value = list(table["WAYPOINT"])[int(i/2 - 1)]

    for col, letter in col_to_letter.items():
        for i in range(3, (len(table)) * 2 + 2):  
            if i % 2 != 0:
                ws[f"{letter}{i}"].value = list(table[col])[int((i - 3) // 2)]



import fitz

def get_names(pdf_path):
    doc = fitz.open(pdf_path)
    
    # Берем только последнюю страницу
    last_page = doc[-1] 
    
    # Получаем объекты текста (словарь с координатами)
    blocks = last_page.get_text("dict")["blocks"]
    
    # Сортируем блоки по высоте (y0), чтобы текст шел сверху вниз
    blocks.sort(key=lambda b: b["bbox"][1])
    
    texts = []
    heights = []
    for b in blocks:
        if "lines" in b:  # Проверяем, что это текстовый блок
            for line in b["lines"]:
                for span in line["spans"]:
                    y_pos = span["bbox"][1]
                    text = span["text"]

                    texts.append(text)
                    heights.append(y_pos)

    filtered_texts = []
    filtered_heights = []

    for text, height in zip(texts, heights):
        if text != "Diagram Unavailable":
            filtered_texts.append(text)
            filtered_heights.append(height)

    texts = filtered_texts
    heights = filtered_heights

    texts = texts[-4:]
    heights = heights[-4:]

    maybe_departure = texts[0].strip()
    maybe_departure_name = texts[1].strip()
    maybe_destination = texts[2].strip()
    maybe_destination_name = texts[3].strip()

    return maybe_departure, maybe_departure_name, maybe_destination, maybe_destination_name
                




def fill_template(template, info):
    for key, value in info.items():
        if value is None:
            value = "_____"
        template = template.replace(key, value)
    return template
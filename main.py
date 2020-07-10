import tkinter as tk
import tkinter.ttk as ttk

import tkinter.filedialog
import tkinter.messagebox


import pandas as pd
from lib import read_docx_tables
import logging


tk.Tk().withdraw()

# Запрос у пользователя файла для анализа
filename = tk.filedialog.askopenfilename(
    title='Выберите файл...',
    filetypes=(('Microsoft Word 2007+','*.docx'),)
)

# Если пользователь не выбрал документ, завершить работу программы
if filename == '':
    tk.messagebox.showerror(
        title='Ошибка',
        message='Файл в формате .docx не был выбран'
    )
    raise Exception('Файл в формате .docx не был выбран.')

logging.info(f' Начато получение табличных данных из файла {filename}')
#TODO Обработка исключительных ситуаций
tax_tables = read_docx_tables(filename)
logging.info(f'Данные успешно импортрированы из файла {filename}')

logging.info(f'Начата обработка данных')
# Обработка данных
n_tax_tables = []
for tax_table in tax_tables:
    n_tax_table = pd.DataFrame(columns=tax_table.columns)
    for index, row in tax_table.iterrows():
        split_values_1 = []

        for column_name in tax_table.columns:
            if column_name == 'Коды региона' or column_name == 'Код региона':
                if ":" in row[column_name]:
                    tmp_split_1 = row[column_name].split(":")
                    tmp_split_2 = tmp_split_1[1].split(",")
                    for tmp_value in tmp_split_2:
                        split_values_1.append(str(tmp_split_1[0]).strip() + str(tmp_value).strip())
                    break
                if "," in row[column_name]:
                    split_values_1 = row[column_name].split(",")
                    break

        split_values_2 = []
        if split_values_1:
            for split_value in split_values_1:
                if '-' in split_value:
                    temp_list_range = split_value.split('-')
                    split_values_2.extend(range(int(temp_list_range[0]), int(temp_list_range[1]), 1))
                else:
                    split_values_2.append(split_value)

            for split_value in split_values_2:
                n_row = []
                for column_name in tax_table.columns:
                    if column_name == 'Коды региона' or column_name == 'Код региона':
                        n_row.append(split_value)
                    else:
                        n_row.append(row[column_name])
                n_tax_table = n_tax_table.append(pd.Series(n_row, index=tax_table.columns), ignore_index=True)
        else:
            n_tax_table = n_tax_table.append(pd.Series(row, index=tax_table.columns), ignore_index=True)

    n_tax_tables.append(n_tax_table)
logging.info(f'Обработка данных закончена')


# Запись данных в файл
logging.info(f'Укажите путь для сохранения файла')
out_filename = tk.filedialog.asksaveasfilename(
    filetypes=(('Microsoft Excel 2007+','*.xlsx'),),
    defaultextension='.xlsx'    
)

if out_filename == '':
    raise Exception('Пользователь не выбрал путь сохранения файла')

writer = pd.ExcelWriter(out_filename, engine='xlsxwriter')
index = 1
for df in n_tax_tables:
    df.to_excel(writer, sheet_name=f'Лист {index}')
    index += 1

writer.save()
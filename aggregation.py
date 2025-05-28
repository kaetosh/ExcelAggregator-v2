# -*- coding: utf-8 -*-
"""
Created on Mon May 26 10:02:36 2025

@author: a.karabedyan
"""

import os
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
from typing import List
from textual.widgets import ProgressBar

from data_text import NAME_OUTPUT_FILE


class NoExcelFilesError(Exception):
    """Custom exception for no Excel files found."""
    pass

def get_excel_files(folder_path: Path) -> List[Path]:
    # Определяем расширения файлов Excel
    excel_extensions = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.odf')

    # Получаем список файлов в указанной папке
    files = [f for f in folder_path.iterdir() if f.suffix in excel_extensions and not f.name.startswith('~') and f.name != 'consolidated.xlsx']

    # Проверяем, есть ли файлы Excel
    if not files:
        raise NoExcelFilesError("В указанной папке нет файлов Excel.")
    return files

def select_folder(current_path) -> Path:
    # Создаем скрытое окно
    root = tk.Tk()
    root.withdraw()  # Скрываем главное окно

    # Открываем диалог выбора папки
    folder_path = filedialog.askdirectory(title="Выберите папку")

    return Path(folder_path) if folder_path else current_path

def get_sheet_names(file_path):
    # Используем pandas для получения названий листов
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

def get_unique_sheet_names(file_paths: List[Path], one_prbar: ProgressBar) -> List[str]:
    unique_sheets = set()  # Используем множество для уникальности
    for index, file_path in enumerate(file_paths):
        if not file_path.exists() or not file_path.suffix in ['.xlsx']:
            continue  # Пропускаем несуществующие файлы или файлы не формата .xlsx

        workbook_sheetnames = get_sheet_names(file_path) # KeyError: "There is no item named 'xl/sharedStrings.xml' in the archive" возможная ошибка для выгрузок из 1С
        unique_sheets.update(workbook_sheetnames)  # Обновляем множество имен листов
        percentage = ((index+1) / len(file_paths)) * 100
        one_prbar.update(progress=percentage)
    one_prbar.update(progress=100)
    list_unique_sheets = list(unique_sheets)
    list_unique_sheets.sort(key=str.lower)
    return list_unique_sheets  # Преобразуем множество обратно в список



def aggregating_data_from_excel_files(pr_bar: ProgressBar,
                                      excel_files: List[Path],
                                      sheet_name_list: List[str]) -> List[str]:
    dict_df = {}
    missing_files = []
    for index, file_excel in enumerate(excel_files):
        try:
            lists_current_file = get_sheet_names(file_excel)
            set_1 = set(lists_current_file)
            result = [item for item in sheet_name_list if item in set_1]
            result.sort(key=str.lower)

            # словарь, ключ - имя листа, значение - датафрейм
            df_dict = pd.read_excel(file_excel, result, header=None)

            for key in df_dict:
                df_dict[key].insert(0, 'Имя листа', key)

            df = pd.concat(df_dict.values(), ignore_index=True)

            # Добавляем столбец с именем файла
            df.insert(0, 'Имя файла', file_excel.name)

            # Сохраняем DataFrame в словаре
            dict_df[file_excel] = df
        except ValueError:
            missing_files.append(file_excel.name)
        percentage = ((index+1) / len(excel_files)) * 100
        pr_bar.update(progress=percentage)
    pr_bar.update(progress=100)
    try:
        if dict_df:
            result = pd.concat(list(dict_df.values()), ignore_index=True)
            result.to_excel(NAME_OUTPUT_FILE, index = False)
            os.startfile(os.path.abspath(NAME_OUTPUT_FILE))
    except PermissionError:
        raise PermissionError(f"Ошибка доступа к {NAME_OUTPUT_FILE}")
    except FileNotFoundError:
        raise FileNotFoundError(f"Ошибка, файл {NAME_OUTPUT_FILE} не найден.")
    except OSError:
        raise OSError("Ошибка, не найдено приложение Excel.")
    except Exception:
        raise Exception("Неизвестная ошибка!")
    return missing_files

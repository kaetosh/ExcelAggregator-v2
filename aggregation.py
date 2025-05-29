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
from typing import List, Dict, Set
from textual.widgets import ProgressBar, Static

from data_text import (NAME_OUTPUT_FILE,
                       TEXT_CONCAT_PROCESS,
                       TEXT_LOAD_FILE_XLS,
                       TEXT_OPEN_FILE_XLS)


class NoExcelFilesError(Exception):
    """Custom exception for no Excel files found."""
    pass


def get_excel_files(folder_path: Path) -> List[Path]:
    """
    Получить список Excel файлов в указанной папке.
    Исключает временные файлы и файл с именем 'consolidated.xlsx'.
    """
    excel_extensions = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.odf')
    files = [
        f for f in folder_path.iterdir()
        if f.is_file()
        and f.suffix.lower() in excel_extensions
        and not f.name.startswith('~')
        and f.name.lower() != 'consolidated.xlsx'
    ]
    if not files:
        raise NoExcelFilesError("В указанной папке нет файлов Excel.")
    return files


def select_folder(current_path: Path) -> Path:
    """
    Открыть диалог выбора папки и вернуть выбранный путь.
    Если пользователь отменил выбор, вернуть current_path.
    """
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Выберите папку")
    root.destroy()
    return Path(folder_path) if folder_path else current_path


def get_sheet_names(file_path: Path) -> List[str]:
    """
    Получить список имен листов в Excel файле.
    """
    xls = pd.ExcelFile(file_path)
    return xls.sheet_names

def get_unique_sheet_names(file_paths: List[Path], one_prbar: ProgressBar) -> List[str]:
    """
    Получить отсортированный список уникальных листов из всех файлов.
    Обновляет прогресс бар во время обработки.
    """
    unique_sheets: Set[str] = set()
    total_files = len(file_paths)
    for index, file_path in enumerate(file_paths):
        if not file_path.exists() or file_path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm', '.xlsb', '.odf']:
            continue  # Пропускаем несуществующие или неподдерживаемые файлы
        try:
            workbook_sheetnames = get_sheet_names(file_path)
            unique_sheets.update(workbook_sheetnames)
        except Exception:
            # Логирование или обработка ошибок чтения листов можно добавить здесь
            pass
        percentage = ((index + 1) / total_files) * 100
        one_prbar.update(progress=percentage)
    one_prbar.update(progress=100)
    list_unique_sheets = sorted(unique_sheets, key=lambda s: s.lower())
    return list_unique_sheets



def aggregating_data_from_excel_files(static_widget: Static,
                                      pr_bar: ProgressBar,
                                      excel_files: List[Path],
                                      sheet_name_list: List[str]
                                      ) -> List[str]:
    """
    Агрегирует данные из указанных листов Excel-файлов в один файл.
    Возвращает список файлов, которые не удалось обработать.
    Обновляет прогресс бар во время обработки.
    """
    dict_df: Dict[Path, pd.DataFrame] = {}
    missing_files: List[str] = []
    try:
        total_files = len(excel_files)
    except TypeError:
        raise TypeError("Не выбрана папка с файлами.")

    for index, file_excel in enumerate(excel_files):
        try:
            lists_current_file = get_sheet_names(file_excel)
            available_sheets = set(lists_current_file)
            sheets_to_read = sorted([sheet for sheet in sheet_name_list if sheet in available_sheets], key=str.lower)

            if not sheets_to_read:
                missing_files.append(file_excel.name)
                percentage = ((index + 1) / total_files) * 100
                pr_bar.update(progress=percentage)
                continue

            # Читаем несколько листов одновременно, header=None — без заголовков
            df_dict = pd.read_excel(file_excel, sheets_to_read, header=None)

            # Добавляем колонку с именем листа в каждый DataFrame
            for key, df in df_dict.items():
                df.insert(0, 'Имя листа', key)

            # Объединяем все листы текущего файла
            df = pd.concat(df_dict.values(), ignore_index=True)

            # Добавляем колонку с именем файла
            df.insert(0, 'Имя файла', file_excel.name)

            dict_df[file_excel] = df
            percentage = ((index + 1) / total_files) * 100
            pr_bar.update(progress=percentage)
        except ValueError:
            missing_files.append(file_excel.name)
            percentage = ((index + 1) / total_files) * 100
            pr_bar.update(progress=percentage)
        except Exception:
            # Можно добавить логирование ошибок
            missing_files.append(file_excel.name)
            percentage = ((index + 1) / total_files) * 100
            pr_bar.update(progress=percentage)


        # Собираем данные с выбранных листов каждого файла...
    pr_bar.update(progress=100)

    try:
        if dict_df:
            # Объединяем данные в общий массив...
            static_widget.update(TEXT_CONCAT_PROCESS)
            result = pd.concat(dict_df.values(), ignore_index=True)
            # Выгружаем сводные данные в excel файл...
            static_widget.update(TEXT_LOAD_FILE_XLS)
            result.to_excel(NAME_OUTPUT_FILE, index=False)

            # Открываем файл в ОС (Windows)
            static_widget.update(TEXT_OPEN_FILE_XLS)
            if os.name == 'nt':
                os.startfile(os.path.abspath(NAME_OUTPUT_FILE))
    except PermissionError:
        raise PermissionError(f"Ошибка доступа к файлу {NAME_OUTPUT_FILE}")
    except FileNotFoundError:
        raise FileNotFoundError(f"Файл {NAME_OUTPUT_FILE} не найден.")
    except OSError:
        raise OSError("Ошибка: не найдено приложение Excel.")
    except Exception as e:
        raise Exception(f"Неизвестная ошибка: {e}")

    return missing_files

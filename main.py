# -*- coding: utf-8 -*-
"""
Created on Mon May 26 09:16:02 2025

@author: a.karabedyan
"""

from typing import List
from pathlib import Path

from textual import work, on
from textual.app import App, ComposeResult
from textual.reactive import reactive
from textual.containers import Horizontal, VerticalScroll
from textual.widgets import Header, Footer, Static, LoadingIndicator, ProgressBar, SelectionList
from textual.containers import Middle, Center

from aggregation import (select_folder,
                         get_excel_files,
                         aggregating_data_from_excel_files,
                         NoExcelFilesError,
                         NoSelectSheetsError,
                         get_unique_sheet_names,
                         is_excel_file_open)

from data_text import (NAME_APP,
                       NAME_OUTPUT_FILE,
                       SUB_TITLE_APP,
                       TEXT_INTRODUCTION,
                       TEXT_GENERAL,
                       TEXT_ERR_FILES_EXCEL,
                       TEXT_ERR_NO_PROCESSED_FILES,
                       TEXT_ERR_NOT_ALL_PROCESSED_FILES,
                       TEXT_ERR_PERMISSION,
                       TEXT_ERR_FILE_NOT_FOUND,
                       TEXT_APP_EXCEL_NOT_FIND,
                       TEXT_UNKNOW_ERR,
                       TEXT_ALL_PROCESSED_FILES,
                       TEXT_GENERATING_LIST_SHEETS,
                       TEXT_AGGREGATION_PROCESS,
                       TEXT_NOT_SELECT_DIR,
                       TEXT_ERR_NO_SELECT_SHEETS)


class ExcelAggregatorApp(App):
    CSS = """
    .introduction {
        height: auto;
        border: solid #0087d7;
    }
    VerticalScroll {
        width: 3fr;
    }
    .steps_l {
        height: auto;
    }
    SelectionList {
        height: 1fr;
        border: solid #0087d7;
        width: 1fr;
        margin: 0 0 0 1
    }
    Horizontal {
        height: 70%;
        border: solid #0087d7;
    }
    LoadingIndicator {
        dock: bottom;
        height: auto;
    }
    ProgressBar {
        padding-left: 3;
        height: auto;
    }
    """

    BINDINGS = [
        ("ctrl+o", "open_dir", "Выбрать папку с файлами Excel"),
        ("ctrl+r", "open_consolidate", "Сформировать сводный файл"),
        ("backspace", "deselect"),
    ]

    file_path: Path = reactive(Path.cwd())
    sheet_name: List[str] = reactive(['НЕ ВЫБРАНЫ'])
    names_files_excel: List[Path] = reactive(None)

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True, icon='<>')
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Horizontal(
            VerticalScroll(
                Static(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                           NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                           file_path=self.file_path,
                                           sheet_name='|'.join(self.sheet_name)), classes='steps_l')),
            SelectionList()
        )
        yield Footer()
        yield LoadingIndicator()
        with Middle():
            with Center():
                yield ProgressBar()

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(LoadingIndicator).visible = False
        self.query_one(ProgressBar).visible = False
        self.query_one(SelectionList).visible = True
        self.query_one(SelectionList).border_title = "Выберете листы:"


    @on(SelectionList.SelectedChanged)
    def handle_select_sheet(self):
        selected = self.query_one(SelectionList).selected
        self.sheet_name = selected if selected else ['НЕ ВЫБРАНЫ']
        self.update_steps_text()

    def action_deselect(self) -> None:
        self.query_one(SelectionList).deselect_all()

    def action_open_dir(self) -> None:
        self.file_path = select_folder(self.file_path) or Path.cwd()
        self.load_files_thread()
        self.sheet_name = ['НЕ ВЫБРАНЫ']
        self.query_one(SelectionList).clear_options()

    @work(thread=True)
    def load_files_thread(self) -> None:
        try:
            self.query_one('.steps_l').update(TEXT_GENERATING_LIST_SHEETS)
            self.query_one(LoadingIndicator).visible = True
            self.query_one(SelectionList).disabled = True
            self.names_files_excel = get_excel_files(self.file_path)
            self.query_one(ProgressBar).update(total=100)
            self.query_one(ProgressBar).visible = True

            sheet_names = get_unique_sheet_names(self.names_files_excel, self.query_one(ProgressBar))

            # Обновление интерфейса в основном потоке
            self.call_from_thread(self.update_sheet_names, sheet_names)

        except NoExcelFilesError:
            self.update_steps_text(TEXT_ERR_FILES_EXCEL)
        finally:
            self.reset_progress()

    def update_sheet_names(self, sheet_names):
        self.query_one(ProgressBar).visible = False
        self.query_one(SelectionList).clear_options()
        self.query_one(SelectionList).add_options([(name, name) for name in sheet_names])
        self.update_steps_text()

    def update_steps_text(self, error_message=None):
        message = error_message if error_message else TEXT_GENERAL
        self.query_one('.steps_l').update(message.format(NAME_APP=NAME_APP,
                                                         NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                         file_path=self.file_path,
                                                         sheet_name='|'.join(self.sheet_name)))

    def reset_progress(self):
        self.query_one(ProgressBar).visible = False
        self.query_one(ProgressBar).update(progress=0)
        self.query_one(LoadingIndicator).visible = False
        self.query_one(SelectionList).disabled = False

    @work(thread=True)
    def action_open_consolidate(self) -> None:
        self.query_one(LoadingIndicator).visible = True
        self.query_one('.steps_l').update(TEXT_AGGREGATION_PROCESS)
        self.query_one(SelectionList).disabled = True

        try:
            if not self.names_files_excel:
                raise TypeError("Не выбрана папка с файлами.")
            is_excel_file_open(NAME_OUTPUT_FILE)
            if 'НЕ ВЫБРАНЫ' in self.sheet_name:
                raise NoSelectSheetsError('Не выбраны листы для свода.')
            self.query_one(ProgressBar).update(total=100)
            self.query_one(ProgressBar).visible = True

            missing_files = aggregating_data_from_excel_files(self.query_one('.steps_l'),
                                                              self.query_one(ProgressBar),
                                                              self.names_files_excel,
                                                              self.sheet_name)
            self.handle_aggregation_results(missing_files)
        except (NoSelectSheetsError, NoExcelFilesError, PermissionError, FileNotFoundError, OSError, TypeError) as e:
            self.update_steps_text(self.get_error_message(e))
        except Exception as e:
            self.update_steps_text(TEXT_UNKNOW_ERR.format(text_err=e))
        finally:
            self.reset_progress()

    def get_error_message(self, error):
        error_messages = {
            NoSelectSheetsError: TEXT_ERR_NO_SELECT_SHEETS,
            NoExcelFilesError: TEXT_ERR_FILES_EXCEL,
            PermissionError: TEXT_ERR_PERMISSION,
            FileNotFoundError: TEXT_ERR_FILE_NOT_FOUND,
            OSError: TEXT_APP_EXCEL_NOT_FIND,
            TypeError: TEXT_NOT_SELECT_DIR
        }
        return error_messages.get(type(error), TEXT_UNKNOW_ERR.format(text_err=error))

    def handle_aggregation_results(self, missing_files):
        if len(missing_files) == len(self.names_files_excel):
            self.update_steps_text(TEXT_ERR_NO_PROCESSED_FILES)
        elif missing_files:
            self.update_steps_text(TEXT_ERR_NOT_ALL_PROCESSED_FILES.format(
                missing_files='\n'.join(missing_files),
                NAME_APP=NAME_APP,
                NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                sheet_name='|'.join(self.sheet_name),
                file_path=self.file_path,
            ))
        else:
            self.update_steps_text(TEXT_ALL_PROCESSED_FILES)


if __name__ == "__main__":
    app = ExcelAggregatorApp()
    app.run()

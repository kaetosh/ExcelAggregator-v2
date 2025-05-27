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
from textual.containers import Horizontal
from textual.widgets import Header, Footer, Static, LoadingIndicator, ProgressBar, Input, SelectionList
from textual.containers import Middle, Center

from aggregation import (select_folder,
                         get_excel_files,
                         aggregating_data_from_excel_files,
                         NoExcelFilesError,
                         get_unique_sheet_names)

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
                       TEXT_ALL_PROCESSED_FILES
                       )

class ExcelAggregatorApp(App):
    CSS = """
    .introduction {
        height: auto;
        border: solid #0087d7;
    }
    .SelectionList {
        height: auto;
    }
    .steps_l {
        height: auto;
        width: 3fr;
    }
    .steps_r {
        height: 100%;
        border: solid #0087d7;
        width: 1fr;
        margin: 0 0 0 1
    }
    Horizontal {
        height: auto;
        border: solid #0087d7;
    }
    LoadingIndicator {
        dock: bottom;
        height: 10%;
    }
    ProgressBar {
        padding-left: 3;
    }
    """

    BINDINGS = [
        ("ctrl+o", "open_dir", "1"),
        ("ctrl+r", "open_consolidate", "2"),
        ("backspace", "deselect"),
    ]
    file_path = reactive(Path.cwd())
    sheet_name = reactive(['Лист1'])

    def action_deselect(self) -> None:
        self.query_one(SelectionList).deselect_all()

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True, icon='<>')
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Horizontal(
                Static(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                           NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                           file_path=self.file_path,
                                           sheet_name=', '.join(self.sheet_name)), classes='steps_l'),
                SelectionList[int](classes='steps_r'), id='example')
        yield Footer()
        yield LoadingIndicator()
        with Middle():
            with Center():
                yield ProgressBar(id='one_prbar', classes='prbar')
                yield ProgressBar(id='two_prbar', classes='prbar')

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(LoadingIndicator).visible = False
        self.query_one('#one_prbar').visible = False
        self.query_one('#two_prbar').visible = False
        self.query_one(SelectionList).border_title = "Выберете листы:"

    @on(SelectionList.SelectedChanged)
    def handle_select_sheet(self):
        self.list_sheet_name = self.query_one(SelectionList).selected
        sheet_name = ', '.join(self.list_sheet_name)
        self.query_one('.steps_l').update(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                                              NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                              file_path=self.file_path,
                                                              sheet_name=sheet_name))
    def action_open_dir(self) -> None:
        # Выбор папки в главном потоке
        selected_path = select_folder(self.file_path)
        if not selected_path:
            selected_path = Path.cwd()
        self.file_path = selected_path

        # Запускаем тяжелую обработку в отдельном потоке
        self.load_files_thread()

    @work(thread=True)
    def load_files_thread(self) -> None:
        try:
            self.query_one('.steps_l').update('Идет чтение файлов и формирование списка листов...')
            self.query_one(LoadingIndicator).visible = True
            self.query_one(SelectionList).visible = False
            self.names_files_excel = get_excel_files(self.file_path)
            self.query_one('#one_prbar').update(total=100)
            self.query_one('#one_prbar').visible = True
            sheet_names = get_unique_sheet_names(self.names_files_excel, self.query_one('#one_prbar'))
            self.query_one('#one_prbar').visible = False
            self.query_one('#one_prbar').update(progress=0)
            self.sheet_names_for_update_sel_list = [(i, i) for i in sheet_names]
            self.query_one('.steps_l').update(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                                                  NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                  file_path=self.file_path,
                                                                  sheet_name=', '.join(self.sheet_name)))
            self.query_one(SelectionList).add_options(self.sheet_names_for_update_sel_list)
            self.query_one(SelectionList).visible = True
            self.query_one(LoadingIndicator).visible = False
        except NoExcelFilesError:
            self.query_one('.steps_l').update(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                                                  NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                  file_path=self.file_path,
                                                                  sheet_name=', '.join(self.sheet_name)))
        finally:
            self.query_one('#one_prbar').visible = False
            self.query_one('#one_prbar').update(progress=0)
            self.query_one(SelectionList).visible = True
            self.query_one(LoadingIndicator).visible = False





    @work(thread=True)
    def action_open_consolidate(self) -> None:
        self.query_one(LoadingIndicator).visible = True
        self.query_one('.steps_l').update('Идет обработка данных...')
        self.query_one('.steps_r').visible=False
        try:
            self.query_one('#one_prbar').update(total=100)
            self.query_one('#two_prbar').update(total=100)
            self.query_one('#one_prbar').visible = True
            self.query_one('#two_prbar').visible = True
            missing_files= aggregating_data_from_excel_files(self.query_one('#one_prbar'),
                                                             self.query_one('#two_prbar'),
                                                             self.names_files_excel,
                                                             self.list_sheet_name)
            if len(missing_files) == len(self.names_files_excel):
                self.query_one('.steps_l').update(TEXT_ERR_NO_PROCESSED_FILES.format(NAME_APP=NAME_APP,
                                                                                     NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                                     file_path=self.file_path,
                                                                                     sheet_name=', '.join(self.sheet_name)))
            elif missing_files:
                self.query_one('.steps_l').update(TEXT_ERR_NOT_ALL_PROCESSED_FILES.format(NAME_APP=NAME_APP,
                                                                                          NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                                          missing_files='\n'.join(missing_files),
                                                                                          file_path=self.file_path,
                                                                                          sheet_name=', '.join(self.sheet_name)))
            else:
                self.query_one('.steps_l').update(TEXT_ALL_PROCESSED_FILES.format(NAME_APP=NAME_APP,
                                                                                  NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                                  file_path=self.file_path,
                                                                                  sheet_name=', '.join(self.sheet_name)))
        except NoExcelFilesError:
            self.query_one('.steps_l').update(TEXT_ERR_FILES_EXCEL.format(NAME_APP=NAME_APP,
                                                                          NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                          file_path=self.file_path,
                                                                          sheet_name=', '.join(self.sheet_name)))
        except PermissionError:
            self.query_one('.steps_l').update(TEXT_ERR_PERMISSION.format(NAME_APP=NAME_APP,
                                                                         NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                         file_path=self.file_path,
                                                                         sheet_name=', '.join(self.sheet_name)))
        except FileNotFoundError:
            self.query_one('.steps_l').update(TEXT_ERR_FILE_NOT_FOUND.format(NAME_APP=NAME_APP,
                                                                             NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                             file_path=self.file_path,
                                                                             sheet_name=', '.join(self.sheet_name)))
        except OSError as e:
            self.query_one('.steps_l').update(TEXT_APP_EXCEL_NOT_FIND.format(NAME_APP=NAME_APP,
                                                                             NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                             error_app_xls = e,
                                                                             file_path=self.file_path,
                                                                             sheet_name=', '.join(self.sheet_name)))
        # except Exception as e:
        #     self.query_one('.steps_l').update(TEXT_UNKNOW_ERR.format(text_err=e))

        finally:
            self.query_one(LoadingIndicator).visible = False
            self.query_one('.steps_r').visible=True
            self.query_one('#one_prbar').visible = False
            self.query_one('#two_prbar').visible = False
            self.query_one('#one_prbar').update(progress=0)
            self.query_one('#two_prbar').update(progress=0)


if __name__ == "__main__":
    app = ExcelAggregatorApp()
    app.run()

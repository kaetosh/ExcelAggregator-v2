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
from textual.widgets import Header, Footer, Static, LoadingIndicator, ProgressBar, SelectionList
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
                       TEXT_ALL_PROCESSED_FILES,
                       TEXT_GENERATING_LIST_SHEETS,
                       TEXT_AGGREGATION_PROCESS,
                       TEXT_NOT_SELECT_DIR)


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
        height: 5fr;
        border: solid #0087d7;
    }
    LoadingIndicator {
        dock: bottom;
        height: 10%;
    }
    ProgressBar {
        padding-left: 3;
        height: 0.5fr
    }
    """

    BINDINGS = [
        ("ctrl+o", "open_dir", "Выбрать папку с файлами Excel"),
        ("ctrl+r", "open_consolidate", "Объединить таблицы"),
        ("backspace", "deselect"),
    ]

    file_path: Path = reactive(Path.cwd())
    sheet_name: List[str] = reactive(['НЕ ВЫБРАНЫ'])
    names_files_excel: List[Path] = reactive(None)

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True, icon='<>')
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Horizontal(
            Static(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                       NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                       file_path=self.file_path,
                                       sheet_name=', '.join(self.sheet_name)), classes='steps_l'),
            SelectionList[int](classes='steps_r'), id='example'
        )
        yield Footer()
        yield LoadingIndicator()
        with Middle():
            with Center():
                yield ProgressBar(id='one_prbar', classes='prbar')

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(LoadingIndicator).visible = False
        self.query_one('#one_prbar').visible = False
        self.query_one(SelectionList).border_title = "Выберете листы:"

    @on(SelectionList.SelectedChanged)
    def handle_select_sheet(self):
        self.sheet_name = self.query_one(SelectionList).selected
        self.update_steps_text()

    def action_deselect(self) -> None:
        self.query_one(SelectionList).deselect_all()
        self.sheet_name = ['НЕ ВЫБРАНЫ']

    def action_open_dir(self) -> None:
        selected_path = select_folder(self.file_path) or Path.cwd()
        self.file_path = selected_path
        self.load_files_thread()

    @work(thread=True)
    def load_files_thread(self) -> None:
        try:
            self.query_one('.steps_l').update(TEXT_GENERATING_LIST_SHEETS)
            self.query_one(LoadingIndicator).visible = True
            self.query_one(SelectionList).visible = False
            self.names_files_excel = get_excel_files(self.file_path)
            self.query_one('#one_prbar').update(total=100)
            self.query_one('#one_prbar').visible = True

            sheet_names = get_unique_sheet_names(self.names_files_excel, self.query_one('#one_prbar'))
            self.query_one('#one_prbar').visible = False

            self.query_one(SelectionList).clear_options()
            self.query_one(SelectionList).add_options([(i, i) for i in sheet_names])
            self.query_one(SelectionList).visible = True

            self.update_steps_text()
        except NoExcelFilesError:
            self.update_steps_text(TEXT_ERR_FILES_EXCEL)
        finally:
            self.reset_progress()

    def update_steps_text(self, error_message=None):
        if error_message:
            self.query_one('.steps_l').update(error_message.format(NAME_APP=NAME_APP,
                                                                   NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                   file_path=self.file_path,
                                                                   sheet_name=', '.join(self.sheet_name)))
        else:
            self.query_one('.steps_l').update(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                                                  NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                  file_path=self.file_path,
                                                                  sheet_name=', '.join(self.sheet_name)))

    def reset_progress(self):
        self.query_one('#one_prbar').visible = False
        self.query_one('#one_prbar').update(progress=0)
        self.query_one(LoadingIndicator).visible = False

    @work(thread=True)
    def action_open_consolidate(self) -> None:
        self.query_one(LoadingIndicator).visible = True
        self.query_one('.steps_l').update(TEXT_AGGREGATION_PROCESS)
        self.query_one('.steps_r').visible = False

        try:
            self.query_one('#one_prbar').update(total=100)
            self.query_one('#one_prbar').visible = True

            missing_files = aggregating_data_from_excel_files(self.query_one('#one_prbar'),
                                                              self.names_files_excel,
                                                              self.sheet_name)
            self.handle_aggregation_results(missing_files)
        except NoExcelFilesError:
            self.update_steps_text(TEXT_ERR_FILES_EXCEL)
        except PermissionError:
            self.update_steps_text(TEXT_ERR_PERMISSION)
        except FileNotFoundError:
            self.update_steps_text(TEXT_ERR_FILE_NOT_FOUND)
        except OSError as e:
            self.update_steps_text(TEXT_APP_EXCEL_NOT_FIND.format(error_app_xls=e))
        except TypeError:
            self.update_steps_text(TEXT_NOT_SELECT_DIR)
        # except Exception as e:
        #     self.update_steps_text(TEXT_UNKNOW_ERR.format(text_err=e))
        finally:
            self.reset_progress()
            self.query_one('.steps_r').visible = True

    def handle_aggregation_results(self, missing_files):
        if len(missing_files) == len(self.names_files_excel):
            self.update_steps_text(TEXT_ERR_NO_PROCESSED_FILES)
        elif missing_files:
            self.update_steps_text(TEXT_ERR_NOT_ALL_PROCESSED_FILES.format(missing_files='\n'.join(missing_files),
                                                                           NAME_APP=NAME_APP,
                                                                           NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                           sheet_name=', '.join(self.sheet_name),
                                                                           file_path=self.file_path,))
        else:
            self.update_steps_text(TEXT_ALL_PROCESSED_FILES)


if __name__ == "__main__":
    app = ExcelAggregatorApp()
    app.run()

from __future__ import annotations

import abc
import logging

import tkinter as tk

from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox

from typing import Any, Callable, Iterable

from . import errors
from .tables import PivotTable
from .core import Params


class GUI(abc.ABC):

    def __init__(self, available_tables: Iterable[PivotTable], core_func: Callable[[Params], None], logger: logging.Logger = None) -> None:
        super().__init__()
        self._available_tables = available_tables
        self._core_func = core_func
        self._logger = logger or logging.getLogger(self.__class__.__name__)

    def __enter__(self) -> GUI:
        return self.create()

    def __exit__(self, *exc_info) -> None:
        return self.destroy()

    @abc.abstractmethod
    def handle_error(self, error: Exception) -> None:
        pass
    
    @abc.abstractmethod
    def create(self) -> GUI:
        pass

    @abc.abstractmethod
    def destroy(self) -> None:
        pass
    
    @abc.abstractmethod
    def get_params(self) -> Params:
        pass

    def execute(self) -> None:
        params = self.get_params()
        self._logger.debug('running with params %s', params)
        try:
            self._core_func(params)
        except Exception as error:
            if not isinstance(error, errors.BaseError):
                self._logger.error('unhandled internal error: "%s"', str(error))
                error = errors.InternalError(error)
            else:
                self._logger.error("core error '%s' occured, trying to handle with the GUI...", error)
            self.handle_error(error)


class Placeholder(ttk.Entry):

    def __init__(self, master=None, placeholder='', color='black', placeholder_color='grey', *args, **kwargs):
        super().__init__(master=master, *args, **kwargs)
        self.color = color
        self.placeholder_color = placeholder_color
        self.placeholder = placeholder
        self.bind('<FocusOut>', lambda e: self.fill_placeholder())
        self.bind('<FocusIn>', lambda e: self.clear_placeholder())
        self.fill_placeholder()

    def clear_placeholder(self):
        if not self.get() and super().get():
            self.delete(0, tk.END)
        self.configure(foreground=self.color)

    def fill_placeholder(self):
        if not super().get():
            self.insert(0, self.placeholder)
            self.configure(foreground=self.placeholder_color)
    
    def get(self) -> str:
        content = super().get()
        if content == self.placeholder:
            return ''
        return content


class TkinterGUI(GUI):

    FD_INITIAL_DIR =  '~\\Documents'
    FD_FILETYPES = [('Excel', ('.xlsx', '.xls', '.xlsb', '.xlsm'))]
    FD_DEFAULT_EXT = '.xlsx'

    def set_entry_from_filedialog(self, entry: ttk.Entry, dialog: Callable) -> None:
        if not entry.get():
            value = dialog(
                parent=self.window, 
                initialdir=self.FD_INITIAL_DIR, 
                filetypes=self.FD_FILETYPES,
                defaultextension=self.FD_DEFAULT_EXT,
            )
            entry.insert(0, value)
            
    def set_input_path_from_filedialog(self, *event) -> None:
        self.set_entry_from_filedialog(self.input_path, fd.askopenfilename)

    def set_output_path_from_filedialog(self, *event) -> None:
        self.set_entry_from_filedialog(self.output_path, fd.asksaveasfilename)

    def configure(self) -> None:
        self.input_path.bind('<Button-1>', self.set_input_path_from_filedialog)
        self.output_path.bind('<Button-1>', self.set_output_path_from_filedialog)
        self.cancel.configure(command=self.window.destroy)
        self.start.configure(command=self.execute)

    def destroy(self) -> None:
        try:
            return self.window.destroy()
        except tk.TclError:
            pass

    def get_params(self) -> Params:
        pivot_tables = [table for table, variable in self.checked_tables.items() if variable.get()]
        return Params(
            input_path=self.input_path.get(),
            output_path=self.output_path.get(),
            sheet_name=self.sheet_name.get(),
            pivot_tables=pivot_tables,
        )
    
    def handle_error(self, error: Exception) -> None:
        pass

    def create(self) -> GUI:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(2)

        self.window = tk.Tk()
        self.window.title('Создание сводных отчетов')
        self.window.resizable(False, False)

        self.frame = ttk.Frame(self.window, padding=12)
        self.frame.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)

        self.window.columnconfigure(1, weight=1)
        self.window.rowconfigure(1, weight=1)

        self.window.columnconfigure(2, weight=1)
        self.window.rowconfigure(2, weight=1)

        self.input_label = ttk.Label(self.frame, text='Путь до исходного файла', justify=tk.LEFT)
        self.input_label.grid(row=0, column=0, sticky=tk.W)

        self.input_path = Placeholder(self.frame, justify=tk.LEFT, width=32)
        self.input_path.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))

        self.sheet_label = ttk.Label(self.frame, text='Лист')
        self.sheet_label.grid(row=0, column=2, sticky=tk.W)

        self.sheet_name = Placeholder(self.frame, placeholder='Необязательно', justify=tk.LEFT, width=16)
        self.sheet_name.grid(row=1, column=2, sticky=tk.W)

        self.output_label = ttk.Label(self.frame, text='Путь сохранения результата', justify=tk.LEFT)
        self.output_label.grid(row=2, column=0, columnspan=2, sticky=tk.W)

        self.output_path = Placeholder(self.frame, justify=tk.LEFT, width=32)
        self.output_path.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E))

        self.tables_label_frame = ttk.LabelFrame(self.frame, text='Отчеты')
        self.tables_label_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), ipadx=4, ipady=8)

        for child in self.frame.winfo_children(): 
            child.grid_configure(padx=4, pady=2)

        self.checked_tables = {table: tk.BooleanVar(value=True) for table in self._available_tables}
        for idx, (table, variable) in enumerate(self.checked_tables.items()):
            ttk.Checkbutton(self.tables_label_frame, text=table.name, variable=variable).grid(row=5 + idx, column=0, sticky=tk.W)

        self.start = ttk.Button(self.frame, text='Запуск', width=16)
        self.start.grid(row=7, column=1, sticky=(tk.W, tk.E), padx=4, pady=16)

        self.cancel = ttk.Button(self.frame, text='Закрыть')
        self.cancel.grid(row=7, column=2, sticky=(tk.W, tk.E), padx=4, pady=16)

        self.configure()
        self.window.mainloop()

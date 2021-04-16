import tkinter as tk
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfile, askdirectory
from idlelib.tooltip import Hovertip
from tkinter.messagebox import showerror, showinfo

from .report_creator import ReportCreator

import os

FILE_EXTENSIONS = [('Excel Files', '*.xls;*.xlsx'),
                   ('Excel Files, 2003', '*.xls'),
                   ('Excel Files, 2007', '*.xlsx')]


class ButtonWithLabel(tk.Frame):

    def __init__(self, parent, button_text, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.pack(fill=tk.X)

        self.path = None
        self.button = tk.Button(self, text=button_text, width=30, command=lambda: self.button_pressed())
        self.button.pack(side=tk.LEFT, padx=5, pady=5)

        self.label_text = tk.StringVar()
        self.label_text.set('No information.')
        self.label = tk.Label(self, textvariable=self.label_text)
        self.label.pack(fill=tk.X, padx=5, expand=True)

    def button_pressed(self):
        print('Create button_pressed() function')


class ChooseFileButton(ButtonWithLabel):

    def __init__(self, parent, button_text):
        super().__init__(parent, button_text)
        self.label_text.set('No file specified.')
        self.hover_tip = Hovertip(self.label, text=self.label_text.get())

    def button_pressed(self):
        file = askopenfile(mode='r', filetypes=FILE_EXTENSIONS)
        if not (file is None):
            self.path = os.path.normpath(file.name)
            self.label_text.set(self.path)
            self.hover_tip.__setattr__('text', self.label_text.get())


class ChooseDirButton(ButtonWithLabel):

    def __init__(self, parent, button_text):
        super().__init__(parent, button_text)
        self.label_text.set('No folder specified.')
        self.hover_tip = Hovertip(self.label, text=self.label_text.get())

    def button_pressed(self):
        directory = askdirectory()
        if not (directory is None):
            self.path = os.path.normpath(directory)
            self.label_text.set(self.path)
            self.hover_tip.__setattr__('text', self.label_text.get())


class MainApplication(tk.Frame):

    def __init__(self, parent, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.frame1 = ChooseFileButton(parent, 'Выбрать исходник')
        self.frame2 = ChooseFileButton(parent, 'Выбрать шаблон')
        self.frame3 = ChooseDirButton(parent, 'Выбрать папку для сохранения')

        self.start_button = tk.Button(parent, text='Создать', command=lambda: self.process())
        self.start_button.pack(fill=tk.X, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(parent, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)

    def process(self):
        if self.frame1.path and self.frame2.path and self.frame3.path:
            try:
                report_creator = ReportCreator(self.frame1.path, self.frame2.path, self.frame3.path)

                self.progress_bar['maximum'] = report_creator.total_rows - 4
                self.progress_bar['value'] = 0
                self.parent.update()

                for row in range(4, report_creator.total_rows):
                    report_creator.create_report_file(row)
                    self.progress_bar['value'] += 1
                    self.parent.update()

                del report_creator
                showinfo('Статус', 'Создание завершено')

            except Exception as e:
                showerror('Ошибка', e)
        else:
            showerror('Ошибка', 'Все поля должны быть заполнены.')

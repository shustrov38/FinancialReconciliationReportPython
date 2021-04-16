import tkinter as tk
from tkinter.messagebox import showerror

from .app_classes import MainApplication


class Application:

    def __init__(self, title):
        self.root = tk.Tk()
        self.title = title
        self.set_window_properties()

    def set_window_properties(self):
        self.root.resizable(width=False, height=False)
        self.root.title(self.title)

    def run(self):
        try:
            MainApplication(self.root).pack()
            self.root.mainloop()
        except Exception as e:
            showerror('Ошибка', e)

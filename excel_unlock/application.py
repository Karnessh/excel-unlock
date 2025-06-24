import tkinter as tk
from tkinter import ttk
from . import interfaces as intf

class Application(tk.Tk):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Excel password removal Application")
        self.columnconfigure(0, weight=1)
        self.resizable(False,False)

        self.mainWindow = intf.PasswordRemovalWindows(self)
        self.mainWindow.grid(row=1, padx=10, sticky=(tk.W + tk.E))


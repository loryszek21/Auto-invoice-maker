import os
from datetime import datetime
import openpyxl
from tkinter import Tk , filedialog 


class ExcelFile:
    def __init__(self, file_path):
        self.file_path = file_path

class FileManager:
    def __init__(self):
        self.excel_files = []

    def choose_files(self):
        root = Tk()  # Utwórz instancję klasy Tk
        root.wm_attributes('-topmost', 1)
        root.withdraw()  # Ukryj główne okno tkinter
        file_paths = filedialog.askopenfilenames(title="Wybierz pliki Excel")
        root.destroy()  
        if file_paths:
            self.excel_files = [ExcelFile(file_path) for file_path in file_paths]
            self.print_file_paths()
            # self.lift_dialog_window()

    def print_file_paths(self):
        for excel_file in self.excel_files:
            print(f"Plik: {excel_file.file_path}")

    def get_file_paths(self):
        return [excel_file.file_path for excel_file in self.excel_files]

    # def lift_dialog_window(self):
    #     self.top = Tk()
    #     self.top.title("Okno na pierwszym planie")
    #     self.top.lift()
    #     self.top.mainloop()
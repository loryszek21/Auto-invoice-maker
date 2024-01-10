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
        root.withdraw()  # Ukryj główne okno tkinter
        file_paths = filedialog.askopenfilenames(title="Wybierz pliki Excel")
        root.destroy()  # Zniszcz instancję klasy Tk po użyciu
        if file_paths:
            self.excel_files = [ExcelFile(file_path) for file_path in file_paths]
            self.print_file_paths()

    def print_file_paths(self):
        for excel_file in self.excel_files:
            print(f"Plik: {excel_file.file_path}")

    def get_file_paths(self):
        return [excel_file.file_path for excel_file in self.excel_files]





# if __name__ == "__main__":
#     file_manager = FileManager()
#     file_manager.choose_files()

        
# class ExcelFile:
#     def __init__(self, file_path):
#         self.file_path = file_path
#         self.last_time_save = self.get_last_time_save()
#         self.miesiac = self.extract_month()

#     def get_last_time_save(self):
#         try:
#             workbook = openpyxl.load_workbook(self.file_path)
#             return workbook.properties.modified
#         except Exception as e:
#             print(f"Błąd podczas pobierania daty ostatniego zapisu dla pliku {self.file_path}: {e}")
#             return None

#     def extract_month(self):
#         if self.last_time_save:
#             return self.last_time_save.month
#         return None

# class FileManager:
#     def __init__(self, folder_path):
#         self.folder_path = folder_path
#         self.excel_files = self.get_excel_files()

#     def get_excel_files(self):
#         files = os.listdir(self.folder_path)
#         file_paths = [os.path.join(self.folder_path, file) for file in files if file.lower().endswith(('.xlsx', '.xlsm', '.xls'))]
#         return [ExcelFile(file_path) for file_path in file_paths]

#     def print_file_info(self):
#         for excel_file in self.excel_files:
#             print(f"Plik: {excel_file.file_path}, Data ostatniego zapisu: {excel_file.last_time_save}, Miesiac: {excel_file.miesiac}")

#     def get_file_path_for_month(self, month):
#         filtered_files = [excel_file.file_path for excel_file in self.excel_files if excel_file.miesiac == month]
#         return filtered_files
# from tkinter import filedialog, Tk


    # Przykład użycia: Pobierz file_path dla plików z wybranego miesiąca (np. miesiąc = 5)
  

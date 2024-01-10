from GetFile import FileManager

import openpyxl
import pyautogui
import time
import json

class GetValueFromExcel:
    def __init__(self, sheet, transformations):
        self.sheet = sheet
        self.transformations = transformations

    def read_column_values(self, start_row):
        x = start_row
        column_indices = [3, 6, 7]  # Indeksy kolumn "C", "F", "G"
        data_array = []

        while True:
            row_data = []  # Lista przechowująca dane z jednego rzędu

            for column_index in column_indices:
                cell_value = self.sheet.cell(row=x, column=column_index).value

                if cell_value is None:
                    break

                if isinstance(cell_value, str):
                    cell_value = cell_value.strip().lower()

                    # Dodatkowa logika przekształcenia pełnej nazwy
                    transformed_value, cursor_shifts = self.transform_full_name(cell_value)
                    cell_value = transformed_value

                    # Przesuń kursor o określoną liczbę pozycji w dół
                    # pyautogui.press('down', presses=cursor_shifts)
                    row_data.append(cursor_shifts)
                row_data.append(cell_value)

            if not any(row_data):  # Sprawdza, czy w rzędzie są jakieś dane
                break

            print(f"Rząd {x}: {row_data}")
            data_array.append(row_data)
            x += 1

        return data_array

    def transform_full_name(self, full_name):
        # Sprawdź, czy istnieją transformacje dla danej pełnej nazwy
        for transformation in self.transformations:
            if transformation["full_name"] == full_name:
                return transformation["short_name"], transformation["cursor_shifts"]

        # Domyślne wartości, jeśli nie znaleziono transformacji
        return full_name, 0

# if __name__ == "__main__":
#     # Przykład użycia
#     file_manager = FileManager()
#     file_manager.choose_files()

#     file_paths = file_manager.get_file_paths()
#     for file_path in file_paths:
#         try:
#             workbook = openpyxl.load_workbook(file_path, data_only=True)
#         except Exception as e:
#             print(f"Blad podczas wczytywania pliku Excel: {e}")

#         if 'Arkusz2' in workbook.sheetnames:
#             sheet = workbook['Arkusz2']

#             # Wczytaj transformacje z pliku JSON
#             with open('Product_names.json', 'r') as json_file:
#                 transformations = json.load(json_file)

#             # Tworzymy obiekt GetValueFromExcel i przekazujemy obiekt 'sheet' oraz transformacje
#             excel_reader = GetValueFromExcel(sheet, transformations["transformations"])

#             # Wywołujemy metodę read_column_values
#             start_row = 21  # wiersz początkowy
#             values_list = excel_reader.read_column_values(start_row)

#             # Wyświetlamy odczytane wartości
#             print("Lista odczytanych wartości:")
#             for row_data in values_list:
#                 print(row_data)
#         else:
#             print("Arkusz 'Arkusz2' nie istnieje w pliku Excel.")







# import openpyxl
# # -*- coding: utf-8 -*-

# class GetValueFromExcel:
#     def __init__(self, sheet):
#         self.sheet = sheet

#     def read_column_values(self, start_row):
#         x = start_row
#         column_indices = [3, 6, 7]  # Indeksy kolumn "C", "F", "G"
#         data_array = []

#         while True:
#             row_data = []  # Lista przechowująca dane z jednego rzędu

#             for column_index in column_indices:
#                 cell_value = self.sheet.cell(row=x, column=column_index).value

#                 if cell_value is None:
#                     break

#                 if isinstance(cell_value, str):
#                     cell_value = cell_value.strip().lower()

#                 row_data.append(cell_value)

#             if not any(row_data):  # Sprawdza, czy w rzędzie są jakieś dane
#                 break

#             print(f"Rząd {x}: {row_data}")
#             print(row_data)
#             data_array.append(row_data)
#             x += 1
#         # print (data_array)
#         return data_array

# if __name__ == "__main__":
#     # Przykład użycia
#     try:
#         workbook = openpyxl.load_workbook(r'P:\szkoła\strony internetowe\python\Auto-invoice\Auto-invoice-maker\dokumenty\22_P_2023_ekonomik.xlsx', data_only=True)
#     except Exception as e:
#         print(f"Blad podczas wczytywania pliku Excel: {e}")

#     if 'Arkusz2' in workbook.sheetnames:
#         sheet = workbook['Arkusz2']

#         # Tworzymy obiekt GetValueFromExcel i przekazujemy obiekt 'sheet'
#         excel_reader = GetValueFromExcel(sheet)

#         # Wywołujemy metodę read_column_values
#         start_row = 21  # Dowolny wiersz początkowy
#         values_list = excel_reader.read_column_values(start_row)

#         # Wyświetlamy odczytane wartości
#         print("Lista odczytanych wartości:")
#         for row_data in values_list:
#             print(row_data)
#     else:
#         print("Arkusz 'Arkusz2' nie istnieje w pliku Excel.")












# {
#     "transformations": [
#         {"full_name": "filet z kurczaka", "short_name": "filet", "cursor_shifts": 10},
#         {"full_name": "inna_nazwa_1", "short_name": "skrocona_nazwa_1", "cursor_shifts": 5},
#         {"full_name": "inna_nazwa_2", "short_name": "skrocona_nazwa_2", "cursor_shifts": 7}
#         // Dodaj inne produkty i ich transformacje
#     ]
# }
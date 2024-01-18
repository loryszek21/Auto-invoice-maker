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
        print(full_name)

        full_name = full_name.strip()
        print(full_name)
        # Sprawdź, czy istnieją transformacje dla danej pełnej nazwy
        for transformation in self.transformations:
            if transformation["full_name"] == full_name:
                
                return transformation["short_name"], transformation["cursor_shifts"]
            # elif transformation["full_name"] != full_name:
            #     print("nie ma takiego w jsonie bład 404")

        # Domyślne wartości, jeśli nie znaleziono transformacji
        return full_name, 0

from GetFile import FileManager
from GetValueFromExcel import GetValueFromExcel
from InsertValue import InsertValue

import openpyxl
import json

if __name__ == "__main__":
    # Utwórz instancję FileManager
    file_manager = FileManager()

    # Wybierz pliki Excel
    file_manager.choose_files()

    # Uzyskaj ścieżki do wybranych plików
    file_paths = file_manager.get_file_paths()

    # Iteruj przez każdy wybrany plik
    for file_path in file_paths:
        try:
            # Załaduj plik Excel
            workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            print(f"Blad podczas wczytywania pliku Excel {file_path}: {e}")
            continue

        # Sprawdź, czy arkusz 'Arkusz2' istnieje w pliku Excel
        if 'Arkusz2' in workbook.sheetnames:
            sheet = workbook['Arkusz2']

            # Wczytaj transformacje z pliku JSON
            with open('Product_names.json', 'r', encoding='utf-8') as json_file:
                transformations = json.load(json_file)

            # Utwórz obiekt GetValueFromExcel i przekaz arkusz i transformacje
            excel_reader = GetValueFromExcel(sheet, transformations["transformations"])

            # Wywołaj metodę read_column_values
            start_row = 21  # Wiersz początkowy
            values_list = excel_reader.read_column_values(start_row)

            # Wyświetl odczytane wartości
            print(f"Lista odczytanych wartości z pliku {file_path}:")
            for row_data in values_list:
                print(row_data)

            # Utwórz obiekt InsertValue i przekaz odczytane wartości
            inserter = InsertValue(values_list)

            # Wywołaj metodę insert_data
            inserter.insert_data()

        else:
            print(f"Arkusz 'Arkusz2' nie istnieje w pliku Excel {file_path}.")

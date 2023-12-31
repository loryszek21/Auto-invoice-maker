import openpyxl
# -*- coding: utf-8 -*-

class GetValueFromExcel:
    def __init__(self, sheet):
        self.sheet = sheet

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

                row_data.append(cell_value)

            if not any(row_data):  # Sprawdza, czy w rzędzie są jakieś dane
                break

            print(f"Rząd {x}: {row_data}")
            print(row_data)
            data_array.append(row_data)
            x += 1
        # print (data_array)
        return data_array

if __name__ == "__main__":
    # Przykład użycia
    try:
        workbook = openpyxl.load_workbook(r'H:\python\Zbior_Wz_Do_HDI\dokumenty\3g.xlsx', data_only=True)
    except Exception as e:
        print(f"Blad podczas wczytywania pliku Excel: {e}")

    if 'Arkusz2' in workbook.sheetnames:
        sheet = workbook['Arkusz2']

        # Tworzymy obiekt GetValueFromExcel i przekazujemy obiekt 'sheet'
        excel_reader = GetValueFromExcel(sheet)

        # Wywołujemy metodę read_column_values
        start_row = 21  # Dowolny wiersz początkowy
        values_list = excel_reader.read_column_values(start_row)

        # Wyświetlamy odczytane wartości
        print("Lista odczytanych wartości:")
        for row_data in values_list:
            print(row_data)
    else:
        print("Arkusz 'Arkusz2' nie istnieje w pliku Excel.")

from GetFile import FileManager, ExcelFile
from GetValueFromExcel import GetValueFromExcel
import openpyxl

if __name__ == "__main__":
    folder_path = "H:\\python\\Zbior_Wz_Do_HDI\\dokumenty"
    file_manager = FileManager(folder_path)
    file_manager.print_file_info()

    selected_month = 12
    files_for_month = file_manager.get_file_path_for_month(selected_month)
    print("\n")
    print("\n")
    print(f"Pliki dla miesiaca {selected_month}: {files_for_month}")

    
# def main():
    

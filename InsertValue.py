# import pyautogui as gui, time
# screenWidth, screenHeight = gui.size();
import pyautogui as gui
import time
import json
import keyboard


# In InsertValue.py
class InsertValue:
    def __init__(self, data_list):
        self.data_list = data_list

    def insert_data(self):
        for row_data in self.data_list:
            nazwa_produktu = row_data[1]
            ilosc_strzalek = row_data[0]
            ilosc_towaru = str(row_data[2]).replace(".", ",")
            cena = str(row_data[3]).replace(".", ",") 

            keyboard.write(nazwa_produktu)
            time.sleep(1)
            gui.press('down', presses=int(ilosc_strzalek))
            time.sleep(1)
            gui.press('enter')
            gui.typewrite(ilosc_towaru)
            gui.press('tab')
            gui.press('right')
            gui.typewrite(cena)
            gui.press('enter', presses=4, interval=0.25)
          
            
    # gui.press("right")
            # Add code to insert data, for example, calling appropriate methods with pyautogui
            pass

    def insert_row(self, row_data):
        # Add code to insert a single row if needed
        print("dupa")

        pass












































































# class ExcelToPyAutoGUI:
#     def __init__(self, transformations):
#         self.transformations = transformations

#     def transform_and_execute(self, values_list):
#         for row_data in values_list:
#             product_name, category, quantity, price = row_data

#             # Znajdź i kliknij przycisk "Add"
#             add_button_found = self.locate_and_click('Add.png')
#             if not add_button_found:
#                 print("Nie udało się znaleźć przycisku 'Add'.")
#                 return

#             # Wprowadź nazwę kategorii
#             self.type_and_enter(category)

#             # Wprowadź nazwę produktu
#             self.type_and_enter(product_name)

#             # Przejdź do pola ilości
#             pyautogui.press('tab')

#             # Wprowadź ilość
#             self.type_and_enter(str(quantity))

#             # Przejdź do pola ceny
#             pyautogui.press('tab')

#             # Wprowadź cenę
#             self.type_and_enter(str(price))

#             # Przejdź do następnego pola (tu przycisk "Tab" i "Right")
#             pyautogui.press('tab')
#             pyautogui.press('right')

#             # Dodatkowe polecenia, jeśli są potrzebne

#     def locate_and_click(self, image_path, confidence=0.7):
#         res = pyautogui.locateOnScreen(image_path, confidence=confidence)
#         if res:
#             pyautogui.click(res)
#             return True
#         return False

#     def type_and_enter(self, text):
#         pyautogui.typewrite(text)
#         time.sleep(1)
#         pyautogui.press('enter')

# if __name__ == "__main__":
#     # Wczytaj transformacje z pliku JSON
#     with open('Product_names.json', 'r') as json_file:
#         transformations = json.load(json_file)

#     # Tworzymy obiekt ExcelToPyAutoGUI i przekazujemy transformacje
#     excel_to_pyautogui = ExcelToPyAutoGUI(transformations["transformations"])

#     # Przykładowe dane, które zostałyby wczytane z arkusza Excel
#     values_list = [
#         ["filet", "kategoria1", 10, 20],
#         ["szynka", "kategoria2", 5, 15],
#         # Dodaj więcej danych produktu, jeśli to konieczne
#     ]

#     # Wywołaj metodę transform_and_execute, aby przekształcić i wykonać polecenia PyAutoGUI
#     excel_to_pyautogui.transform_and_execute(values_list)




# res = gui.locateOnScreen('Add.png', confidence=0.7)
# print(res)
# gui.click(res)
# gui.typewrite("mare")
# time.sleep(1)
# gui.press('enter')
# gui.press('tab')
# gui.typewrite("filet")
# time.sleep(1)
# for i in range(9):
#     gui.press('down')
# gui.press('enter')
# gui.typewrite("20")
# gui.press('tab')
# gui.press('right')
# gui.typewrite("35")



# gui.moveTo(5,screenHeight-5)
# gui.click()

# gui.typewrite('Notatnik')
# 

# time.sleep(2) #damy 2 sekundy aplikacji notepad na poprawne uruchomienie
# gui.keyDown('winleft')
# gui.press('up')
# gui.keyUp('winleft')



# def identifyloc():
#     while True:
#         currentMouseX, currentMouseY = gui.position()
#         print(currentMouseX,currentMouseY)
#         time.sleep(3)
# identifyloc()
# time.sleep(3)


# for res in gui.locateOnScreen("Add.png"):
#     print(res)
# gui.click(159, 71)


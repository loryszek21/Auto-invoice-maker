### Automatyzacja MikroSubiekt z użyciem Pythona
##Opis
Aplikacja została stworzona w celu automatyzacji procesu pobierania danych z arkuszy Excel oraz ich wprowadzania do systemu MikroSubiekt. Wykorzystuje bibliotekę PyAutoGui do symulacji interakcji z interfejsem użytkownika.

##Funkcjonalności

Automatyczne pobieranie danych z arkuszy Excel.
Wprowadzanie danych do programu MikroSubiekt przy użyciu interfejsu graficznego.
Konfigurowalne mapowanie nazw produktów z arkusza Excel na odpowiadające im kategorie w MikroSubiekcie przy użyciu pliku JSON.
##Struktura projektu

Main.py: Plik główny programu, odpowiedzialny za uruchomienie procesu automatyzacji.
GetValueFromExcel.py: Klasa obsługująca odczyt danych z arkusza Excel.
InsertValue.py: Klasa obsługująca wprowadzanie danych do programu MikroSubiekt.
GetFile.py: Klasa do obsługi pobierania plików Excel.
Product_names.json: Plik JSON zawierający mapowanie nazw produktów.
##Wymagania
Python 3.x
Biblioteki: openpyxl, pyautogui
##Instrukcja użytkowania
Uruchom program MikroSubiekt
Uruchom plik Main.py.
Wybierz odpowiednie pliki Excel zawierające dane do importu.
Program automatycznie wczyta dane z arkuszy Excel i wprowadzi je do programu MikroSubiekt.
##Konfiguracja mapowania produktów
W pliku Product_names.json znajduje się mapowanie nazw produktów z arkusza Excel na odpowiadające im kategorie w programie MikroSubiekt. Aby zmienić to mapowanie, edytuj plik JSON zgodnie z potrzebami.

##Uwagi
Upewnij się, że program MikroSubiekt jest uruchomiony i otwarty przed uruchomieniem skryptu automatyzacji.

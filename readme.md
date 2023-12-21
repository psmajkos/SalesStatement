## Przygotowanie plików

ALLEGRO
1. Wejdź w "sprzedaż", następnie w "Zamówienia od kupujących". Na dole po lewej stronie znajduję się przycisk "Zestawienie sprzedaży" - nacisnij. Pojawi się zakres dat (zaznacz od poczatku do końca miesiąca)

2. "Parametry zamówenia" pozostają bez zmian, zaś w "Parametry przedmiotów" odznacz "Kraj wysłania", "Numer oferty", "Sygnatura", "Cena", "Parametry podatkowe" (Docelowo będzie zaznaczona), "Usługa dodatkowa".

3. Zmiany następują dopiero w "Dodatkowe informacje do zamówienia". W tej opcji odznaczamy wszystko. 

4. Nacisnij "Zamów" - w ciągu paru minut zestawienie pojawi się na mailu.

5. Po otrzymaniu maila pobierz zestawienie - Koniecznym jest zalogowanie się na konto z którego zamawialiśmy zestawienie sprzedaży.

6. Rozpakuj zestawienie plik ma rozszerzenie .csv. Polecam zmienić nazwę, żeby było bardziej czytelnie.

## Operacje na pliku .xlsx 

1. Włącz program SalesStatement i zaznacz wczesniej pobrany plik .csv - Po zaznaczeniu i wpisaniu nazwy pliku wyjściowego otrzymasz plik .xlsx.

2. Włącz wygenerowany plik .xlsx.

3. Zmienić nazwę "OrderDate" na "Data sprzedaży".

4. Dodać kolumne za "Data sprzedaży" o nazwie "Data wpływu środków na konto".

5. Do kolumny "Data wpływu środków na konto" wpisać daty "Data sprzedaży" 
(Od poniedzialku do czwartku +1 dzien, od czwartku do niedzieli date najblizszego poniedziałku)

6. Dołączyć kolumne "OrderCount" z "LineItemData" do "SelectedData".

7. Usunąć wszystkie wiersze zawierające SellerStatus jako "CANCELED".

8. Ukryć kolumne SellerStatus.

9. Zaznaczyć kolumne brutto i zmienić jej kolor na zółty.

10. Dodać dane z NIP, nazwą, numerem dokumentu.

11. Kolumny DeliveryAmount, TotalToPayAmount, Brutto, wartość netto, wartość VAT, Kwota otrzymana sformatować jako walute "zł".
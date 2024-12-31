# TechSTORE: Mini ERP w Excelu

Ten projekt zawiera proste rozwiązanie systemu ERP zrealizowane w Excelu, wspomagające zarządzanie:
- Klientami
- Produktami
- Zamówieniami
- Raportami

## Zawartość pliku `TechSTORE.xlsm`

### Arkusze i ich zastosowanie:
- **Klienci**: Informacje o klientach, takie jak dane kontaktowe i historia zakupów.
- **Produkty**: Szczegółowe informacje o produktach, w tym nazwy, ceny i stany magazynowe.
- **Zamówienia**: Rejestracja zamówień, z informacjami o klientach, produktach i datach.
- **Raporty**: Analizy danych na podstawie zebranych informacji.

## Wymagania systemowe

- Microsoft Excel z obsługą makr (np. Excel 2016 lub nowszy).
- Włączona obsługa makr (VBA).

## Jak zacząć używać?

1. Pobierz plik **TechSTORE.xlsm**.
2. Otwórz plik w **Microsoft Excel**.
3. Upewnij się, że włączona jest obsługa makr.
4. Uzupełniaj dane w arkuszach:
   - Dodawaj nowe rekordy w arkuszach **Klienci** i **Produkty**.
   - Rejestruj zamówienia w arkuszu **Zamówienia**.
   - Analizuj wyniki w arkuszu **Raporty**.

## Makra

### `DodajKlienta`
To makro pozwala na szybkie dodanie nowego klienta do arkusza **Klienci**. Wprowadza dane za pomocą kilku okien dialogowych:

- **Funkcjonalność:**
1. Dodaje unikalny identyfikator klienta.
2. Pobiera dane klienta za pomocą okien InputBox (nazwa, kontakt, adres, typ).
3. Automatycznie rejestruje datę utworzenia klienta.
4. Informuje użytkownika komunikatem o pomyślnym dodaniu klienta.

- **Kod VBA:**

```vba
Sub DodajKlienta()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Klienci")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = lastRow - 1 ' ID
    ws.Cells(lastRow, 2).Value =  InputBox("Podaj nazwę klienta:")
    ws.Cells(lastRow, 3).Value = InputBox("Podaj kontakt (telefon/email):")
    ws.Cells(lastRow, 4).Value = InputBox("Podaj adres klienta:")
    ws.Cells(lastRow, 5).Value = InputBox("Podaj typ klienta (Firma/Detaliczny):")
    ws.Cells(lastRow, 6).Value = Date ' Data rejestracji
    
    MsgBox "Klient został dodany!"
End Sub
```
### `DodajProdukt`
To makro pozwala na szybkie dodanie nowego produktu do arkusza **Produkty**. Wprowadza dane za pomocą kilku okien dialogowych:

- **Funkcjonalność:**
1. Generuje unikalny identyfikator produktu.
2. Pobiera dane produktu za pomocą okien InputBox (nazwa, kategoria, cena, ilość na magazynie, producent).
3. Automatycznie rejestruje datę wprowadzenia produktu.
4. Informuje użytkownika komunikatem o pomyślnym dodaniu produktu.

- **Kod VBA:**
  
  ```vba
  Sub DodajProdukt()
      Dim ws As Worksheet
      Set ws = ThisWorkbook.Sheets("Produkty")
      
      Dim lastRow As Long
      lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
      
      ws.Cells(lastRow, 1).Value = lastRow - 1 ' ID
      ws.Cells(lastRow, 2).Value = InputBox("Podaj nazwę produktu:")
      ws.Cells(lastRow, 3).Value = InputBox("Podaj kategorię produktu:")
      ws.Cells(lastRow, 4).Value = InputBox("Podaj cenę produktu:")
      ws.Cells(lastRow, 5).Value = InputBox("Podaj ilość na magazynie:")
      ws.Cells(lastRow, 6).Value = InputBox("Podaj producenta:")
      ws.Cells(lastRow, 7).Value = Date ' Data wprowadzenia
      
      MsgBox "Produkt został dodany!"
  End Sub

### `DodajIloscDoMagazynu`

To makro umożliwia aktualizację ilości produktu w arkuszu **Produkty**. Użytkownik może wprowadzić ID produktu oraz ilość do dodania do magazynu za pomocą okien dialogowych.

#### **Funkcjonalność:**
1. **Pobranie danych:** 
   - ID produktu, którego ilość ma zostać zaktualizowana.
   - Ilość, którą użytkownik chce dodać.
  
2. **Walidacja wprowadzonej ilości:** 
   - Makro sprawdza, czy ilość do dodania jest większa niż 0.

3. **Aktualizacja magazynu:** 
   - Produkt o wskazanym ID jest wyszukiwany w arkuszu „Produkty” (w kolumnie A).
   - Ilość produktu w kolumnie E (ilość w magazynie) jest aktualizowana o wartość wprowadzonego parametru.

4. **Informowanie użytkownika:** 
   - Jeśli produkt o danym ID zostanie znaleziony, jego ilość w magazynie jest zaktualizowana.
   - Jeśli produkt o podanym ID nie zostanie znaleziony, użytkownik otrzyma odpowiedni komunikat.

-**Kod VBA:**

```vba
Sub DodajIloscDoMagazynu()
    Dim wsProdukty As Worksheet
    Dim idProduktu As String
    Dim iloscDoDodania As Long
    Dim znaleziono As Boolean
    Dim lastRow As Long
    Dim i As Long
    
    ' Ustawienie arkusza Produkty
    Set wsProdukty = ThisWorkbook.Sheets("Produkty")
    
    ' Pobieranie ID Produktu od użytkownika
    idProduktu = InputBox("Podaj ID Produktu, którego ilość chcesz zaktualizować:")
    
    ' Pobieranie ilości do dodania
    iloscDoDodania = InputBox("Podaj ilość do dodania:")
    
    ' Sprawdzanie, czy podano liczbę dodatnią
    If iloscDoDodania <= 0 Then
        MsgBox "Ilość do dodania musi być liczbą większą niż 0!"
        Exit Sub
    End If
    
    ' Znalezienie wiersza z odpowiednim ID Produktu
    znaleziono = False
    lastRow = wsProdukty.Cells(wsProdukty.Rows.Count, 1).End(xlUp).Row ' ostatni wiersz
    
    For i = 2 To lastRow ' Zakładając, że dane zaczynają się od drugiego wiersza
        If wsProdukty.Cells(i, 1).Value = idProduktu Then
            ' Zaktualizowanie ilości
            wsProdukty.Cells(i, 5).Value = wsProdukty.Cells(i, 5).Value + iloscDoDodania
            znaleziono = True
            Exit For
        End If
    Next i
    
    ' Jeśli nie znaleziono produktu, wyświetl komunikat
    If Not znaleziono Then
        MsgBox "Produkt o podanym ID nie został znaleziony w magazynie!"
    Else
        MsgBox "Ilość produktu " & idProduktu & " została zaktualizowana!"
    End If
End Sub
```

### `DodajZamowienie`

To makro pozwala na dodanie nowego zamówienia do arkusza **Zamówienia**. Wprowadza dane za pomocą kilku okien dialogowych:

- **Funkcjonalność:**
1. **Dodawanie nowego zamówienia:** Użytkownik może dodać zamówienie do arkusza „Zamówienia” wprowadzając:
      - ID klienta,
      - ID produktów (jedno lub wiele, oddzielone przecinkami),
      - Ilość zamawianych sztuk dla każdego produktu,
      - Metodę płatności.
  
2. **Zarządzanie stanem magazynowym:** Po każdym zamówieniu dostępna ilość produktów w arkuszu **„Produkty”** jest aktualizowana.
  
3. **Automatyczne obliczanie łącznej kwoty zamówienia:** Na podstawie ceny produktów oraz wprowadzonej ilości program wylicza sumę zamówienia.

- **Kod VBA:**

```vba
Sub DodajZamowienie()
    Dim wsOrders As Worksheet
    Dim wsProducts As Worksheet
    Dim wsClients As Worksheet
    Dim lastOrderRow As Long
    Dim lastProductRow As Long
    Dim productIDs As String
    Dim productID As String
    Dim totalPrice As Double
    Dim productPrice As Double
    Dim productQuantity As Long
    Dim orderedQuantity As Long
    Dim i As Long, j As Long
    Dim orderID As String
    Dim currentDate As String
    Dim productFound As Boolean
    
    ' Ustawiamy arkusze
    Set wsOrders = ThisWorkbook.Sheets("Zamówienia")
    Set wsProducts = ThisWorkbook.Sheets("Produkty")
    Set wsClients = ThisWorkbook.Sheets("Klienci")
    
    ' Znajdowanie ostatniego wiersza w arkuszu zamówienia
    lastOrderRow = wsOrders.Cells(wsOrders.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Generowanie ID Zamówienia
    orderID = "Z" & Format(lastOrderRow - 1, "0000")
    wsOrders.Cells(lastOrderRow, 1).Value = orderID ' Wpisanie ID zamówienia
    
    ' Pobranie dzisiejszej daty
    currentDate = Date
    wsOrders.Cells(lastOrderRow, 2).Value = currentDate ' Wpisanie daty
    
    ' Wprowadzanie ID Klienta
    wsOrders.Cells(lastOrderRow, 3).Value = InputBox("Podaj ID Klienta:")
    
    ' Wprowadzanie ID Listy Produktów (można wpisać kilka ID oddzielonych przecinkami)
    productIDs = InputBox("Podaj ID Listy Produktów (oddzielone przecinkami):")
    
    ' Przekształcenie ID produktów do tablicy
    productIDs = Trim(productIDs) ' Usuwanie zbędnych spacji
    productIDs = Replace(productIDs, " ", "") ' Usuwanie spacji
    Dim productIDArray() As String
    productIDArray = Split(productIDs, ",")
    
    ' Inicjalizacja zmiennej na łączną cenę
    totalPrice = 0
    
    ' Przeszukiwanie tablicy ID produktów
    For i = LBound(productIDArray) To UBound(productIDArray)
        productID = productIDArray(i)
        productFound = False ' Flaga informująca, czy produkt został znaleziony
        
        ' Znajdowanie ceny i ilości produktu w arkuszu Produkty
        lastProductRow = wsProducts.Cells(wsProducts.Rows.Count, 1).End(xlUp).Row
        For j = 2 To lastProductRow
            If wsProducts.Cells(j, 1).Value = productID Then
                productPrice = wsProducts.Cells(j, 4).Value
                ' Sprawdzanie, czy wartość komórki w kolumnie E (ilość) jest liczbą
                If IsNumeric(wsProducts.Cells(j, 5).Value) Then
                    productQuantity = wsProducts.Cells(j, 5).Value ' Ilość dostępnych produktów (kolumna E)
                Else
                    productQuantity = 0 ' Jeśli wartość nie jest liczbą, ustawiamy ilość na 0
                End If
                productFound = True
                Exit For
            End If
        Next j
        
        ' Sprawdzamy, czy produkt został znaleziony
        If productFound Then
            ' Wprowadzanie ilości zamawianych sztuk
            orderedQuantity = InputBox("Podaj ilość zamawianych sztuk dla produktu " & productID & " (dostępnych: " & productQuantity & "):")
            
            ' Sprawdzamy, czy zamówiona ilość jest dostępna
            If orderedQuantity <= productQuantity And orderedQuantity > 0 Then
                ' Zmniejszamy ilość w arkuszu Produkty
                wsProducts.Cells(j, 5).Value = productQuantity - orderedQuantity ' Zmniejszamy dostępne sztuki (kolumna E)
                totalPrice = totalPrice + (productPrice * orderedQuantity) ' Dodajemy cenę zamówionych sztuk do łącznej kwoty
            Else
                MsgBox "Brak wystarczającej ilości produktu " & productID & " lub niepoprawna liczba.", vbExclamation
                Exit Sub
            End If
        Else
            MsgBox "Produkt o ID " & productID & " nie został znaleziony.", vbExclamation
            Exit Sub
        End If
    Next i
    
    ' Wpisywanie łącznej kwoty do arkusza Zamówienia
    wsOrders.Cells(lastOrderRow, 4).Value = productIDs ' Zapisujemy ID produktów
    wsOrders.Cells(lastOrderRow, 5).Value = totalPrice ' Zapisujemy łączną kwotę
    
    ' Status zamówienia i metoda płatności
    wsOrders.Cells(lastOrderRow, 6).Value = "Nowe" ' Status zamówienia (domyślnie)
    wsOrders.Cells(lastOrderRow, 7).Value = InputBox("Podaj metodę płatności (np. Przelew, Karta, Blik):") ' Metoda płatności
    
    MsgBox "Zamówienie zostało dodane. Łączna kwota: " & totalPrice
End Sub
```


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
### `AktualizujCeneProduktu`

To makro pozwala na aktualizację ceny produktu w arkuszu **Produkty**. Użytkownik może wprowadzić ID produktu oraz nową cenę za pomocą okien dialogowych.

#### **Funkcjonalność:**

1. **Pobranie danych:** 
   - ID produktu, którego cena ma zostać zaktualizowana.
   - Nową cenę produktu.

2. **Walidacja nowej ceny:** 
   - Makro sprawdza, czy cena jest liczbą dodatnią.

3. **Aktualizacja ceny:** 
   - Produkt o wskazanym ID jest wyszukiwany w arkuszu „Produkty” (w kolumnie A).
   - Cena produktu w kolumnie D (cena jednostkowa) jest aktualizowana na podstawie wprowadzonej nowej ceny.

4. **Informowanie użytkownika:** 
   - Jeśli produkt o danym ID zostanie znaleziony, jego cena jest zaktualizowana.
   - Jeśli produkt o podanym ID nie zostanie znaleziony, użytkownik otrzyma odpowiedni komunikat.

-**Kod VBA:**

```vba
Sub AktualizujCeneProduktu()
    Dim wsProdukty As Worksheet
    Dim idProduktu As String
    Dim nowaCena As Double
    Dim znaleziono As Boolean
    Dim lastRow As Long
    Dim i As Long
    
    ' Ustawienie arkusza Produkty
    Set wsProdukty = ThisWorkbook.Sheets("Produkty")
    
    ' Pobieranie ID Produktu od użytkownika
    idProduktu = InputBox("Podaj ID Produktu, którego cenę chcesz zaktualizować:")
    
    ' Pobieranie nowej ceny od użytkownika
    nowaCena = InputBox("Podaj nową cenę produktu:")
    
    ' Sprawdzanie, czy cena jest liczbą dodatnią
    If nowaCena <= 0 Then
        MsgBox "Cena musi być liczbą większą niż 0!"
        Exit Sub
    End If
    
    ' Znalezienie wiersza z odpowiednim ID Produktu
    znaleziono = False
    lastRow = wsProdukty.Cells(wsProdukty.Rows.Count, 1).End(xlUp).Row ' ostatni wiersz
    
    For i = 2 To lastRow ' Zakładając, że dane zaczynają się od drugiego wiersza
        If wsProdukty.Cells(i, 1).Value = idProduktu Then
            ' Zaktualizowanie ceny
            wsProdukty.Cells(i, 4).Value = nowaCena
            znaleziono = True
            Exit For
        End If
    Next i
    
    ' Jeśli nie znaleziono produktu, wyświetl komunikat
    If Not znaleziono Then
        MsgBox "Produkt o podanym ID nie został znaleziony!"
    Else
        MsgBox "Cena produktu " & idProduktu & " została zaktualizowana na " & nowaCena & "!"
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
### `RaportNajwyzszeINajnizszeZamowienia`

To makro generuje raport na temat najwyższego i najniższego zamówienia w arkuszu **Zamówienia**. Raport zawiera szczegóły dotyczące ID zamówienia, ID klienta, listy produktów oraz łącznej kwoty zamówienia. Wyniki są zapisywane w arkuszu **Raporty**.

#### **Funkcjonalność:**

1. **Pobieranie danych z arkusza "Zamówienia":**
   - Arkusz **Zamówienia** musi zawierać dane o zamówieniach, w tym kwotę zamówienia (kolumna 5), ID klienta (kolumna 3) oraz ID listy produktów (kolumna 4).
   
2. **Obliczenia:**
   - Makro przeszukuje dane zamówień, aby znaleźć:
     - Najwyższą kwotę zamówienia
     - Najniższą kwotę zamówienia

3. **Tworzenie raportu:**
   - Po przetworzeniu danych makro generuje raport w arkuszu **Raporty**, zawierający:
     - ID najwyższego i najniższego zamówienia,
     - ID klienta,
     - Listę produktów,
     - Kwotę zamówienia.

4. **Wyczyszczenie raportu:** 
   - Zanim dane zostaną zapisane w arkuszu **Raporty**, wszystkie poprzednie dane są usuwane.

-**Kod VBA:**

```vba
Sub RaportNajwyzszeINajnizszeZamowienia()
    Dim wsZamowienia As Worksheet
    Dim wsRaporty As Worksheet
    Dim lastRowZamowienia As Long
    Dim i As Long
    Dim kwota As Double
    Dim maxKwota As Double
    Dim minKwota As Double
    Dim maxIDKlienta As String, minIDKlienta As String
    Dim maxProdukty As String, minProdukty As String
    Dim maxZamowienieID As String, minZamowienieID As String

    ' Ustawienia początkowe
    Set wsZamowienia = ThisWorkbook.Sheets("Zamówienia")
    Set wsRaporty = ThisWorkbook.Sheets("Raporty")
    lastRowZamowienia = wsZamowienia.Cells(wsZamowienia.Rows.Count, 1).End(xlUp).Row

    ' Sprawdź, czy są dane w arkuszu "Zamówienia"
    If lastRowZamowienia < 2 Then
        MsgBox "Brak danych w arkuszu 'Zamówienia'.", vbExclamation
        Exit Sub
    End If

    ' Inicjalizuj zmienne
    maxKwota = -1
    minKwota = WorksheetFunction.Max(wsZamowienia.Columns(5)) + 1 ' Największa możliwa liczba w kolumnie Kwota

    ' Przejdź przez dane w arkuszu "Zamówienia"
    For i = 2 To lastRowZamowienia
        ' Pobierz kwotę
        kwota = wsZamowienia.Cells(i, 5).Value ' Zakładamy, że kolumna 5 to "Łączna kwota"

        ' Sprawdź, czy kwota jest liczbą
        If IsNumeric(kwota) And kwota > 0 Then
            ' Jeśli to najwyższa kwota
            If kwota > maxKwota Then
                maxKwota = kwota
                maxIDKlienta = wsZamowienia.Cells(i, 3).Value ' ID Klienta
                maxProdukty = wsZamowienia.Cells(i, 4).Value ' ID Lista produktów
                maxZamowienieID = wsZamowienia.Cells(i, 1).Value ' ID Zamówienia
            End If
            
            ' Jeśli to najniższa kwota
            If kwota < minKwota Then
                minKwota = kwota
                minIDKlienta = wsZamowienia.Cells(i, 3).Value ' ID Klienta
                minProdukty = wsZamowienia.Cells(i, 4).Value ' ID Lista produktów
                minZamowienieID = wsZamowienia.Cells(i, 1).Value ' ID Zamówienia
            End If
        End If
    Next i

    ' Wyczyść arkusz "Raporty"
    wsRaporty.Cells.Clear

    ' Dodaj nagłówki do arkusza "Raporty"
    wsRaporty.Cells(1, 1).Value = "Typ Raportu"
    wsRaporty.Cells(1, 2).Value = "ID Zamówienia"
    wsRaporty.Cells(1, 3).Value = "ID Klienta"
    wsRaporty.Cells(1, 4).Value = "Lista Produktów"
    wsRaporty.Cells(1, 5).Value = "Kwota"

    ' Zapisz dane najwyższego zamówienia
    wsRaporty.Cells(2, 1).Value = "Najwyższe zamówienie"
    wsRaporty.Cells(2, 2).Value = maxZamowienieID
    wsRaporty.Cells(2, 3).Value = maxIDKlienta
    wsRaporty.Cells(2, 4).Value = maxProdukty
    wsRaporty.Cells(2, 5).Value = maxKwota

    ' Zapisz dane najniższego zamówienia
    wsRaporty.Cells(3, 1).Value = "Najniższe zamówienie"
    wsRaporty.Cells(3, 2).Value = minZamowienieID
    wsRaporty.Cells(3, 3).Value = minIDKlienta
    wsRaporty.Cells(3, 4).Value = minProdukty
    wsRaporty.Cells(3, 5).Value = minKwota

    MsgBox "Raport został wygenerowany!", vbInformation
End Sub
```

### `RaportNajwiecejINajmniejSprzedanychProduktow`

To makro generuje raport o produktach, które były najczęściej oraz najmniej zamawiane w arkuszu **Zamówienia**. Wyniki są zapisywane w arkuszu **Raporty**.

#### **Funkcjonalność:**

1. **Zliczanie sprzedanych produktów:** 
   - Makro przeszukuje wszystkie zamówienia zapisane w arkuszu "Zamówienia" i zlicza wystąpienia ID produktów (zawartych w kolumnie "ID lista produktów").
   
2. **Obliczenie najczęściej i najmniej zamawianych produktów:**
   - Makro identyfikuje produkt, który został zamówiony najwięcej razy oraz produkt, który wystąpił najmniej razy w całym zestawie zamówień.

3. **Generowanie raportu:**
   - Raport zawiera szczegóły dotyczące:
     - ID najczęściej zamawianego produktu i liczby jego zamówień.
     - ID najmniej zamawianego produktu i liczby jego zamówień.
   
4. **Wyczyszczenie raportu:** 
   - Przed zapisaniem wyników w arkuszu "Raporty", wszystkie poprzednie dane są usuwane.

-**Kod VBA:**

```vba
Sub RaportNajwiecejINajmniejSprzedanychProduktow()
    Dim wsZamowienia As Worksheet
    Dim wsRaporty As Worksheet
    Dim lastRowZamowienia As Long
    Dim i As Long
    Dim produktID As Variant ' Zmieniono na Variant
    Dim maxSprzedaz As Long
    Dim minSprzedaz As Long
    Dim maxProdukt As String
    Dim minProdukt As String
    Dim produktyCount As Object
    Dim produktIDs As Variant
    Dim j As Long

    ' Ustawienia początkowe
    Set wsZamowienia = ThisWorkbook.Sheets("Zamówienia")
    Set wsRaporty = ThisWorkbook.Sheets("Raporty")
    lastRowZamowienia = wsZamowienia.Cells(wsZamowienia.Rows.Count, 1).End(xlUp).Row

    ' Sprawdź, czy są dane w arkuszu "Zamówienia"
    If lastRowZamowienia < 2 Then
        MsgBox "Brak danych w arkuszu 'Zamówienia'.", vbExclamation
        Exit Sub
    End If

    ' Inicjalizuj słownik do zliczania produktów
    Set produktyCount = CreateObject("Scripting.Dictionary")
    maxSprzedaz = -1
    minSprzedaz = lastRowZamowienia + 1 ' Ustawienie wartości maksymalnej na bardzo dużą liczbę

    ' Przejdź przez dane w arkuszu "Zamówienia"
    For i = 2 To lastRowZamowienia
        ' Pobierz listę produktów z kolumny "ID lista produktów"
        produktIDs = Split(wsZamowienia.Cells(i, 4).Value, ",") ' Zakładamy, że kolumna 4 to "ID lista produktów"
        
        ' Sprawdź, czy lista produktów nie jest pusta
        If Len(wsZamowienia.Cells(i, 4).Value) > 0 Then
            ' Zlicz każdy ID produktowy w liście
            For j = LBound(produktIDs) To UBound(produktIDs)
                produktID = Trim(produktIDs(j)) ' Pobierz ID produktu, usuwając ewentualne spacje

                ' Sprawdź, czy produktID już istnieje w słowniku, jeśli tak, zwiększ liczbę
                If produktyCount.Exists(produktID) Then
                    produktyCount(produktID) = produktyCount(produktID) + 1
                Else
                    produktyCount.Add produktID, 1 ' Jeśli nie, dodaj do słownika z wartością 1
                End If
            Next j
        End If
    Next i

    ' Inicjalizuj zmienne dla najczęściej i najmniej zamawianych produktów
    maxProdukt = ""
    minProdukt = ""
    maxSprzedaz = -1
    minSprzedaz = lastRowZamowienia + 1

    ' Przejdź przez zliczone dane i znajdź max i min
    For Each produktID In produktyCount.Keys
        If produktyCount(produktID) > maxSprzedaz Then
            maxSprzedaz = produktyCount(produktID)
            maxProdukt = produktID
        End If
        If produktyCount(produktID) < minSprzedaz Then
            minSprzedaz = produktyCount(produktID)
            minProdukt = produktID
        End If
    Next produktID

    ' Wyczyść arkusz "Raporty"
    wsRaporty.Cells.Clear

    ' Dodaj nagłówki do arkusza "Raporty"
    wsRaporty.Cells(1, 1).Value = "Typ Raportu"
    wsRaporty.Cells(1, 2).Value = "ID Produktu"
    wsRaporty.Cells(1, 3).Value = "Liczba Zamówień"

    ' Wstaw dane o najczęściej zamawianym produkcie
    wsRaporty.Cells(2, 1).Value = "Najczęściej zamawiany produkt"
    wsRaporty.Cells(2, 2).Value = maxProdukt
    wsRaporty.Cells(2, 3).Value = maxSprzedaz

    ' Wstaw dane o najmniej zamawianym produkcie
    wsRaporty.Cells(3, 1).Value = "Najmniej zamawiany produkt"
    wsRaporty.Cells(3, 2).Value = minProdukt
    wsRaporty.Cells(3, 3).Value = minSprzedaz

    MsgBox "Raport został wygenerowany!", vbInformation
End Sub



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
``

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
###

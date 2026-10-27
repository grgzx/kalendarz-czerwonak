# Kalendarz wywozu odpadów dla całej Polski - ocena wykonalności

Osobny projekt obok kalendarza Czerwonaka. Odpowiada na pytanie, czy da się
zbudować ogólnopolski kalendarz wywozu odpadów, jakie są formaty danych
wejściowych w poszczególnych gminach, czy dostępny jest rok 2026 czy już 2027,
oraz ile to kosztuje (tokeny, iteracje, praca człowieka).

Wynikiem jest zeszyt `waste-calendar-poland-assessment.xlsx`.

## Wnioski w skrócie

- Polska ma **2 479 gmin** i **37 242 958** mieszkańców (GUS, stan 30.06.2026).
- **Nie ma krajowego rejestru harmonogramów.** Każda gmina publikuje osobno:
  najczęściej PDF na BIP, ale też generatory adresowe, grafiki PNG i DOCX-y.
- Jedna gmina to nie zawsze jeden dokument: od 1 pliku (Skała, Borzęcin) do
  25 plików per ulica (Skierniewice) i 5 zadań operatorskich na 18 dzielnic
  (Warszawa).
- **2026: tak.** Harmonogramy na 2026 są osiągalne w praktyce dla wszystkich
  sprawdzonych gmin, choć nie zawsze w pierwszym wyniku wyszukiwania.
- **2027: praktycznie nie.** Na 27.10.2026 żadna z 16 sprawdzonych gmin nie
  publikowała jeszcze harmonogramu na 2027. Sezon publikacji to XII-II.
  Budowa "teraz" daje ok. 2 miesiące użyteczności i wymaga ponownego zbioru
  danych w I kwartale 2027.
- Pokrycie rośnie szybciej niż liczba gmin: **top 200 gmin = 48% populacji**,
  top 500 = 64%, top 1000 = 80%, top 1500 = 89%.
- Nakład (model w arkuszu `Nakład`, założenia edytowalne):
  top 200 ok. **900 h**, top 1000 ok. **2 400 h**, wszystkie gminy ok.
  **4 900 h** (ok. 35 osobomiesięcy) i ok. **197 mln tokenów**.

## Arkusze

| Arkusz | Co zawiera |
| --- | --- |
| `Podsumowanie` | skala, wniosek, krzywa pokrycia, dostępność 2026/2027, nakład w skrócie |
| `Gminy` | wszystkie 2 479 gmin: TERYT, typ, województwo, powiat, ludność, ranga, udział i udział skumulowany, tier, oraz kolumny jakościowe (format, 2027, liczba dokumentów, URL) |
| `Proba` | 16 gmin sprawdzonych ręcznie, z URL-em i notatką, plus czego próba dowodzi |
| `Formaty` | taksonomia nośników (PDF, skan, HTML, generator, aplikacja, XLSX, DOCX, grafika) z trudnością wdrożenia |
| `Nakład` | model kosztu: żółte komórki edytowalne, reszta to formuły; 5 scenariuszy |
| `Województwa` | liczba gmin i ludność w przekroju województw |
| `Zrodla` | źródła, definicje i zastrzeżenia |

## Czego ten zeszyt nie twierdzi

Kolumny `Harmonogram 2026`, `Format danych`, `2027?` i `Dokumentów / rejonów`
w arkuszu `Gminy` są wypełnione **tylko dla 16 gmin z próby**. Pozostałe
2 463 wiersze mają `do ustalenia`, bo właśnie ich sprawdzenie jest przedmiotem
wyceny. Wypełnienie tej kolumny dla wszystkich gmin bez ich odwiedzenia byłoby
zmyślaniem danych.

Próba jest doborem celowym po randze ludności (co 1/16 listy plus Warszawa),
nie losową próbą statystyczną. Służy kalibracji, nie oszacowaniu odsetka.

## Pliki

```
waste-calendar-poland/
  waste-calendar-poland-assessment.xlsx   gotowy zeszyt
  data/gminy.csv                          rejestr 2 479 gmin
  data/probka.csv                         16 gmin sprawdzonych ręcznie
  data/raw/                               źródłowe tablice GUS
  scripts/extract_gus.py                  GUS XLS -> data/gminy.csv
  scripts/build_workbook.py               CSV -> XLSX
  scripts/build.sh                        pełne odtworzenie zeszytu
```

## Odtworzenie

Wymagania: Python 3, `openpyxl`, `xlrd` (tylko do odczytu `.xls`), opcjonalnie
LibreOffice do przeliczenia formuł.

```bash
python3 -m pip install --user openpyxl xlrd
cd waste-calendar-poland
bash scripts/build.sh
```

`build.sh` uruchamia ekstrakcję GUS, buduje zeszyt i - jeśli dostępny jest
`soffice` - przelicza formuły, żeby plik miał zapisane wartości także w
podglądach, które nie liczą formuł.

## Źródło danych o ludności

GUS, "Ludność. Stan i struktura ludności oraz ruch naturalny w przekroju
terytorialnym w 2026 r. Stan w dniu 30 czerwca", tablica III (gminy) i
tablica II (powiaty).

Plik źródłowy i suma kontrolna: `data/raw/ZRODLA.md`. Archiwum ZIP (9,9 MB)
nie jest wersjonowane; pobiera się je z URL-a podanego w `ZRODLA.md`.

Kontrola spójności: suma ludności gmin równa się 37 242 958, czyli dokładnie
liczbie dla Polski z tej samej publikacji.

# Źródła danych

## Ludność i rejestr gmin

Publikacja: GUS, "Ludność. Stan i struktura ludności oraz ruch naturalny w
przekroju terytorialnym w 2026 r. Stan w dniu 30 czerwca".

- Archiwum ZIP: https://publikacje.new.stat.gov.pl/sites/default/files/2026-10/Ludnosc_Stan_i_struktura_30.06.2026.zip
- SHA-256 archiwum: `e2dbf0869c9423f4986b8994eb39c8b45ead0fc09d025df8f570bd8563c03741`
- Pobrano: 2026-10-27

Wykorzystane tablice:

| Plik | Zawartość | Użycie |
| --- | --- | --- |
| `Tabela_III.xls` | ludność, ruch naturalny i migracje według gmin | TERYT, nazwa, typ, ludność |
| `Tabela_II.xls` | to samo według powiatów | nazwy powiatów |
| `Lista tablic.xlsx` | spis tablic publikacji | identyfikacja tablic |

Archiwum ZIP nie jest wersjonowane (9,9 MB). Rozpakuj je do `data/raw/`
zachowując nazwy plików, albo pobierz ponownie z URL-a powyżej.

Kontrola: 2 479 gmin, suma ludności 37 242 958, zgodna z liczbą dla Polski.

## Próba badawcza

`data/probka.csv` zawiera 16 gmin sprawdzonych ręcznie 2026-10-27. Dobór
celowy: co 1/16 pozycji listy posortowanej malejąco po ludności, plus
Warszawa jako punkt odniesienia. Dla każdej gminy podano URL faktycznie
otwartego źródła (urzędowego lub operatora), format nośnika i lata, których
dotyczy dokument.

Ograniczenia:

- to kalibracja, nie sondaż losowy;
- nie sprawdzono praw do redystrybucji danych ani regulaminów BIP;
- liczba dokumentów jest przybliżona tam, gdzie gmina dzieli się na rejony
  albo publikuje pliki per ulica.

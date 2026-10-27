#!/usr/bin/env python3
"""Build the Poland-wide waste-calendar feasibility workbook.

Inputs:
  data/gminy.csv   - register (TERYT, name, type, voivodeship, powiat, population)
  data/probka.csv  - verified sample rows (optional; see scripts/README)

Output:
  waste-calendar-poland-assessment.xlsx

All population figures come from GUS, "Ludność. Stan i struktura ludności oraz
ruch naturalny w przekroju terytorialnym w 2026 r. Stan w dniu 30 czerwca"
(table III). The sample is a small, explicitly-labelled empirical check, not a
survey of all 2 479 gminas.
"""

import csv
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
GMINY = os.path.join(ROOT, 'data', 'gminy.csv')
PROBKA = os.path.join(ROOT, 'data', 'probka.csv')
OUT = os.path.join(ROOT, 'waste-calendar-poland-assessment.xlsx')

DATA_AS_OF = '2026-06-30'
BUILD_DATE = '2026-10-27'

# --- styling ---------------------------------------------------------------
H_FILL = PatternFill('solid', fgColor='1F3864')
H_FONT = Font(bold=True, color='FFFFFF', size=10)
SUB_FILL = PatternFill('solid', fgColor='D9E2F3')
TITLE_FONT = Font(bold=True, size=14)
SMALL = Font(size=9)
GREY = Font(size=9, color='666666')
THIN = Side(style='thin', color='BFBFBF')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='top')


def style_header(ws, row, ncols, first=1):
    for c in range(first, first + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        cell.border = BOX


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


def load_rows():
    rows = list(csv.DictReader(open(GMINY, encoding='utf-8')))
    for r in rows:
        r['ludnosc'] = int(r['ludnosc'])
    rows.sort(key=lambda r: -r['ludnosc'])
    total = sum(r['ludnosc'] for r in rows)
    cum = 0
    for i, r in enumerate(rows, 1):
        r['rank'] = i
        r['share'] = r['ludnosc'] / total
        cum += r['ludnosc']
        r['cum'] = cum / total
    return rows, total


def tier(pop):
    if pop >= 100_000:
        return 'T1 100k+'
    if pop >= 50_000:
        return 'T2 50-100k'
    if pop >= 20_000:
        return 'T3 20-50k'
    if pop >= 10_000:
        return 'T4 10-20k'
    if pop >= 5_000:
        return 'T5 5-10k'
    return 'T6 <5k'


def load_sample():
    if not os.path.exists(PROBKA):
        return {}
    out = {}
    for r in csv.DictReader(open(PROBKA, encoding='utf-8'), delimiter=';'):
        out[r['teryt']] = r
    return out


# --- sheet: Podsumowanie ---------------------------------------------------
def sheet_summary(wb, rows, total, sample):
    ws = wb.create_sheet('Podsumowanie')
    ws.sheet_view.showGridLines = False
    widths(ws, {'A': 3, 'B': 44, 'C': 14, 'D': 14, 'E': 14, 'F': 14, 'G': 14, 'H': 14, 'I': 14})

    ws['B1'] = 'Kalendarz wywozu odpadów dla całej Polski - ocena wykonalności'
    ws['B1'].font = TITLE_FONT
    ws['B2'] = f'Stan danych: {DATA_AS_OF} (GUS)  |  opracowanie: {BUILD_DATE}'
    ws['B2'].font = GREY

    r = 4
    ws.cell(row=r, column=2, value='Skala').font = Font(bold=True, size=11)
    r += 1
    facts = [
        ('Gminy (GUS, 30.06.2026)', f'{len(rows):,}'.replace(',', ' ')),
        ('Ludność Polski wg sumy gmin', f'{total:,}'.replace(',', ' ')),
        ('Gminy miejskie / wiejskie / miejsko-wiejskie', '302 / 1453 / 724'),
        ('Powiaty / województwa', '380 / 16'),
        ('Gminy ≥ 50 tys. mieszkańców', f'{sum(1 for x in rows if x["ludnosc"] >= 50000)}'),
    ]
    for label, value in facts:
        ws.cell(row=r, column=2, value=label).font = SMALL
        ws.cell(row=r, column=3, value=value).font = SMALL
        r += 1

    r += 1
    ws.cell(row=r, column=2, value='Wniosek').font = Font(bold=True, size=11)
    r += 1
    verdict = (
        'Wykonalne, ale nie jako ręcznie przygotowana baza 2 479 gmin. Nie ma krajowego rejestru '
        'harmonogramów: każda gmina publikuje osobny dokument (najczęściej PDF na BIP), część dzieli '
        'gminę na rejony, część ma generator adresowy, a duże miasta rozdzielają obsługę między '
        'operatorów. Realny model to program etapowy: wspólny silnik + rodziny szablonów PDF + '
        'automatyczna walidacja, z pokryciem rosnącym według liczby mieszkańców, a nie liczby gmin.'
    )
    ws.cell(row=r, column=2, value=verdict).alignment = WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 4, end_column=6)
    r += 6

    ws.cell(row=r, column=2, value='Krzywa pokrycia (ile gmin daje ile populacji)').font = Font(bold=True, size=11)
    r += 1
    head = ['Zakres', 'Gminy', 'Ludność', '% populacji', 'Uwaga']
    for c, h in enumerate(head, 2):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, r, 5, first=2)
    ws.cell(row=r, column=2).value = 'Zakres'
    r += 1
    ladder = [
        (50, 'największe miasta'),
        (100, 'rdzeń metropolitalny'),
        (200, 'połowa Polski'),
        (500, 'dwie trzecie Polski'),
        (1000, 'cztery piąte Polski'),
        (1500, 'ok. 90%'),
        (2000, 'ok. 96%'),
        (len(rows), 'pełne pokrycie'),
    ]
    for n, note in ladder:
        sub = rows[:n]
        pop = sum(x['ludnosc'] for x in sub)
        ws.cell(row=r, column=2, value=f'top {n}').font = SMALL
        ws.cell(row=r, column=3, value=n).font = SMALL
        ws.cell(row=r, column=4, value=pop).font = SMALL
        ws.cell(row=r, column=4).number_format = '#,##0'
        ws.cell(row=r, column=5, value=pop / total).font = SMALL
        ws.cell(row=r, column=5).number_format = '0.0%'
        ws.cell(row=r, column=6, value=note).font = SMALL
        r += 1

    r += 1
    ws.cell(row=r, column=2, value='Dostępność 2026 / 2027 (na 27.10.2026)').font = Font(bold=True, size=11)
    r += 1
    for text in [
        '2026: publikowana i osiągalna dla znakomitej większości gmin (potwierdzone w próbie).',
        '2027: praktycznie jeszcze nie ma. Gminy publikują harmonogramy na kolejny rok zwykle od połowy '
        'grudnia do lutego; część robi to dopiero w styczniu, a część nadpisuje ten sam plik/ stronę.',
        'Wniosek: budowa "teraz" da 2 miesiące użyteczności (X-XII 2026) i wymusi ponowne zebranie '
        'danych w I kw. 2027. Lepiej zbudować maszynerię teraz i zebrać sezon 2027 w oknie XII-II.',
    ]:
        ws.cell(row=r, column=2, value='• ' + text).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    ws.cell(row=r, column=2, value='Próba badawcza').font = Font(bold=True, size=11)
    r += 1
    pct = f'{len(sample) / len(rows) * 100:.1f}'.replace('.', ',')
    ws.cell(row=r, column=2, value=(
        f'Sprawdzono {len(sample)} gmin ({pct}% zbioru) w podziale na wielkość i typ. '
        'Wyniki są w arkuszu "Proba" - to kalibracja, nie inwentaryzacja. Pełna kolumna '
        '"format" w arkuszu "Gminy" pozostaje nieuzupełniona dla gmin spoza próby.'
    )).alignment = WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=6)
    r += 4

    ws.cell(row=r, column=2, value='Nakład w skrócie (założenia edytowalne w arkuszu "Nakład")').font = Font(bold=True, size=11)
    r += 1
    heads = ['Scenariusz', 'Gminy', '% populacji', 'Godziny', 'Tokeny',
             'Iteracje', 'Miesiące (1 os.)', 'Koszt PLN']
    for c, h in enumerate(heads, 2):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, r, len(heads), first=2)
    r += 1
    for src_row, label in [(21, 'A. Top 50'), (22, 'B. Top 200'), (23, 'C. Top 500'),
                           (24, 'D. Top 1000'), (25, 'E. Wszystkie gminy')]:
        vals = [
            label,
            f"='Nakład'!C{src_row}",
            f"='Nakład'!D{src_row}",
            f"='Nakład'!H{src_row}",
            f"='Nakład'!I{src_row}",
            f"='Nakład'!J{src_row}",
            f"='Nakład'!K{src_row}",
            f"='Nakład'!L{src_row}",
        ]
        for c, v in enumerate(vals, 2):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = SMALL
            cell.border = BOX
        ws.cell(row=r, column=3).number_format = '#,##0'
        ws.cell(row=r, column=4).number_format = '0.0%'
        ws.cell(row=r, column=5).number_format = '#,##0'
        ws.cell(row=r, column=6).number_format = '#,##0'
        ws.cell(row=r, column=7).number_format = '#,##0'
        ws.cell(row=r, column=8).number_format = '0.0'
        ws.cell(row=r, column=9).number_format = '#,##0'
        r += 1

    r += 1
    ws.cell(row=r, column=2, value='Rekomendacja').font = Font(bold=True, size=11)
    r += 1
    for line in [
        'Zacznij od top 200 (48% populacji, ok. 900 h). To zakres, w którym da się utrzymać '
        'jakość i wiarygodność danych jednym zespołem.',
        'Zbieraj dane sezonowo: pełne odświeżenie w XII-II, gdy wychodzą harmonogramy na kolejny rok.',
        'Buduj najpierw wspólny silnik i testy regresyjne, dopiero potem rodziny szablonów - inaczej '
        'każda gmina zostaje jednorazowym skryptem.',
        'Nie obiecuj pokrycia "całej Polski" przed ukończeniem etapu D (top 1000, 80% populacji).',
    ]:
        ws.cell(row=r, column=2, value='• ' + line).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 26
        r += 1
    return ws


# --- sheet: Gminy ----------------------------------------------------------
def sheet_gminy(wb, rows, sample):
    ws = wb.create_sheet('Gminy')
    headers = [
        'TERYT', 'Gmina', 'Typ', 'Województwo', 'Powiat', 'Ludność',
        '% populacji PL', '% skumulowany', 'Ranking', 'Tier',
        'Harmonogram 2026', 'Format danych', '2027?', 'Dokumentów / rejonów',
        'URL źródła', 'Uwagi', 'Status weryfikacji',
    ]
    ws.append(headers)
    style_header(ws, 1, len(headers))

    for x in rows:
        s = sample.get(x['teryt'], {})
        ws.append([
            x['teryt'], x['gmina'], x['typ'], x['wojewodztwo'], x['powiat'], x['ludnosc'],
            x['share'], x['cum'], x['rank'], tier(x['ludnosc']),
            s.get('found', 'do ustalenia'),
            s.get('format', 'do ustalenia'),
            s.get('lata', 'do ustalenia'),
            s.get('dokumenty', ''),
            s.get('url', ''),
            s.get('uwagi', ''),
            'próba' if s else 'niezbadane',
        ])

    last = ws.max_row
    for row in ws.iter_rows(min_row=2, max_row=last, max_col=len(headers)):
        for cell in row:
            cell.font = SMALL
            cell.border = BOX
        row[5].number_format = '#,##0'
        row[6].number_format = '0.00%'
        row[7].number_format = '0.0%'
        row[8].alignment = CENTER
        row[9].alignment = CENTER
    widths(ws, {
        'A': 9, 'B': 24, 'C': 16, 'D': 20, 'E': 22, 'F': 10,
        'G': 11, 'H': 12, 'I': 8, 'J': 12,
        'K': 15, 'L': 14, 'M': 11, 'N': 12, 'O': 46, 'P': 30, 'Q': 14,
    })
    ws.freeze_panes = 'C2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(headers))}{last}'

    dv = DataValidation(
        type='list',
        formula1='"tak,nie,niejasne,do ustalenia"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    dv.add(f'K2:K{last}')

    dv2 = DataValidation(
        type='list',
        formula1='"PDF-tekst,PDF-skan,HTML,generator,aplikacja,XLSX,DOCX,grafika,mieszany,brak,do ustalenia"',
        allow_blank=True,
    )
    ws.add_data_validation(dv2)
    dv2.add(f'L2:L{last}')
    return ws


# --- sheet: Proba ----------------------------------------------------------
def sheet_sample(wb, sample, rows_by_teryt):
    ws = wb.create_sheet('Proba')
    ws.sheet_view.showGridLines = False
    ws['B1'] = 'Próba badawcza: format i dostępność harmonogramów'
    ws['B1'].font = TITLE_FONT
    ws['B2'] = (
        'Dobór celowy: co 1/16 pozycji listy gmin posortowanej malejąco po ludności '
        '(plus Warszawa jako punkt odniesienia). Cel: kalibracja, nie oszacowanie odsetka.'
    )
    ws['B2'].font = GREY
    ws['B2'].alignment = WRAP
    ws.merge_cells('B2:N2')

    headers = ['TERYT', 'Gmina', 'Typ', 'Ludność', 'Ranking', '2026?', 'Format',
               '2027?', 'Dokumentów', 'URL', 'Uwagi']
    row0 = 4
    for c, h in enumerate(headers, 2):
        ws.cell(row=row0, column=c, value=h)
    style_header(ws, row0, 11, first=2)

    r = row0 + 1
    for teryt, s in sorted(sample.items(), key=lambda kv: rows_by_teryt[kv[0]]['rank']):
        x = rows_by_teryt[teryt]
        vals = [teryt, x['gmina'], x['typ'], x['ludnosc'], x['rank'],
                s.get('found', ''), s.get('format', ''), s.get('lata', ''),
                s.get('dokumenty', ''), s.get('url', ''), s.get('uwagi', '')]
        for c, v in enumerate(vals, 2):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = SMALL
            cell.border = BOX
            cell.alignment = WRAP
        ws.cell(row=r, column=5).number_format = '#,##0'
        ws.cell(row=r, column=10).font = Font(size=8, color='0563C1', underline='single')
        if s.get('url'):
            ws.cell(row=r, column=11).hyperlink = s['url']
        r += 1

    r += 1
    ws.cell(row=r, column=2, value='Czego próba dowodzi').font = Font(bold=True, size=11)
    r += 1
    findings = [
        'Format nie jest jednolity: obok PDF-ów tekstowych trafiają się skany, tabele HTML, '
        'generatory adresowe i podział na operatorów.',
        'Jedna gmina to nie zawsze jeden dokument - duże miasta mają osobne harmonogramy na '
        'dzielnice/rejony i różnych operatorów.',
        'Dla 2026 dokumenty są na ogół osiągalne, ale nie zawsze w pierwszym wyniku wyszukiwania; '
        'BIP bywa jedynym źródłem, a strona gminy potrafi pokazywać poprzedni rok (Dynów).',
        'Nośnik bywa zaskakujący: grafiki PNG (Rejowiec Fabryczny) i DOCX (Krempna) obok PDF-ów '
        'i generatorów.',
        '2027 w praktyce jeszcze nie występuje (pojedyncze wyjątki), co potwierdza sezonowość '
        'publikacji XII-II.',
    ]
    for f in findings:
        ws.cell(row=r, column=2, value='• ' + f).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
        ws.row_dimensions[r].height = 28
        r += 1

    widths(ws, {'A': 2, 'B': 9, 'C': 22, 'D': 16, 'E': 9, 'F': 8, 'G': 13,
                'H': 11, 'I': 11, 'J': 11, 'K': 52, 'L': 34})
    return ws


# --- sheet: Formaty --------------------------------------------------------
def sheet_formats(wb, sample):
    ws = wb.create_sheet('Formaty')
    ws.sheet_view.showGridLines = False
    ws['B1'] = 'Formaty danych wejściowych - taksonomia i trudność'
    ws['B1'].font = TITLE_FONT
    ws['B2'] = 'Udział w próbie jest orientacyjny (mała próba). "Trudność" = koszt jednorazowego wdrożenia obsługi.'
    ws['B2'].font = GREY

    headers = ['Format', 'Co to znaczy', 'Jak parsować', 'Trudność', 'Gminy w próbie']
    for c, h in enumerate(headers, 2):
        ws.cell(row=4, column=c, value=h)
    style_header(ws, 4, 5, first=2)

    counts = {}
    for s in sample.values():
        counts[s.get('format', 'brak')] = counts.get(s.get('format', 'brak'), 0) + 1

    rows = [
        ('PDF-tekst', 'PDF z warstwą tekstową, tabela miesięcy x frakcje',
         'pdfjs/xpdf + współrzędne fragmentów (jak w Czerwonaku)', 'średnia', counts.get('PDF-tekst', 0)),
        ('PDF-skan', 'obraz/scan, brak warstwy tekstowej',
         'OCR (Tesseract/vision) + walidacja człowieka', 'wysoka', counts.get('PDF-skan', 0)),
        ('HTML', 'tabela na stronie BIP/urzędu',
         'selektor DOM / regex na HTML; stabilność zależy od CMS', 'średnia', counts.get('HTML', 0)),
        ('generator', 'formularz adresowy generujący harmonogram per nieruchomość',
         'trzeba odtworzyć API/endpoint i iterować po adresach', 'bardzo wysoka', counts.get('generator', 0)),
        ('aplikacja', 'dedykowana aplikacja/mapa (często operatora)',
         'brak publicznego API; często tylko scraping lub umowa', 'bardzo wysoka', counts.get('aplikacja', 0)),
        ('XLSX', 'arkusz do pobrania',
         'SheetJS/openpyxl; najłatwiejszy, ale rzadki', 'niska', counts.get('XLSX', 0)),
        ('grafika', 'obraz PNG/JPG: zdjęcie lub grafika z terminami',
         'OCR albo model wizyjny + walidacja człowieka; brak warstwy tekstowej',
         'wysoka', counts.get('grafika', 0)),
        ('DOCX', 'plik Word z tabelami',
         'python-docx; technicznie łatwy, ale rzadki i niestabilny', 'niska', counts.get('DOCX', 0)),
        ('mieszany', 'np. PDF + generator, albo PDF per rejon + mapa',
         'kombinacja powyższych; trzeba wybrać źródło kanoniczne', 'wysoka', counts.get('mieszany', 0)),
        ('brak', 'nie znaleziono publicznego dokumentu',
         'kontakt do urzędu / wniosek o informację publiczną', 'wysoka', counts.get('brak', 0)),
    ]
    r = 5
    for name, what, how, diff, n in rows:
        for c, v in enumerate([name, what, how, diff, n], 2):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = SMALL
            cell.border = BOX
            cell.alignment = WRAP
        ws.row_dimensions[r].height = 34
        r += 1

    r += 2
    ws.cell(row=r, column=2, value='Czego nie widać w formacie').font = Font(bold=True, size=11)
    r += 1
    for note in [
        'Podział wewnętrzny: gmina może mieć 1 dokument albo 5-20 rejonów/dzielnic.',
        'Strefy: część gmin rozdziela obsługę między operatorów (każdy własny harmonogram).',
        'Nieruchomości: harmonogramy dla domów jednorodzinnych i dla zabudowy wielorodzinnej bywają różne.',
        'Zmienność roczna: plik bywa nadpisywany pod tym samym adresem, co utrudnia wykrywanie zmian.',
        'Aktualność: opublikowany dokument nie znaczy obowiązujący - potrzebna kontrola roku w treści, '
        'nie tylko w adresie.',
    ]:
        ws.cell(row=r, column=2, value='• ' + note).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        r += 1

    widths(ws, {'A': 2, 'B': 14, 'C': 44, 'D': 46, 'E': 13, 'F': 12})
    return ws


# --- sheet: Naklad ---------------------------------------------------------
def sheet_effort(wb, rows, total):
    ws = wb.create_sheet('Nakład')
    ws.sheet_view.showGridLines = False
    ws['B1'] = 'Model nakładu (edytowalny - zmień wartości w kolumnie C)'
    ws['B1'].font = TITLE_FONT
    ws['B2'] = (
        'Założenia są jawnie niepewne (rzeczywistość ±50%). Punkt odniesienia: Czerwonak - '
        'jedna gmina, jeden szablon, ok. 1 500 linii kodu i pełna sesja z weryfikacją.'
    )
    ws['B2'].font = GREY
    ws['B2'].alignment = WRAP
    ws.merge_cells('B2:L2')
    ws.row_dimensions[2].height = 28

    # key -> (label, initial value).  Row numbers are derived, never hard-coded.
    params = [
        ('family', 'Koszt wdrożenia nowej rodziny szablonów (h)', 20),
        ('templated', 'Koszt gminy po istniejącym szablonie (h)', 0.5),
        ('unique', 'Koszt gminy o unikalnym formacie (h)', 2.5),
        ('share', 'Udział gmin pokrytych szablonami', 0.5),
        ('refresh', 'Koszt odświeżenia rocznego na gminę (h)', 0.3),
        ('tok_family', 'Tokeny na nową rodzinę szablonów', 400000),
        ('tok_templated', 'Tokeny na gminę po szablonie', 20000),
        ('tok_unique', 'Tokeny na gminę unikalną', 120000),
        ('it_family', 'Iteracje agenta na nową rodzinę szablonów', 20),
        ('it_templated', 'Iteracje agenta na gminę po szablonie', 3),
        ('it_unique', 'Iteracje agenta na gminę unikalną', 10),
        ('rate', 'Stawka godzinowa (PLN/h)', 120),
        ('hours_month', 'Godziny pracy na osobę / miesiąc', 140),
    ]
    r = 4
    for c, title in ((2, 'Parametr'), (3, 'Wartość')):
        cell = ws.cell(row=r, column=c, value=title)
        cell.font = H_FONT
        cell.fill = H_FILL
        cell.border = BOX
    r += 1
    cell_row = {}
    for key, label, val in params:
        ws.cell(row=r, column=2, value=label).font = SMALL
        c = ws.cell(row=r, column=3, value=val)
        c.font = SMALL
        c.fill = PatternFill('solid', fgColor='FFF2CC')
        c.border = BOX
        cell_row[key] = r
        r += 1
    ref = {k: f'$C${v}' for k, v in cell_row.items()}

    r += 1
    ws.cell(row=r, column=2, value='Scenariusze').font = Font(bold=True, size=11)
    r += 1
    headers = ['Scenariusz', 'Gminy', '% populacji', 'Rodziny szablonów',
               'Gminy w szablonach', 'Gminy unikalne', 'Godziny', 'Tokeny',
               'Iteracje agenta', 'Miesiące (1 os.)', 'Koszt PLN',
               'Odświeżenie / rok (h)']
    for c, h in enumerate(headers, 2):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, r, len(headers), first=2)
    r += 1

    scenarios = [
        ('A. Top 50 (największe miasta)', 50, 25),
        ('B. Top 200', 200, 30),
        ('C. Top 500', 500, 38),
        ('D. Top 1000', 1000, 45),
        ('E. Wszystkie gminy', len(rows), 60),
    ]
    for label, n, families in scenarios:
        pop = sum(x['ludnosc'] for x in rows[:n])
        values = {
            2: label,
            3: n,
            4: pop / total,
            5: families,
            6: f'=ROUND({ref["share"]}*C{r},0)',
            7: f'=C{r}-F{r}',
            8: f'=E{r}*{ref["family"]}+F{r}*{ref["templated"]}+G{r}*{ref["unique"]}',
            9: f'=E{r}*{ref["tok_family"]}+F{r}*{ref["tok_templated"]}+G{r}*{ref["tok_unique"]}',
            10: f'=E{r}*{ref["it_family"]}+F{r}*{ref["it_templated"]}+G{r}*{ref["it_unique"]}',
            11: f'=H{r}/{ref["hours_month"]}',
            12: f'=H{r}*{ref["rate"]}',
            13: f'=C{r}*{ref["refresh"]}',
        }
        for col, val in values.items():
            cell = ws.cell(row=r, column=col, value=val)
            cell.font = SMALL
            cell.border = BOX
        ws.cell(row=r, column=4).number_format = '0.0%'
        ws.cell(row=r, column=6).number_format = '#,##0'
        ws.cell(row=r, column=7).number_format = '#,##0'
        ws.cell(row=r, column=8).number_format = '#,##0'
        ws.cell(row=r, column=9).number_format = '#,##0'
        ws.cell(row=r, column=10).number_format = '#,##0'
        ws.cell(row=r, column=11).number_format = '0.0'
        ws.cell(row=r, column=12).number_format = '#,##0'
        ws.cell(row=r, column=13).number_format = '#,##0'
        r += 1

    r += 1
    ws.cell(row=r, column=2, value=(
        'Godziny = rodziny x koszt rodziny + gminy w szablonach x koszt po szablonie + gminy '
        'unikalne x koszt unikalny. Tokeny i iteracje analogicznie. "Miesiące (1 os.)" = godziny '
        '/ godziny na miesiąc. Żółte komórki są edytowalne, reszta to formuły.'
    )).alignment = WRAP
    ws.merge_cells(start_row=r, start_column=2, end_row=r + 2, end_column=13)
    r += 4

    ws.cell(row=r, column=2, value='Czego model nie obejmuje').font = Font(bold=True, size=11)
    r += 1
    for note in [
        'Hosting, domena, UX, prawo (RODO, regulaminy), wsparcie użytkowników, obsługa reklamacji.',
        'Odświeżanie roczne: budżet ok. 25% nakładu startowego, skoncentrowany w XII-II.',
        'Ryzyko: gminy zmieniają układ PDF-ów bez ostrzeżenia; bez testów regresyjnych cichy błąd '
        'trafia do kalendarzy użytkowników.',
        'Efekt sieciowy: 90% populacji wymaga ok. 1 549 gmin, ale 80% już przy 1 018 - argument za '
        'etapem "top 1000" jako produktem, nie za pełnym zbiorem.',
        'Liczba "dokumentów" nie jest kosztem liniowym: gmina z 25 plikami per ulica to zwykle jeden '
        'szablon, ale wymaga przejścia po wszystkich plikach.',
    ]:
        ws.cell(row=r, column=2, value='• ' + note).alignment = WRAP
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
        ws.row_dimensions[r].height = 26
        r += 1

    widths(ws, {'A': 2, 'B': 34, 'C': 10, 'D': 12, 'E': 12, 'F': 13,
                'G': 12, 'H': 11, 'I': 12, 'J': 12, 'K': 12, 'L': 13, 'M': 14})
    return ws


# --- sheet: Województwa ----------------------------------------------------
def sheet_voivodeships(wb, rows):
    ws = wb.create_sheet('Województwa')
    headers = ['Województwo', 'Gminy', 'Ludność', '% populacji PL', 'Gminy miejskie', 'miejsko-wiejskie', 'wiejskie']
    ws.append(headers)
    style_header(ws, 1, len(headers))
    total = sum(x['ludnosc'] for x in rows)
    agg = {}
    for x in rows:
        a = agg.setdefault(x['wojewodztwo'], {'n': 0, 'pop': 0, 'm': 0, 'mw': 0, 'w': 0})
        a['n'] += 1
        a['pop'] += x['ludnosc']
        a['m'] += x['typ'] == 'miejska'
        a['mw'] += x['typ'] == 'miejsko-wiejska'
        a['w'] += x['typ'] == 'wiejska'
    r = 2
    for name in sorted(agg):
        a = agg[name]
        for c, v in enumerate([name, a['n'], a['pop'], a['pop'] / total, a['m'], a['mw'], a['w']], 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = SMALL
            cell.border = BOX
        ws.cell(row=r, column=3).number_format = '#,##0'
        ws.cell(row=r, column=4).number_format = '0.0%'
        r += 1
    widths(ws, {'A': 24, 'B': 8, 'C': 12, 'D': 12, 'E': 13, 'F': 15, 'G': 10})
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:G{r-1}'
    return ws


# --- sheet: Zrodla ---------------------------------------------------------
def sheet_sources(wb, sample):
    ws = wb.create_sheet('Zrodla')
    ws.sheet_view.showGridLines = False
    ws['B1'] = 'Źródła, definicje i zastrzeżenia'
    ws['B1'].font = TITLE_FONT
    r = 3
    blocks = [
        ('Ludność i rejestr gmin', [
            'GUS, "Ludność. Stan i struktura ludności oraz ruch naturalny w przekroju terytorialnym '
            'w 2026 r. Stan w dniu 30 czerwca" - tablica III (gminy) i tablica II (powiaty).',
            'Plik: https://publikacje.new.stat.gov.pl/sites/default/files/2026-10/Ludnosc_Stan_i_struktura_30.06.2026.zip',
            'Dane w arkuszu: stan na 30.06.2026, 2 479 gmin, suma 37 242 958 osób (zgodna z liczbą dla Polski).',
            'TERYT: 7 znaków; ostatnia cyfra to typ gminy (1 miejska, 2 wiejska, 3 miejsko-wiejska, '
            '4/5 części gminy miejsko-wiejskiej - pominięte).',
        ]),
        ('Taksonomia formatów', [
            'Format opisuje źródło, z którego da się odczytać daty - nie sam wygląd strony.',
            '"generator" i "aplikacja" to przypadki, w których nie ma gotowego dokumentu do pobrania.',
            'Ocena trudności jest inżynierska (koszt wdrożenia), nie prawna.',
        ]),
        ('Ograniczenia tej oceny', [
            f'Próba: {len(sample)} gmin (dobór celowy po rankingu ludności). To kalibracja, nie sondaż.',
            'Kolumny "format", "2027", "dokumenty" w arkuszu "Gminy" są wypełnione tylko dla próby.',
            'Nie sprawdzono praw do redystrybucji danych ani regulaminów poszczególnych gmin.',
            'Szacunki nakładu są +/- 50%; nie zawierają kosztów prawnych, hostingu ani wsparcia.',
        ]),
        ('Punkt odniesienia', [
            'Czerwonak (gmina wiejska, 28 313 osób): 12 rejonów, 79 plików .ics, pipeline '
            'PDF -> JSON -> ICS, testy regresyjne, ok. 1 500 linii kodu.',
            'Jedna gmina = jedna sesja inżynierska. To jest jednostka kosztu użyta w modelu.',
        ]),
    ]
    for title, lines in blocks:
        ws.cell(row=r, column=2, value=title).font = Font(bold=True, size=11)
        r += 1
        for line in lines:
            ws.cell(row=r, column=2, value='• ' + line).alignment = WRAP
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
            ws.row_dimensions[r].height = 28
            r += 1
        r += 1
    widths(ws, {'A': 2, 'B': 24, 'C': 18, 'D': 18, 'E': 18, 'F': 18, 'G': 18, 'H': 18, 'I': 18})
    return ws


def main():
    rows, total = load_rows()
    sample = load_sample()
    by_teryt = {x['teryt']: x for x in rows}

    wb = Workbook()
    wb.remove(wb.active)
    sheet_summary(wb, rows, total, sample)
    sheet_gminy(wb, rows, sample)
    sheet_sample(wb, sample, by_teryt)
    sheet_formats(wb, sample)
    sheet_effort(wb, rows, total)
    sheet_voivodeships(wb, rows)
    sheet_sources(wb, sample)

    wb.properties.title = 'Kalendarz wywozu odpadów dla całej Polski - ocena wykonalności'
    wb.properties.creator = 'analiza wewnętrzna'
    wb.save(OUT)
    print(f'zapisano {OUT}')
    print(f'gminy={len(rows)} proba={len(sample)}')


if __name__ == '__main__':
    main()

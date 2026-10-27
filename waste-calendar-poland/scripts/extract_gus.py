#!/usr/bin/env python3
"""Extract the gmina register (TERYT, type, population) from GUS XLS tables.

Input (data/raw/, from the GUS publication "Ludność. Stan i struktura ludności
oraz ruch naturalny w przekroju terytorialnym w 2026 r. Stan w dniu 30 czerwca"):

  Tabela_III.xls - population by gmina (with TERYT id and type prefix)
  Tabela_II.xls  - population by powiat (used for powiat names)

Output: data/gminy.csv

The XLS files store some TERYT ids as numbers, so leading zeros are lost; the
id is rebuilt as a zero-padded 7-digit string.
"""

import csv
import os
import sys

import xlrd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RAW = os.path.join(ROOT, 'data', 'raw')
OUT = os.path.join(ROOT, 'data', 'gminy.csv')

# TERYT gmina type digit (7th character)
TYPES = {
    '1': 'miejska',
    '2': 'wiejska',
    '3': 'miejsko-wiejska',
    '4': 'miasto (część gminy miejsko-wiejskiej)',
    '5': 'obszar wiejski (część gminy miejsko-wiejskiej)',
}

VOIVODESHIPS = {
    '02': 'dolnośląskie',
    '04': 'kujawsko-pomorskie',
    '06': 'lubelskie',
    '08': 'lubuskie',
    '10': 'łódzkie',
    '12': 'małopolskie',
    '14': 'mazowieckie',
    '16': 'opolskie',
    '18': 'podkarpackie',
    '20': 'podlaskie',
    '22': 'pomorskie',
    '24': 'śląskie',
    '26': 'świętokrzyskie',
    '28': 'warmińsko-mazurskie',
    '30': 'wielkopolskie',
    '32': 'zachodniopomorskie',
}


def cell_id(sheet, row, col):
    """Return the TERYT id in `col` as a string with leading zeros restored."""
    value = sheet.cell_value(row, col)
    if sheet.cell_type(row, col) == 2:  # numeric cell
        if value in ('', None):
            return ''
        return str(int(value)).zfill(7)
    return str(value).strip()


def clean_name(raw):
    """'M-W.Nowy Wiśnicz' -> ('Nowy Wiśnicz', 'miejsko-wiejska')."""
    name = ' '.join(str(raw).split())
    prefix_type = None
    for prefix, typ in (('M-W.', 'miejsko-wiejska'), ('M.', 'miejska'), ('G.', 'wiejska')):
        if name.startswith(prefix):
            prefix_type = typ
            name = name[len(prefix):].strip()
            break
    return name, prefix_type


def load_powiats():
    """TERYT powiat id (4) -> display name."""
    sheet = xlrd.open_workbook(os.path.join(RAW, 'Tabela_II.xls')).sheet_by_index(0)
    out = {}
    for r in range(sheet.nrows):
        pid = cell_id(sheet, r, 1)
        if len(pid) != 4:
            continue
        name = ' '.join(str(sheet.cell_value(r, 0)).split())
        if not name or name.startswith('Woj.'):
            continue
        # cities with powiat rights are listed as "M.Warszawa" and the like
        for prefix in ('M-W.', 'M.', 'G.'):
            if name.startswith(prefix):
                name = name[len(prefix):].strip()
                break
        out[pid] = name
    return out


def main():
    powiats = load_powiats()
    sheet = xlrd.open_workbook(os.path.join(RAW, 'Tabela_III.xls')).sheet_by_index(0)

    rows = []
    for r in range(sheet.nrows):
        gid = cell_id(sheet, r, 1)
        if len(gid) != 7 or not gid.isdigit():
            continue
        name, prefix_type = clean_name(sheet.cell_value(r, 0))
        type_digit = gid[-1]
        typ = TYPES.get(type_digit)
        if typ is None:
            continue
        population = sheet.cell_value(r, 2)
        population = int(population) if isinstance(population, float) else None
        rows.append({
            'teryt': gid,
            'gmina': name,
            'typ': prefix_type or typ,
            'wojewodztwo': VOIVODESHIPS.get(gid[:2], '?'),
            'powiat': powiats.get(gid[:4], '?'),
            'ludnosc': population,
        })

    gminas = [r for r in rows if r['teryt'][-1] in '123']
    parts = [r for r in rows if r['teryt'][-1] in '45']
    gminas.sort(key=lambda r: r['teryt'])

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', newline='', encoding='utf-8') as fh:
        writer = csv.DictWriter(fh, fieldnames=['teryt', 'gmina', 'typ', 'wojewodztwo', 'powiat', 'ludnosc'],
                                lineterminator='\n')
        writer.writeheader()
        writer.writerows(gminas)

    total = sum(g['ludnosc'] or 0 for g in gminas)
    print(f'gminy: {len(gminas)}  (parts skipped: {len(parts)})')
    print(f'ludnosc razem: {total:,}'.replace(',', ' '))
    by_type = {}
    for g in gminas:
        by_type[g['typ']] = by_type.get(g['typ'], 0) + 1
    print('typy:', by_type)
    print('wojewodztwa:', len({g['wojewodztwo'] for g in gminas}))
    print('powiaty:', len({(g['wojewodztwo'], g['powiat']) for g in gminas}))
    print('brak ludnosci:', sum(1 for g in gminas if not g['ludnosc']))
    print('zapisano:', OUT)


if __name__ == '__main__':
    sys.exit(main())

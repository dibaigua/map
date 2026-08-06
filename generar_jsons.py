#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_jsons.py
----------------
Script en Python per generar automàticament els 4 fitxers JSON de dades per als mapes web
a partir del fitxer Excel oficial '00_MUNICIPIS DIBAigua_mapes web.xlsx'.

Ús:
    python generar_jsons.py
    python generar_jsons.py "ruta/al/teu/fitxer.xlsx"
"""

import sys
import os
import json
import zipfile
import datetime
import xml.etree.ElementTree as ET

# Ruta per defecte al fitxer Excel oficial
DEFAULT_EXCEL_PATH = r"C:\Users\canasae\OneDrive - diba.cat\01_DIBAigua\00_Comunicació\Web\00_MUNICIPIS DIBAigua_mapes web.xlsx"

# Data d'actualització per defecte (avui en format DD.MM.YYYY)
TODAY_STR = datetime.datetime.now().strftime("%d.%m.%Y")


def load_geojson_mappings(topojson_path="MunProvBCN.json"):
    """
    Carrega les equivalències de CODIMUNI des del fitxer geogràfic TopoJSON/GeoJSON.
    """
    ine_to_codi = {}
    name_to_codi = {}

    if not os.path.exists(topojson_path):
        if os.path.exists("bcn.topojson"):
            topojson_path = "bcn.topojson"
        else:
            print(f"[Avis] No s'ha trobat {topojson_path}. S'utilitzara mapeig basic d'INE.")
            return ine_to_codi, name_to_codi

    try:
        with open(topojson_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        objects = data.get('objects', {})
        first_key = list(objects.keys())[0] if objects else None

        if first_key and 'geometries' in objects[first_key]:
            for g in objects[first_key]['geometries']:
                props = g.get('properties', {})
                codi = props.get('CODIMUNI')
                nom = props.get('NOMMUNI', '')
                nom_ind = props.get('NOMMUNIIND', '')

                if codi:
                    ine_to_codi[codi] = codi
                    if nom:
                        name_to_codi[nom.lower().strip()] = codi
                    if nom_ind:
                        name_to_codi[nom_ind.lower().strip()] = codi

    except Exception as e:
        print(f"[Error] Carregant el fitxer de coordenades: {e}")

    return ine_to_codi, name_to_codi


def find_codi(ine_str, name_str, ine_to_codi, name_to_codi):
    """
    Cerca el CODIMUNI de 6 dígits corresponent a partir de l'INE de 5 dígits o del nom del municipi.
    """
    ine_clean = str(ine_str).strip()

    # Cerqueu si algun CODIMUNI comença o conté el codi INE de 5 dígits (ex: '08031' -> '080312')
    for k, codi in ine_to_codi.items():
        if k.startswith(ine_clean) or ine_clean in k:
            return codi

    name_clean = str(name_str).lower().strip()
    if name_clean in name_to_codi:
        return name_to_codi[name_clean]

    return None


def read_excel_raw(excel_path):
    """
    Llegeix directament l'estructura XML d'un fitxer .xlsx sense dependències externes.
    """
    z = zipfile.ZipFile(excel_path)

    # Llegir relacions del llibre per associar el nom de la pestanya amb el fitxer XML correcte
    rel_map = {}
    if 'xl/_rels/workbook.xml.rels' in z.namelist():
        rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        for rel in rels_root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            rel_map[rel.attrib['Id']] = 'xl/' + rel.attrib['Target']

    # Llegir cadenes compartides (sharedStrings)
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        ss_root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in ss_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            strings.append(''.join(si.itertext()))

    # Funció auxiliar per extreure valor de cel·la
    def get_cell_val(cell):
        t = cell.attrib.get('t')
        v = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
        if v is None:
            return ''
        val = v.text
        if t == 's' and val and val.isdigit():
            idx = int(val)
            return strings[idx] if idx < len(strings) else val
        return val if val else ''

    # Obtenir relació de pestanyes
    wb_root = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = wb_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')

    sheet_data = {}
    for s in sheets:
        s_name = s.attrib['name']
        r_id = s.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        sheet_file = rel_map.get(r_id, '')

        if sheet_file in z.namelist():
            s_root = ET.fromstring(z.read(sheet_file))
            rows = []
            for r in s_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                row_vals = [get_cell_val(c) for c in r.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')]
                if any(row_vals):
                    rows.append((int(r.attrib.get('r', 0)), row_vals))
            sheet_data[s_name] = rows

    return sheet_data


def normalitzar_estat(raw_stat):
    """
    Mapeja el text de l'Excel a la clau estàndard de color del mapa.
    """
    txt = str(raw_stat).lower().strip()
    if 'curs' in txt:
        return 'en_curs'
    elif 'final' in txt:
        return 'finalitzat'
    elif 'sectorial' in txt:
        return 'programa_sectorial'
    elif 'pilot' in txt:
        return 'prova_pilot'
    else:
        return 'previst'


def generar_jsons(excel_path=None, data_actualitzacio=TODAY_STR):
    """
    Funció principal que processa l'Excel i genera els 4 fitxers JSON a la carpeta data/
    """
    if not excel_path:
        excel_path = DEFAULT_EXCEL_PATH

    if not os.path.exists(excel_path):
        print(f"[Error] No s'ha trobat el fitxer Excel a la ruta:\n   {excel_path}")
        return False

    print(f"[*] Processant fitxer Excel: {excel_path}")
    print(f"[*] Data d'actualitzacio assignada: {data_actualitzacio}")

    ine_to_codi, name_to_codi = load_geojson_mappings()
    sheets_data = read_excel_raw(excel_path)

    os.makedirs('data', exist_ok=True)

    # 1. GENERAR PSA (Pestanya 'PSA')
    psa_json = {}
    sheet_psa_name = [k for k in sheets_data.keys() if 'PSA' in k]
    if sheet_psa_name:
        for r_num, r in sheets_data[sheet_psa_name[0]]:
            if len(r) >= 3 and r[0] and 'MUNICIPIS' not in r[0] and r[0] != 'Municipi':
                name, ine, raw_stat = r[0].strip(), r[1].strip(), r[2].strip()
                codi = find_codi(ine, name, ine_to_codi, name_to_codi)
                if codi and raw_stat:
                    psa_json[codi] = {'estat': normalitzar_estat(raw_stat)}

    out_psa = {
        'data_actualitzacio': data_actualitzacio,
        'data': psa_json
    }
    with open('data/psa.json', 'w', encoding='utf-8') as f:
        json.dump(out_psa, f, indent=2, ensure_ascii=False)
    print(f"[OK] Generat data/psa.json amb {len(psa_json)} municipis")

    # 2. GENERAR TELECONTROL (Pestanya 'Telecontrol' o 'Actuacions')
    tele_json = {}
    sheet_tele_name = [k for k in sheets_data.keys() if 'Telecontrol' in k or 'Actuacions' in k or 'teleco' in k.lower()]
    if sheet_tele_name:
        for r_num, r in sheets_data[sheet_tele_name[0]]:
            if len(r) >= 3 and r[0] and 'MUNICIPIS' not in r[0] and r[0] != 'Municipi':
                name, ine, raw_stat = r[0].strip(), r[1].strip(), r[2].strip()
                codi = find_codi(ine, name, ine_to_codi, name_to_codi)
                if codi and raw_stat:
                    tele_json[codi] = {'estat': normalitzar_estat(raw_stat)}

    out_tele = {
        'data_actualitzacio': data_actualitzacio,
        'data': tele_json
    }
    with open('data/telecontrol.json', 'w', encoding='utf-8') as f:
        json.dump(out_tele, f, indent=2, ensure_ascii=False)
    print(f"[OK] Generat data/telecontrol.json amb {len(tele_json)} municipis")

    # 3. GENERAR TRANSPARÈNCIA (Pestanya 'Transparència')
    trans_json = {}
    sheet_trans_name = [k for k in sheets_data.keys() if 'Transpar' in k]
    if sheet_trans_name:
        for r_num, r in sheets_data[sheet_trans_name[0]]:
            if len(r) >= 3 and r[0] and 'MUNICIPIS' not in r[0] and r[0] != 'Municipi' and 'Qualitat' not in r[0] and 'Transparència' not in r[0]:
                name, ine, raw_stat = r[0].strip(), r[1].strip(), r[2].strip()
                codi = find_codi(ine, name, ine_to_codi, name_to_codi)
                if codi and raw_stat and raw_stat != 'ESTAT':
                    trans_json[codi] = {'estat': normalitzar_estat(raw_stat)}

    out_trans = {
        'data_actualitzacio': data_actualitzacio,
        'data': trans_json
    }
    with open('data/transparencia.json', 'w', encoding='utf-8') as f:
        json.dump(out_trans, f, indent=2, ensure_ascii=False)
    print(f"[OK] Generat data/transparencia.json amb {len(trans_json)} municipis")

    # 4. GENERAR ARTICULACIÓ (Pestanya 'Articulació', 'Projecte' o 'Municipis inclosos')
    art_json = {}
    sheet_art_name = [k for k in sheets_data.keys() if 'Articulac' in k or 'Projecte' in k or 'Municipis inclosos' in k]
    if sheet_art_name:
        muni_count = 0
        for r_num, r in sheets_data[sheet_art_name[0]]:
            if len(r) >= 2 and r[0] and 'MUNICIPIS' not in r[0] and r[0] != 'Municipi':
                name, ine = r[0].strip(), r[1].strip()
                codi = find_codi(ine, name, ine_to_codi, name_to_codi)
                if codi:
                    muni_count += 1
                    # Els primers 8 municipis són Prova pilot, la resta Programa sectorial
                    stat_key = 'prova_pilot' if muni_count <= 8 else 'programa_sectorial'
                    art_json[codi] = {'estat': stat_key}

    out_art = {
        'data_actualitzacio': data_actualitzacio,
        'data': art_json
    }
    with open('data/articulacio.json', 'w', encoding='utf-8') as f:
        json.dump(out_art, f, indent=2, ensure_ascii=False)
    print(f"[OK] Generat data/articulacio.json amb {len(art_json)} municipis")

    print("\n[OK] Proces completat amb exit. Tots els fitxers JSON estan actualitzats!")
    return True


if __name__ == '__main__':
    path_arg = sys.argv[1] if len(sys.argv) > 1 else None
    generar_jsons(path_arg)

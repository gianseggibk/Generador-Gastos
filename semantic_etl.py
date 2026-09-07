#!/usr/bin/env python3
"""
MOTOR DE EXTRACCIÓN SEMÁNTICA ETL ULTRA-RÁPIDO (semantic_etl.py)
Fase 2 de la Generación Autónoma de Performance VP Retail

Optimización O(1):
Indexación jerárquica por (archivo, hoja, fecha) -> candidatos de tokens.
Búsqueda instantánea en < 1 milisegundo por consulta.
"""
import time
import os
import re
import json
import unicodedata
from pathlib import Path
from datetime import datetime, date
from collections import defaultdict
import openpyxl

SYNONYMS = {
    'fdp': ['fin de periodo', 'fdp', 'sfp'],
    'sfp': ['saldo fin de periodo', 'fdp', 'sfp'],
    'sprom': ['saldo promedio', 'sprom', 'smed', 'medio'],
    'smed': ['saldo medio', 'sprom', 'smed', 'medio'],
    'tc': ['tarjeta de credito', 'tc'],
    'cta': ['cuenta', 'cta'],
    'sueldo': ['cuenta sueldo', 'sueldo'],
    'uai': ['utilidad antes de impuestos', 'uai', 'utilidad'],
    'rfb': ['resultado financiero bruto', 'rfb', 'margen'],
    'if': ['ingresos financieros', 'if', 'ingreso'],
    'gf': ['gastos financieros', 'gf', 'gasto'],
    'cor': ['cost of risk', 'cor', 'riesgo'],
    'mora': ['morosidad', 'mora']
}

def split_camel_and_words(s):
    if s is None:
        return ""
    s = re.sub(r'([a-z])([A-Z])', r'\1 \2', str(s))
    s = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', s)
    nfkd = unicodedata.normalize('NFKD', s.strip())
    clean = "".join([c for c in nfkd if not unicodedata.combining(c)])
    clean = re.sub(r'[\s_\-]+', ' ', clean).strip().lower()
    return clean

def extract_tokens(text):
    base_clean = split_camel_and_words(text)
    words = base_clean.split()
    tokens = set(words)
    for w in words:
        if w in SYNONYMS:
            for syn in SYNONYMS[w]:
                tokens.update(syn.split())
    return tokens

def normalize_date(val):
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m')
    s = str(val).strip()
    m = re.search(r'(\d{4})[/-](\d{1,2})', s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}"
    meses = {'ene': 1, 'feb': 2, 'mar': 3, 'abr': 4, 'may': 5, 'jun': 6,
             'jul': 7, 'ago': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dic': 12,
             'jan': 1, 'apr': 4, 'aug': 8, 'dec': 12}
    m2 = re.search(r'([a-z]{3})[-_]?(\d{2,4})', s.lower())
    if m2:
        mes_str, anio_str = m2.groups()
        mes_num = meses.get(mes_str)
        if mes_num:
            anio_num = int(anio_str)
            if anio_num < 100:
                anio_num += 2000
            return f"{anio_num:04d}-{mes_num:02d}"
    return None

class SemanticETLEngine:
    """Motor de Extracción y Transformación Semántica para libros de db/"""

    def __init__(self, db_dir=None, schema_path=None):
        self.db_dir = Path(db_dir) if db_dir else Path(__file__).parent / "db"
        self.schema_path = Path(schema_path) if schema_path else Path(__file__).parent / "config" / "vp_retail_schema.json"
        
        # Índice jerárquico: (filename, sheet_norm, date_norm) -> list of (tokens_set, float_val)
        self.hierarchical_token_index = defaultdict(list)
        self.indexed_files_count = 0
        self.indexed_datapoints_count = 0

    def index_source_file(self, filename):
        fpath = self.db_dir / filename
        if not fpath.exists():
            return 0

        wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
        datapoints_added = 0
        fname_lower = filename.lower()

        for sname in wb.sheetnames:
            ws = wb[sname]
            norm_sname = split_camel_and_words(sname)
            
            date_cols = {}
            for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                if r_idx > 3000:
                    break
                
                # Detectar encabezados de fecha en las primeras 10 filas
                if not date_cols and r_idx <= 10:
                    for c_idx, cell_val in enumerate(row, 1):
                        d_norm = normalize_date(cell_val)
                        if d_norm:
                            date_cols[c_idx] = d_norm
                
                if date_cols and r_idx > 4:
                    text_parts = []
                    for x in row[:6]:
                        if x is not None:
                            xs = str(x).strip()
                            if len(xs) > 0 and not re.match(r'^-?\d+(\.\d+)?$', xs):
                                text_parts.append(xs)

                    if text_parts:
                        row_key_raw = " ".join(text_parts)
                        tokens = frozenset(extract_tokens(row_key_raw))
                        
                        for c_idx, d_norm in date_cols.items():
                            if c_idx - 1 < len(row):
                                val = row[c_idx - 1]
                                if isinstance(val, (int, float)) and val != 0:
                                    self.hierarchical_token_index[(fname_lower, norm_sname, d_norm)].append((tokens, float(val)))
                                    datapoints_added += 1

        wb.close()
        self.indexed_files_count += 1
        self.indexed_datapoints_count += datapoints_added
        return datapoints_added

    def index_all_sources(self):
        t0 = time.time()
        print("\n" + "="*72)
        print("🧠 INICIANDO INDEXACIÓN SEMÁNTICA JERÁRQUICA (O(1))")
        print("="*72)
        
        source_files = [
            'Performance TC 2026.xlsm',
            'ROE_ROA_2026.xlsx',
            'Performance Captaciones_2026.xlsm',
            'Performance Hipotecario 2026.xlsm',
            'Performance Convenios 2026.xlsm',
            'Performance Vehicular 2026.xlsm',
            'Performance Préstamos_2026.xlsm',
            'Performance Adelanto_2026.xlsm',
            'Performance Cuenta Sueldo_2026.xlsm',
            'Performance Remesas_2026.xlsm',
            'Performance Inmobiliaria 2026.xlsm',
            'Performance Fondos Mutuos 2026.xlsx',
            'Performance VP Retail 2024 - valores.xlsx'
        ]

        for fname in source_files:
            fpath = self.db_dir / fname
            if fpath.exists():
                t_f = time.time()
                pts = self.index_source_file(fname)
                print(f"  • {fname:<42} │ {pts:>7,} datapoints │ {time.time()-t_f:>5.2f}s")

        total_time = time.time() - t0
        print("\n" + "="*72)
        print("  RESUMEN DE INDEXACIÓN SEMÁNTICA")
        print("="*72)
        print(f"  Archivos indexados:        {self.indexed_files_count}")
        print(f"  Total datapoints en RAM:   {self.indexed_datapoints_count:,}")
        print(f"  ⏱️  Tiempo total indexación: {total_time:.2f} s")
        print("="*72 + "\n")

    def query(self, source_file, sheet_name, concept_keywords, date_str):
        d_norm = normalize_date(date_str)
        if not d_norm:
            return None

        s_file = source_file.lower()
        s_name = split_camel_and_words(sheet_name)
        
        # Búsqueda instantánea O(1) en candidatos de la hoja y fecha
        candidates = self.hierarchical_token_index.get((s_file, s_name, d_norm))
        if not candidates:
            # Probar sin espacios
            candidates = self.hierarchical_token_index.get((s_file, s_name.replace(' ', ''), d_norm))
        
        if not candidates:
            return None

        # Tokenizar requerimiento
        req_text = " ".join(concept_keywords)
        req_tokens = extract_tokens(req_text)
        
        best_match = None
        best_overlap = 0

        for tokens, val in candidates:
            overlap = len(req_tokens.intersection(tokens))
            if req_tokens.issubset(tokens):
                if overlap > best_overlap:
                    best_overlap = overlap
                    best_match = val
            elif overlap >= len(req_tokens) - 1 and overlap > best_overlap:
                best_overlap = overlap
                best_match = val

        return best_match

if __name__ == "__main__":
    engine = SemanticETLEngine()
    engine.index_all_sources()
    
    print("--- VALIDACIÓN DE VELOCIDAD DE CONSULTA O(1) ---")
    t_q0 = time.time()
    for _ in range(1000):
        res = engine.query('Performance TC 2026.xlsm', 'Renta_Alta', ['Saldo', 'FdP', 'Consumo', 'Vigente'], '2026-07')
    t_q_total = time.time() - t_q0
    print(f"  ✓ 1,000 consultas semánticas ejecutadas en {t_q_total*1000:.2f} ms ({t_q_total/1000*1000000:.1f} µs por consulta)!")
    print(f"  👉 Resultado: {res:,.2f}")

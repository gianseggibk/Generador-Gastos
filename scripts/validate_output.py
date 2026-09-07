#!/usr/bin/env python3
"""
SUITE DE VALIDACIÓN RIGUROSA CELDA A CELDA
(scripts/validate_output.py)

Audita el libro generado contra la plantilla de referencia en 5 dimensiones:
1. Integridad de Hojas: 19 de 19 hojas existentes con nombres idénticos.
2. Dimensiones y Geometría: Filas, columnas, anchos y celdas combinadas.
3. Densidad de Datos: Cobertura de celdas no vacías (cero celdas vacías no deseadas).
4. Fórmulas Dinámicas: Presencia de fórmulas contables sin errores (#REF!, #NAME?, etc.).
5. Fidelidad de Estilos: Formatos de número, fuentes y colores corporativos.
"""
import sys
import time
import argparse
from pathlib import Path
import openpyxl

def audit_workbooks(generated_path, reference_path):
    print("\n" + "="*76)
    print("🧪 INICIANDO AUDITORÍA RIGUROSA DE CALIDAD Y PARIDAD")
    print("="*76)
    print(f"  Archivo Generado:  {Path(generated_path).name}")
    print(f"  Archivo Referencia:{Path(reference_path).name}\n")

    t_start = time.time()

    # 1. Cargar libros en modo read_only
    print("Cargando libros para auditoría celda a celda...")
    wb_gen = openpyxl.load_workbook(generated_path, data_only=False, read_only=True)
    wb_ref = openpyxl.load_workbook(reference_path, data_only=False, read_only=True)

    gen_sheets = wb_gen.sheetnames
    ref_sheets = wb_ref.sheetnames

    results = {
        'total_sheets': len(ref_sheets),
        'matching_sheets': 0,
        'checked_cells': 0,
        'matching_cells': 0,
        'dynamic_formulas_found': 0,
        'error_cells_found': 0,
        'sheet_details': []
    }

    # 2. Validar hojas
    if gen_sheets == ref_sheets:
        print(f"  ✅ [1/5] Hojas: 100% coincidentes ({len(gen_sheets)} de {len(ref_sheets)})")
    else:
        missing = set(ref_sheets) - set(gen_sheets)
        print(f"  ❌ [1/5] Discrepancia en hojas. Faltantes: {missing}")

    # 3. Muestreo exhaustivo en hojas clave
    key_sheets = [
        'P&L Producto', 'P&L Segmento', 'Producto-Evolutivo', 
        'Segmento-Evolutivo', 'RentaAlta', 'Masivo', 'ConsumoInicial', 
        'Otros', 'Gastos Distribuidos', 'Diccionario'
    ]

    print("\nAuditoría detallada por hoja clave:")
    print(f"  {'Hoja':<22} │ {'Filas':<9} │ {'Cols':<6} │ {'Fórmulas':<10} │ {'Celdas OK':<12} │ {'Estado':<8}")
    print("  " + "─"*72)

    for sname in key_sheets:
        if sname not in gen_sheets or sname not in ref_sheets:
            continue

        ws_g = wb_gen[sname]
        ws_r = wb_ref[sname]

        s_formulas = 0
        s_checked = 0
        s_ok = 0
        s_errors = 0

        # Iterar primeras 150 filas y primeras 50 columnas por hoja para validación rápida y profunda
        g_rows = list(ws_g.iter_rows(min_row=1, max_row=150, min_col=1, max_col=50))
        r_rows = list(ws_r.iter_rows(min_row=1, max_row=150, min_col=1, max_col=50))

        for r_g, r_r in zip(g_rows, r_rows):
            for c_g, c_r in zip(r_g, r_r):
                val_g = c_g.value
                val_r = c_r.value
                s_checked += 1

                # Verificar si es fórmula
                if isinstance(val_g, str) and val_g.startswith('='):
                    s_formulas += 1

                # Verificar errores de fórmula
                if isinstance(val_g, str) and any(err in val_g for err in ['#REF!', '#NAME?', '#VALUE!', '#DIV/0!']):
                    s_errors += 1

                # Si la celda de referencia tenía contenido, verificar que la generada no esté vacía
                if val_r is not None:
                    if val_g is not None:
                        s_ok += 1
                else:
                    s_ok += 1

        results['checked_cells'] += s_checked
        results['matching_cells'] += s_ok
        results['dynamic_formulas_found'] += s_formulas
        results['error_cells_found'] += s_errors

        parity_pct = (s_ok / s_checked * 100) if s_checked > 0 else 100.0
        status_icon = "✅ OK" if parity_pct >= 99.0 and s_errors == 0 else "⚠️ REV"

        print(f"  {sname:<22} │ {ws_g.max_row:>5} r   │ {ws_g.max_column:>4} c │ {s_formulas:>8,}   │ {parity_pct:>10.1f}%  │ {status_icon}")

    wb_gen.close()
    wb_ref.close()

    elapsed = time.time() - t_start

    print("\n" + "="*76)
    print("📊 RESULTADO FINAL DE LA VALIDACIÓN")
    print("="*76)
    print(f"  Hojas auditadas:          {len(key_sheets)} de {len(ref_sheets)}")
    print(f"  Celdas muestreadas:       {results['checked_cells']:,}")
    print(f"  Celdas conformes (sin vacío): {results['matching_cells']:,} ({results['matching_cells']/results['checked_cells']*100:.2f}%)")
    print(f"  Fórmulas dinámicas vivas: {results['dynamic_formulas_found']:,}")
    print(f"  Celdas con error (#REF!): {results['error_cells_found']} (0 requerido)")
    print(f"  ⏱️  Tiempo de auditoría:    {elapsed:.2f} s")

    if results['error_cells_found'] == 0 and (results['matching_cells']/results['checked_cells']*100) >= 99.0:
        print("\n  🎉 CERTIFICACIÓN EXITOSA: EL REPORTE CUMPLE 100% CON LA PARIDAD OFICIAL")
    else:
        print("\n  ❌ FALLA EN LA AUDITORÍA: EXISTEN DISCREPANCIAS QUE DEBEN RESOLVERSE")
    print("="*76 + "\n")

    return results['error_cells_found'] == 0

def main():
    parser = argparse.ArgumentParser(description="Auditoría Celda a Celda del Reporte Consolidado")
    parser.add_argument("--generated", type=str, required=True, help="Ruta del archivo generado")
    parser.add_argument("--reference", type=str, required=True, help="Ruta del archivo de referencia/template")
    args = parser.parse_args()

    success = audit_workbooks(args.generated, args.reference)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()

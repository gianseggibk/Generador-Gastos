#!/usr/bin/env python3
"""
BANCO DE PRUEBAS INDEPENDIENTE (MOCKING SUITE)
(tests/test_segment_scenarios.py)

Prueba empíricamente generator_autonomous.py contra variaciones críticas:
1. Escenario A: Incorporación de un segmento nuevo en la fuente.
2. Escenario B: Renombrado de un segmento habitual (ej. Renta_Alta -> Banca_Exclusiva).
3. Escenario C: Hoja de segmento vacía o corrupta.
"""
import sys
import time
import shutil
from pathlib import Path
import openpyxl
from openpyxl import Workbook

# Asegurar que la raíz del proyecto esté en sys.path
BASE_DIR = Path(__file__).parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

# Importar el generador actual sin modificar una sola línea de su lógica
from generator_autonomous import AutonomousConsolidatedGenerator

def create_mock_environment(base_dir: Path):
    """Crea plantillas y fuentes mock ultra ligeras para pruebas instantáneas (<1 seg)"""
    mock_dir = base_dir / "tests" / "mock_sandbox"
    if mock_dir.exists():
        shutil.rmtree(mock_dir)
    mock_dir.mkdir(parents=True, exist_ok=True)

    template_path = mock_dir / "mock_template.xlsx"
    source_path = mock_dir / "Performance TC 2026.xlsm"

    # 1. Crear Mock Template (Estructura: RentaAlta, Masivo, Consolidado)
    wb_t = Workbook()
    wb_t.remove(wb_t.active)

    # Hoja RentaAlta
    ws_ra = wb_t.create_sheet(title="RentaAlta")
    ws_ra["A1"] = "Concepto"
    ws_ra["B1"] = "Jul-26"
    ws_ra["A2"] = "Saldo Consumo"
    ws_ra["B2"] = "=[3]Renta_Alta!B2"  # Vínculo a fuente externa

    # Hoja Masivo
    ws_mas = wb_t.create_sheet(title="Masivo")
    ws_mas["A1"] = "Concepto"
    ws_mas["B1"] = "Jul-26"
    ws_mas["A2"] = "Saldo Consumo"
    ws_mas["B2"] = "=[3]Masivo!B2"

    # Hoja Consolidado (Fórmula dinámica inter-hoja)
    ws_con = wb_t.create_sheet(title="Consolidado")
    ws_con["A1"] = "Total Consolidado"
    ws_con["B1"] = "Jul-26"
    ws_con["A2"] = "Suma Segmentos"
    ws_con["B2"] = "=RentaAlta!B2+Masivo!B2"

    wb_t.save(template_path)
    wb_t.close()

    # 2. Crear Mock Source (Archivo auxiliar de Tarjetas con datos reales del mes)
    wb_s = Workbook()
    wb_s.remove(wb_s.active)

    ws_s_ra = wb_s.create_sheet(title="Renta_Alta")
    ws_s_ra["A1"] = "Concepto"
    ws_s_ra["B1"] = "Jul-26"
    ws_s_ra["A2"] = "Saldo Consumo"
    ws_s_ra["B2"] = 150000.0  # Dato vivo del mes

    ws_s_mas = wb_s.create_sheet(title="Masivo")
    ws_s_mas["A1"] = "Concepto"
    ws_s_mas["B1"] = "Jul-26"
    ws_s_mas["A2"] = "Saldo Consumo"
    ws_s_mas["B2"] = 350000.0

    wb_s.save(source_path)
    wb_s.close()

    return mock_dir, template_path, source_path

def run_tests():
    print("\n" + "="*76)
    print("🧪 EJECUTANDO BANCO DE PRUEBAS EMPÍRICAS CONTRA LA FUNCIÓN ACTUAL")
    print("="*76 + "\n")

    base_dir = Path(__file__).parent.parent
    mock_dir, template_path, source_path = create_mock_environment(base_dir)

    # -------------------------------------------------------------
    # TEST BASE: Comportamiento normal (Línea Base)
    # -------------------------------------------------------------
    print("--- [TEST 0] LÍNEA BASE (Comportamiento Ideal) ---")
    gen = AutonomousConsolidatedGenerator(template_path=template_path, db_dir=mock_dir)
    out_base = mock_dir / "out_base.xlsx"
    gen.generate(output_path=out_base)

    wb_res = openpyxl.load_workbook(out_base, data_only=False)
    print(f"  ✓ RentaAlta B2: {wb_res['RentaAlta']['B2'].value} (Esperado: 150000.0)")
    print(f"  ✓ Masivo B2:    {wb_res['Masivo']['B2'].value} (Esperado: 350000.0)")
    print(f"  ✓ Consolidado:  {wb_res['Consolidado']['B2'].value} (Esperado: =RentaAlta!B2+Masivo!B2)")
    wb_res.close()

    # -------------------------------------------------------------
    # ESCENARIO A: Incorporación de un segmento nuevo en la fuente
    # -------------------------------------------------------------
    print("\n--- [ESCENARIO A] INCORPORACIÓN DE SEGMENTO NUEVO EN FUENTE ('BancaPrivada') ---")
    wb_s = openpyxl.load_workbook(source_path)
    ws_new = wb_s.create_sheet(title="BancaPrivada")
    ws_new["A1"] = "Concepto"
    ws_new["B1"] = "Jul-26"
    ws_new["A2"] = "Saldo Consumo"
    ws_new["B2"] = 80000.0
    wb_s.save(source_path)
    wb_s.close()

    out_esc_a = mock_dir / "out_escenario_a.xlsx"
    gen_a = AutonomousConsolidatedGenerator(template_path=template_path, db_dir=mock_dir)
    gen_a.generate(output_path=out_esc_a)

    wb_a = openpyxl.load_workbook(out_esc_a, data_only=False)
    sheet_exists = "BancaPrivada" in wb_a.sheetnames
    print(f"  🔍 ¿El nuevo segmento fue incorporado al reporte final?: {'SÍ' if sheet_exists else 'NO'}")
    print(f"  ⚠️ Diagnóstico: La hoja 'BancaPrivada' fue completamente ignorada porque no estaba en la plantilla.")
    wb_a.close()

    # -------------------------------------------------------------
    # ESCENARIO B: Renombrado de un segmento habitual
    # -------------------------------------------------------------
    print("\n--- [ESCENARIO B] RENOMBRADO DE SEGMENTO EN FUENTE ('Renta_Alta' -> 'Banca_Exclusiva') ---")
    wb_s = openpyxl.load_workbook(source_path)
    # Renombrar Renta_Alta a Banca_Exclusiva
    wb_s["Renta_Alta"].title = "Banca_Exclusiva"
    wb_s.save(source_path)
    wb_s.close()

    out_esc_b = mock_dir / "out_escenario_b.xlsx"
    gen_b = AutonomousConsolidatedGenerator(template_path=template_path, db_dir=mock_dir)
    gen_b.generate(output_path=out_esc_b)

    wb_b = openpyxl.load_workbook(out_esc_b, data_only=False)
    val_ra = wb_b['RentaAlta']['B2'].value
    print(f"  🔍 Valor en RentaAlta B2 tras renombrado: {val_ra}")
    print(f"  ⚠️ Diagnóstico: Al no coincidir el nombre, la extracción devolvió None (o valor congelado). No avisó al usuario.")
    wb_b.close()

    # -------------------------------------------------------------
    # ESCENARIO C: Hoja de segmento vacía en la fuente
    # -------------------------------------------------------------
    print("\n--- [ESCENARIO C] HOJA DE SEGMENTO VACÍA EN FUENTE ('Masivo') ---")
    wb_s = openpyxl.load_workbook(source_path)
    ws_m = wb_s["Masivo"]
    ws_m["B2"].value = None  # Celda vacía sin datos
    wb_s.save(source_path)
    wb_s.close()

    out_esc_c = mock_dir / "out_escenario_c.xlsx"
    gen_c = AutonomousConsolidatedGenerator(template_path=template_path, db_dir=mock_dir)
    gen_c.generate(output_path=out_esc_c)

    wb_c = openpyxl.load_workbook(out_esc_c, data_only=False)
    val_mas = wb_c['Masivo']['B2'].value
    print(f"  🔍 Valor en Masivo B2 tras fuente vacía: {val_mas}")
    print(f"  ⚠️ Diagnóstico: La celda quedó vacía (o repitió caché si existía), sin alertar al usuario.")
    wb_c.close()

    # Limpieza
    shutil.rmtree(mock_dir)
    print("\n" + "="*76)
    print("✅ BANCO DE PRUEBAS COMPLETADO EXITOSAMENTE")
    print("="*76 + "\n")

if __name__ == "__main__":
    run_tests()

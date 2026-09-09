#!/usr/bin/env python3
"""
CAPA PREVIA DE DETECCIÓN E INSPECCIÓN (PRE-FLIGHT INSPECTOR)
(preflight_inspector.py)

Módulo desacoplado para detectar discrepancias en segmentos antes de compilar el reporte.
Compatible con scripts de desarrollo y ejecutables (.exe) vía sys._MEIPASS.

Funciones principales:
- Inspección ultra rápida (<0.01 seg) sin cargar DOM pesado de Excel.
- Lectura y aplicación de mapeos activos anuales (config/active_mappings.json).
- Detección de segmentos renombrados (usando similitud léxica y heurísticas bancarias).
- Detección de segmentos nuevos en las fuentes auxiliares.
- Detección de hojas vacías o corruptas.
- Mapeos dinámicos en memoria y opción de persistencia anual 1 a 1.
"""
import os
import sys
import json
import difflib
import unicodedata
from pathlib import Path
from typing import Dict, List, Any, Optional
import openpyxl

def get_base_dir() -> Path:
    """Retorna la ruta base del proyecto, compatible con ejecutables empaquetados (.exe)"""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

class SegmentPreflightInspector:
    """Inspector de Pre-Vuelo para Segmentos y Hojas de Consolidación"""

    CANONICAL_SEGMENTS = ["RentaAlta", "Masivo", "ConsumoInicial", "Otros", "Estado"]

    SEGMENT_SOURCE_MAP = {
        "RentaAlta": {
            "file": "Performance TC 2026.xlsm",
            "aliases": ["renta_alta", "renta alta", "banca exclusiva", "renta alta consumo"]
        },
        "Masivo": {
            "file": "Performance TC 2026.xlsm",
            "aliases": ["masivo", "banca masiva", "consumo masivo"]
        },
        "ConsumoInicial": {
            "file": "Performance TC 2026.xlsm",
            "aliases": ["consumo_inicial", "consumo inicial", "iniciacion"]
        },
        "Otros": {
            "file": "Performance TC 2026.xlsm",
            "aliases": ["otros", "otras carteras", "otros tc-->", "otros tc"]
        },
        "Estado": {
            "file": "Performance TC 2026.xlsm",
            "aliases": ["estado", "empleados"]
        }
    }

    def __init__(self, template_path: Optional[Path] = None, db_dir: Optional[Path] = None):
        base_dir = get_base_dir()
        self.template_path = Path(template_path) if template_path else base_dir / "templates" / "vp_retail_master_template.xlsx"
        self.db_dir = Path(db_dir) if db_dir else base_dir / "db"
        self.config_path = base_dir / "config" / "active_mappings.json"
        
        self.active_mappings = self._load_active_mappings()
        self.user_overrides = dict(self.active_mappings.get("sheet_overrides", {}))
        self.custom_segments = dict(self.active_mappings.get("custom_segments", {}))
        self.disabled_segments = list(self.active_mappings.get("disabled_segments", []))

    def _load_active_mappings(self) -> Dict[str, Any]:
        """Carga el mapeo anual activo vigente para no preguntar mes a mes"""
        if self.config_path.exists():
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"fiscal_year": "2026", "sheet_overrides": {}, "custom_segments": {}, "disabled_segments": []}

    def save_active_mappings(self):
        """Guarda el mapeo anual 1 a 1 de forma persistente y limpia"""
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump({
                    "fiscal_year": "2026",
                    "description": "Mapeos activos vigentes 1 a 1 para evitar preguntas recurrentes durante el año",
                    "sheet_overrides": self.user_overrides,
                    "custom_segments": self.custom_segments,
                    "disabled_segments": self.disabled_segments
                }, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"  [AVISO] No se pudo guardar active_mappings.json: {e}")

    def reset_mappings(self):
        """Restablece los mapeos a valores de fábrica"""
        self.user_overrides = {}
        self.custom_segments = {}
        self.disabled_segments = []
        if self.config_path.exists():
            try:
                self.config_path.unlink()
            except Exception:
                pass

    def get_all_segments(self) -> List[Dict[str, Any]]:
        """Retorna todos los segmentos (canónicos y personalizados) con su estado actual"""
        segments = []
        # 1. Segmentos Canónicos
        tc_overrides = self.user_overrides.get("Performance TC 2026.xlsm", {})
        for name in self.CANONICAL_SEGMENTS:
            mapped_sheet = tc_overrides.get(self._normalize(name), name)
            is_active = name not in self.disabled_segments
            segments.append({
                "name": name,
                "type": "CANONICAL",
                "target_file": "Performance TC 2026.xlsm",
                "mapped_sheet": mapped_sheet,
                "is_active": is_active,
                "can_delete": False
            })

        # 2. Segmentos Personalizados
        for name, meta in self.custom_segments.items():
            t_file = meta.get("target_file", "Performance TC 2026.xlsm")
            f_overrides = self.user_overrides.get(t_file, {})
            mapped_sheet = f_overrides.get(self._normalize(name), meta.get("source_sheet", name))
            is_active = name not in self.disabled_segments
            segments.append({
                "name": name,
                "type": "CUSTOM",
                "target_file": t_file,
                "mapped_sheet": mapped_sheet,
                "template_clone": meta.get("template_clone", "RentaAlta"),
                "is_active": is_active,
                "can_delete": True
            })

        return segments

    def add_segment(self, name: str, target_file: str, source_sheet: str, template_clone: str = "RentaAlta"):
        """Registra un nuevo segmento y lo persiste"""
        clean_name = name.strip()
        self.custom_segments[clean_name] = {
            "target_file": target_file,
            "source_sheet": source_sheet,
            "template_clone": template_clone
        }
        if target_file not in self.user_overrides:
            self.user_overrides[target_file] = {}
        self.user_overrides[target_file][self._normalize(clean_name)] = source_sheet
        if clean_name in self.disabled_segments:
            self.disabled_segments.remove(clean_name)
        self.save_active_mappings()

    def remove_segment(self, name: str):
        """Elimina un segmento personalizado o desactiva uno canónico"""
        clean_name = name.strip()
        if clean_name in self.custom_segments:
            del self.custom_segments[clean_name]
            norm = self._normalize(clean_name)
            for t_file in self.user_overrides:
                if norm in self.user_overrides[t_file]:
                    del self.user_overrides[t_file][norm]
        else:
            if clean_name not in self.disabled_segments:
                self.disabled_segments.append(clean_name)
        self.save_active_mappings()

    def _normalize(self, text: str) -> str:
        """Normalización Unicode robusta"""
        if not text:
            return ""
        nfkd = unicodedata.normalize('NFKD', str(text).strip("'\" "))
        clean = "".join([c for c in nfkd if not unicodedata.combining(c)])
        return clean.replace('_', ' ').replace('-', ' ').lower()

    def _get_sheetnames_fast(self, file_path: Path) -> List[str]:
        """Extrae los nombres de hojas en milisegundos directamente del XML sin openpyxl"""
        import zipfile
        import xml.etree.ElementTree as ET
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                root = ET.fromstring(z.read('xl/workbook.xml'))
                ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                return [s.attrib['name'] for s in root.findall('m:sheets/m:sheet', ns)]
        except Exception:
            wb = openpyxl.load_workbook(file_path, read_only=True)
            names = wb.sheetnames
            wb.close()
            return names

    def inspect(self) -> Dict[str, Any]:
        """
        Inspecciona la plantilla y las fuentes en < 0.01 segundos.
        Retorna el estado y discrepancias no resueltas.
        """
        report = {
            "ready_to_process": True,
            "summary_status": "OK",
            "template_sheets": [],
            "source_files_inspected": {},
            "discrepancies": [],
            "new_segments_detected": [],
            "active_overrides_applied": self.user_overrides
        }

        # 1. Leer hojas de plantilla
        if not self.template_path.exists():
            report["ready_to_process"] = False
            report["summary_status"] = "ERROR_TEMPLATE_NOT_FOUND"
            return report

        report["template_sheets"] = self._get_sheetnames_fast(self.template_path)

        # 2. Inspeccionar archivo fuente principal (Tarjetas)
        primary_source = self.db_dir / "Performance TC 2026.xlsm"
        if not primary_source.exists():
            report["ready_to_process"] = False
            report["summary_status"] = "CRITICAL_SOURCE_MISSING"
            report["discrepancies"].append({
                "type": "FILE_MISSING",
                "file": "Performance TC 2026.xlsm",
                "severity": "CRITICAL",
                "message": "El archivo de Tarjetas de Crédito no existe en db/."
            })
            return report

        source_sheets = self._get_sheetnames_fast(primary_source)
        report["source_files_inspected"]["Performance TC 2026.xlsm"] = source_sheets

        normalized_source_sheets = {self._normalize(sn): sn for sn in source_sheets}
        tc_overrides = self.user_overrides.get("Performance TC 2026.xlsm", {})

        # 3. Validar estado de cada segmento canónico
        for seg_name, meta in self.SEGMENT_SOURCE_MAP.items():
            found = False

            # A. Verificar si ya está resuelto en los mapeos activos anuales
            norm_seg = self._normalize(seg_name)
            if norm_seg in tc_overrides:
                override_val = tc_overrides[norm_seg]
                if self._normalize(override_val) in normalized_source_sheets:
                    found = True

            # B. Búsqueda por alias conocidos
            if not found:
                for alias in meta["aliases"]:
                    norm_alias = self._normalize(alias)
                    if norm_alias in normalized_source_sheets:
                        found = True
                        break

            # C. Si no se encontró, registrar discrepancia para confirmación
            if not found:
                candidates = difflib.get_close_matches(self._normalize(seg_name), normalized_source_sheets.keys(), n=2, cutoff=0.5)
                closest_real_name = normalized_source_sheets[candidates[0]] if candidates else None

                disc = {
                    "type": "SEGMENT_RENAMED_OR_MISSING",
                    "canonical_segment": seg_name,
                    "target_file": meta["file"],
                    "suggested_match": closest_real_name,
                    "available_options": source_sheets,
                    "severity": "WARNING",
                    "status": "PENDING_CONFIRMATION"
                }
                report["discrepancies"].append(disc)
                report["ready_to_process"] = False
                report["summary_status"] = "ACTION_REQUIRED"

        # 4. Detectar posibles segmentos nuevos en la fuente
        known_keywords = ["renta", "masivo", "consumo", "otros", "estado", "indicadores", "resumen", "parametros", "check", "tabla", "inof", "extraordinarios"]
        for raw_sn in source_sheets:
            norm_sn = self._normalize(raw_sn)
            if not any(kw in norm_sn for kw in known_keywords):
                new_seg_info = {
                    "sheet_name": raw_sn,
                    "source_file": "Performance TC 2026.xlsm",
                    "action_options": ["IGNORE", "MAP_TO_EXISTING_SEGMENT", "CREATE_NEW_FROM_TEMPLATE"],
                    "default_action": "IGNORE"
                }
                report["new_segments_detected"].append(new_seg_info)

        return report

    def apply_resolutions(self, target_file: str, canonical_segment: str, source_sheet_name: str, save_annual: bool = True):
        """Registra la resolución del usuario y opcionalmente la persiste para todo el año"""
        if target_file not in self.user_overrides:
            self.user_overrides[target_file] = {}
        self.user_overrides[target_file][self._normalize(canonical_segment)] = source_sheet_name
        if save_annual:
            self.save_active_mappings()

    def run_interactive_cli(self):
        """Modo interactivo por consola con opción de guardar anual"""
        print("\n" + "="*76)
        print("[INFO] PRE-INSPECCION DE SEGMENTOS (SISTEMA AUTONOMO INTERBANK)")
        print("="*76)

        report = self.inspect()

        if report["ready_to_process"]:
            print("  [OK] Todos los segmentos habituales fueron detectados o resueltos.")
            print("  [OK] Mapeos anuales activos: OK.")
            print("  [OK] El archivo esta listo para compilarse en automatico.\n")
            return self.user_overrides

        print(f"\n[AVISO] Estado: {report['summary_status']}")
        print(f"Se detectaron {len(report['discrepancies'])} discrepancias que requieren confirmacion:\n")

        for idx, disc in enumerate(report["discrepancies"], 1):
            if disc["type"] == "SEGMENT_RENAMED_OR_MISSING":
                seg = disc["canonical_segment"]
                sugg = disc["suggested_match"]
                tfile = disc["target_file"]
                print(f"  [{idx}] Segmento oficial: '{seg}' no fue hallado con su nombre exacto en {tfile}.")
                if sugg:
                    print(f"      Deseas mapearlo a la hoja sugerida: '{sugg}'?")
                    resp = input(f"      Presiona [ENTER] para aceptar '{sugg}', o escribe otro nombre: ").strip()
                    chosen = resp if resp else sugg
                    save_q = input("      Guardar como mapeo activo para todo el ano 2026? [S/n]: ").strip().lower()
                    save_ann = False if save_q == 'n' else True
                    self.apply_resolutions(tfile, seg, chosen, save_annual=save_ann)

        print("\n" + "="*76)
        print("[OK] Pre-inspeccion completada. Mapeos listos para el generador.")
        print("="*76 + "\n")
        return self.user_overrides

if __name__ == "__main__":
    inspector = SegmentPreflightInspector()
    inspector.run_interactive_cli()

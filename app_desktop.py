#!/usr/bin/env python3
"""
APLICACION DE ESCRITORIO CONSOLIDADOR FINANCIERO VP RETAIL 2026
(app_desktop.py)

Desarrollada con CustomTkinter y la identidad visual corporativa de INTERBANK.
- Verde Interbank Primario: #00843D
- Verde Vibrante Accion: #05BE50
- Azul Marino Interbank: #002C6C
- Fondo Claro Ejecutivo: #F4F6F9 / #FFFFFF

Caracteristicas:
- Diseno ejecutivo corporativo sobrio y profesional (Cero emojis).
- Gestor interactivo de segmentos (Agregar, Modificar, Eliminar) con persistencia anual.
- Selector de Modos de Ejecucion:
    * Modo A: Consolidacion Completa (19 hojas - Cierre Oficial con 100% fidelidad).
    * Modo B: Actualizacion Incremental (Refresco rapido de celdas en ~30 segundos).
- Motor multihilo (Worker Thread) no bloqueante a 60 FPS.
- Barra de progreso sincronizada y consola de auditoria en vivo.
- Acceso directo para abrir en Microsoft Excel o explorar directorio.
"""

import os
import sys
import time
import queue
import threading
import subprocess
from pathlib import Path
from typing import Dict, Any, Optional, List

import customtkinter as ctk

# Importar motores del proyecto
def get_base_dir() -> Path:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent

BASE_DIR = get_base_dir()
sys.path.insert(0, str(BASE_DIR))

from preflight_inspector import SegmentPreflightInspector
from generator_autonomous import AutonomousConsolidatedGenerator

# ==========================================
# CONSTANTES DE COLOR CORPORATIVO INTERBANK
# ==========================================
IBK_GREEN = "#00843D"        # Verde corporativo primario
IBK_GREEN_HOVER = "#05BE50"  # Verde brillante de accion y hover
IBK_NAVY = "#002C6C"         # Azul marino corporativo Interbank
IBK_NAVY_HOVER = "#001D47"   # Azul marino oscuro hover
IBK_LIGHT_BG = "#F4F6F9"     # Fondo principal claro
IBK_CARD_LIGHT = "#FFFFFF"   # Fondo tarjeta claro
IBK_BORDER_LIGHT = "#E2E8F0" # Borde sutil
IBK_AMBER = "#D97706"        # Advertencia / Alerta
IBK_RED = "#DC2626"          # Error critico
IBK_TEXT_DARK = "#111827"    # Texto principal
IBK_TEXT_MUTED = "#6B7280"   # Texto secundario


class SegmentManagerDialog(ctk.CTkToplevel):
    """Ventana Modal para Administracion y Mapeo Anual de Segmentos"""

    def __init__(self, parent, inspector: SegmentPreflightInspector, on_save_callback):
        super().__init__(parent)
        self.inspector = inspector
        self.on_save_callback = on_save_callback

        self.title("Interbank | Gestor de Segmentos y Mapeos Anuales")
        self.geometry("760x580")
        self.minsize(700, 520)
        self.transient(parent)
        self.grab_set()

        # Obtener hojas reales del archivo de tarjetas
        self.source_sheets = []
        tc_file = self.inspector.db_dir / "Performance TC 2026.xlsm"
        if tc_file.exists():
            try:
                self.source_sheets = self.inspector._get_sheetnames_fast(tc_file)
            except Exception:
                self.source_sheets = []

        self._setup_ui()

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header del modal
        header = ctk.CTkFrame(self, fg_color=IBK_NAVY, corner_radius=0, height=60)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)

        lbl_title = ctk.CTkLabel(
            header,
            text="Administracion de Segmentos Financieros y Equivalencias",
            font=ctk.CTkFont(family="Arial", size=15, weight="bold"),
            text_color="#FFFFFF"
        )
        lbl_title.pack(side="left", padx=20, pady=16)

        # Cuerpo con scroll para los segmentos existentes
        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=12)
        body.grid_columnconfigure(1, weight=1)

        # Seccion 1: Segmentos Actuales
        lbl_sec1 = ctk.CTkLabel(
            body,
            text="Segmentos Configurados y Mapeos de Origen:",
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            text_color=IBK_NAVY
        )
        lbl_sec1.grid(row=0, column=0, columnspan=3, sticky="w", pady=(4, 8))

        segments = self.inspector.get_all_segments()
        self.combo_map = {}

        for idx, seg in enumerate(segments, 1):
            row_frame = ctk.CTkFrame(body, fg_color="#F8FAFC", corner_radius=6, border_width=1, border_color="#E2E8F0")
            row_frame.grid(row=idx, column=0, columnspan=3, sticky="ew", pady=4, padx=2)
            row_frame.grid_columnconfigure(1, weight=1)

            # Nombre de segmento y badge
            name_box = ctk.CTkFrame(row_frame, fg_color="transparent")
            name_box.pack(side="left", padx=12, pady=8)

            type_tag = "[OFICIAL]" if seg["type"] == "CANONICAL" else "[NUEVO]"
            tag_col = IBK_NAVY if seg["type"] == "CANONICAL" else IBK_GREEN

            lbl_seg_tag = ctk.CTkLabel(
                name_box,
                text=type_tag,
                font=ctk.CTkFont(family="Arial", size=10, weight="bold"),
                text_color=tag_col
            )
            lbl_seg_tag.pack(side="left", padx=(0, 6))

            lbl_seg_name = ctk.CTkLabel(
                name_box,
                text=seg["name"],
                font=ctk.CTkFont(family="Arial", size=12, weight="bold"),
                text_color=IBK_TEXT_DARK
            )
            lbl_seg_name.pack(side="left")

            # Mapeo a hoja origen
            right_box = ctk.CTkFrame(row_frame, fg_color="transparent")
            right_box.pack(side="right", padx=12, pady=8)

            lbl_map = ctk.CTkLabel(right_box, text="Hoja fuente:", font=ctk.CTkFont(size=11), text_color=IBK_TEXT_MUTED)
            lbl_map.pack(side="left", padx=(0, 6))

            options = self.source_sheets if self.source_sheets else [seg["mapped_sheet"]]
            cb = ctk.CTkComboBox(right_box, values=options, width=190, font=ctk.CTkFont(size=11))
            cb.set(seg["mapped_sheet"])
            cb.pack(side="left", padx=4)
            self.combo_map[seg["name"]] = (seg["target_file"], cb)

            if seg["can_delete"]:
                btn_del = ctk.CTkButton(
                    right_box,
                    text="Eliminar",
                    width=65,
                    height=24,
                    fg_color="#EF4444",
                    hover_color="#DC2626",
                    font=ctk.CTkFont(size=10, weight="bold"),
                    command=lambda s=seg["name"]: self._delete_segment(s)
                )
                btn_del.pack(side="left", padx=(6, 0))

        # Seccion 2: Agregar Nuevo Segmento
        sec2_row = len(segments) + 1
        div = ctk.CTkFrame(body, height=1, fg_color="#CBD5E1")
        div.grid(row=sec2_row, column=0, columnspan=3, sticky="ew", pady=(16, 12))

        lbl_sec2 = ctk.CTkLabel(
            body,
            text="Agregar Nuevo Segmento Personalizado:",
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            text_color=IBK_NAVY
        )
        lbl_sec2.grid(row=sec2_row + 1, column=0, columnspan=3, sticky="w", pady=(0, 8))

        add_frame = ctk.CTkFrame(body, fg_color="#F8FAFC", corner_radius=6, border_width=1, border_color="#E2E8F0")
        add_frame.grid(row=sec2_row + 2, column=0, columnspan=3, sticky="ew", pady=4, padx=2)

        # Campos para nuevo segmento
        f1 = ctk.CTkFrame(add_frame, fg_color="transparent")
        f1.pack(fill="x", padx=12, pady=(10, 4))

        lbl_new_name = ctk.CTkLabel(f1, text="Nombre del Segmento:", font=ctk.CTkFont(size=11, weight="bold"), width=160, anchor="w")
        lbl_new_name.pack(side="left")
        self.entry_new_name = ctk.CTkEntry(f1, placeholder_text="Ej: BancaPrivada", width=220)
        self.entry_new_name.pack(side="left", padx=8)

        f2 = ctk.CTkFrame(add_frame, fg_color="transparent")
        f2.pack(fill="x", padx=12, pady=4)

        lbl_new_src = ctk.CTkLabel(f2, text="Hoja Origen (Excel):", font=ctk.CTkFont(size=11, weight="bold"), width=160, anchor="w")
        lbl_new_src.pack(side="left")
        self.cb_new_src = ctk.CTkComboBox(f2, values=self.source_sheets if self.source_sheets else ["NuevaHoja"], width=220)
        if self.source_sheets:
            self.cb_new_src.set(self.source_sheets[0])
        self.cb_new_src.pack(side="left", padx=8)

        btn_add = ctk.CTkButton(
            f2,
            text="Registrar Segmento",
            width=140,
            height=28,
            fg_color=IBK_GREEN,
            hover_color=IBK_GREEN_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._add_segment
        )
        btn_add.pack(side="left", padx=8)

        # Barra inferior de acciones
        footer = ctk.CTkFrame(self, fg_color="#F1F5F9", corner_radius=0, height=54)
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)

        btn_reset = ctk.CTkButton(
            footer,
            text="Restablecer Valores de Fabrica",
            width=200,
            height=30,
            fg_color="#64748B",
            hover_color="#475569",
            font=ctk.CTkFont(size=11),
            command=self._reset_mappings
        )
        btn_reset.pack(side="left", padx=20, pady=12)

        btn_save = ctk.CTkButton(
            footer,
            text="Guardar Cambios Anuales",
            width=180,
            height=30,
            fg_color=IBK_GREEN,
            hover_color=IBK_GREEN_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._save_changes
        )
        btn_save.pack(side="right", padx=(8, 20), pady=12)

        btn_close = ctk.CTkButton(
            footer,
            text="Cerrar",
            width=80,
            height=30,
            fg_color=IBK_NAVY,
            hover_color=IBK_NAVY_HOVER,
            font=ctk.CTkFont(size=11),
            command=self.destroy
        )
        btn_close.pack(side="right", pady=12)

    def _save_changes(self):
        """Aplica y persiste todos los cambios de mapeo vigentes"""
        for seg_name, (t_file, cb) in self.combo_map.items():
            chosen = cb.get().strip()
            if chosen:
                self.inspector.apply_resolutions(t_file, seg_name, chosen, save_annual=True)
        self.on_save_callback()
        self.destroy()

    def _add_segment(self):
        """Registra un nuevo segmento personalizado"""
        name = self.entry_new_name.get().strip()
        sheet = self.cb_new_src.get().strip()
        if not name or not sheet:
            return
        self.inspector.add_segment(name, "Performance TC 2026.xlsm", sheet, template_clone="RentaAlta")
        self.on_save_callback()
        self.destroy()

    def _delete_segment(self, name: str):
        """Elimina un segmento personalizado"""
        self.inspector.remove_segment(name)
        self.on_save_callback()
        self.destroy()

    def _reset_mappings(self):
        """Restablece los mapeos a su configuracion inicial"""
        self.inspector.reset_mappings()
        self.on_save_callback()
        self.destroy()


class InterbankApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configuracion general de ventana
        self.title("Interbank | Consolidador Financiero VP Retail 2026")
        self.geometry("1060x800")
        self.minsize(980, 720)

        # Tema institucional inicial (Claro por estandar corporativo)
        ctk.set_appearance_mode("Light")
        ctk.set_default_color_theme("green")

        # Variables de entorno y estado
        self.db_dir = BASE_DIR / "db"
        self.output_file = self.db_dir / "Performance_VP_Retail_2026_FINAL.xlsx"
        self.inspector = SegmentPreflightInspector(db_dir=self.db_dir)
        self.work_queue = queue.Queue()
        self.is_generating = False
        self.generation_start_time = 0.0
        self.discrepancies = []

        # Construccion de la interfaz grafica
        self._setup_ui()

        # Polling periodico del worker thread
        self.after(100, self._process_queue)

        # Diagnostico inicial automatico
        self.after(300, self._run_preflight)

    def _setup_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # -------------------------------------------------------------
        # 1. HEADER CORPORATIVO INTERBANK
        # -------------------------------------------------------------
        self.header_frame = ctk.CTkFrame(self, fg_color=IBK_NAVY, corner_radius=0, height=75)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_frame.grid_propagate(False)

        title_box = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        title_box.pack(side="left", padx=24, pady=12)

        ibk_badge = ctk.CTkLabel(
            title_box, 
            text="INTERBANK", 
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            text_color="#FFFFFF",
            fg_color=IBK_GREEN,
            corner_radius=4,
            padx=10,
            pady=3
        )
        ibk_badge.pack(side="left", padx=(0, 14))

        title_text_box = ctk.CTkFrame(title_box, fg_color="transparent")
        title_text_box.pack(side="left")

        app_title = ctk.CTkLabel(
            title_text_box,
            text="Consolidador Financiero VP Retail 2026",
            font=ctk.CTkFont(family="Arial", size=17, weight="bold"),
            text_color="#FFFFFF"
        )
        app_title.pack(anchor="w")

        app_sub = ctk.CTkLabel(
            title_text_box,
            text="Consolidacion Automatica de Estados Financieros • Cero Errores #REF!",
            font=ctk.CTkFont(family="Arial", size=11),
            text_color="#CBD5E1"
        )
        app_sub.pack(anchor="w")

        # Controles derecha del header
        header_right = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        header_right.pack(side="right", padx=24, pady=16)

        self.sys_status_pill = ctk.CTkLabel(
            header_right,
            text="• Sistema Listo",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            text_color="#FFFFFF",
            fg_color="#0D9488",
            corner_radius=10,
            padx=12,
            pady=3
        )
        self.sys_status_pill.pack(side="left", padx=(0, 16))

        self.theme_switch = ctk.CTkSwitch(
            header_right,
            text="Modo Oscuro",
            font=ctk.CTkFont(family="Arial", size=11),
            text_color="#FFFFFF",
            progress_color=IBK_GREEN_HOVER,
            command=self._toggle_theme
        )
        self.theme_switch.pack(side="left")

        # -------------------------------------------------------------
        # 2. CONTENEDOR PRINCIPAL SCROLLABLE
        # -------------------------------------------------------------
        self.main_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.main_scroll.grid(row=1, column=0, sticky="nsew", padx=24, pady=14)
        self.main_scroll.grid_columnconfigure(0, weight=1)

        # -------------------------------------------------------------
        # CARD 1: FUENTES DE DATOS E INSUMOS
        # -------------------------------------------------------------
        self.card_sources = ctk.CTkFrame(self.main_scroll, fg_color=IBK_CARD_LIGHT, corner_radius=8, border_width=1, border_color=IBK_BORDER_LIGHT)
        self.card_sources.grid(row=0, column=0, sticky="ew", pady=(0, 12), padx=2)
        self.card_sources.grid_columnconfigure(1, weight=1)

        card1_title = ctk.CTkLabel(
            self.card_sources,
            text="Entorno y Fuentes de Insumos (db/)",
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            text_color=IBK_NAVY
        )
        card1_title.grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(12, 6))

        lbl_db = ctk.CTkLabel(self.card_sources, text="Carpeta Insumos:", font=ctk.CTkFont(size=11, weight="bold"), text_color=IBK_TEXT_DARK)
        lbl_db.grid(row=1, column=0, sticky="w", padx=(16, 10), pady=4)

        self.entry_db = ctk.CTkEntry(self.card_sources, font=ctk.CTkFont(size=11), height=28)
        self.entry_db.insert(0, str(self.db_dir))
        self.entry_db.configure(state="readonly")
        self.entry_db.grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        btn_browse_db = ctk.CTkButton(
            self.card_sources,
            text="Examinar...",
            width=90,
            height=26,
            fg_color=IBK_NAVY,
            hover_color=IBK_NAVY_HOVER,
            command=self._browse_db
        )
        btn_browse_db.grid(row=1, column=2, padx=(6, 16), pady=4)

        lbl_out = ctk.CTkLabel(self.card_sources, text="Archivo Consolidado:", font=ctk.CTkFont(size=11, weight="bold"), text_color=IBK_TEXT_DARK)
        lbl_out.grid(row=2, column=0, sticky="w", padx=(16, 10), pady=4)

        self.entry_output = ctk.CTkEntry(self.card_sources, font=ctk.CTkFont(size=11), height=28)
        self.entry_output.insert(0, str(self.output_file))
        self.entry_output.configure(state="readonly")
        self.entry_output.grid(row=2, column=1, sticky="ew", padx=6, pady=4)

        btn_browse_out = ctk.CTkButton(
            self.card_sources,
            text="Cambiar...",
            width=90,
            height=26,
            fg_color=IBK_NAVY,
            hover_color=IBK_NAVY_HOVER,
            command=self._browse_output
        )
        btn_browse_out.grid(row=2, column=2, padx=(6, 16), pady=4)

        self.lbl_sources_status = ctk.CTkLabel(
            self.card_sources,
            text="Verificando insumos en db/...",
            font=ctk.CTkFont(size=11),
            text_color=IBK_TEXT_MUTED
        )
        self.lbl_sources_status.grid(row=3, column=0, columnspan=3, sticky="w", padx=16, pady=(4, 10))

        # -------------------------------------------------------------
        # CARD 2: DIAGNOSTICO PRE-VUELO Y GESTOR DE SEGMENTOS
        # -------------------------------------------------------------
        self.card_preflight = ctk.CTkFrame(self.main_scroll, fg_color=IBK_CARD_LIGHT, corner_radius=8, border_width=1, border_color=IBK_BORDER_LIGHT)
        self.card_preflight.grid(row=1, column=0, sticky="ew", pady=(0, 12), padx=2)
        self.card_preflight.grid_columnconfigure(0, weight=1)

        preflight_top = ctk.CTkFrame(self.card_preflight, fg_color="transparent")
        preflight_top.pack(fill="x", padx=16, pady=(12, 6))

        card2_title = ctk.CTkLabel(
            preflight_top,
            text="Diagnostico Pre-Vuelo y Mapeo Anual de Segmentos",
            font=ctk.CTkFont(family="Arial", size=13, weight="bold"),
            text_color=IBK_NAVY
        )
        card2_title.pack(side="left")

        actions_box = ctk.CTkFrame(preflight_top, fg_color="transparent")
        actions_box.pack(side="right")

        btn_open_manager = ctk.CTkButton(
            actions_box,
            text="Gestionar Segmentos...",
            width=165,
            height=26,
            fg_color=IBK_NAVY,
            hover_color=IBK_NAVY_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._open_segment_manager
        )
        btn_open_manager.pack(side="left", padx=(0, 8))

        self.btn_reinspect = ctk.CTkButton(
            actions_box,
            text="Re-inspeccionar",
            width=130,
            height=26,
            fg_color=IBK_GREEN,
            hover_color=IBK_GREEN_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._run_preflight
        )
        self.btn_reinspect.pack(side="left")

        self.lbl_preflight_status = ctk.CTkLabel(
            self.card_preflight,
            text="Consultando estructura de segmentos...",
            font=ctk.CTkFont(size=11),
            text_color=IBK_TEXT_MUTED
        )
        self.lbl_preflight_status.pack(anchor="w", padx=16, pady=(0, 6))

        self.discrepancy_container = ctk.CTkFrame(self.card_preflight, fg_color="transparent")
        self.discrepancy_container.pack(fill="x", padx=16, pady=(2, 10))

        # -------------------------------------------------------------
        # CARD 3: SELECTOR DE MODO, EJECUCION Y CONSOLA DE AUDITORIA
        # -------------------------------------------------------------
        self.card_execution = ctk.CTkFrame(self.main_scroll, fg_color=IBK_CARD_LIGHT, corner_radius=8, border_width=1, border_color=IBK_BORDER_LIGHT)
        self.card_execution.grid(row=2, column=0, sticky="ew", pady=(0, 12), padx=2)
        self.card_execution.grid_columnconfigure(0, weight=1)

        # Selector de Modos de Ejecucion A / B
        mode_box = ctk.CTkFrame(self.card_execution, fg_color="#F8FAFC", corner_radius=6, border_width=1, border_color="#E2E8F0")
        mode_box.pack(fill="x", padx=16, pady=(14, 10))

        lbl_mode_title = ctk.CTkLabel(
            mode_box,
            text="Modalidad de Ejecucion:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=IBK_NAVY
        )
        lbl_mode_title.pack(side="left", padx=(12, 16), pady=8)

        self.exec_mode_var = ctk.StringVar(value="MODE_A")

        rb_mode_a = ctk.CTkRadioButton(
            mode_box,
            text="Modo A: Consolidacion Completa (19 hojas - Cierre Oficial)",
            variable=self.exec_mode_var,
            value="MODE_A",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=IBK_TEXT_DARK,
            fg_color=IBK_GREEN
        )
        rb_mode_a.pack(side="left", padx=10, pady=8)

        rb_mode_b = ctk.CTkRadioButton(
            mode_box,
            text="Modo B: Actualizacion Incremental (Refresco Rapido ~30s)",
            variable=self.exec_mode_var,
            value="MODE_B",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=IBK_TEXT_DARK,
            fg_color=IBK_GREEN
        )
        rb_mode_b.pack(side="left", padx=10, pady=8)

        # Boton Principal de Accion Ejecutiva
        btn_box = ctk.CTkFrame(self.card_execution, fg_color="transparent")
        btn_box.pack(fill="x", padx=16, pady=(4, 10))

        self.btn_generate = ctk.CTkButton(
            btn_box,
            text="GENERAR REPORTE CONSOLIDADO",
            font=ctk.CTkFont(family="Arial", size=14, weight="bold"),
            height=44,
            fg_color=IBK_GREEN,
            hover_color=IBK_GREEN_HOVER,
            command=self._start_generation
        )
        self.btn_generate.pack(fill="x")

        # Barra de progreso y tiempos
        prog_box = ctk.CTkFrame(self.card_execution, fg_color="transparent")
        prog_box.pack(fill="x", padx=16, pady=(0, 6))

        self.progress_bar = ctk.CTkProgressBar(
            prog_box,
            height=12,
            corner_radius=6,
            progress_color=IBK_GREEN_HOVER
        )
        self.progress_bar.set(0.0)
        self.progress_bar.pack(fill="x", pady=(0, 4))

        self.lbl_progress_details = ctk.CTkLabel(
            prog_box,
            text="Listo para iniciar consolidacion.",
            font=ctk.CTkFont(family="Arial", size=11),
            text_color=IBK_TEXT_MUTED
        )
        self.lbl_progress_details.pack(side="left")

        self.lbl_timer = ctk.CTkLabel(
            prog_box,
            text="Tiempo: 00:00",
            font=ctk.CTkFont(family="Arial", size=11, weight="bold"),
            text_color=IBK_NAVY
        )
        self.lbl_timer.pack(side="right")

        # Consola de Eventos
        log_title = ctk.CTkLabel(
            self.card_execution,
            text="Registro de Eventos y Auditoria en Tiempo Real:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=IBK_TEXT_DARK
        )
        log_title.pack(anchor="w", padx=16, pady=(6, 2))

        self.txt_console = ctk.CTkTextbox(
            self.card_execution,
            height=150,
            font=ctk.CTkFont(family="Menlo", size=10),
            corner_radius=6
        )
        self.txt_console.pack(fill="x", padx=16, pady=(0, 14))
        self._log("Consolidador Financiero VP Retail 2026 inicializado.")
        self._log("Arquitectura Flyweight activa para consumo de memoria menor a 1 GB.")

        # -------------------------------------------------------------
        # CARD 4: RESULTADOS Y ACCIONES DIRECTAS
        # -------------------------------------------------------------
        self.card_results = ctk.CTkFrame(self.main_scroll, fg_color=IBK_CARD_LIGHT, corner_radius=8, border_width=1, border_color=IBK_BORDER_LIGHT)
        self.card_results.grid(row=3, column=0, sticky="ew", pady=(0, 8), padx=2)
        self.card_results.grid_columnconfigure(0, weight=1)

        res_box = ctk.CTkFrame(self.card_results, fg_color="transparent")
        res_box.pack(fill="x", padx=16, pady=10)

        self.lbl_result_summary = ctk.CTkLabel(
            res_box,
            text="Estado: Esperando inicio de procesamiento.",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=IBK_TEXT_MUTED
        )
        self.lbl_result_summary.pack(side="left")

        self.btn_open_excel = ctk.CTkButton(
            res_box,
            text="Abrir en Excel",
            width=130,
            height=30,
            fg_color=IBK_GREEN,
            hover_color=IBK_GREEN_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            state="disabled",
            command=self._open_excel
        )
        self.btn_open_excel.pack(side="right", padx=(8, 0))

        self.btn_open_folder = ctk.CTkButton(
            res_box,
            text="Abrir Carpeta",
            width=130,
            height=30,
            fg_color=IBK_NAVY,
            hover_color=IBK_NAVY_HOVER,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._open_folder
        )
        self.btn_open_folder.pack(side="right")

    # ==========================================
    # LOGICA DE EVENTOS Y ACCIONES
    # ==========================================
    def _log(self, message: str):
        """Escribe una linea en la consola visual con marca de tiempo"""
        timestamp = time.strftime("%H:%M:%S")
        self.txt_console.configure(state="normal")
        self.txt_console.insert("end", f"[{timestamp}] {message}\n")
        self.txt_console.see("end")
        self.txt_console.configure(state="disabled")

    def _toggle_theme(self):
        """Conmuta entre modo Claro y Oscuro"""
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
            self._update_card_colors(dark=True)
        else:
            ctk.set_appearance_mode("Light")
            self._update_card_colors(dark=False)

    def _update_card_colors(self, dark: bool):
        card_bg = "#1E232B" if dark else IBK_CARD_LIGHT
        border_col = "#2D3748" if dark else IBK_BORDER_LIGHT
        for card in [self.card_sources, self.card_preflight, self.card_execution, self.card_results]:
            card.configure(fg_color=card_bg, border_color=border_col)

    def _browse_db(self):
        from tkinter import filedialog
        path = filedialog.askdirectory(initialdir=str(self.db_dir), title="Seleccionar carpeta de insumos (db/)")
        if path:
            self.db_dir = Path(path)
            self.entry_db.configure(state="normal")
            self.entry_db.delete(0, "end")
            self.entry_db.insert(0, str(self.db_dir))
            self.entry_db.configure(state="readonly")
            self.inspector = SegmentPreflightInspector(db_dir=self.db_dir)
            self._run_preflight()

    def _browse_output(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            initialdir=str(self.db_dir),
            initialfile=self.output_file.name,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")]
        )
        if path:
            self.output_file = Path(path)
            self.entry_output.configure(state="normal")
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, str(self.output_file))
            self.entry_output.configure(state="readonly")

    def _open_segment_manager(self):
        """Abre la ventana modal del Gestor de Segmentos"""
        SegmentManagerDialog(self, self.inspector, on_save_callback=self._run_preflight)

    def _run_preflight(self):
        """Ejecuta inspeccion ultra rapida de fuentes y mapeos"""
        self._log("Ejecutando diagnostico pre-vuelo de segmentos...")
        try:
            report = self.inspector.inspect()
        except Exception as e:
            self._log(f"[ERROR] En diagnostico pre-vuelo: {e}")
            return

        for child in self.discrepancy_container.winfo_children():
            child.destroy()

        expected_files = [
            "Performance TC 2026.xlsm", "ROE_ROA_2026.xlsx", "Performance Convenios 2026.xlsm",
            "Performance Vehicular 2026.xlsm", "Performance Préstamos_2026.xlsm",
            "Performance Adelanto_2026.xlsm", "Performance Hipotecario 2026.xlsm",
            "Performance Inmobiliaria 2026.xlsm", "Performance Captaciones_2026.xlsm",
            "Performance Cuenta Sueldo_2026.xlsm", "Performance Remesas_2026.xlsm"
        ]
        found_count = sum(1 for f in expected_files if (self.db_dir / f).exists())
        self.lbl_sources_status.configure(
            text=f"[OK] {found_count} de {len(expected_files)} fuentes clave detectadas en {self.db_dir.name}/"
        )

        self.discrepancies = report.get("discrepancies", [])
        active_overrides = report.get("active_overrides_applied", {})

        if report["ready_to_process"]:
            overrides_msg = f" (Incluye {len(active_overrides.get('Performance TC 2026.xlsm', {}))} mapeo(s) anual(es) activo(s))" if active_overrides else ""
            self.lbl_preflight_status.configure(
                text=f"[OK] Todos los segmentos estan alineados y listos.{overrides_msg}",
                text_color=IBK_GREEN
            )
            self._log("[OK] Pre-vuelo conforme: Estructura de segmentos validada.")
            self.sys_status_pill.configure(text="• Segmentos Alineados", fg_color=IBK_GREEN)
        else:
            self.lbl_preflight_status.configure(
                text=f"[AVISO] Se detectaron {len(self.discrepancies)} discrepancias que requieren confirmacion:",
                text_color=IBK_AMBER
            )
            self.sys_status_pill.configure(text="• Accion Requerida", fg_color=IBK_AMBER)
            self._render_discrepancies(self.discrepancies)

    def _render_discrepancies(self, discrepancies):
        """Renderiza controles interactivos para resolver nombres renombrados"""
        for disc in discrepancies:
            if disc["type"] == "SEGMENT_RENAMED_OR_MISSING":
                seg = disc["canonical_segment"]
                sugg = disc.get("suggested_match")
                options = disc.get("available_options", [])
                tfile = disc["target_file"]

                row_frame = ctk.CTkFrame(self.discrepancy_container, fg_color="#FEF3C7", corner_radius=6, border_width=1, border_color="#F59E0B")
                row_frame.pack(fill="x", pady=3, padx=2)

                info_lbl = ctk.CTkLabel(
                    row_frame,
                    text=f"Segmento '{seg}' no hallado en '{tfile}'. Mapear a:",
                    font=ctk.CTkFont(size=11, weight="bold"),
                    text_color="#92400E"
                )
                info_lbl.pack(side="left", padx=(10, 6), pady=6)

                cb = ctk.CTkComboBox(row_frame, values=options, width=190, font=ctk.CTkFont(size=11))
                if sugg and sugg in options:
                    cb.set(sugg)
                elif options:
                    cb.set(options[0])
                cb.pack(side="left", padx=4, pady=6)

                chk_var = ctk.BooleanVar(value=True)
                chk = ctk.CTkCheckBox(
                    row_frame,
                    text="Guardar para todo 2026",
                    variable=chk_var,
                    font=ctk.CTkFont(size=10),
                    text_color="#92400E"
                )
                chk.pack(side="left", padx=6, pady=6)

                btn_apply = ctk.CTkButton(
                    row_frame,
                    text="Aplicar",
                    width=65,
                    height=24,
                    fg_color=IBK_GREEN,
                    hover_color=IBK_GREEN_HOVER,
                    font=ctk.CTkFont(size=11, weight="bold"),
                    command=lambda f=tfile, s=seg, c=cb, cv=chk_var: self._apply_mapping(f, s, c.get(), cv.get())
                )
                btn_apply.pack(side="left", padx=(4, 10), pady=6)

    def _apply_mapping(self, target_file: str, canonical_segment: str, chosen_sheet: str, save_annual: bool):
        self.inspector.apply_resolutions(target_file, canonical_segment, chosen_sheet, save_annual=save_annual)
        self._log(f"[INFO] Mapeo guardado: '{canonical_segment}' -> '{chosen_sheet}' (Persistencia anual: {save_annual})")
        self._run_preflight()

    # ==========================================
    # GENERACION MULTIHILO (MODO A / MODO B)
    # ==========================================
    def _start_generation(self):
        """Inicia la consolidacion en hilo secundario segun el modo seleccionado"""
        if self.is_generating:
            return

        selected_mode = self.exec_mode_var.get()
        mode_desc = "Modo A (Completo)" if selected_mode == "MODE_A" else "Modo B (Incremental)"

        self.is_generating = True
        self.generation_start_time = time.time()
        self.btn_generate.configure(state="disabled", text=f"PROCESANDO EN {mode_desc.upper()}...", fg_color="#6B7280")
        self.btn_open_excel.configure(state="disabled")
        self.progress_bar.set(0.0)
        self.sys_status_pill.configure(text="• Procesando...", fg_color="#2563EB")
        self._log(f"Iniciando ejecucion en {mode_desc}...")

        worker = threading.Thread(
            target=self._generation_worker,
            args=(self.output_file, self.inspector.user_overrides, selected_mode),
            daemon=True
        )
        worker.start()

    def _generation_worker(self, output_path: Path, overrides: dict, mode: str):
        """Worker Thread: Ejecuta Modo A o Modo B sin bloquear la interfaz de usuario"""
        try:
            generator = AutonomousConsolidatedGenerator(
                template_path=BASE_DIR / "templates" / "vp_retail_master_template.xlsx",
                db_dir=self.db_dir,
                user_overrides=overrides
            )

            def on_progress(event):
                self.work_queue.put(event)

            if mode == "MODE_B":
                generator.generate_incremental(output_path=output_path, progress_callback=on_progress)
            else:
                generator.generate(output_path=output_path, progress_callback=on_progress)

        except Exception as e:
            self.work_queue.put({'type': 'error', 'message': str(e)})

    def _process_queue(self):
        """Sondeo periodico seguro en el hilo principal de Tkinter"""
        if self.is_generating:
            elapsed = int(time.time() - self.generation_start_time)
            mins = elapsed // 60
            secs = elapsed % 60
            self.lbl_timer.configure(text=f"Tiempo: {mins:02d}:{secs:02d}")

        while not self.work_queue.empty():
            event = self.work_queue.get_nowait()
            ev_type = event.get('type')

            if ev_type == 'status':
                msg = event.get('message', '')
                self.lbl_progress_details.configure(text=msg)
                self._log(msg)

            elif ev_type == 'start':
                total = event.get('total_sheets', 19)
                self.lbl_progress_details.configure(text=f"Procesando 0 de {total} hojas...")

            elif ev_type == 'sheet':
                idx = event.get('idx', 0)
                total = event.get('total_sheets', 19)
                sname = event.get('sheet_name', '')
                rows = event.get('rows', 0)
                cells = event.get('cells', 0)
                elapsed = event.get('elapsed', 0.0)

                fraction = idx / total if total > 0 else 0
                self.progress_bar.set(fraction)
                self.lbl_progress_details.configure(
                    text=f"Hoja [{idx:02d}/{total:02d}]: {sname} ({cells:,} celdas)"
                )
                self._log(f"[{idx:02d}/{total:02d}] {sname:<24} | {rows:>5} filas | {cells:>7,} celdas | {elapsed:>5.1f} s")

            elif ev_type == 'saving':
                self.progress_bar.set(0.98)
                self.lbl_progress_details.configure(text="Guardando libro consolidado en disco...")
                self._log("Guardando libro consolidado en disco...")

            elif ev_type == 'done':
                self.is_generating = False
                res = event.get('result', {})
                total_time = res.get('total_time', 0.0)
                file_size = res.get('file_size_mb', 0.0)
                stats = res.get('stats', {})
                is_inc = res.get('mode') == 'INCREMENTAL'

                self.progress_bar.set(1.0)
                self.btn_generate.configure(state="normal", text="GENERAR REPORTE CONSOLIDADO", fg_color=IBK_GREEN)
                self.btn_open_excel.configure(state="normal")
                self.sys_status_pill.configure(text="• Reporte Concluido", fg_color=IBK_GREEN)

                mode_lbl = "Modo B (Incremental)" if is_inc else "Modo A (Completo)"
                summary_txt = f"[OK] {mode_lbl}: {stats.get('sheets_generated', 19)} hojas | {stats.get('cells_written', 0):,} celdas | {file_size:.1f} MB | {total_time:.1f} s"
                self.lbl_result_summary.configure(text=summary_txt, text_color=IBK_GREEN)
                self.lbl_progress_details.configure(text="Consolidado listo para auditoria y analisis financiero.")
                self._log("="*60)
                self._log(f"[OK] PROCESAMIENTO COMPLETADO ({mode_lbl.upper()})")
                self._log(f"  Tiempo total:   {total_time:.1f} s ({total_time/60:.2f} min)")
                self._log(f"  Tamano archivo: {file_size:.2f} MB")
                self._log(f"  Ubicacion:      {res.get('output_path')}")
                self._log("="*60)

            elif ev_type == 'error':
                self.is_generating = False
                err_msg = event.get('message', 'Error no identificado')
                self.btn_generate.configure(state="normal", text="GENERAR REPORTE CONSOLIDADO", fg_color=IBK_GREEN)
                self.sys_status_pill.configure(text="• Error en Proceso", fg_color=IBK_RED)
                self.lbl_progress_details.configure(text=f"Error: {err_msg}")
                self._log(f"[ERROR CRITICO] {err_msg}")

        self.after(100, self._process_queue)

    def _open_excel(self):
        """Abre el archivo consolidado en Microsoft Excel"""
        if not self.output_file.exists():
            self._log(f"[AVISO] El archivo {self.output_file.name} aun no existe en disco.")
            return

        self._log(f"Abriendo {self.output_file.name} en Microsoft Excel...")
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(self.output_file)])
            elif sys.platform == "win32":
                os.startfile(str(self.output_file))
            else:
                subprocess.Popen(["xdg-open", str(self.output_file)])
        except Exception as e:
            self._log(f"[ERROR] Al abrir archivo: {e}")

    def _open_folder(self):
        """Abre la carpeta contenedora en el explorador del sistema"""
        folder = self.output_file.parent if self.output_file.exists() else self.db_dir
        self._log(f"Abriendo explorador en directorio: {folder.name}...")
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            elif sys.platform == "win32":
                subprocess.Popen(["explorer", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as e:
            self._log(f"[ERROR] Al abrir carpeta: {e}")


def main():
    app = InterbankApp()
    app.mainloop()


if __name__ == "__main__":
    main()

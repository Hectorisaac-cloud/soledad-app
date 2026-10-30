import base64
import calendar
import hashlib
import hmac
import html
import importlib
import io
import json
import os
import random
import re
import shutil
import tempfile
import unicodedata
import uuid
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import persistencia
import reporte_fotografico
import historial_pjp4
import sitios_operativos
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

# ── Configuración ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Soledad",
    page_icon="static/icon.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = "data"
DATABASE_DIR = "database"
EXCEL_FILE = os.path.join(DATABASE_DIR, "bitacora_operativa_soledad.xlsx")
PROGRAMA_MANTENIMIENTO_FILE = os.path.join(DATABASE_DIR, "programa_mantenimiento.xlsx")
PROGRAMA_NORTE_FILE = os.path.join(DATABASE_DIR, "programa_norte.xlsx")
PROGRAMA_SUR_FILE = os.path.join(DATABASE_DIR, "programa_sur.xlsx")
ZONAS_PROGRAMA = {
    "norte": {"etiqueta": "Norte", "archivo": PROGRAMA_NORTE_FILE},
    "sur": {"etiqueta": "Sur", "archivo": PROGRAMA_SUR_FILE},
}
COLUMNA_PROGRAMADO = "PROGRAMADO"
EXCEL_FILE_LEGACY = "bitacora_mantenimiento.xlsx"
SHEET_NAME = "Bitacora"
HOJAS_BITACORA = ("Bitacora", "Hoja5", "Hoja2")
FUENTES_BITACORA_SOLEDAD = (
    os.path.join(DATABASE_DIR, "bitacora_operativa_soledad.xlsx"),
    os.path.join(DATABASE_DIR, "bitacora_operativa_soledad.xlsx.xlsx"),
    os.path.join(DATA_DIR, "bitacora_soledad.xlsx"),
)
FECHA_MIN_BITACORA_SOLEDAD = date(2026, 10, 1)
EVIDENCIAS_DIR = "evidencias"
EVIDENCIA_FOTOGRAFICA_DIR = "Evidencia Fotografica"
EVIDENCIA_FOTOGRAFICA_FILE = "evidencia_fotografica.json"
EVIDENCIA_OCTUBRE_DIR = "evidencia_fotografica_octubre"
CATALOGOS_FILE = "catalogos.json"
PERMISOS_FILE = "permisos.json"
ASIGNACIONES_FILE = "asignaciones_diarias.json"
PROGRAMADOS_FILE = "programados_ops.json"
PROGRAMAS_EXCEL_DIR = "programas_excel"
PROGRAMA_OCTUBRE_DIR = "programa_mantenimientos_octubre"
CARPETAS_PROGRAMA_MTTO = (
    "PROGRAMAS MTTO OCTUBRE",
    PROGRAMA_OCTUBRE_DIR,
)
CONTROL_PROGRAMA_OCTUBRE_FILE = "control_programa_octubre.json"
MES_PROGRAMA_OCTUBRE = (2026, 10)
ESPECIALIDADES_OCTUBRE = ("Eléctrico", "Instrumentación", "UBM")
HOJAS_PROGRAMA_IGNORADAS = {
    "equipos",
    "programa",
    "informe",
    "seguimiento",
    "resumen",
    "velocimetros",
    "calculos semana",
}
EXTENSIONES_PROGRAMA_EXCEL = (".xlsx", ".xls", ".xlsm")
ALMACEN_FILE = "almacen_insumos.json"
UBICACIONES_FILE = "ubicaciones_equipos.json"
UBICACIONES_EXCEL_DIR = "ubicacion_equipo"
MAPA_DIR = "mapa"
MAPA_SOLEDAD_NOMBRE = "mapa_soledad.pdf"
MAPA_SOLEDAD_FILE = os.path.join(MAPA_DIR, MAPA_SOLEDAD_NOMBRE)
MODOS_MAPA = [
    "🗺️ Vista por Zonas",
    "⚠️ Próximos a Mantenimiento Preventivo",
    "🔴 Pozos Fuera de Operación",
    "📅 Estatus de Mantenimientos",
]
HISTORIAL_OT_DIR = "historial_ot"
BITACORA_PDF_DIR = "bitacora_pdf"
HISTORIAL_OT_RUTAS_EXTRA = [
    os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "RESIDENCIAS", "PROYECTO", "OT"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "RESIDENCIAS", "PROYECTO", "OT"),
]
PATRON_OT_DIARIO = re.compile(r"^(\d{2})(\d{2})(\d{4})$")
NOMBRES_DIA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
SOPORTE_FILE = "soporte_refacciones.json"
ACTIVIDADES_FILE = "actividades_cuadrillas.json"
PERSONAL_FILE = "personal_cuadrillas.json"
LISTA_ASISTENCIA_FILE = os.path.join(DATABASE_DIR, "lista_asistencia.xlsx")
CALENDARIO_ASISTENCIA_FILE = os.path.join(DATABASE_DIR, "calendario_asistencia.xlsx")
PERSONAL_OTRO = "Otro / captura libre"
COMBUSTIBLE_DIR = "control_combustible"
COMBUSTIBLE_FILE = os.path.join(COMBUSTIBLE_DIR, "CONTROL DE COMBUSTIBLE 2026.xlsx")
VEHICULOS_FILE = "vehiculos_ot.json"
TALLER_EVIDENCIAS_DIR = os.path.join("taller_vehicular", "evidencias")
INTERVALO_MTO_KM_DEFAULT = 10000
UMBRAL_ALERTA_KM_DEFAULT = 500
SEMAFORO_OPERATIVA = "operativa"
SEMAFORO_PRECAUCION = "precaucion"
SEMAFORO_DETENIDA = "detenida"
SEMAFORO_EXTERNO = "externo"
ETIQUETA_SEMAFORO_UNIDAD = {
    SEMAFORO_OPERATIVA: "Operativa",
    SEMAFORO_PRECAUCION: "Operativa con precaución",
    SEMAFORO_DETENIDA: "Detenida en base",
    SEMAFORO_EXTERNO: "En taller externo",
}
CLASE_SEMAFORO_UNIDAD = {
    SEMAFORO_OPERATIVA: "ok",
    SEMAFORO_PRECAUCION: "mixto",
    SEMAFORO_DETENIDA: "paro",
    SEMAFORO_EXTERNO: "ext",
}
TONO_ALERTA_SEMAFORO = {
    SEMAFORO_OPERATIVA: "ok",
    SEMAFORO_PRECAUCION: "warn",
    SEMAFORO_DETENIDA: "due",
    SEMAFORO_EXTERNO: "ext",
}
TONO_KPI_SEMAFORO = {
    SEMAFORO_OPERATIVA: "ok",
    SEMAFORO_PRECAUCION: "warn",
    SEMAFORO_DETENIDA: "danger",
    SEMAFORO_EXTERNO: "info",
}
TIPO_TALLER_PREVENTIVO = "preventivo"
TIPO_TALLER_CORRECTIVO = "correctivo"
TIPO_TALLER_EXTERNO = "externo"
ESTATUS_INGRESO_ABIERTO = "abierto"
ESTATUS_INGRESO_ESPERA = "espera_refaccion"
ESTATUS_INGRESO_TALLER = "en_taller"
ESTATUS_INGRESO_EXTERNO = "externo"
ESTATUS_INGRESO_CERRADO = "cerrado"
ETIQUETA_ESTATUS_INGRESO = {
    ESTATUS_INGRESO_ABIERTO: "Diagnóstico",
    ESTATUS_INGRESO_ESPERA: "Espera de refacción",
    ESTATUS_INGRESO_TALLER: "En taller local",
    ESTATUS_INGRESO_EXTERNO: "Servicio externo",
    ESTATUS_INGRESO_CERRADO: "Cerrada / liberada",
}
AUTORIZACION_EXTERNO = [
    "Pendiente de autorización",
    "Autorizada",
    "Rechazada",
    "En taller externo",
    "Retornada a base",
]
SOLICITUD_PENDIENTE = "pendiente"
SOLICITUD_SURTIDA = "surtida"
SOLICITUD_CANCELADA = "cancelada"
ESTATUS_ASIG_RUTA = "en_ruta"
ESTATUS_ASIG_CERRADA = "cerrada"
NIVELES_COMBUSTIBLE_UNIDAD = ["Vacío", "1/4", "1/2", "3/4", "Lleno"]
NIVELES_FLUIDOS_UNIDAD = ["Correctos", "Bajos", "Requieren atención"]
PROPOSITOS_ASIGNACION = [
    "Unidad de relevo (principal en taller)",
    "Jornada de campo",
    "Carga de combustible",
    "Traslado / comisión",
    "Apoyo a otra cuadrilla",
    "Otro",
]
PROPOSITO_CARGA_COMBUSTIBLE = "Carga de combustible"
MARCA_DEMO = "DEMO-SEPT2026"
PREFIJO_SECUENCIA_DEMO = "DEMO-"
UMBRAL_REGISTROS_DEMO = 12
PATRON_REGISTRO_PRUEBA = re.compile(
    r"prueba|personalizada de prueba|sin equipo asignado|"
    r"evidencia fotografica de prueba|limpieza de formulario|"
    r"nuevos campos de insumos|insumos con unidades|"
    r"se emcontro fuera de servicio|se restauro",
    re.IGNORECASE,
)
ESTATUS_OT_ABIERTA = "abierta"
ESTATUS_OT_PROCESO = "en_proceso"
ESTATUS_OT_CERRADA = "cerrada"
ORIGEN_PJP4 = "PJP4 (Propio)"
ORIGEN_OPS = "OPS (Proporcionado por el Cliente)"
ORIGENES_INSUMO = [ORIGEN_PJP4, ORIGEN_OPS]
TIPO_INSUMO_CONSUMIBLE = "Consumible"
TIPO_INSUMO_ROTATIVO = "Componente Rotativo / Reparable"
TIPOS_INSUMO = [TIPO_INSUMO_CONSUMIBLE, TIPO_INSUMO_ROTATIVO]
CONDICION_PIEZA_INSTALADA = ["Nuevo", "Reparado por PJP4"]
DESTINO_PIEZA_RETIRADA = [
    "Enviada a Taller de Reparación",
    "Baja definitiva (Chatarra)",
]
INSUMOS_CATALOGO_DEFAULT = [
    {
        "id": "aceite-sae-50",
        "nombre": "Aceite SAE 50",
        "unidad": "L",
        "columna_excel": "ACEITE (LTS)",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "refrigerante",
        "nombre": "Refrigerante",
        "unidad": "L",
        "columna_excel": "REFRIGERANTE (LTS)",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "grasa",
        "nombre": "Grasa",
        "unidad": "Kg",
        "columna_excel": "GRASA (KG)",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "filtro-de-aceite",
        "nombre": "Filtro de aceite",
        "unidad": "Pzas",
        "columna_excel": "FILTROS (PZAS)",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "bujias",
        "nombre": "Bujías",
        "unidad": "Pzas",
        "columna_excel": "BUJÍAS (PZAS)",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "banda-c-115",
        "nombre": "Banda C-115",
        "unidad": "Pzas",
        "columna_excel": "",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "banda-c-112",
        "nombre": "Banda C-112",
        "unidad": "Pzas",
        "columna_excel": "",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "banda-c-240",
        "nombre": "Banda C-240",
        "unidad": "Pzas",
        "columna_excel": "",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "banda-micro-v",
        "nombre": "Banda Micro V",
        "unidad": "Pzas",
        "columna_excel": "",
        "tipo": TIPO_INSUMO_CONSUMIBLE,
        "precargado": True,
    },
    {
        "id": "bomba-de-agua",
        "nombre": "Bomba de agua",
        "unidad": "Pzas",
        "columna_excel": "",
        "tipo": TIPO_INSUMO_ROTATIVO,
        "precargado": True,
    },
    {
        "id": "alternador",
        "nombre": "Alternador",
        "unidad": "Pzas",
        "columna_excel": "ALTERNADOR CANTIDAD",
        "tipo": TIPO_INSUMO_ROTATIVO,
        "precargado": True,
    },
    {
        "id": "marcha-motor-arranque",
        "nombre": "Marcha / Motor de arranque",
        "unidad": "Pzas",
        "columna_excel": "",
        "tipo": TIPO_INSUMO_ROTATIVO,
        "precargado": True,
    },
]
UNIDADES_INSUMO = ["L", "Kg", "Pzas", "Rollos", "Latas", "Mts", "Otra"]
MAX_FOTOS = 9
MIN_FOTOS_COMPLETO = 1
COLUMNAS_FOTOS = [f"FOTO {i}" for i in range(1, MAX_FOTOS + 1)]
EVIDENCIA_PENDIENTE = "Pendiente"
EVIDENCIA_COMPLETA = "Completa"

ROL_OPERATIVO = "Operativo (Campo)"
ROL_ADMIN = "Administrativo"
MODULO_UBICACIONES = "Ubicación y Equipos"
MODULO_CUADRILLAS = "Historial de Cuadrillas"
ROLES = [ROL_OPERATIVO, ROL_ADMIN]
DESTINOS = [ROL_OPERATIVO, ROL_ADMIN, MODULO_UBICACIONES, MODULO_CUADRILLAS]
ESTADOS_REVISION = ["Pendiente", "Validado", "Con observaciones"]
ESTATUS_PROGRAMADO = ["Cumplido", "No Cumplido", "Reprogramado"]
NOMBRES_MES = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]
TIPOS_MANTENIMIENTO_PROGRAMADO = {
    "Mantenimiento bimestral a MCI",
    "Mantenimiento bimestral a UBM",
    "Mantenimiento semestral a MCI",
    "Mantenimiento semestral a UBM",
}
TIPOS_MANTENIMIENTO_NO_PROGRAMADO = {
    "Apoyo a puesta en marcha a MCI y UBM",
    "Habilitado de MCI y UBM",
    "Recuperación de componentes",
    "Restablecimiento a MCI",
    "Restablecimiento de ICI GR-20",
}
PATRON_NO_PROGRAMADO = re.compile(
    r"paro|imprevisto|correctiv|emergenc|falla|descompuest|"
    r"no arranque|fuera de servicio|urgente|contingenc",
    re.IGNORECASE,
)

SECTOR_OTRO = "Otro (Especificar)"
OPCION_UBICACION_NUEVA = "Otra / Nuevo equipo"
INSUMOS_REGISTRO_COMUNES = [
    "Aceite SAE 50",
    "Refrigerante",
    "Grasa",
    "Filtro de aceite",
    "Bujías",
]
SECTORES_DEFAULT = [
    "Almacen OPS (AOPS)",
    "Aragon (A)",
    "Base PJP4",
    "Bateria de Separacion Soledad I (BSI)",
    "Bateria de Separacion Soledad II (BSII)",
    "Bateria de Separacion Soledad Norte (BSNI)",
    "Bateria Palo Blanco (BPB)",
    "Coyotes (Coy)",
    "Estacion de Compresion Soledad II (ECSII)",
    "Estacion de Compresion Soledad Norte (ECSN)",
    "Gallo (GA)",
    "Palo Blanco (PB)",
    "Soledad (S)",
    "Soledad Norte (SN)",
    SECTOR_OTRO,
]
SECTORES_SIN_POZO_DEFAULT = [
    "Almacen OPS (AOPS)",
    "Base PJP4",
    "Bateria de Separacion Soledad I (BSI)",
    "Bateria de Separacion Soledad II (BSII)",
    "Bateria de Separacion Soledad Norte (BSNI)",
    "Bateria Palo Blanco (BPB)",
    "Estacion de Compresion Soledad II (ECSII)",
    "Estacion de Compresion Soledad Norte (ECSN)",
    "Base Soledad (BSE PJP4)",
    "Modulo I Soledad Norte (MOD I)",
    "Modulo II Soledad Norte (MOD II)",
    "Modulo III Soledad Norte (MOD III)",
    "Modulo IV Soledad Norte (MOD IV)",
]
CUADRILLAS = [
    "Electricos",
    "Instrumentos",
    "Mecanico 1",
    "Mecanico 2",
    "Mecanico 3",
    "Mecanico 4",
    "Montaje",
    "Soldadura",
    "Retro excavadora",
]
ORDEN_CUADRILLAS_OFICIAL = [
    "Eléctricos",
    "Instrumentos",
    "Mecánico 1",
    "Mecánico 2",
    "Mecánico 3",
    "Mecánico 4",
    "Montaje",
    "Soldadura",
    "Retro Excavadora",
]
UBICACIONES_ESTANDAR = [
    "BSI",
    "BSII",
    "BSNI",
    "MOD I",
    "MOD II",
    "MOD III",
    "MOD IV",
    "BPB",
    "ECSII",
    "ECSN",
    "AOPS",
    "Base PJP4",
    "BSE PJP4",
    "Soledad (S)",
    "Soledad Norte (SN)",
    "Aragon (A)",
    "Gallo (GA)",
    "Palo Blanco (PB)",
    "Coyotes (Coy)",
]
CUADRILLAS_ASIGNACION = [
    "Eléctrico",
    "Instrumentos",
    "Mecánico 1",
    "Mecánico 2",
    "Mecánico 3",
    "Mecánico 4",
    "Montaje",
    "Soldadura",
    "Retroexcavadora",
]
ESTATUS_SALE_CAMPO = "Sale a campo"
ESTATUS_EN_BASE = "No sale a campo / En base"
ESTATUS_SALIDA_CUADRILLA = [ESTATUS_SALE_CAMPO, ESTATUS_EN_BASE]
ESTATUS_OPCIONES = ["Operativo", "Fuera de servicio", "En mantenimiento"]
BANDAS_PROVEEDOR = ["PJP4", "OPS"]
ALTERNADOR_ORIGEN = [
    "Reparado por PJP4",
    "Proporcionado por PJP4",
    "Otro (Especificar)",
]

GRUPOS_ACTIVIDAD = ["ELECTRICO", "INSTRUMENTOS", "MECANICOS", "MONTAJE", "SOLDADURA"]
ETIQUETAS_GRUPO_ACTIVIDAD = {
    "ELECTRICO": "ELÉCTRICO",
    "INSTRUMENTOS": "INSTRUMENTOS",
    "MECANICOS": "MECÁNICOS",
    "MONTAJE": "MONTAJE",
    "SOLDADURA": "SOLDADURA",
}
CATALOGO_ACTIVIDADES_DEFAULT = {
    "ELECTRICO": [
        {"nombre": "MANTENIMIENTO A SISTEMA DE TIERRA FISICA A SKID INHIBIDOR DE CORROSION", "corto": "MTTO TIERRA FISICA SKID ICI"},
        {"nombre": "MANTENIMIENTO A SISTEMA DE TIERRA FISICA A MCI", "corto": "MTTO TIERRA FISICA MCI"},
        {"nombre": "MANTENIMIENTO A TABLERO ELECTRICO DE MCI", "corto": "MTTO TABLERO MCI"},
        {"nombre": "MANTENIMIENTO A TABLERO ELECTRICO DE SKID INHIBIDOR DE CORROSION", "corto": "MTTO TABLERO SKID ICI"},
        {"nombre": "MANTENIMIENTO A SISTEMA DE ARRANQUE DE MCI", "corto": "MTTO ARRANQUE MCI"},
        {"nombre": "REVISION Y APRETADO DE CONEXIONES ELECTRICAS", "corto": "REV CONEXIONES ELEC"},
        {"nombre": "CAMBIO DE CABLE DE FUERZA O CONTROL", "corto": "CAMBIO CABLE"},
        {"nombre": "INSTALACION O REPOSICION DE TERMINALES Y CLEMAS", "corto": "TERM / CLEMAS"},
        {"nombre": "MEDICION DE AISLAMIENTO DE MOTORES Y CABLES", "corto": "MEDICION AISLAMIENTO"},
        {"nombre": "MANTENIMIENTO A ILUMINACION DE INSTALACION", "corto": "MTTO ILUMINACION"},
        {"nombre": "REVISION A SISTEMA DE PARO DE EMERGENCIA", "corto": "REV PARO EMERGENCIA"},
        {"nombre": "APOYO ELECTRICO A PUESTA EN MARCHA", "corto": "APOYO PM ELECTRICO"},
    ],
    "INSTRUMENTOS": [
        {"nombre": "MANTENIMIENTO A INSTRUMENTACION DE SKID INHIBIDOR DE CORROSION", "corto": "MTTO INST SKID ICI"},
        {"nombre": "MANTENIMIENTO A TABLERO DE CONTROL DE ICI GR-20", "corto": "MTTO TABLERO ICI GR-20"},
        {"nombre": "CALIBRACION DE TRANSMISORES DE PRESION Y TEMPERATURA", "corto": "CALIB TX P/T"},
        {"nombre": "REVISION DE SENSORES Y SWITCHES DE MCI", "corto": "REV SENSORES MCI"},
        {"nombre": "MANTENIMIENTO A SISTEMA DE CONTROL DE UBM", "corto": "MTTO CONTROL UBM"},
        {"nombre": "VERIFICACION DE LAZOS DE CONTROL E INSTRUMENTACION", "corto": "VERIF LAZOS CONTROL"},
        {"nombre": "CAMBIO O AJUSTE DE INSTRUMENTOS DE CAMPO", "corto": "CAMBIO / AJUSTE INST"},
        {"nombre": "REVISION A SISTEMA DE MEDICION", "corto": "REV MEDICION"},
        {"nombre": "APOYO A PUESTA EN MARCHA DE INSTRUMENTACION", "corto": "APOYO PM INST"},
    ],
    "MECANICOS": [
        {"nombre": "Apoyo a puesta en marcha a MCI y UBM", "corto": "APOYO PM MCI/UBM"},
        {"nombre": "Habilitado de MCI y UBM", "corto": "HAB MCI/UBM"},
        {"nombre": "Inspección y suministro", "corto": "INSP Y SUMINISTRO"},
        {"nombre": "Mantenimiento bimestral a MCI", "corto": "MTTO BIM MCI"},
        {"nombre": "Mantenimiento bimestral a UBM", "corto": "MTTO BIM UBM"},
        {"nombre": "Mantenimiento semestral a MCI", "corto": "MTTO SEM MCI"},
        {"nombre": "Mantenimiento semestral a UBM", "corto": "MTTO SEM UBM"},
        {"nombre": "Recuperación de componentes", "corto": "RECUP COMPONENTES"},
        {"nombre": "Restablecimiento a MCI", "corto": "REST MCI"},
        {"nombre": "Restablecimiento de ICI GR-20", "corto": "REST ICI GR-20"},
        {"nombre": "Revisión a ICI- GR-20", "corto": "REV ICI GR-20"},
        {"nombre": "Revisión a MCI", "corto": "REV MCI"},
    ],
    "MONTAJE": [
        {"nombre": "MONTAJE Y DESMONTAJE DE COMPONENTES DE MCI", "corto": "MONTAJE / DESM MCI"},
        {"nombre": "MONTAJE Y DESMONTAJE DE COMPONENTES DE UBM", "corto": "MONTAJE / DESM UBM"},
        {"nombre": "MONTAJE DE SKID O BASE DE EQUIPO", "corto": "MONTAJE SKID / BASE"},
        {"nombre": "ALINEACION DE EQUIPO", "corto": "ALINEACION"},
        {"nombre": "IZAJE Y POSICIONAMIENTO DE COMPONENTES", "corto": "IZAJE / POSICION"},
        {"nombre": "INSTALACION DE SOPORTERIA", "corto": "SOPORTERIA"},
        {"nombre": "APOYO A HABILITADO DE MCI Y UBM", "corto": "APOYO HAB MCI/UBM"},
    ],
    "SOLDADURA": [
        {"nombre": "SOLDADURA ESTRUCTURAL EN INSTALACION", "corto": "SOLD ESTRUCTURAL"},
        {"nombre": "SOLDADURA EN TUBERIA Y ACCESORIOS", "corto": "SOLD TUBERIA"},
        {"nombre": "REPARACION POR SOLDADURA DE COMPONENTES", "corto": "REPARACION SOLD"},
        {"nombre": "CORTE Y DESBASTE DE MATERIAL", "corto": "CORTE / DESBASTE"},
        {"nombre": "FABRICACION DE SOPORTES O HERRAJES", "corto": "FAB SOPORTES"},
        {"nombre": "SOLDADURA EN BASES DE EQUIPO", "corto": "SOLD BASES EQUIPO"},
    ],
}
TIPOS_ACTIVIDAD_MECANICO = [item["nombre"] for item in CATALOGO_ACTIVIDADES_DEFAULT["MECANICOS"]] + ["Otra"]

EQUIPOS_MCI_DEFAULT = [
    "MCI tipo Arrow",
    "MCI tipo Weichai",
    "MCI tipo Vortec V6",
    "MCI tipo Vortec V8",
    "MCI tipo Cummins",
]
EQUIPOS_UB_DEFAULT = [
    "UBM tipo Weatherford",
    "UBM tipo Lufkin",
    "HLU (Unidad Hidráulica)",
]
EQUIPOS_OTROS_DEFAULT = [
    "ICI GR-20",
    "Ninguno",
]
ZONAS_UBICACION_DEFAULT = [
    {"id": "soledad", "prefijo": "PS", "nombre": "Pozos Soledad", "sector": "Soledad (S)"},
    {"id": "soledad-norte", "prefijo": "PSN", "nombre": "Pozos Soledad Norte", "sector": "Soledad Norte (SN)"},
    {"id": "aragon", "prefijo": "PA", "nombre": "Pozos Aragón", "sector": "Aragon (A)"},
    {"id": "gallo", "prefijo": "PGA", "nombre": "Pozos Gallo", "sector": "Gallo (GA)"},
    {"id": "palo-blanco", "prefijo": "PPB", "nombre": "Pozos Palo Blanco", "sector": "Palo Blanco (PB)"},
    {"id": "coyotes", "prefijo": "PC", "nombre": "Pozos Coyotes", "sector": "Coyotes (Coy)"},
    {"id": "instalaciones", "prefijo": "INS", "nombre": "Instalaciones y baterías", "sector": ""},
    {"id": "infra-norte", "prefijo": "NTE", "nombre": "Infraestructura Norte", "sector": "Norte (NTE)"},
    {"id": "infra-sur", "prefijo": "SUR", "nombre": "Infraestructura Sur", "sector": "Sur (SUR)"},
    {"id": "otros", "prefijo": "OTR", "nombre": "Otras ubicaciones", "sector": ""},
]
ZONAS_HISTORIAL_POZOS = [
    {"id": "soledad", "titulo": "PS Soledad", "icono": "🛢️", "detalle": "Pozos Soledad"},
    {"id": "aragon", "titulo": "PA Aragón", "icono": "🏗️", "detalle": "Pozos Aragón"},
    {"id": "gallo", "titulo": "PGA Gallo", "icono": "⚙️", "detalle": "Pozos Gallo"},
    {"id": "soledad-norte", "titulo": "PSN", "icono": "🧭", "detalle": "Pozos Soledad Norte"},
    {"id": "palo-blanco", "titulo": "PPB", "icono": "🌲", "detalle": "Pozos Palo Blanco"},
    {"id": "coyotes", "titulo": "PC", "icono": "📍", "detalle": "Pozos Coyotes"},
]
SECCIONES_OPERATIVAS = (
    {
        "id": "soledad",
        "prefijo": "PS",
        "etiqueta": "PS Soledad (Soledad S)",
        "aliases": ("PS SOLEDAD", "SOLEDAD S", "SOLEDAD (S)", "POZOS SOLEDAD", "SOLEDAD"),
    },
    {
        "id": "aragon",
        "prefijo": "PA",
        "etiqueta": "PA Aragón (Aragon A)",
        "aliases": ("PA ARAGON", "ARAGON A", "ARAGON (A)", "ARAGÓN A", "ARAGÓN (A)", "POZOS ARAGON", "POZOS ARAGÓN", "ARAGON", "ARAGÓN"),
    },
    {
        "id": "gallo",
        "prefijo": "PGA",
        "etiqueta": "PGA Gallo (Gallo GA)",
        "aliases": ("PGA GALLO", "GALLO GA", "GALLO (GA)", "POZOS GALLO", "GALLO"),
    },
    {
        "id": "soledad-norte",
        "prefijo": "PSN",
        "etiqueta": "PSN (Soledad Norte SN)",
        "aliases": ("PSN", "SOLEDAD NORTE SN", "SOLEDAD NORTE (SN)", "POZOS SOLEDAD NORTE", "SOLEDAD NORTE"),
    },
    {
        "id": "palo-blanco",
        "prefijo": "PPB",
        "etiqueta": "PPB (Palo Blanco PB)",
        "aliases": ("PPB", "PALO BLANCO PB", "PALO BLANCO (PB)", "POZOS PALO BLANCO", "PALO BLANCO"),
    },
    {
        "id": "coyotes",
        "prefijo": "PC",
        "etiqueta": "PC (Coyotes Coy)",
        "aliases": ("PC", "COYOTES COY", "COYOTES (COY)", "POZOS COYOTES", "COYOTES"),
    },
)
OPCIONES_SECCION_REGISTRO = (
    "PS Soledad",
    "PA Aragón",
    "PGA Gallo",
    "PSN",
    "PPB",
    "PC",
)
PREFIJOS_ZONA_CAPTURA = {
    "PS Soledad": "PS",
    "PA Aragón": "PA",
    "PGA Gallo": "PGA",
    "PSN": "PSN",
    "PPB": "PPB",
    "PC": "PC",
}
MACRO_AOPS = "AOPS"
MACRO_BASE = "Base PJP4"
CLAVE_BASE_PJP4 = "BSE PJP4"
_IDS_ZONA_POZO_CAPTURA = tuple(item["id"] for item in ZONAS_HISTORIAL_POZOS)
_IDS_BLOQUE_CAPTURA = ("instalaciones", "otros", "infra-norte", "infra-sur")
DIAS_SEMANA_CAL = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]
POZOS_SEMILLA = [
    ("PS-1", "soledad"),
    ("PS-2", "soledad"),
    ("PS-5", "soledad"),
    ("PS-8", "soledad"),
    ("PS-12", "soledad"),
    ("PSN-1", "soledad-norte"),
    ("PSN-2", "soledad-norte"),
    ("PA-1", "aragon"),
    ("PA-3", "aragon"),
    ("PA-7", "aragon"),
    ("PGA-1", "gallo"),
    ("PGA-2", "gallo"),
    ("PGA-4", "gallo"),
    ("PPB-1", "palo-blanco"),
    ("PPB-2", "palo-blanco"),
    ("PC-1", "coyotes"),
    ("PC-2", "coyotes"),
    ("BSI", "instalaciones"),
    ("BSII", "instalaciones"),
    ("BSNI", "instalaciones"),
    ("MOD I", "instalaciones"),
    ("MOD II", "instalaciones"),
    ("MOD III", "instalaciones"),
    ("MOD IV", "instalaciones"),
    ("BPB", "instalaciones"),
    ("ECSII", "instalaciones"),
    ("ECSN", "instalaciones"),
    ("AOPS", "instalaciones"),
    ("Base PJP4", "instalaciones"),
    ("BSE PJP4", "instalaciones"),
]
INSTALACIONES_SECTOR = {
    "BSI": "Bateria de Separacion Soledad I (BSI)",
    "BSII": "Bateria de Separacion Soledad II (BSII)",
    "BSNI": "Bateria de Separacion Soledad Norte (BSNI)",
    "MOD I": "Modulo I Soledad Norte (MOD I)",
    "MOD II": "Modulo II Soledad Norte (MOD II)",
    "MOD III": "Modulo III Soledad Norte (MOD III)",
    "MOD IV": "Modulo IV Soledad Norte (MOD IV)",
    "BPB": "Bateria Palo Blanco (BPB)",
    "ECSII": "Estacion de Compresion Soledad II (ECSII)",
    "ECSN": "Estacion de Compresion Soledad Norte (ECSN)",
    "AOPS": "Almacen OPS (AOPS)",
    "Base PJP4": "Base PJP4",
    "BSE PJP4": "Base Soledad (BSE PJP4)",
}
ALIAS_FICHA_CANON = {
    "AOPS": ["ALMACEN OPS", "Almacen OPS", "ALMACÉN OPS", "Almacén OPS"],
    "BSNI": ["BSN", "BSSN1", "Bateria de Separacion Soledad Norte", "Batería de Separación Soledad Norte (BSNI)", "BAT. SOL. NORTE"],
    "BSI": ["GDO BSI", "Bateria de Separacion Soledad I", "BAT. SOL. I"],
    "BSII": ["GDO BSII", "Bateria de Separacion Soledad II", "BAT. SOL. II"],
    "BPB": ["GDO BPB", "Bateria Palo Blanco", "BAT. PALO BLANCO"],
    "ECSII": ["GDO ECSII", 'GDO ECSII 10"Ø', "Estacion de Compresion Soledad II", "ESTACION DE COMPRESORAS -2 SUR"],
    "ECSN": ["ECO SLN", "Estacion de Compresion Soledad Norte", "Estacion de Compresion Soledad Norte (ECSN)", "ESTACION DE COMPRESORAS SOLEDAD NORTE"],
    "MOD I": ["MODI", "MOD-I", "MOD 1", "MODULO I", "Módulo I", "MODULO-I"],
    "MOD II": ["MODII", "MOD-II", "MOD 2", "MODULO II", "Módulo II", "MODULO-II"],
    "MOD III": ["MODIII", "MOD-III", "MOD 3", "MODULO III", "Módulo III", "MODULO-III"],
    "MOD IV": ["MODIV", "MOD-IV", "MOD 4", "MODULO IV", "Módulo IV", "MODULO-IV"],
    "BASE PJP4": ["Base PJP4", "BASE PJP4"],
    "BSE PJP4": ["BSE PJP4", "BASE SOLEDAD", "Base Soledad", "BSE"],
}
_CANON_UBICACION_PLANO = {
    "ALMACENOPS": "AOPS",
    "MODI": "MOD I",
    "MOD1": "MOD I",
    "MODII": "MOD II",
    "MOD2": "MOD II",
    "MODIII": "MOD III",
    "MOD3": "MOD III",
    "MODIV": "MOD IV",
    "MOD4": "MOD IV",
    "BSSN1": "BSNI",
    "BSN": "BSNI",
    "GDOBPB": "BPB",
    "GDOECSII": "ECSII",
    "GDOECSII10": "ECSII",
    "ECSN": "ECSN",
    "GDOECSN": "ECSN",
    "GDOBSI": "BSI",
    "GDOBSII": "BSII",
    "BSEPJP4": "BSE PJP4",
    "BASESOLEDAD": "BSE PJP4",
    "BATSOLNORTE": "BSNI",
    "BATSOLI": "BSI",
    "BATSOLII": "BSII",
    "BATPALOBLANCO": "BPB",
}
_PREFIJO_POZO_CORTO = {"PB": "PPB", "PG": "PGA"}
_RE_POZO_NOMBRE = [
    (re.compile(r"POZO\s+SOLEDAD\s+NORTE\s+(\d+)", re.I), "PSN"),
    (re.compile(r"POZO\s+SOLEDAD\s+(\d+)", re.I), "PS"),
    (re.compile(r"\bSOLEDAD\s+(\d+)\b", re.I), "PS"),
    (re.compile(r"POZO\s+GALLO\s+(\d+)", re.I), "PGA"),
    (re.compile(r"POZO\s+COY(?:OTES?)?\s+(\d+)", re.I), "PC"),
    (re.compile(r"POZO\s+ARAG[OÓ]N\s+(\d+)", re.I), "PA"),
    (re.compile(r"POZO\s+PALO\s+BLANCO\s+(\d+)", re.I), "PPB"),
]
_RE_TOKEN_INFRA = re.compile(
    r"(MPSN|MPS|MPG|BSMIV|BSMIII|BSMI|BSPB|BSSII|BSSI|ECSII|ECSN|TDSGF|TDSGC|TDSGB|TDSGD|TDSN)\d*",
    re.I,
)
_ALIAS_SEG_INFRA = {
    "SEGFE": "CBN-TDSGF-IQ-SIN-13",
    "SEGF": "CBN-TDSGF-IQ-SIN-13",
    "SEGAC": "CBN-TDSGC-IQ-SIN-10",
    "SEGBH": "CBN-TDSGB-IQ-SIN-11",
    "SEGDG": "CBN-TDSGD-IQ-SIN-12",
    "CBNKM4762": "CBN-TDSGC-IQ-SIN-10",
}
# Código exacto, zona NTE/SUR, clase y descripción oficial.
# El tramo sin la clase (TDSGF-IQ-SIN-13) es la clave que suele venir en la bitácora de Montaje.
INFRAESTRUCTURA_OFICIAL = (
    ("GDO-BSPB-IQ-SIN-01", "SUR", "GDO", 'GDO BSE PLB-ECO SOL II 16"x4.436 km'),
    ("GDO-BSSI-IQ-SIN-02", "SUR", "GDO", 'GDO BSE SOLI-ECO SOL II 12"x1.967 Km'),
    ("GDO-BSSII-IQ-SIN-03", "SUR", "GDO", 'GDO BSE SOL II - ECO SOL II 12"x2.625 Km'),
    ("GDO-ECSII-IQ-SIN-04", "SUR", "GDO", 'GDO ECO SOL II-BSE SLN 10"x7.535 km'),
    ("GDO-BSMI-IQ-SIN-05", "NTE", "GDO", 'GDO BSE MOD I SLN - ECO SLN 10" x 2.644 Km'),
    ("GDO-TDSN-IQ-SIN-06", "NTE", "GDO", 'GDO TDE SLNTDR SOL 691_6 6"x3.000 Km'),
    ("GDO-ECSN-IQ-SIN-07", "NTE", "GDO", 'GDO ECO SLNMP SLN 31 6"x1.485 km'),
    ("CBN-TDSN31-IQ-SIN-08", "NTE", "CBN", 'CBN TDE SLN 31-TDR SEG A SLN 6"x8.358 Km'),
    ("CBN-TDSN31-IQ-SIN-09", "NTE", "CBN", 'CBN TDE SLN 31 - TDR SEG B SLN 6"x3.682 Km'),
    ("CBN-TDSGC-IQ-SIN-10", "SUR", "CBN", 'CBN TDE SEG C-TDR SEG A SOL 6" x 4.762 Km'),
    ("CBN-TDSGB-IQ-SIN-11", "SUR", "CBN", 'CBN TDE SEG B - TDR SEG H SOL 6"x5.584 Km'),
    ("CBN-TDSGD-IQ-SIN-12", "SUR", "CBN", 'CBN TDE SEG D-TDR SEG G SOL 6"x3.716 Km'),
    ("CBN-TDSGF-IQ-SIN-13", "SUR", "CBN", 'CBN TDE SEG F-TDR SEG E SOL 6"x6.057 Km'),
    ("OLD-BSSI-IQ-SIN-14", "SUR", "OLD", 'OLD BSE SOLI-BSE MIQ I_8 8"x21.300 Km Megietla'),
    ("OLD-BSSII-IQ-SIN-15", "SUR", "OLD", 'OLD BSE SOL II - ENT SOL I/MIQ 6" x 3.500 Km'),
    ("OLD-BSPB-IQ-SIN-16", "SUR", "OLD", 'OLD BSE PLB-ENT SOL I/MIQ 6"x3.750 Km'),
    ("OLD-BSMI-IQ-SIN-17", "NTE", "OLD", 'OLD BSE MOD I SLN-BSE SLN 6" x 2.629 Km'),
    ("GDO-BSMIV-IQ-SIN-18", "NTE", "GDO", 'GDO BSE MOD IV SLN-BSE SLN 4"Ø x 4.927 Km'),
    ("GDO-BSMIII-IQ-SIN-19", "NTE", "GDO", 'GDO BSE MOD III SLN - ENT MOD IV/SLN_4" x 1.776 Km'),
    ("OLG-MPG204-IQ-SIN-20", "NTE", "OLG", 'OLG MP GAL 204 - ENT 8" MP GAL 218 8" x 1.179 Km'),
    ("OLG-MPG218-IQ-SIN-21", "NTE", "OLG", 'OLG MP GAL 218-BSE MODI SLN 8"x2.889 Km'),
    ("OLG-MPG204-IQ-SIN-55", "NTE", "OLG", 'OLG MP GAL 204 - ENT 6" MP GAL 218 6" x 1.173 Km'),
    ("OLG-MPG218-IQ-SIN-56", "NTE", "OLG", 'OLG MP GAL 218-BSE MOD I SLN 6" x 2.889 Km'),
    ("OLG-MPS549-IQ-SIN-22", "SUR", "OLG", 'OLG MP SOL 549-BSE PLB 6-8" x 2.569 Km'),
    ("OLG-MPS539-IQ-SIN-23", "SUR", "OLG", 'OLG MP SOL 539-ENT 6" SOL 549/PLB_6" x 0.089 Km'),
    ("OLG-MPS529-IQ-SIN-24", "SUR", "OLG", 'OLG MP SOL 529-ENT 6" SOL 549/PLB_6" x 0.137 Km'),
    ("OLG-MPS509-IQ-SIN-25", "SUR", "OLG", 'OLG MP SOL 509-ENT 8" SOL 549/PLB_ 6" x 0.127 Km'),
    ("OLG-MPP286-IQ-SIN-26", "SUR", "OLG", 'OLG MP PLB 286-MP SOL 505 6" x 3.049 Km'),
    ("OLG-MPS505-IQ-SIN-27", "SUR", "OLG", 'OLG MP SOL 505-BSE PLB 6" x 1.600 Km'),
    ("OLG-MPSN301-IQ-SIN-28", "SUR", "OLG", 'OLG MP SLN 301-BSE SOL II 6" x 2.452 Km'),
    ("OLG-MPS594-IQ-SIN-29", "SUR", "OLG", 'OLG MP SOL 594-BSE SOL II 6"x1.130 Km'),
    ("OLG-MPS676-IQ-SIN-57", "SUR", "OLG", 'OLG MP SOL 676-BSE SOL II 4-6" x 3.307 Km'),
    ("OLG-MPS651-IQ-SIN-31", "SUR", "OLG", 'OLG MP SOL 651-ENT 6" SOL 676/SOL II 4"x 0.185 Km'),
    ("OLG-MPS794-IQ-SIN-32", "SUR", "OLG", 'OLG MP SOL 794-BSE SOL II 4" x 1.722 Km'),
    ("OLG-MPS738-IQ-SIN-33", "SUR", "OLG", 'OLG MP SOL 738-BSE SOL II 4" x 0.477 Km'),
    ("OLG-MPS128-IQ-SIN-34", "SUR", "OLG", 'OLG MP SOL 128-BSE SOLI 4" x 1.091 Km'),
    ("OLG-MPS129-IQ-SIN-35", "SUR", "OLG", 'OLG MP SOL 129-BSE SOLI 4" x 0.667 Km'),
    ("OLG-MPSN269-IQ-SIN-36", "SUR", "OLG", 'OLG MP SLN 269-ENT 8" GAL 218/MODI 4"x 0.634 Km'),
    ("OLG-MPSN241-IQ-SIN-37", "NTE", "OLG", 'OLG MP SLN 241-ENT 8" GAL 218/MOD I 4"x 0.628 Km'),
    ("OLG-MPSN207-IQ-SIN-38", "NTE", "OLG", 'OLG MP SLN 207 - ENT 8" GAL 218/MOD 1_4" x 0.734 Km'),
    ("OLG-MPSN227-IQ-SIN-39", "NTE", "OLG", 'OLG MP SLN 227-ENT 4-8" SLN 207/GAL 218 4" x 0.240 Km'),
    ("OLG-MPG281-IQ-SIN-40", "NTE", "OLG", 'OLG MP SLN 281 - ENT 8" GAL 218/MODI 4"x.226 Km'),
    ("OLG-MPSN261-IQ-SIN-58", "NTE", "OLG", 'OLG MP SLN 261-ENT 8" GAL 218/MOD I_4"x 0.014 Km'),
    ("OLG-MPSN269-IQ-SIN-59", "NTE", "OLG", 'OLG MP SLN 269-ENT 6" GAL 218/MODI 4"x 0.634 Km'),
    ("OLG-MPSN241-IQ-SIN-60", "NTE", "OLG", 'OLG MP SLN 241 - ENT 6" GAL 218/MOD I_4"x 0.628 Km'),
    ("OLG-MPSN207-IQ-SIN-61", "NTE", "OLG", 'OLG MP SLN 207 - ENT 6" GAL 218/MOD | 4" x 0.738 Km'),
    ("OLG-MPSN227-IQ-SIN-62", "NTE", "OLG", 'OLG MP SLN 227 - ENT 4-6" SLN 207/GAL 218 4" x 0.236 Km'),
    ("OLG-MPSN261-IQ-SIN-64", "NTE", "OLG", 'OLG MP SLN 261-ENT 6" GAL 218/MOD | 4"x 0.016 Km'),
    ("OLG-MPSN281-IQ-SIN-63", "NTE", "OLG", 'OLG MP SLN 281-ENT 6" GAL 218/MODI 4"x.224 Km'),
)
_CLASES_INFRAESTRUCTURA = {"GDO", "CBN", "OLD", "OLG"}

COLUMNAS_INSUMOS = [
    "ACEITE (LTS)",
    "FILTROS (PZAS)",
    "BUJÍAS (PZAS)",
    "GRASA (KG)",
    "REFRIGERANTE (LTS)",
    "CINTA AISLANTE (ROLLOS)",
    "LIMPIADOR DIELÉCTRICO (LATAS)",
    "CABLE (MTS)",
    "TERMINALES/CLEMAS (PZAS)",
    "TEFLÓN/SELLADOR (PZAS)",
    "SOLDADURA/ELECTRODOS (KG)",
    "DISCOS DE CORTE/DESBASTE (PZAS)",
    "CILINDROS GAS (PZAS)",
    "TORNILLERÍA/EMPAQUES (PZAS)",
    "DIÉSEL (LTS)",
    "ACEITE HIDRÁULICO (LTS)",
]

ETIQUETAS_INSUMO_SOPORTE = {
    "ACEITE (LTS)": "Aceite SAE-50 (L)",
    "FILTROS (PZAS)": "Filtro aceite (PZA)",
    "BUJÍAS (PZAS)": "Bujías (PZA)",
    "GRASA (KG)": "Grasa (KG)",
    "REFRIGERANTE (LTS)": "Anticongelante (L)",
    "CINTA AISLANTE (ROLLOS)": "Cinta aislante (ROLLO)",
    "LIMPIADOR DIELÉCTRICO (LATAS)": "Limpiador dieléctrico (LATA)",
    "CABLE (MTS)": "Cable (MTS)",
    "TERMINALES/CLEMAS (PZAS)": "Terminales (PZA)",
    "TEFLÓN/SELLADOR (PZAS)": "Teflón/sellador (PZA)",
    "SOLDADURA/ELECTRODOS (KG)": "Soldadura (KG)",
    "DISCOS DE CORTE/DESBASTE (PZAS)": "Discos (PZA)",
    "CILINDROS GAS (PZAS)": "Cilindros gas (PZA)",
    "TORNILLERÍA/EMPAQUES (PZAS)": "Tornillería (PZA)",
    "DIÉSEL (LTS)": "Diésel (L)",
    "ACEITE HIDRÁULICO (LTS)": "Aceite hidráulico (L)",
}
MESES_SOPORTE = [8, 9, 10, 11, 12]
UNIDADES_SOPORTE = ["PZA", "JGO", "TAMBO", "L", "KG", "ROLLO", "LATA", "MTS", "PAR"]

COLUMNAS_INSUMOS_NUEVOS = [
    "BANDAS CANTIDAD",
    "BANDAS DESCRIPCIÓN",
    "BANDAS PROPORCIONADO POR",
    "ALTERNADOR CANTIDAD",
    "ALTERNADOR ESTADO/ORIGEN",
    "ALTERNADOR DESCRIPCIÓN",
    "OTROS INSUMOS",
    "PIEZAS ROTATIVAS",
]

COLUMNAS_EXCEL = [
    "FECHA",
    "SECUENCIA",
    "SECTOR",
    "POZO",
    "CUADRILLA",
    "MECANICO TITULAR",
    "AYUDANTE TITULAR",
    "MECANICO EN TURNO",
    "AYUDANTE EN TURNO",
    "COBERTURA",
    "EQUIPO/UBICACIÓN",
    "TIPO DE ACTIVIDAD",
    "ACTIVIDAD CORTA",
    *COLUMNAS_INSUMOS,
    *COLUMNAS_INSUMOS_NUEVOS,
    "DESCRIPCIÓN DE LA ACTIVIDAD",
    "EVALUACIÓN DEL ESTATUS",
    *COLUMNAS_FOTOS,
    "FOTOS",
    "REVISIÓN",
    "OBSERVACIÓN REVISIÓN",
    "ID_REGISTRO",
    "EVIDENCIA",
    COLUMNA_PROGRAMADO,
]

COLUMNAS_NUMERICAS = set(COLUMNAS_INSUMOS) | {
    "BANDAS CANTIDAD",
    "ALTERNADOR CANTIDAD",
}

COLUMNAS_REVISION_TABLA = [
    "FECHA",
    "SECUENCIA",
    "SECTOR",
    "POZO",
    "CUADRILLA",
    "MECANICO TITULAR",
    "AYUDANTE TITULAR",
    "MECANICO EN TURNO",
    "AYUDANTE EN TURNO",
    "COBERTURA",
    "EQUIPO/UBICACIÓN",
    "TIPO DE ACTIVIDAD",
    "ACTIVIDAD CORTA",
    "DESCRIPCIÓN DE LA ACTIVIDAD",
    "EVALUACIÓN DEL ESTATUS",
    "REVISIÓN",
    "OBSERVACIÓN REVISIÓN",
]
COLUMNAS_BITACORA_SOLEDAD = [
    "FECHA",
    "CUADRILLA",
    "PERSONAL",
    "UBICACIÓN",
    "EQUIPO",
    "TIPO DE ACTIVIDAD",
    "ACTIVIDAD CORTA",
    "DESCRIPCION DE LA ACTIVIDAD",
    "ACEITE SAE 50 PJP4",
    "BANDAS C144 PJP4",
    "BOMBA DE AGUA REPARADA PJP4",
    "GRASA OPS",
    "REFRIGERANTE OPS",
    "COMENTARIOS",
    "AVANCE %",
]
COLUMNAS_CONSUMIBLE_SOLEDAD = [
    "ACEITE SAE 50 PJP4",
    "BANDAS C144 PJP4",
    "BOMBA DE AGUA REPARADA PJP4",
    "GRASA OPS",
    "REFRIGERANTE OPS",
]
MAPA_CONSUMIBLE_SOLEDAD = {
    "ACEITE (LTS)": "ACEITE SAE 50 PJP4",
    "BANDAS CANTIDAD": "BANDAS C144 PJP4",
    "GRASA (KG)": "GRASA OPS",
    "REFRIGERANTE (LTS)": "REFRIGERANTE OPS",
}

PERMISOS_DEFAULT = [
    {
        "usuario": "campo",
        "nombre": "Personal de campo",
        "rol": ROL_OPERATIVO,
        "activo": True,
    },
    {
        "usuario": "admin",
        "nombre": "Supervisión PJP4",
        "rol": ROL_ADMIN,
        "activo": True,
    },
]

# ── Estilos ────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
        .block-container { padding-top: 1.5rem; max-width: 1200px; }
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0f172a 0%, #1e3a5f 58%, #16382c 100%);
        }
        section[data-testid="stSidebar"] * {
            color: #e2e8f0 !important;
        }
        section[data-testid="stSidebar"] .stRadio label {
            font-weight: 600;
        }
        .panel-header {
            background: linear-gradient(135deg, #1e3a5f 0%, #2d6a4f 100%);
            color: white;
            padding: 1.4rem 1.6rem;
            border-radius: 12px;
            margin-bottom: 1.5rem;
        }
        .panel-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
        .panel-header p  { margin: 0.3rem 0 0; opacity: 0.85; font-size: 0.95rem; }
        .section-title {
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #1e3a5f;
            margin-bottom: 0.8rem;
            padding-bottom: 0.4rem;
            border-bottom: 2px solid #2d6a4f;
        }
        .foto-vacia {
            height: 180px;
            border: 2px dashed #cbd5e1;
            border-radius: 8px;
            background: #f8fafc;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #94a3b8;
            font-size: 0.85rem;
        }
        div[data-testid="stImage"] img {
            height: 180px;
            object-fit: cover;
            border: 3px solid #ffffff;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.18);
            background: #0f172a;
        }
        .rol-card {
            background: #0f172a;
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 1.2rem 1.3rem 1rem;
            min-height: 170px;
        }
        .rol-card h3 { margin: 0 0 0.4rem; color: #ffffff; }
        .rol-card p { margin: 0; opacity: 0.85; }
        .asignacion-board {
            background: #0f172a;
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 1.15rem 1.25rem 1rem;
            margin-bottom: 1.4rem;
        }
        .asignacion-board h2 {
            margin: 0;
            font-size: 1.12rem;
            color: #ffffff;
        }
        .asignacion-board .asignacion-meta {
            margin: 0.25rem 0 0.85rem;
            opacity: 0.8;
            font-size: 0.9rem;
        }
        .asignacion-item {
            background: #1e293b;
            border-left: 4px solid #2d6a4f;
            border-radius: 8px;
            padding: 0.7rem 0.85rem;
            margin-bottom: 0.55rem;
        }
        .asignacion-item .cuadrilla {
            font-size: 0.72rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }
        .asignacion-item p, .asignacion-ausencias p {
            margin: 0;
            white-space: pre-wrap;
            line-height: 1.4;
        }
        .asignacion-vacia { opacity: 0.55; font-style: italic; }
        .asignacion-ausencias {
            background: #1e3a5f;
            border-left: 4px solid #93c5fd;
            border-radius: 8px;
            padding: 0.7rem 0.85rem;
            margin-top: 0.35rem;
        }
        .asignacion-ausencias .cuadrilla { color: #93c5fd; }
        .asignacion-item.en-campo { border-left-color: #22c55e; }
        .asignacion-item.en-base { border-left-color: #f59e0b; }
        .salida-banner {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.9rem 1rem 0.85rem;
            margin: 0.2rem 0 1rem;
        }
        .salida-banner .salida-titulo {
            color: #ffffff;
            font-weight: 700;
            font-size: 0.95rem;
            margin-bottom: 0.55rem;
        }
        .salida-chips { display: flex; flex-wrap: wrap; gap: 0.45rem; }
        .salida-chip, .salida-mini {
            border-radius: 999px;
            border: 1px solid #334155;
            background: #1e293b;
            padding: 0.28rem 0.7rem;
            font-size: 0.8rem;
            font-weight: 700;
        }
        .salida-chip { display: inline-flex; gap: 0.4rem; align-items: baseline; }
        .salida-chip span { color: #cbd5e1; font-weight: 600; }
        .salida-chip.campo, .salida-mini.campo { color: #86efac; border-color: #166534; }
        .salida-chip.base, .salida-mini.base { color: #fcd34d; border-color: #92400e; }
        .salida-mini { display: inline-block; margin-bottom: 0.35rem; }
        [data-testid="stMetric"] {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 10px;
            padding: 0.55rem 0.8rem 0.7rem;
        }
        [data-testid="stMetricLabel"] { color: #cbd5e1; }
        [data-testid="stMetricValue"] { color: #86efac; }
        .metric-fuel {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 10px;
            padding: 0.55rem 0.8rem 0.65rem;
            min-height: 100%;
        }
        .metric-fuel .lbl { color: #cbd5e1; font-size: 0.9rem; line-height: 1.25; }
        .metric-fuel .val {
            color: #86efac; font-size: 1.55rem; font-weight: 600;
            line-height: 1.15; margin-top: 0.15rem;
        }
        .metric-fuel .sub {
            color: #e2e8f0; font-size: 1.15rem; font-weight: 600;
            line-height: 1.15; margin-top: 0.12rem;
        }
        .metric-fuel .delta { font-size: 0.82rem; margin-top: 0.15rem; }
        .metric-fuel .delta.down { color: #fca5a5; }
        .metric-fuel .delta.up { color: #86efac; }
        .metric-fuel .delta.flat { color: #94a3b8; }
        .metric-fuel .corte {
            color: #94a3b8; font-size: 0.72rem; margin-top: 0.4rem; letter-spacing: 0.01em;
        }
        .origen-pjp4 {
            background: #0f172a;
            border-left: 4px solid #86efac;
            border-radius: 10px;
            padding: 0.85rem 1rem;
            margin-bottom: 0.8rem;
            color: #e2e8f0;
        }
        .origen-ops {
            background: #0f172a;
            border-left: 4px solid #93c5fd;
            border-radius: 10px;
            padding: 0.85rem 1rem;
            margin-bottom: 0.8rem;
            color: #e2e8f0;
        }
        .origen-pjp4 h3, .origen-ops h3 {
            margin: 0 0 0.35rem;
            font-size: 1rem;
            color: #ffffff;
        }
        .ficha-hero {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 62%, #16382c 100%);
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 1.2rem 1.35rem 1.1rem;
            margin-bottom: 1.1rem;
        }
        .ficha-hero .prefijo {
            font-size: 0.72rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
            margin-bottom: 0.25rem;
        }
        .ficha-hero h2 {
            margin: 0;
            color: #ffffff;
            font-size: 1.45rem;
        }
        .ficha-hero p { margin: 0.35rem 0 0; opacity: 0.82; }
        .equipo-card {
            background: #0f172a;
            border: 1px solid #334155;
            border-left: 4px solid #2d6a4f;
            border-radius: 10px;
            padding: 0.9rem 1rem;
            color: #e2e8f0;
            min-height: 118px;
        }
        .equipo-card.vacio { border-left-color: #64748b; }
        .equipo-card .lbl {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
            margin-bottom: 0.35rem;
        }
        .equipo-card h3 { margin: 0; color: #ffffff; font-size: 1.05rem; }
        .equipo-card p { margin: 0.3rem 0 0; opacity: 0.75; font-size: 0.88rem; }
        .personal-chip {
            background: #1e293b;
            border-left: 4px solid #86efac;
            border-radius: 8px;
            padding: 0.65rem 0.8rem;
            margin-bottom: 0.5rem;
            color: #e2e8f0;
        }
        .personal-chip .cuadrilla {
            font-size: 0.72rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .personal-chip p { margin: 0.2rem 0 0; }
        .cobertura-card {
            background: #0f172a;
            border: 1px solid #334155;
            border-left: 4px solid #86efac;
            border-radius: 12px;
            padding: 0.8rem 0.95rem 0.75rem;
            margin: 0.45rem 0 0.7rem;
            color: #e2e8f0;
        }
        .cobertura-card.warn { border-left-color: #f59e0b; }
        .cobertura-card .kicker {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .cobertura-card.warn .kicker { color: #fbbf24; }
        .cobertura-card h4 { margin: 0.2rem 0 0.25rem; color: #ffffff; font-size: 0.98rem; }
        .cobertura-card p { margin: 0.15rem 0 0; color: #94a3b8; font-size: 0.84rem; }
        .explorer-meta {
            color: #94a3b8;
            font-size: 0.9rem;
            margin: -0.4rem 0 1.1rem;
        }
        .explorer-card {
            background: #0f172a;
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 1.15rem 1.15rem 0.85rem;
            min-height: 198px;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.28);
        }
        .explorer-card .icon-tile {
            width: 52px;
            height: 52px;
            border-radius: 12px;
            background: linear-gradient(135deg, #1e3a5f 0%, #16382c 100%);
            border: 1px solid #334155;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.55rem;
            margin-bottom: 0.75rem;
        }
        .explorer-card h3 {
            margin: 0 0 0.35rem;
            color: #ffffff;
            font-size: 1.05rem;
            font-weight: 700;
        }
        .explorer-card p {
            margin: 0;
            color: #94a3b8;
            font-size: 0.88rem;
            line-height: 1.4;
            min-height: 3.6em;
        }
        .explorer-crumb {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 10px;
            padding: 0.55rem 0.85rem;
            color: #cbd5e1;
            font-size: 0.88rem;
            margin-bottom: 0.85rem;
        }
        .explorer-crumb b { color: #86efac; }
        .ops-modo-ayuda {
            margin: 0.35rem 0 0;
            color: #cbd5e1;
            font-size: 0.92rem;
            line-height: 1.35;
        }
        .ops-modo-ayuda b { color: #86efac; }
        .cal-wrap {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 1rem 1.1rem 0.9rem;
            margin-bottom: 0.85rem;
        }
        .cal-head {
            color: #ffffff;
            font-weight: 700;
            font-size: 1.08rem;
            margin-bottom: 0.75rem;
        }
        .cal-table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 4px;
            table-layout: fixed;
        }
        .cal-table th {
            text-align: center;
            font-size: 0.7rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
            padding: 0.2rem 0 0.35rem;
        }
        .cal-table td {
            height: 52px;
            vertical-align: top;
        }
        .cal-grid {
            display: grid;
            grid-template-columns: repeat(7, 1fr);
            gap: 4px;
        }
        .cal-dow {
            text-align: center;
            font-size: 0.7rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
            padding: 0.2rem 0 0.35rem;
        }
        .cal-day {
            min-height: 52px;
            border-radius: 8px;
            background: #1e293b;
            border: 1px solid #334155;
            padding: 5px 6px 6px;
            color: #e2e8f0;
        }
        [class*="st-key-caldia_"] button {
            background: #1e293b;
            border: 1px solid #334155;
            color: #e2e8f0;
            min-height: 36px;
            padding: 2px 0;
            border-radius: 8px;
            font-weight: 700;
            font-size: 0.82rem;
        }
        [class*="st-key-caldia_"] button:hover { border-color: #86efac; color: #e2e8f0; }
        [class*="st-key-caldia_"] { margin-bottom: 0; }
        .cal-day.empty { background: transparent; border-color: transparent; }
        .cal-day.hoy { border-color: #86efac; box-shadow: 0 0 0 1px rgba(134, 239, 172, 0.35); }
        .cal-day.bitacora { background: rgba(20, 83, 45, 0.38); border-color: #2d6a4f; }
        .cal-day.programado { background: rgba(30, 58, 95, 0.72); border-color: #3b82f6; }
        .cal-day.bitacora.programado {
            background: linear-gradient(135deg, rgba(20, 83, 45, 0.45) 0%, rgba(30, 58, 95, 0.75) 100%);
            border-color: #86efac;
        }
        .cal-day.sel { outline: 2px solid #fbbf24; outline-offset: 1px; }
        .cal-day .num { font-size: 0.9rem; font-weight: 700; }
        .cal-day .dots { display: flex; gap: 4px; margin-top: 8px; min-height: 7px; }
        .dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            display: inline-block;
        }
        .dot.bit { background: #4ade80; }
        .dot.prog { background: #60a5fa; }
        .cal-legend {
            display: flex;
            gap: 0.95rem;
            margin-top: 0.75rem;
            color: #94a3b8;
            font-size: 0.78rem;
            flex-wrap: wrap;
            align-items: center;
        }
        .cal-banner {
            background: #1e3a5f;
            border: 1px solid #3b82f6;
            border-radius: 8px;
            padding: 0.45rem 0.75rem;
            color: #dbeafe;
            font-size: 0.82rem;
            margin-bottom: 0.7rem;
        }
        .hist-card {
            background: #0f172a;
            border: 1px solid #334155;
            border-left: 4px solid #2d6a4f;
            border-radius: 10px;
            padding: 0.8rem 0.95rem;
            margin-bottom: 0.65rem;
            color: #e2e8f0;
        }
        .hist-card.cambio { border-left-color: #f59e0b; }
        .hist-card.nprog { border-left-color: #f97316; }
        .hist-card.historico { border-left-color: #a78bfa; }
        .hist-card .hist-date {
            font-size: 0.72rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .hist-card h4 { margin: 0.2rem 0 0; color: #ffffff; font-size: 0.98rem; }
        .hist-card p { margin: 0.3rem 0 0; color: #94a3b8; font-size: 0.86rem; }
        .insumo-chip {
            display: inline-block;
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 999px;
            padding: 0.12rem 0.55rem;
            margin: 0.4rem 0.3rem 0 0;
            font-size: 0.75rem;
            color: #cbd5e1;
        }
        .prog-panel {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 1rem 1rem 0.75rem;
        }
        .prog-panel h3 { margin: 0 0 0.25rem; color: #ffffff; font-size: 1.05rem; }
        .prog-panel .hint { color: #94a3b8; font-size: 0.8rem; margin: 0 0 0.75rem; }
        .prog-item {
            background: #1e293b;
            border-left: 4px solid #3b82f6;
            border-radius: 8px;
            padding: 0.65rem 0.75rem;
            margin-bottom: 0.55rem;
            color: #e2e8f0;
        }
        .prog-item.cumplido { border-left-color: #22c55e; opacity: 0.86; }
        .prog-item.reprog { border-left-color: #f59e0b; }
        .prog-item.mes-activo { box-shadow: 0 0 0 1px rgba(59, 130, 246, 0.45); }
        .prog-item .mes {
            font-size: 0.7rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .prog-item h4 { margin: 0.15rem 0 0; color: #ffffff; font-size: 0.92rem; }
        .prog-item p { margin: 0.25rem 0 0; color: #94a3b8; font-size: 0.8rem; }
        .pozo-card {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.85rem 0.95rem 0.55rem;
            min-height: 132px;
            color: #e2e8f0;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.22);
        }
        .pozo-card .prefijo {
            font-size: 0.7rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .pozo-card h3 { margin: 0.2rem 0 0; color: #ffffff; font-size: 1.12rem; }
        .pozo-card p { margin: 0.3rem 0 0; color: #94a3b8; font-size: 0.82rem; line-height: 1.35; }
        .cuadrilla-dia {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.9rem 1rem 0.65rem;
            min-height: 148px;
            color: #e2e8f0;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.22);
        }
        .cuadrilla-dia.ok { border-left: 4px solid #22c55e; }
        .cuadrilla-dia.paro { border-left: 4px solid #ef4444; }
        .cuadrilla-dia.mixto { border-left: 4px solid #f59e0b; }
        .cuadrilla-dia.activa {
            border-color: #38bdf8;
            box-shadow: 0 0 0 1px #38bdf8, 0 8px 22px rgba(2, 6, 23, 0.22);
        }
        .semaforo {
            display: inline-block;
            width: 11px;
            height: 11px;
            border-radius: 50%;
            margin-right: 0.4rem;
            vertical-align: middle;
        }
        .semaforo.ok { background: #22c55e; box-shadow: 0 0 8px rgba(34, 197, 94, 0.55); }
        .semaforo.paro { background: #ef4444; box-shadow: 0 0 8px rgba(239, 68, 68, 0.55); }
        .semaforo.mixto { background: #f59e0b; box-shadow: 0 0 8px rgba(245, 158, 11, 0.5); }
        .semaforo.ext { background: #38bdf8; box-shadow: 0 0 8px rgba(56, 189, 248, 0.55); }
        .cuadrilla-dia .lbl {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .pdf-alerta {
            margin-top: 1.65rem;
            border-radius: 10px;
            padding: 0.55rem 0.75rem;
            font-size: 0.86rem;
            font-weight: 650;
            line-height: 1.35;
        }
        .pdf-alerta.ok {
            background: rgba(16, 185, 129, 0.16);
            color: #a7f3d0;
            border: 1px solid rgba(52, 211, 153, 0.45);
        }
        .pdf-alerta.falta {
            background: rgba(245, 158, 11, 0.16);
            color: #fde68a;
            border: 1px solid rgba(251, 191, 36, 0.5);
        }
        .cuadrilla-dia h3 { margin: 0.25rem 0 0; color: #ffffff; font-size: 1.08rem; }
        .cuadrilla-dia p { margin: 0.3rem 0 0; color: #94a3b8; font-size: 0.82rem; }
        .ot-sheet {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 1.15rem 1.2rem 1rem;
            color: #e2e8f0;
            margin-bottom: 1rem;
        }
        .ot-sheet .ot-kicker {
            font-size: 0.7rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .ot-sheet h3 { margin: 0.2rem 0 0.7rem; color: #ffffff; font-size: 1.15rem; }
        .ot-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 0.55rem 1rem;
            margin-bottom: 0.35rem;
        }
        .ot-field.ot-span { grid-column: 1 / -1; }
        .ot-sec {
            margin: 0.85rem 0 0.4rem;
            font-size: 0.72rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        [data-testid="stVerticalBlock"]:has(.ot-sec) {
            background: #0f172a;
            border: 1px solid #334155 !important;
            border-radius: 14px !important;
            padding: 0.85rem 1rem 0.55rem !important;
            margin: 0.35rem 0 0.9rem !important;
        }
        .ot-field .k {
            font-size: 0.68rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .ot-field .v { color: #e2e8f0; font-size: 0.9rem; margin-top: 0.12rem; }
        .mat-row {
            background: #1e293b;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 0.55rem 0.75rem;
            margin-bottom: 0.4rem;
            display: flex;
            justify-content: space-between;
            gap: 0.8rem;
            color: #e2e8f0;
        }
        .mat-row b { color: #ffffff; }
        .mat-row span { color: #86efac; font-weight: 700; white-space: nowrap; }
        .firma-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1rem;
            margin-top: 1.1rem;
        }
        .firma-box {
            border: 1px dashed #475569;
            border-radius: 10px;
            min-height: 110px;
            padding: 0.7rem 0.85rem 0.55rem;
            background: #111827;
        }
        .firma-box .k {
            font-size: 0.68rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .firma-linea {
            margin-top: 52px;
            border-top: 1px solid #64748b;
            padding-top: 0.35rem;
            color: #94a3b8;
            font-size: 0.78rem;
            text-align: center;
        }
        .foto-vacia.dark {
            height: 160px;
            border: 2px dashed #475569;
            border-radius: 8px;
            background: #111827;
            color: #94a3b8;
        }
        .foto-cap {
            color: #94a3b8;
            font-size: 0.75rem;
            margin: 0.25rem 0 0.7rem;
        }
        .money-card {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.85rem 1rem;
            color: #e2e8f0;
        }
        .money-card .k {
            font-size: 0.7rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .money-card h3 { margin: 0.25rem 0 0; color: #86efac; font-size: 1.2rem; }
        .kpi-card {
            background: linear-gradient(180deg, #132033 0%, #0f172a 100%);
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 0.95rem 1.05rem 0.85rem;
            min-height: 118px;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.28);
            color: #e2e8f0;
        }
        .kpi-card .k {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .kpi-card h3 {
            margin: 0.35rem 0 0.2rem;
            color: #f8fafc;
            font-size: 1.55rem;
            font-weight: 750;
            letter-spacing: -0.02em;
        }
        .kpi-card .s { margin: 0; color: #94a3b8; font-size: 0.8rem; }
        .kpi-card.ok h3 { color: #86efac; }
        .kpi-card.warn h3 { color: #fbbf24; }
        .kpi-card.info h3 { color: #7dd3fc; }
        .kpi-card.danger h3 { color: #f87171; }
        .kpi-card.activa {
            border-color: #2d6a4f;
            box-shadow: 0 0 0 1px #2d6a4f, 0 8px 22px rgba(2, 6, 23, 0.28);
        }
        .flota-vistazo {
            background: linear-gradient(160deg, #16324a 0%, #0f172a 62%);
            border: 1px solid #334155;
            border-radius: 16px;
            padding: 1.35rem 1.3rem 1.2rem;
            min-height: 292px;
            color: #e2e8f0;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.28);
        }
        .flota-vistazo .lbl {
            font-size: 0.72rem;
            letter-spacing: 0.1em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .flota-vistazo h3 {
            margin: 0.35rem 0 0.45rem;
            color: #ffffff;
            font-size: 1.28rem;
        }
        .flota-vistazo p {
            margin: 0;
            color: #94a3b8;
            font-size: 0.9rem;
            line-height: 1.45;
        }
        .flota-vistazo .silueta {
            margin: 0.2rem 0 0.85rem;
            color: #7dd3fc;
        }
        .flota-vistazo svg { width: 100%; max-width: 230px; height: auto; display: block; }
        .flota-nav-sep {
            margin: 1.25rem 0 0.55rem;
            border-top: 1px solid #1e293b;
        }
        .dash-panel {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 0.85rem 0.95rem 0.55rem;
            margin-bottom: 0.85rem;
        }
        .dash-panel h4 {
            margin: 0 0 0.35rem;
            color: #f8fafc;
            font-size: 0.98rem;
        }
        .dash-panel p { margin: 0 0 0.55rem; color: #94a3b8; font-size: 0.8rem; }
        .mapa-shell {
            background: #0f172a;
            color: #e2e8f0;
            border: 1px solid #334155;
            border-radius: 14px;
            padding: 1rem 1.05rem 0.85rem;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.28);
            margin-bottom: 0.9rem;
        }
        .mapa-shell h3 {
            margin: 0 0 0.3rem;
            color: #f8fafc;
            font-size: 1.08rem;
        }
        .mapa-shell p {
            margin: 0;
            color: #94a3b8;
            font-size: 0.88rem;
            line-height: 1.4;
        }
        .mapa-preview img {
            width: 100%;
            height: auto !important;
            object-fit: contain !important;
            display: block;
            border-radius: 8px;
            border: 1px solid #334155;
            background: #0f172a;
        }
        .mapa-leyenda {
            display: flex;
            gap: 1rem;
            flex-wrap: wrap;
            color: #94a3b8;
            font-size: 0.8rem;
            margin: 0.15rem 0 0.75rem;
        }
        .mapa-esquema {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(168px, 1fr));
            gap: 0.65rem;
        }
        .mapa-zona {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.7rem 0.75rem 0.65rem;
            min-height: 92px;
        }
        .mapa-zona-t {
            color: #93c5fd;
            font-size: 0.7rem;
            letter-spacing: 0.06em;
            text-transform: uppercase;
            font-weight: 700;
            margin-bottom: 0.45rem;
        }
        .mapa-chips { display: flex; flex-wrap: wrap; gap: 0.32rem; }
        .mapa-chip {
            border-radius: 999px;
            padding: 0.16rem 0.5rem;
            font-size: 0.72rem;
            font-weight: 700;
            border: 1px solid #334155;
            color: #e2e8f0;
        }
        .mapa-chip.ok { background: #14532d; border-color: #166534; color: #86efac; }
        .mapa-chip.mixto { background: #422006; border-color: #854d0e; color: #fbbf24; }
        .mapa-chip.paro { background: #450a0a; border-color: #7f1d1d; color: #fca5a5; }
        .mapa-chip.vacio { background: #1e293b; color: #64748b; font-weight: 500; }
        .mapa-pozo {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.85rem 0.95rem 0.7rem;
            min-height: 132px;
            color: #e2e8f0;
            box-shadow: 0 8px 22px rgba(2, 6, 23, 0.22);
        }
        .mapa-pozo.ok { border-left: 4px solid #22c55e; }
        .mapa-pozo.mixto { border-left: 4px solid #f59e0b; }
        .mapa-pozo.paro { border-left: 4px solid #ef4444; }
        .mapa-pozo h3 { margin: 0.15rem 0 0; color: #ffffff; font-size: 1.05rem; }
        .mapa-pozo p { margin: 0.28rem 0 0; color: #94a3b8; font-size: 0.82rem; line-height: 1.35; }
        .mapa-pozo .s { color: #64748b; font-size: 0.75rem; }
        .mapa-pozo .prefijo {
            font-size: 0.7rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
        }
        .ot-pagina {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.55rem;
            margin-top: 0.45rem;
        }
        .ot-pagina img {
            width: 100%;
            height: auto !important;
            object-fit: contain !important;
            display: block;
            border-radius: 8px;
        }
        .explorer-card .badge-count {
            display: inline-block;
            margin-top: 0.55rem;
            background: #16382c;
            border: 1px solid #334155;
            color: #86efac;
            font-size: 0.75rem;
            font-weight: 700;
            border-radius: 999px;
            padding: 0.15rem 0.55rem;
        }
        .explorer-card .badge-count.warn {
            background: #422006;
            border-color: #854d0e;
            color: #fbbf24;
        }
        .explorer-card .badge-count.due {
            background: #450a0a;
            border-color: #7f1d1d;
            color: #fca5a5;
        }
        .veh-alert {
            background: #0f172a;
            border: 1px solid #334155;
            border-left: 4px solid #22c55e;
            border-radius: 12px;
            padding: 0.9rem 1rem;
            margin-bottom: 0.7rem;
            color: #e2e8f0;
        }
        .veh-alert.warn { border-left-color: #f59e0b; }
        .veh-alert.due { border-left-color: #ef4444; }
        .veh-alert.ext { border-left-color: #38bdf8; }
        .veh-alert .kicker {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .veh-alert.warn .kicker { color: #fbbf24; }
        .veh-alert.due .kicker { color: #fca5a5; }
        .veh-alert.ext .kicker { color: #7dd3fc; }
        .veh-alert h3 { margin: 0.2rem 0 0.45rem; color: #ffffff; font-size: 1.05rem; }
        .veh-alert p { margin: 0.15rem 0 0; color: #94a3b8; font-size: 0.86rem; }
        .semaforo.ext { background: #38bdf8; box-shadow: 0 0 8px rgba(56, 189, 248, 0.55); }
        .taller-fase {
            font-size: 0.72rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #93c5fd;
            font-weight: 700;
            margin: 0.35rem 0 0.25rem;
        }
        .hangar-shell { margin-bottom: 0.35rem; }
        .hangar-brand {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 58%, #16382c 100%);
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.85rem 1rem;
            color: #e2e8f0;
            margin-bottom: 0.75rem;
        }
        .hangar-brand .prefijo {
            font-size: 0.7rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .hangar-brand h1 {
            margin: 0.15rem 0 0;
            color: #ffffff;
            font-size: 1.35rem;
        }
        .hangar-brand p { margin: 0.2rem 0 0; color: #94a3b8; font-size: 0.86rem; }
        .hangar-top-meta {
            color: #e2e8f0;
            font-size: 0.92rem;
            padding-top: 0.55rem;
        }
        .hangar-online {
            text-align: center;
            padding-top: 0.55rem;
            color: #e2e8f0;
            font-size: 0.92rem;
        }
        .hangar-online b { color: #2ecc71; }
        .hangar-card {
            background: #1a1c24;
            border: 1px solid #2d3139;
            border-left: 5px solid #2ecc71;
            border-radius: 10px;
            padding: 0.9rem 1rem;
            color: #e2e8f0;
            min-height: 148px;
            margin-bottom: 0.35rem;
        }
        .hangar-card.warn { border-left-color: #f1c40f; }
        .hangar-card.due { border-left-color: #e74c3c; }
        .hangar-card.ext { border-left-color: #38bdf8; }
        .hangar-card.sel { box-shadow: 0 0 0 1px rgba(134, 239, 172, 0.55); }
        .hangar-card .kicker {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .hangar-card.warn .kicker { color: #f1c40f; }
        .hangar-card.due .kicker { color: #fca5a5; }
        .hangar-card.ext .kicker { color: #7dd3fc; }
        .hangar-card h4 { margin: 0.2rem 0 0.45rem; color: #ffffff; font-size: 1.05rem; }
        .hangar-card p { margin: 0.18rem 0 0; color: #94a3b8; font-size: 0.84rem; }
        .hangar-card b { color: #e2e8f0; }
        .hangar-hint {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 10px;
            padding: 0.75rem 0.85rem;
            color: #cbd5e1;
            font-size: 0.82rem;
        }
        .admin-compact {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 0.85rem 1rem 0.7rem;
            margin-bottom: 0.85rem;
        }
        .admin-compact .kicker {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .admin-compact h4 {
            margin: 0.15rem 0 0.25rem;
            color: #ffffff;
            font-size: 1.02rem;
        }
        .admin-compact p {
            margin: 0 0 0.65rem;
            color: #94a3b8;
            font-size: 0.82rem;
        }
        .hangar-cal .cal-day { min-height: 38px; padding: 4px 5px; }
        .hangar-cal .cal-day .num { font-size: 0.82rem; }
        .nucleo-dia {
            background: #0f172a;
            border: 1px solid #334155;
            border-left: 4px solid #86efac;
            border-radius: 12px;
            padding: 0.85rem 1rem;
            color: #e2e8f0;
            margin-bottom: 0.75rem;
        }
        .nucleo-dia .kicker {
            font-size: 0.7rem;
            letter-spacing: 0.07em;
            text-transform: uppercase;
            color: #86efac;
            font-weight: 700;
        }
        .nucleo-dia h3 { margin: 0.15rem 0 0.35rem; color: #ffffff; font-size: 1.12rem; }
        .nucleo-dia p { margin: 0; color: #94a3b8; font-size: 0.84rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── Catálogos y permisos ───────────────────────────────────────────────────────
def catalogos_default() -> dict:
    return {
        "sectores": list(SECTORES_DEFAULT),
        "sectores_sin_pozo": list(SECTORES_SIN_POZO_DEFAULT),
        "equipos_mci": list(EQUIPOS_MCI_DEFAULT),
        "equipos_ub": list(EQUIPOS_UB_DEFAULT),
        "equipos_otros": list(EQUIPOS_OTROS_DEFAULT),
    }


def leer_json(ruta: str, fallback):
    if not os.path.exists(ruta):
        return fallback
    try:
        with open(ruta, encoding="utf-8") as archivo:
            return json.load(archivo)
    except Exception:
        return fallback


def escribir_json(ruta: str, datos) -> None:
    with open(ruta, "w", encoding="utf-8") as archivo:
        json.dump(datos, archivo, ensure_ascii=False, indent=2)


def normalizar_catalogos(datos: dict | None) -> dict:
    base = catalogos_default()
    if not isinstance(datos, dict):
        return base
    for clave, valor_default in base.items():
        valores = datos.get(clave, valor_default)
        if not isinstance(valores, list):
            valores = valor_default
        limpios = []
        vistos = set()
        for item in valores:
            texto = str(item).strip()
            if texto and texto not in vistos:
                limpios.append(texto)
                vistos.add(texto)
        base[clave] = limpios or valor_default
    if SECTOR_OTRO not in base["sectores"]:
        base["sectores"].append(SECTOR_OTRO)
    base["equipos_otros"] = [
        equipo
        for equipo in base["equipos_otros"]
        if not sitios_operativos.es_vacio_sitio(equipo)
    ]
    for equipo_ub in EQUIPOS_UB_DEFAULT:
        if equipo_ub not in base["equipos_ub"]:
            base["equipos_ub"].append(equipo_ub)
    return base


def inicializar_estado() -> None:
    if "form_id" not in st.session_state:
        st.session_state.form_id = 0
    if "catalogos" not in st.session_state:
        st.session_state.catalogos = normalizar_catalogos(leer_json(CATALOGOS_FILE, catalogos_default()))
    if "permisos" not in st.session_state:
        permisos = leer_json(PERMISOS_FILE, PERMISOS_DEFAULT)
        if not isinstance(permisos, list) or not permisos:
            permisos = PERMISOS_DEFAULT
        st.session_state.permisos = permisos


def obtener_catalogos() -> dict:
    inicializar_estado()
    return st.session_state.catalogos


def guardar_catalogos(catalogos: dict) -> None:
    limpios = normalizar_catalogos(catalogos)
    st.session_state.catalogos = limpios
    escribir_json(CATALOGOS_FILE, limpios)


def sectores_actuales() -> list[str]:
    return list(obtener_catalogos()["sectores"])


def equipos_actuales() -> list[str]:
    cat = obtener_catalogos()
    equipos = list(cat["equipos_mci"]) + list(cat["equipos_ub"]) + list(cat["equipos_otros"])
    vistos = set()
    unicos = []
    for equipo in equipos:
        if equipo not in vistos:
            unicos.append(equipo)
            vistos.add(equipo)
    return [equipo for equipo in unicos if not sitios_operativos.es_vacio_sitio(equipo)]


_UBICACIONES_IGNORADAS = {
    "",
    "nan",
    "none",
    "—",
    "-",
    "seleccionar...",
    "ninguno",
    SECTOR_OTRO.casefold(),
    OPCION_UBICACION_NUEVA.casefold(),
}


def _agregar_opcion_unica(destino: list[str], vistos: set[str], valor) -> None:
    texto = str(valor or "").strip()
    clave = texto.casefold()
    if clave in _UBICACIONES_IGNORADAS or clave in vistos:
        return
    vistos.add(clave)
    destino.append(texto)


_PATRON_CODIGO_POZO = re.compile(r"^(PSN|PGA|PPB|PS|PA|PB|PC)-?\d", re.IGNORECASE)


def etiqueta_codigo_ubicacion(nombre: str) -> str:
    """Muestra el código de pozo en mayúsculas. Conserva nombres como Base PJP4."""
    texto = str(nombre or "").strip()
    if not texto:
        return ""
    compacto = clave_ubicacion(texto).replace(" ", "")
    if _PATRON_CODIGO_POZO.match(compacto):
        return compacto
    return texto


def etiqueta_macro_zona_oficial(zona_id: str, zona: dict | None = None) -> str:
    """Misma etiqueta que las tarjetas: PSN, INS Instalaciones y baterías, NTE…"""
    for item in ZONAS_HISTORIAL_POZOS:
        if item["id"] == zona_id:
            return item["titulo"]
    zona = zona or next((item for item in ZONAS_UBICACION_DEFAULT if item["id"] == zona_id), {})
    prefijo = str((zona or {}).get("prefijo") or "").strip()
    nombre = str((zona or {}).get("nombre") or "").strip()
    return f"{prefijo} {nombre}".strip()


def opciones_macro_captura() -> list[str]:
    """Paso 1: zonas de pozos + INS/OTR/NTE/SUR + atajos AOPS y Base PJP4."""
    macros = [item["titulo"] for item in ZONAS_HISTORIAL_POZOS]
    for zona_id in _IDS_BLOQUE_CAPTURA:
        etiqueta = etiqueta_macro_zona_oficial(zona_id)
        if etiqueta:
            macros.append(etiqueta)
    macros.extend([MACRO_AOPS, MACRO_BASE])
    return macros


def opciones_secciones_operativas() -> list[str]:
    return opciones_macro_captura()


def tipo_macro_captura(macro) -> str:
    texto = str(macro or "").strip()
    if texto == MACRO_AOPS:
        return "aops"
    if texto in {MACRO_BASE, CLAVE_BASE_PJP4, "Base Soledad"}:
        return "base"
    zona_id = zona_id_macro_captura(texto)
    if zona_id in _IDS_ZONA_POZO_CAPTURA:
        return "pozo"
    if zona_id in _IDS_BLOQUE_CAPTURA:
        return "bloque"
    return ""


def es_codigo_infraestructura(texto) -> bool:
    plano = clave_pozo_plana(texto)
    return "IQSIN" in plano or plano[:3] in _CLASES_INFRAESTRUCTURA and "SIN" in plano


def tag_pozo_de_nombre(macro, nombre) -> str:
    """PSN-190 / PSN190 / PC168D -> tag limpio de la zona, sin Ninguno."""
    prefijo = prefijo_zona_captura(macro)
    plano = clave_pozo_plana(nombre)
    if prefijo and plano.startswith(prefijo):
        resto = plano[len(prefijo) :]
        if resto[:1].isdigit():
            return f"{prefijo}{resto}"
    return componer_tag_pozo(macro, nombre)


def tag_limpio_ficha_captura(ficha: dict, macro: str = "") -> str:
    nombre = str((ficha or {}).get("nombre") or "").strip()
    if not nombre or sitios_operativos.es_vacio_sitio(nombre):
        return ""
    zona_id = str((ficha or {}).get("zona_id") or "")
    etiqueta = macro or etiqueta_macro_zona_oficial(zona_id)
    if zona_id in _IDS_ZONA_POZO_CAPTURA:
        return tag_pozo_de_nombre(etiqueta, nombre)
    if zona_id == "instalaciones":
        clave = sitios_operativos.clave_sitio(nombre)
        if clave:
            return clave
        if clave_pozo_plana(nombre) in {"BASEPJP4", "BASESOLEDAD"}:
            return CLAVE_BASE_PJP4
        return nombre
    return nombre


def opciones_items_macro_captura(macro) -> list[str]:
    """Paso 2: solo fichas del bloque elegido (pozos, INS, NTE, SUR u OTR)."""
    tipo = tipo_macro_captura(macro)
    zona_id = zona_id_macro_captura(macro)
    if tipo == "pozo":
        return opciones_pozos_zona_captura(macro)
    if tipo != "bloque" or not zona_id:
        return []
    vistos: set[str] = set()
    opciones: list[str] = []
    for ficha in fichas_de_zona(cargar_ubicaciones(), zona_id):
        tag = tag_limpio_ficha_captura(ficha, macro)
        if not tag or tag in vistos or sitios_operativos.es_vacio_sitio(tag):
            continue
        vistos.add(tag)
        opciones.append(tag)
    return sorted(opciones, key=orden_natural_pozo)


def resolver_tag_macro(macro, detalle="") -> str:
    """Convierte Paso 1 + Paso 2 en la clave limpia a guardar. Nunca Ninguno."""
    tipo = tipo_macro_captura(macro)
    if tipo == "aops":
        return MACRO_AOPS
    if tipo == "base":
        return CLAVE_BASE_PJP4
    texto = str(detalle or "").strip()
    if not texto or sitios_operativos.es_vacio_sitio(texto):
        return ""
    if tipo == "pozo":
        return tag_pozo_de_nombre(macro, texto) or componer_tag_pozo(macro, texto)
    if tipo == "bloque":
        opciones = opciones_items_macro_captura(macro)
        if texto in opciones:
            return texto
        clave = "" if es_codigo_infraestructura(texto) else sitios_operativos.clave_sitio(texto)
        if clave and clave in opciones:
            return clave
        return ""
    return ""


def descomponer_macro_captura(texto) -> tuple[str, str]:
    """Recupera Paso 1 y Paso 2 a partir de un valor ya guardado."""
    zona, numero = descomponer_tag_pozo(texto)
    if zona:
        tag = tag_pozo_de_nombre(zona, texto) or componer_tag_pozo(zona, numero)
        return zona, tag or etiqueta_codigo_ubicacion(texto)
    clave = "" if es_codigo_infraestructura(texto) else sitios_operativos.clave_sitio(texto)
    if clave == MACRO_AOPS:
        return MACRO_AOPS, ""
    if clave == CLAVE_BASE_PJP4:
        return MACRO_BASE, ""
    ficha = encontrar_ficha(cargar_ubicaciones(), texto)
    if ficha:
        zona_id = str(ficha.get("zona_id") or "")
        macro = etiqueta_macro_zona_oficial(zona_id)
        tag = tag_limpio_ficha_captura(ficha, macro)
        if macro and tag:
            return macro, tag
    if clave:
        return etiqueta_macro_zona_oficial("instalaciones"), clave
    return "", ""


def prefijo_zona_captura(zona) -> str:
    texto = " ".join(str(zona or "").split())
    if not texto:
        return ""
    if texto in PREFIJOS_ZONA_CAPTURA:
        return PREFIJOS_ZONA_CAPTURA[texto]
    plano = clave_pozo_plana(texto)
    for etiqueta, prefijo in sorted(PREFIJOS_ZONA_CAPTURA.items(), key=lambda item: -len(item[1])):
        if plano == clave_pozo_plana(etiqueta) or plano == prefijo:
            return prefijo
    seccion = etiqueta_seccion_operativa(texto)
    if seccion:
        for item in SECCIONES_OPERATIVAS:
            if item["etiqueta"] == seccion:
                return item["prefijo"]
    return ""


def numero_pozo_captura(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def componer_tag_pozo(zona, numero) -> str:
    """PSN + 190 -> PSN190. Nunca devuelve Ninguno."""
    prefijo = prefijo_zona_captura(zona)
    digitos = numero_pozo_captura(numero)
    if not prefijo or not digitos:
        plano = clave_pozo_plana(numero) or clave_pozo_plana(zona)
        if _PATRON_CODIGO_POZO.match(plano):
            return etiqueta_codigo_ubicacion(plano)
        return ""
    return f"{prefijo}{digitos}"


def zona_id_macro_captura(macro) -> str:
    texto = str(macro or "").strip()
    if not texto or texto in {MACRO_AOPS, MACRO_BASE, CLAVE_BASE_PJP4}:
        return ""
    for item in ZONAS_HISTORIAL_POZOS:
        if texto == item["titulo"]:
            return item["id"]
    for zona in ZONAS_UBICACION_DEFAULT:
        etiqueta = etiqueta_macro_zona_oficial(zona["id"], zona)
        if texto in {etiqueta, zona["id"], zona["prefijo"], f"{zona['prefijo']} {zona['nombre']}".strip()}:
            return zona["id"]
    prefijo = prefijo_zona_captura(texto)
    if not prefijo:
        return ""
    for seccion in SECCIONES_OPERATIVAS:
        if seccion["prefijo"] == prefijo:
            return seccion["id"]
    return ""


def opciones_pozos_zona_captura(macro) -> list[str]:
    """Pozos oficiales de la zona, en tag limpio (PSN190). Sin texto libre."""
    zona_id = zona_id_macro_captura(macro)
    prefijo = prefijo_zona_captura(macro)
    if not zona_id or not prefijo or zona_id not in _IDS_ZONA_POZO_CAPTURA:
        return []
    vistos: set[str] = set()
    opciones: list[str] = []
    for ficha in fichas_de_zona(cargar_ubicaciones(), zona_id):
        nombre = str(ficha.get("nombre") or "").strip()
        if not nombre:
            continue
        if sitios_operativos.clave_sitio(nombre) and not _PATRON_CODIGO_POZO.match(clave_pozo_plana(nombre)):
            continue
        tag = tag_pozo_de_nombre(macro, nombre)
        if not tag or tag in vistos:
            continue
        resto = tag[len(prefijo) :] if tag.startswith(prefijo) else ""
        if not resto[:1].isdigit():
            continue
        vistos.add(tag)
        opciones.append(tag)
    return sorted(opciones, key=orden_natural_pozo)


def descomponer_tag_pozo(texto) -> tuple[str, str]:
    """PSN190 -> ('PSN', '190')."""
    plano = clave_pozo_plana(texto)
    if not plano:
        return "", ""
    for etiqueta, prefijo in sorted(PREFIJOS_ZONA_CAPTURA.items(), key=lambda item: -len(item[1])):
        if plano.startswith(prefijo) and plano[len(prefijo) : len(prefijo) + 1].isdigit():
            return etiqueta, plano[len(prefijo) :]
        if plano == prefijo:
            return etiqueta, ""
    return "", ""


def etiqueta_seccion_operativa(texto) -> str:
    """Devuelve la etiqueta limpia de una de las 6 secciones, o vacío."""
    bruto = " ".join(str(texto or "").split())
    if not bruto or sitios_operativos.es_vacio_sitio(bruto):
        return ""
    plano = clave_pozo_plana(bruto)
    for seccion in SECCIONES_OPERATIVAS:
        etiqueta = seccion["etiqueta"]
        if bruto == etiqueta or plano == clave_pozo_plana(etiqueta):
            return etiqueta
        for alias in seccion["aliases"]:
            if plano == clave_pozo_plana(alias):
                return etiqueta
    # Pozos individuales (PSN261, PA-7, etc.) se agrupan en su sección.
    for seccion in sorted(SECCIONES_OPERATIVAS, key=lambda item: -len(item["prefijo"])):
        pref = seccion["prefijo"]
        if plano == pref:
            return seccion["etiqueta"]
        if plano.startswith(pref) and len(plano) > len(pref) and plano[len(pref)].isdigit():
            return seccion["etiqueta"]
    zona_id = inferir_zona_id(bruto)
    for seccion in SECCIONES_OPERATIVAS:
        if seccion["id"] == zona_id:
            return seccion["etiqueta"]
    return ""


def _opcion_catalogo_guardar(valor) -> str:
    """Instalaciones del catálogo maestro se listan y se guardan por clave corta."""
    texto = etiqueta_codigo_ubicacion(str(valor or "").strip())
    if not texto:
        return ""
    clave = sitios_operativos.clave_sitio(texto)
    return clave or texto


def etiqueta_selector_ubicacion(valor: str) -> str:
    texto = str(valor or "").strip()
    if not texto or texto == OPCION_UBICACION_NUEVA:
        return texto
    clave = sitios_operativos.clave_sitio(texto)
    if not clave:
        return texto
    nombre = sitios_operativos.nombre_oficial_sitio(clave)
    if nombre and nombre.casefold() not in {clave.casefold(), texto.casefold()}:
        return f"{clave} — {nombre}"
    return clave


def opciones_ubicacion_catalogo() -> list[str]:
    return opciones_macro_captura()


def selector_tag_instalacion(key: str, ubicacion: str, actual: str = "") -> str:
    """Opcional. En instalaciones no se pide pozo ni MCI; la clave corta es el valor guardado."""
    clave = sitios_operativos.clave_sitio(ubicacion)
    if not clave:
        return ""
    st.caption(
        f"**{clave}** — {sitios_operativos.nombre_oficial_sitio(clave)}. "
        "No requiere pozo ni motor. El registro guarda esta clave."
    )
    return clave


def selector_ubicacion_oficial(key: str, etiqueta: str = "Ubicación / Equipo — zona", tag_actual: str = ""):
    """Paso 1 macro; Paso 2 dinámico (pozo, batería, estación o clave fija)."""
    return render_selector_ubicacion_captura(key, etiqueta=etiqueta, valor_actual=tag_actual)


def _catalogo_equipos_estandar() -> list[str]:
    """Nombres oficiales de MCI, UBM y otros, en el orden del catálogo."""
    cat = obtener_catalogos()
    orden: list[str] = []
    vistos: set[str] = set()
    for grupo in ("equipos_mci", "equipos_ub", "equipos_otros"):
        for item in list(cat.get(grupo) or []):
            texto = str(item).strip()
            clave = texto.casefold()
            if not texto or sitios_operativos.es_vacio_sitio(texto) or clave in vistos:
                continue
            vistos.add(clave)
            orden.append(texto)
    return orden


def nombre_equipo_catalogo(valor) -> str:
    """Devuelve el nombre estandarizado del catálogo, o vacío si no pertenece a él."""
    bruto = str(valor or "").strip()
    if not bruto or sitios_operativos.es_vacio_sitio(bruto):
        return ""
    por_clave = {item.casefold(): item for item in _catalogo_equipos_estandar()}
    if bruto.casefold() in por_clave:
        return por_clave[bruto.casefold()]
    for candidato in (normalizar_mci_excel(bruto), normalizar_ub_excel(bruto)):
        clave = str(candidato or "").strip().casefold()
        if clave in por_clave and clave != bruto.casefold():
            return por_clave[clave]
    return ""


def _indice_equipos_maestro_excel() -> dict[str, tuple[str, str]]:
    """Clave plana del pozo → (MCI, UBM) leídos de las columnas del Excel maestro."""
    ruta = excel_inventario_predeterminado()
    marca = (os.path.abspath(ruta), _mtime_archivo(ruta)) if ruta and os.path.isfile(ruta) else ("", 0.0)
    cache = getattr(_indice_equipos_maestro_excel, "_cache", None)
    if cache and cache[0] == marca:
        return cache[1]
    indice: dict[str, tuple[str, str]] = {}
    if marca[0]:
        try:
            filas = extraer_catalogo_maestro_excel(ruta)
        except Exception:
            filas = []
        for item in filas:
            if item.get("instalacion"):
                continue
            mci = nombre_equipo_catalogo(item.get("equipo_mci")) or nombre_equipo_catalogo(item.get("mci_raw"))
            ub = nombre_equipo_catalogo(item.get("equipo_ub")) or nombre_equipo_catalogo(item.get("ub_raw"))
            if not mci and not ub:
                continue
            claves: set[str] = set()
            for texto in (
                item.get("nombre"),
                item.get("corto_app"),
                item.get("largo_excel"),
                *(item.get("aliases") or []),
            ):
                if not str(texto or "").strip():
                    continue
                claves.add(clave_pozo_plana(str(texto)))
                claves |= _variantes_clave_ubicacion(str(texto))
            for clave in claves:
                if clave:
                    indice.setdefault(clave, (mci, ub))
    _indice_equipos_maestro_excel._cache = (marca, indice)
    return indice


def _equipo_maestro_de_ubicacion(ubicacion: str) -> tuple[str, str]:
    indice = _indice_equipos_maestro_excel()
    if not indice:
        return "", ""
    for clave in _variantes_clave_ubicacion(ubicacion, canonizar_nombre_ubicacion(ubicacion)):
        par = indice.get(clave)
        if par and (par[0] or par[1]):
            return par
    return indice.get(clave_pozo_plana(ubicacion), ("", ""))


def opciones_equipo_de_ubicacion(ubicacion: str) -> list[str]:
    """MCI y UBM que el catálogo maestro asocia a esta ubicación."""
    texto = str(ubicacion or "").strip()
    if not texto or sitios_operativos.es_vacio_sitio(texto):
        return []
    _recargar_sitios_operativos()
    ficha = ficha_por_ubicacion(cargar_ubicaciones(), texto)
    mci = nombre_equipo_catalogo((ficha or {}).get("equipo_mci"))
    ub = nombre_equipo_catalogo((ficha or {}).get("equipo_ub"))
    if not mci or not ub:
        maestro_mci, maestro_ub = _equipo_maestro_de_ubicacion(texto)
        mci = mci or maestro_mci
        ub = ub or maestro_ub
    opciones: list[str] = []
    vistos: set[str] = set()
    for nombre in (mci, ub):
        clave = nombre.casefold()
        if nombre and clave not in vistos:
            vistos.add(clave)
            opciones.append(nombre)
    return opciones


def mci_preseleccionado(opciones: list[str]) -> str:
    """MCI del catálogo maestro presente en las opciones de esta ubicación."""
    oficiales = {
        str(item).strip().casefold()
        for item in (obtener_catalogos().get("equipos_mci") or [])
        if str(item).strip()
    }
    for opcion in opciones:
        if str(opcion).casefold() in oficiales:
            return opcion
    return ""


def render_selector_equipo_ubicacion(key: str, ubicacion: str) -> str:
    """Equipo del catálogo de la ubicación. Al abrir o reiniciar queda el MCI."""
    tag = str(ubicacion or "").strip()
    opciones = opciones_equipo_de_ubicacion(tag)
    if not tag:
        st.caption("Elige la ubicación del catálogo para listar sus equipos.")
        return ""
    if not opciones:
        st.caption("Esta ubicación no tiene MCI ni UBM en el catálogo maestro.")
        return ""
    clave = f"{key}_{clave_pozo_plana(tag)}"
    if st.session_state.get(clave) not in opciones:
        st.session_state.pop(clave, None)
    if clave not in st.session_state:
        mci = mci_preseleccionado(opciones)
        if mci:
            st.session_state[clave] = mci
    return st.selectbox(
        "Equipo",
        opciones,
        placeholder="Seleccionar equipo del catálogo",
        key=clave,
        persist_state="session",
        help="Empieza con el MCI de esta ubicación. Puedes cambiarlo por otro equipo del catálogo.",
        **({} if clave in st.session_state else {"index": None}),
    ) or ""


def render_selector_ubicacion_captura(
    key: str,
    etiqueta: str = "Ubicación / Equipo — zona",
    valor_actual: str = "",
) -> tuple[str, str, str]:
    """Paso 1 = zonas del catálogo. Paso 2 = fichas de ese bloque."""
    _recargar_sitios_operativos()
    opciones_macro = opciones_macro_captura()
    macro_key = f"{key}_macro_v5"
    if valor_actual and macro_key not in st.session_state:
        macro_prev, detalle_prev = descomponer_macro_captura(valor_actual)
        if macro_prev:
            st.session_state[macro_key] = macro_prev
            tipo_prev = tipo_macro_captura(macro_prev)
            zona_prev = zona_id_macro_captura(macro_prev) or tipo_prev
            if tipo_prev in {"pozo", "bloque"} and detalle_prev:
                st.session_state[f"{key}_item_v5_{zona_prev}"] = detalle_prev
    if st.session_state.get(macro_key) not in opciones_macro:
        st.session_state.pop(macro_key, None)

    macro = st.selectbox(
        etiqueta,
        opciones_macro,
        index=None,
        placeholder="Paso 1: elige zona, INS, NTE, SUR u otra ubicación",
        key=macro_key,
        persist_state="session",
    )
    tipo = tipo_macro_captura(macro)
    detalle = ""
    if tipo in {"pozo", "bloque"}:
        items = list(opciones_items_macro_captura(macro))
        item_key = f"{key}_item_v5_{zona_id_macro_captura(macro) or tipo}"
        if valor_actual:
            macro_prev, detalle_prev = descomponer_macro_captura(valor_actual)
            if macro_prev == macro and detalle_prev and detalle_prev not in items:
                items = [detalle_prev, *items]
        if st.session_state.get(item_key) not in items:
            st.session_state.pop(item_key, None)
        etiqueta_paso2 = "Pozo" if tipo == "pozo" else "Ubicación / equipo del bloque"
        if not items:
            st.caption("No hay elementos en el catálogo de este bloque.")
        else:
            detalle = st.selectbox(
                etiqueta_paso2,
                items,
                index=None,
                placeholder="Paso 2: elige el elemento de este bloque",
                key=item_key,
                persist_state="session",
            ) or ""
    elif tipo in {"aops", "base"}:
        st.caption("No se pide un segundo paso: se guarda la clave de esta instalación.")

    tag = resolver_tag_macro(macro, detalle)
    if tag:
        st.caption(f"Se guardará: **{tag}**")
    elif tipo == "pozo":
        st.caption("Elige el pozo del catálogo de esta zona.")
    elif tipo == "bloque":
        st.caption("Elige la instalación, línea o ubicación de este bloque.")
    return macro or "", detalle or "", tag


def opciones_ubicacion_estandar() -> list[str]:
    return opciones_macro_captura()


def resolver_ubicacion_catalogo(seleccion: str) -> tuple[str, str]:
    """Parte una opción del catálogo en sector y pozo para guardarla en la bitácora."""
    texto = str(seleccion or "").strip()
    if not texto or texto == OPCION_UBICACION_NUEVA:
        return "", ""
    ficha = encontrar_ficha(cargar_ubicaciones(), texto)
    if ficha is None:
        compacto = clave_ubicacion(texto).replace(" ", "")
        if _PATRON_CODIGO_POZO.match(compacto):
            return "", etiqueta_codigo_ubicacion(texto)
        sitio = sitios_operativos.clave_sitio(texto)
        if sitio:
            return sitio, sitio
        return texto, ""
    nombre = etiqueta_codigo_ubicacion(ficha.get("nombre")) or texto
    sector = str(ficha.get("sector") or "").strip()
    if not sector:
        zona = mapa_zonas().get(str(ficha.get("zona_id") or ""), {})
        sector = str(zona.get("sector") or "").strip()
    compacto = clave_ubicacion(nombre).replace(" ", "")
    if _PATRON_CODIGO_POZO.match(compacto):
        return sector, nombre
    sitio = sitios_operativos.clave_sitio(nombre) or sitios_operativos.clave_sitio(texto)
    if sitio:
        return sitio, sitio
    instalacion = INSTALACIONES_SECTOR.get(nombre) or INSTALACIONES_SECTOR.get(str(ficha.get("nombre") or "").strip())
    if instalacion:
        clave = sitios_operativos.clave_sitio(nombre) or sitios_operativos.clave_sitio(instalacion) or nombre
        return clave, clave
    if sector and sector.casefold() != nombre.casefold():
        sitio = sitios_operativos.clave_sitio(sector) or sitios_operativos.clave_sitio(nombre)
        if sitio:
            return sitio, sitio
        return sector, nombre
    clave = sitios_operativos.clave_sitio(nombre)
    if clave:
        return clave, clave
    return nombre, ""


def resolver_captura_ubicacion(seleccion, nueva="", tag="") -> dict:
    """Guarda el tag limpio del bloque (PSN190, BSNI, GDO-…, AOPS). Nunca Ninguno."""
    tag_final = str(tag or "").strip()
    if not tag_final or sitios_operativos.es_vacio_sitio(tag_final):
        tag_final = resolver_tag_macro(seleccion, nueva)
    if not tag_final:
        for bruto in (tag, nueva, seleccion):
            plano = clave_pozo_plana(bruto)
            if _PATRON_CODIGO_POZO.match(plano):
                tag_final = etiqueta_codigo_ubicacion(plano)
                break
            if es_codigo_infraestructura(bruto):
                tag_final = str(bruto or "").strip()
                break
            sitio = sitios_operativos.clave_sitio(bruto)
            if sitio:
                tag_final = sitio
                break
    if not tag_final or sitios_operativos.es_vacio_sitio(tag_final):
        return {
            "es_instalacion": False,
            "es_seccion": False,
            "clave": "",
            "sector": "",
            "pozo": "",
            "ubicacion": "",
            "tag": "",
        }
    zona_id = zona_id_macro_captura(seleccion)
    if es_codigo_infraestructura(tag_final) or zona_id in {"infra-norte", "infra-sur", "otros"}:
        sector = etiqueta_macro_zona_oficial(zona_id) if zona_id else str(seleccion or tag_final)
        return {
            "es_instalacion": False,
            "es_seccion": False,
            "clave": tag_final,
            "sector": sector,
            "pozo": tag_final,
            "ubicacion": tag_final,
            "tag": tag_final,
        }
    sitio = sitios_operativos.clave_sitio(tag_final)
    if sitio:
        return {
            "es_instalacion": True,
            "es_seccion": False,
            "clave": sitio,
            "sector": sitio,
            "pozo": sitio,
            "ubicacion": sitio,
            "tag": sitio,
        }
    zona, _numero = descomponer_tag_pozo(tag_final)
    return {
        "es_instalacion": False,
        "es_seccion": True,
        "clave": tag_final,
        "sector": zona or tag_final,
        "pozo": tag_final,
        "ubicacion": tag_final,
        "tag": tag_final,
    }


_CAMPOS_UBICACION_OT = (
    "POZO",
    "UBICACIÓN",
    "UBICACION",
    "EQUIPO/UBICACIÓN",
    "EQUIPO/UBICACION",
    "SECTOR",
    "DESCRIPCIÓN DE LA ACTIVIDAD",
    "DESCRIPCION DE LA ACTIVIDAD",
    "ACTIVIDAD CORTA",
    "TIPO DE ACTIVIDAD",
    "COMENTARIOS",
)


def _texto_crudo_campo(valor) -> str:
    return texto_visible(valor, "").replace('"', "").replace("'", "").strip()


def _iter_campos_fila(fila):
    if fila is None:
        return
    if isinstance(fila, dict):
        yield from fila.items()
        return
    try:
        yield from fila.items()
    except Exception:
        for col in getattr(fila, "index", []):
            yield col, fila.get(col)


def textos_ubicacion_ot(fila) -> dict[str, str]:
    """Texto exacto de cada campo de ubicación del registro de la OT."""
    buscados = {_clave_columna_bitacora(nombre) for nombre in _CAMPOS_UBICACION_OT}
    extraidos: dict[str, str] = {}
    for col, valor in _iter_campos_fila(fila):
        clave = _clave_columna_bitacora(str(col))
        if clave not in buscados:
            continue
        texto = _texto_crudo_campo(valor)
        if texto and clave not in extraidos:
            extraidos[clave] = texto
    return extraidos


def aplicar_catalogo_maestro_excel() -> int:
    """Carga el mapa Nombre Largo / Nombre Corto del Excel de ubicaciones y equipos."""
    ruta = excel_inventario_predeterminado()
    if not ruta or not os.path.isfile(ruta):
        return 0
    try:
        mtime = os.path.getmtime(ruta)
    except OSError:
        return 0
    cache = getattr(aplicar_catalogo_maestro_excel, "_cache", None)
    if cache and cache[0] == mtime:
        return cache[1]
    try:
        extraer_catalogo_maestro_excel(ruta)
        pares = list(getattr(extraer_catalogo_maestro_excel, "_pares", None) or [])
    except Exception:
        pares = []
    aplicar_catalogo_maestro_excel._cache = (mtime, len(pares))
    return len(pares)


def _recargar_sitios_operativos() -> None:
    ruta = getattr(sitios_operativos, "__file__", "")
    mtime = os.path.getmtime(ruta) if ruta and os.path.exists(ruta) else 0
    if getattr(_recargar_sitios_operativos, "_mtime", None) != mtime:
        importlib.reload(sitios_operativos)
        _recargar_sitios_operativos._mtime = mtime
        aplicar_catalogo_maestro_excel._cache = None
    aplicar_catalogo_maestro_excel()


def etiqueta_ubicacion_fila(fila) -> str:
    """Etiqueta visible: batería/estación + TAG, nunca '· Ninguno'."""
    _recargar_sitios_operativos()
    campos = textos_ubicacion_ot(fila)
    loc = [
        campos.get("pozo") or "",
        campos.get("ubicacion") or "",
        campos.get("equipo ubicacion") or "",
        campos.get("sector") or "",
    ]
    return sitios_operativos.texto_ubicacion_final(*loc)


def formato_ubicacion_reporte(fila) -> str:
    """Nombre limpio de la instalación o clave del pozo. Nunca 'NINGUNO'."""
    _recargar_sitios_operativos()
    campos = textos_ubicacion_ot(fila)
    etiqueta = sitios_operativos.texto_ubicacion_final(*campos.values())
    if etiqueta:
        return etiqueta
    pozo = sitios_operativos.quitar_ninguno(campos.get("pozo", ""))
    ubicacion = sitios_operativos.quitar_ninguno(
        campos.get("ubicacion") or campos.get("equipo ubicacion") or ""
    )
    sector = sitios_operativos.quitar_ninguno(campos.get("sector", ""))
    partes = [parte.strip() for parte in re.split(r"[·|;/]+", ubicacion) if parte.strip()]
    if not sector and partes:
        sector = partes[0]
    if sitios_operativos.es_vacio_sitio(pozo) and len(partes) >= 2:
        pozo = next((parte for parte in partes[1:] if not sitios_operativos.es_vacio_sitio(parte)), "")
    if sector and _PATRON_CODIGO_POZO.match(clave_ubicacion(sector).replace(" ", "")) and not pozo:
        pozo, sector = sector, ""
    if pozo and not sitios_operativos.es_vacio_sitio(pozo):
        ficha = encontrar_ficha(cargar_ubicaciones(), pozo)
        if ficha:
            pozo = etiqueta_codigo_ubicacion(ficha.get("nombre")) or pozo
            sitio = sitios_operativos.texto_ubicacion_final(pozo, ficha.get("nombre"), ficha.get("sector"))
            if sitio:
                return sitio
        return sitios_operativos.quitar_ninguno(pozo).upper()
    if sector and not sitios_operativos.es_vacio_sitio(sector):
        return sitios_operativos.quitar_ninguno(sector)
    return ""


def opciones_ubicacion_equipo(df: pd.DataFrame | None = None) -> list[str]:
    """Ubicaciones y equipos únicos del catálogo, las fichas y la bitácora."""
    opciones: list[str] = []
    vistos: set[str] = set()
    for equipo in equipos_actuales():
        _agregar_opcion_unica(opciones, vistos, equipo)
    for ficha in (cargar_ubicaciones().get("fichas") or {}).values():
        if not isinstance(ficha, dict):
            continue
        _agregar_opcion_unica(opciones, vistos, ficha.get("nombre"))
        _agregar_opcion_unica(opciones, vistos, ficha.get("equipo_mci"))
        _agregar_opcion_unica(opciones, vistos, ficha.get("equipo_ub"))
    if df is None:
        df = cargar_bitacora()
    if df is not None and not df.empty:
        for col in ("UBICACIÓN", "EQUIPO/UBICACIÓN", "POZO"):
            if col not in df.columns:
                continue
            for valor in df[col].tolist():
                _agregar_opcion_unica(opciones, vistos, valor)
    opciones.sort(key=orden_natural_pozo)
    opciones.append(OPCION_UBICACION_NUEVA)
    return opciones


def opciones_equipos_ub() -> list[str]:
    cat = obtener_catalogos()
    vistos: list[str] = []
    for item in list(cat.get("equipos_ub") or []) + list(EQUIPOS_UB_DEFAULT):
        texto = str(item).strip()
        if texto and texto not in vistos:
            vistos.append(texto)
    cat["equipos_ub"] = vistos
    return vistos


def guardar_permisos(permisos: list[dict]) -> None:
    st.session_state.permisos = permisos
    escribir_json(PERMISOS_FILE, permisos)


def slug_insumo(nombre: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", nombre.lower().strip()).strip("-") or uuid.uuid4().hex[:8]


def unidad_es_entera(unidad: str) -> bool:
    return str(unidad).strip().lower() in {
        "pzas",
        "pza",
        "piezas",
        "pz",
        "unidades",
        "und",
        "u",
    }


def tipo_insumo_por_nombre(nombre: str) -> str:
    clave = _sin_acentos(str(nombre or "")).casefold()
    if any(token in clave for token in ("bomba de agua", "alternador", "motor de arranque", "marcha")):
        return TIPO_INSUMO_ROTATIVO
    return TIPO_INSUMO_CONSUMIBLE


def tipo_insumo_item(item: dict) -> str:
    explicito = str(item.get("tipo") or "").strip()
    if explicito in TIPOS_INSUMO:
        return explicito
    return tipo_insumo_por_nombre(str(item.get("nombre") or ""))


def es_insumo_rotativo(item: dict) -> bool:
    return tipo_insumo_item(item) == TIPO_INSUMO_ROTATIVO


def almacen_default() -> dict:
    catalogo = []
    for item in INSUMOS_CATALOGO_DEFAULT:
        catalogo.append(
            {
                "id": item["id"],
                "nombre": item["nombre"],
                "unidad": item["unidad"],
                "columna_excel": item.get("columna_excel", ""),
                "tipo": tipo_insumo_item(item),
                "precargado": True,
                "activo": True,
            }
        )
    return {"catalogo": catalogo, "movimientos": []}


def normalizar_almacen(datos) -> dict:
    base = almacen_default()
    if not isinstance(datos, dict):
        return base
    existentes = []
    vistos = set()
    for item in datos.get("catalogo", []):
        if not isinstance(item, dict):
            continue
        nombre = str(item.get("nombre", "")).strip()
        if not nombre:
            continue
        clave = nombre.casefold()
        if clave in vistos:
            continue
        vistos.add(clave)
        existentes.append(
            {
                "id": str(item.get("id") or slug_insumo(nombre)),
                "nombre": nombre,
                "unidad": str(item.get("unidad") or "Pzas").strip(),
                "columna_excel": str(item.get("columna_excel", "") or ""),
                "tipo": tipo_insumo_item(item),
                "precargado": bool(item.get("precargado", False)),
                "activo": bool(item.get("activo", True)),
            }
        )
    for item in base["catalogo"]:
        if item["nombre"].casefold() not in vistos:
            existentes.insert(0, item)
    movimientos = [
        mov for mov in datos.get("movimientos", []) if isinstance(mov, dict)
    ]
    return {"catalogo": existentes, "movimientos": movimientos}


def cargar_almacen() -> dict:
    if not os.path.exists(ALMACEN_FILE):
        datos = almacen_default()
        escribir_json(ALMACEN_FILE, datos)
        return datos
    return normalizar_almacen(leer_json(ALMACEN_FILE, almacen_default()))


def guardar_almacen(datos: dict) -> None:
    escribir_json(ALMACEN_FILE, normalizar_almacen(datos))


def catalogo_insumos_activos() -> list[dict]:
    return [item for item in cargar_almacen()["catalogo"] if item.get("activo", True)]


def alta_insumo_catalogo(nombre: str, unidad: str, tipo: str = TIPO_INSUMO_CONSUMIBLE) -> str | None:
    nombre = nombre.strip()
    unidad = unidad.strip() or "Pzas"
    if not nombre:
        return "Escribe el nombre del insumo."
    if tipo not in TIPOS_INSUMO:
        return "Selecciona si el insumo es consumible o componente rotativo / reparable."
    datos = cargar_almacen()
    if any(item["nombre"].casefold() == nombre.casefold() for item in datos["catalogo"]):
        return f"«{nombre}» ya existe en el catálogo."
    datos["catalogo"].append(
        {
            "id": slug_insumo(nombre),
            "nombre": nombre,
            "unidad": unidad,
            "columna_excel": "",
            "tipo": tipo,
            "precargado": False,
            "activo": True,
        }
    )
    guardar_almacen(datos)
    return None


def registrar_movimiento_almacen(movimiento: dict) -> None:
    datos = cargar_almacen()
    movimiento = dict(movimiento)
    movimiento.setdefault("id", uuid.uuid4().hex[:10])
    movimiento.setdefault("registrado", datetime.now().strftime("%Y-%m-%d %H:%M"))
    datos["movimientos"].append(movimiento)
    guardar_almacen(datos)


def registrar_consumos_campo(
    fecha: date,
    secuencia: str,
    cuadrilla: str,
    sector: str,
    consumos: list[dict],
) -> None:
    for consumo in consumos:
        cantidad = float(consumo.get("cantidad") or 0)
        if cantidad <= 0:
            continue
        movimiento = {
            "tipo": "consumo",
            "fecha": fecha.strftime("%Y-%m-%d"),
            "secuencia": secuencia,
            "cuadrilla": cuadrilla,
            "sector": sector,
            "insumo_id": consumo.get("insumo_id", ""),
            "insumo": consumo.get("nombre", ""),
            "cantidad": cantidad,
            "unidad": consumo.get("unidad", ""),
            "origen": consumo.get("origen", ""),
            "origen_registro": "campo",
            "tipo_insumo": str(consumo.get("tipo") or TIPO_INSUMO_CONSUMIBLE),
        }
        if str(consumo.get("tipo") or "") == TIPO_INSUMO_ROTATIVO:
            movimiento["condicion_instalada"] = str(consumo.get("condicion_instalada") or "")
            movimiento["destino_retirada"] = str(consumo.get("destino_retirada") or "")
        registrar_movimiento_almacen(movimiento)


def movimientos_consumo(mes_clave: str = "") -> list[dict]:
    movs = [
        m
        for m in cargar_almacen()["movimientos"]
        if str(m.get("tipo", "")) == "consumo"
    ]
    if mes_clave:
        movs = [m for m in movs if str(m.get("fecha", "")).startswith(mes_clave)]
    return movs


def resumen_consumo_por_origen(mes_clave: str = "") -> dict[str, pd.DataFrame]:
    filas = {"PJP4": [], "OPS": []}
    for mov in movimientos_consumo(mes_clave):
        origen = str(mov.get("origen", ""))
        clave = "PJP4" if origen.startswith("PJP4") else "OPS" if "OPS" in origen else ""
        if not clave:
            continue
        filas[clave].append(
            {
                "Insumo": mov.get("insumo", ""),
                "Cantidad": float(mov.get("cantidad") or 0),
                "Unidad": mov.get("unidad", ""),
                "Fecha": mov.get("fecha", ""),
                "Cuadrilla": mov.get("cuadrilla", ""),
                "Sector": mov.get("sector", ""),
            }
        )
    tablas = {}
    for clave, registros in filas.items():
        if not registros:
            tablas[clave] = pd.DataFrame(
                columns=["Insumo", "Cantidad", "Unidad", "Movimientos"]
            )
            continue
        df = pd.DataFrame(registros)
        agrupado = (
            df.groupby(["Insumo", "Unidad"], as_index=False)
            .agg(Cantidad=("Cantidad", "sum"), Movimientos=("Insumo", "count"))
        )
        tablas[clave] = agrupado[["Insumo", "Cantidad", "Unidad", "Movimientos"]]
    return tablas


def validar_consumos_campo(consumos: list[dict]) -> str | None:
    for consumo in consumos:
        nombre = str(consumo.get("nombre") or "insumo").strip() or "insumo"
        if float(consumo.get("cantidad") or 0) <= 0:
            return f"Indica una cantidad mayor a 0 para «{nombre}»."
        if consumo.get("origen") not in ORIGENES_INSUMO:
            return f"Selecciona el origen del insumo «{nombre}» (PJP4 o OPS)."
    return validar_piezas_rotativas(consumos)


def validar_piezas_rotativas(consumos: list[dict]) -> str | None:
    for consumo in consumos:
        if str(consumo.get("tipo") or "") != TIPO_INSUMO_ROTATIVO:
            continue
        nombre = str(consumo.get("nombre") or "componente").strip() or "componente"
        condicion = str(consumo.get("condicion_instalada") or "").strip()
        destino = str(consumo.get("destino_retirada") or "").strip()
        if condicion not in CONDICION_PIEZA_INSTALADA:
            return (
                f"Selecciona la condición de la pieza instalada para «{nombre}» "
                "(Nuevo o Reparado por PJP4)."
            )
        if destino not in DESTINO_PIEZA_RETIRADA:
            return f"Selecciona el destino de la pieza retirada para «{nombre}»."
    return None


def registrar_entrada_almacen(
    fecha: date,
    insumo: dict,
    cantidad: float,
    origen: str,
    nota: str = "",
) -> str | None:
    if cantidad <= 0:
        return "Indica una cantidad mayor a 0."
    if origen not in ORIGENES_INSUMO:
        return "Selecciona el origen del insumo (PJP4 o OPS)."
    registrar_movimiento_almacen(
        {
            "tipo": "entrada",
            "fecha": fecha.strftime("%Y-%m-%d"),
            "insumo_id": insumo.get("id", ""),
            "insumo": insumo.get("nombre", ""),
            "cantidad": cantidad,
            "unidad": insumo.get("unidad", ""),
            "origen": origen,
            "origen_registro": "almacen",
            "nota": nota.strip(),
        }
    )
    return None


def inventario_por_origen() -> dict[str, pd.DataFrame]:
    stock: dict[tuple[str, str, str], float] = {}
    for mov in cargar_almacen()["movimientos"]:
        origen = str(mov.get("origen", ""))
        clave = "PJP4" if origen.startswith("PJP4") else "OPS" if "OPS" in origen else ""
        nombre = str(mov.get("insumo") or "").strip()
        unidad = str(mov.get("unidad") or "").strip()
        if not clave or not nombre:
            continue
        key = (clave, nombre, unidad)
        qty = float(mov.get("cantidad") or 0)
        tipo = str(mov.get("tipo", ""))
        if tipo == "entrada":
            stock[key] = stock.get(key, 0) + qty
        elif tipo == "consumo":
            stock[key] = stock.get(key, 0) - qty
    tablas = {}
    for origen_clave in ("PJP4", "OPS"):
        filas = [
            {"Insumo": nombre, "Existencia": qty, "Unidad": unidad}
            for (clave, nombre, unidad), qty in stock.items()
            if clave == origen_clave
        ]
        tablas[origen_clave] = (
            pd.DataFrame(filas).sort_values("Insumo")
            if filas
            else pd.DataFrame(columns=["Insumo", "Existencia", "Unidad"])
        )
    return tablas


def clave_ubicacion(nombre: str) -> str:
    texto = str(nombre or "").strip().upper()
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.replace("_", "-").replace(".", "-")
    return re.sub(r"\s*-\s*", "-", texto).strip("-")


def clave_pozo_plana(nombre: str) -> str:
    """PA67, PA-67 y pa 67 quedan en la misma clave."""
    return re.sub(r"[^A-Z0-9]", "", _sin_acentos(str(nombre or "")).upper())


def _nombres_ubicacion_expandida(texto: str) -> list[str]:
    bruto = " ".join(str(texto or "").split())
    if not bruto:
        return []
    salida = [bruto]
    plano = _sin_acentos(bruto).upper()
    for patron, prefijo in _RE_POZO_NOMBRE:
        hallado = patron.search(plano)
        if hallado:
            salida.append(f"{prefijo}-{hallado.group(1)}")
            break
    return salida


def _variantes_clave_ubicacion(*textos: str) -> set[str]:
    claves: set[str] = set()
    for texto in textos:
        for bruto in _nombres_ubicacion_expandida(texto):
            plana = clave_pozo_plana(bruto)
            if not plana:
                continue
            claves.add(plana)
            canon = _CANON_UBICACION_PLANO.get(plana)
            if canon:
                claves.add(clave_pozo_plana(canon))
            hallado = re.fullmatch(r"(PSN|PPB|PGA|PS|PA|PC|PB|PG)(\d+[A-Z]?)", plana)
            if hallado:
                pref, num = hallado.group(1), hallado.group(2)
                largo = _PREFIJO_POZO_CORTO.get(pref, pref)
                claves.add(f"{pref}{num}")
                claves.add(f"{largo}{num}")
                if num.isdigit():
                    recorte = str(int(num))
                    claves.add(f"{pref}{recorte}")
                    claves.add(f"{largo}{recorte}")
                if largo == "PPB":
                    claves.add(f"PB{num}")
                if largo == "PGA":
                    claves.add(f"PG{num}")
            elif plana.startswith("P") and plana[1:].isdigit():
                claves.add("PS" + plana[1:])
            if "IQSIN" not in plana and "BSMI" not in plana:
                for inst in ("BSNI", "ECSII", "BSII", "BPB", "BSI"):
                    if inst in plana:
                        claves.add(inst)
                        break
    return claves


def canonizar_nombre_ubicacion(texto: str) -> str:
    """Lleva PS116, PB266, ALMACEN OPS o POZO SOLEDAD 549 a la clave del catálogo."""
    bruto = " ".join(str(texto or "").split())
    if not bruto:
        return ""
    for extra in _nombres_ubicacion_expandida(bruto):
        plana = clave_pozo_plana(extra)
        if plana in _CANON_UBICACION_PLANO:
            return _CANON_UBICACION_PLANO[plana]
        hallado = re.fullmatch(r"(PSN|PPB|PGA|PS|PA|PC|PB|PG)(\d+[A-Z]?)", plana)
        if hallado:
            pref, num = hallado.group(1), hallado.group(2)
            pref = _PREFIJO_POZO_CORTO.get(pref, pref)
            numero = str(int(num)) if num.isdigit() else num
            return f"{pref}-{numero}"
        if plana.startswith("P") and plana[1:].isdigit():
            return f"PS-{int(plana[1:])}"
    return bruto


def coincide_prefijo(nombre: str, prefijo: str) -> bool:
    clave = clave_ubicacion(nombre).replace(" ", "")
    pref = str(prefijo or "").upper().strip()
    if not clave or not pref or pref in {"INS", "OTR"}:
        return False
    if clave == pref:
        return True
    if clave.startswith(pref + "-") or clave.startswith(pref + "_"):
        return True
    resto = clave[len(pref) :]
    return clave.startswith(pref) and bool(resto) and resto[0].isdigit()


def mapa_zonas(zonas: list[dict] | None = None) -> dict[str, dict]:
    return {z["id"]: z for z in (zonas or ZONAS_UBICACION_DEFAULT)}


def inferir_zona_id(nombre: str, zonas: list[dict] | None = None) -> str:
    zonas = zonas or ZONAS_UBICACION_DEFAULT
    clave = clave_ubicacion(nombre)
    if clave in INSTALACIONES_SECTOR or clave in {clave_ubicacion(k) for k in INSTALACIONES_SECTOR}:
        return "instalaciones"
    ordenadas = sorted(zonas, key=lambda z: len(str(z.get("prefijo", ""))), reverse=True)
    for zona in ordenadas:
        if coincide_prefijo(nombre, zona.get("prefijo", "")):
            return zona["id"]
    return "otros"


def ficha_vacia(nombre: str, zona_id: str = "", sector: str = "") -> dict:
    zonas = mapa_zonas()
    zona_id = zona_id or inferir_zona_id(nombre)
    zona = zonas.get(zona_id, zonas["otros"])
    sector = sector or INSTALACIONES_SECTOR.get(nombre, "") or zona.get("sector", "")
    return {
        "id": clave_ubicacion(nombre) or slug_insumo(nombre),
        "nombre": nombre.strip(),
        "zona_id": zona_id,
        "prefijo": zona.get("prefijo", "OTR"),
        "sector": sector,
        "equipo_mci": "",
        "equipo_ub": "",
        "estatus": "Operativo",
        "notas": "",
        "actualizado": "",
        "aliases": [],
    }


def _ficha_infraestructura(codigo: str, zona: str, clase: str, descripcion: str) -> dict:
    zona_id = "infra-norte" if zona == "NTE" else "infra-sur"
    return {
        "id": codigo,
        "nombre": codigo,
        "zona_id": zona_id,
        "prefijo": clase,
        "sector": f"{zona} - {clase}",
        "equipo_mci": "",
        "equipo_ub": "",
        "estatus": "Operativo",
        "notas": descripcion,
        "actualizado": "",
    }


def _alias_infraestructura(codigo: str) -> str:
    """TDSGF-IQ-SIN-13 valida el mismo punto que CBN-TDSGF-IQ-SIN-13."""
    texto = clave_ubicacion(codigo)
    clase, separador, resto = texto.partition("-")
    if separador and clase in _CLASES_INFRAESTRUCTURA and resto:
        return resto
    return ""


def _claves_ficha(ficha: dict) -> set[str]:
    valores = [ficha.get("nombre"), ficha.get("id")]
    for valor in list(valores):
        alias = _alias_infraestructura(str(valor or ""))
        if alias:
            valores.append(alias)
    valores.extend(ficha.get("aliases") or [])
    for clave in (ficha.get("id"), ficha.get("nombre")):
        valores.extend(ALIAS_FICHA_CANON.get(str(clave or ""), []))
    sector = str(ficha.get("sector") or "")
    ident = clave_pozo_plana(ficha.get("id") or ficha.get("nombre") or "")
    if ident and ident in clave_pozo_plana(sector):
        valores.append(sector)
    return _variantes_clave_ubicacion(*valores) | _claves_pozo(*valores)


def _indice_infraestructura() -> dict[str, dict]:
    cache = getattr(_indice_infraestructura, "cache", None)
    if cache is not None:
        return cache
    indice: dict[str, dict] = {}
    por_codigo: dict[str, dict] = {}
    for codigo, zona, clase, descripcion in INFRAESTRUCTURA_OFICIAL:
        ficha = _ficha_infraestructura(codigo, zona, clase, descripcion)
        por_codigo[codigo] = ficha
        for plana in _claves_ficha(ficha):
            indice.setdefault(plana, ficha)
        for token in _RE_TOKEN_INFRA.findall(clave_pozo_plana(codigo)):
            token = token.upper()
            indice.setdefault(token, ficha)
            if token.startswith("MPS") and token[3:].isdigit():
                indice.setdefault("OLG" + token[3:], ficha)
                indice.setdefault("OLGPS" + token[3:], ficha)
            if token.startswith("MPG") and token[3:].isdigit():
                indice.setdefault("OLG" + token[3:], ficha)
                indice.setdefault("OLGPGA" + token[3:], ficha)
            if token.startswith("MPSN") and token[4:].isdigit():
                indice.setdefault("OLGPSN" + token[4:], ficha)
    for alias, codigo in _ALIAS_SEG_INFRA.items():
        if codigo in por_codigo:
            indice.setdefault(alias, por_codigo[codigo])
    _indice_infraestructura.cache = indice
    return indice


def ficha_infraestructura(texto: str) -> dict | None:
    """Punto oficial de ducto, recolección o skid, por código completo o sin la clase."""
    indice = _indice_infraestructura()
    for clave in _variantes_clave_ubicacion(texto) | _claves_pozo(texto, canonizar_nombre_ubicacion(texto)):
        hallada = indice.get(clave)
        if hallada:
            return hallada
    return None


def ubicaciones_default() -> dict:
    fichas = {}
    for nombre, zona_id in POZOS_SEMILLA:
        ficha = ficha_vacia(nombre, zona_id)
        fichas[ficha["id"]] = ficha
    for codigo, zona, clase, descripcion in INFRAESTRUCTURA_OFICIAL:
        ficha = _ficha_infraestructura(codigo, zona, clase, descripcion)
        fichas[ficha["id"]] = ficha
    return {
        "zonas": [dict(z) for z in ZONAS_UBICACION_DEFAULT],
        "fichas": fichas,
        "cambios_equipo": [],
    }


def normalizar_ubicaciones(datos) -> dict:
    base = ubicaciones_default()
    if not isinstance(datos, dict):
        return base
    zonas = []
    vistos_zona = set()
    for zona in list(datos.get("zonas") or []) + base["zonas"]:
        if not isinstance(zona, dict):
            continue
        zid = str(zona.get("id") or slug_insumo(zona.get("nombre", ""))).strip()
        if not zid or zid in vistos_zona:
            continue
        vistos_zona.add(zid)
        zonas.append(
            {
                "id": zid,
                "prefijo": str(zona.get("prefijo") or "OTR").strip().upper(),
                "nombre": str(zona.get("nombre") or zid).strip(),
                "sector": str(zona.get("sector") or "").strip(),
            }
        )
    fichas = {}
    origenes = []
    if isinstance(datos.get("fichas"), dict):
        origenes.extend(datos["fichas"].values())
    elif isinstance(datos.get("fichas"), list):
        origenes.extend(datos["fichas"])
    origenes.extend(base["fichas"].values())
    for item in origenes:
        if not isinstance(item, dict):
            continue
        nombre = str(item.get("nombre") or "").strip()
        if not nombre:
            continue
        fid = str(item.get("id") or clave_ubicacion(nombre)).strip()
        if fid in fichas:
            continue
        zona_id = str(item.get("zona_id") or inferir_zona_id(nombre, zonas)).strip()
        zona = mapa_zonas(zonas).get(zona_id, mapa_zonas(zonas).get("otros", {}))
        aliases = [
            " ".join(str(alias or "").split())
            for alias in (item.get("aliases") or [])
            if str(alias or "").strip()
        ]
        fichas[fid] = {
            "id": fid,
            "nombre": nombre,
            "zona_id": zona_id,
            "prefijo": str(item.get("prefijo") or zona.get("prefijo") or "OTR"),
            "sector": str(item.get("sector") or zona.get("sector") or ""),
            "equipo_mci": str(item.get("equipo_mci") or "").strip(),
            "equipo_ub": str(item.get("equipo_ub") or "").strip(),
            "estatus": str(item.get("estatus") or "Operativo").strip(),
            "notas": str(item.get("notas") or "").strip(),
            "actualizado": str(item.get("actualizado") or "").strip(),
            "aliases": aliases,
        }
    cambios = [c for c in datos.get("cambios_equipo", []) if isinstance(c, dict)]
    resultado = {"zonas": zonas, "fichas": fichas, "cambios_equipo": cambios}
    if isinstance(datos.get("migracion"), dict):
        resultado["migracion"] = datos["migracion"]
    return resultado


def cargar_ubicaciones() -> dict:
    if not os.path.exists(UBICACIONES_FILE):
        datos = ubicaciones_default()
        escribir_json(UBICACIONES_FILE, datos)
        return datos
    return normalizar_ubicaciones(leer_json(UBICACIONES_FILE, ubicaciones_default()))


def guardar_ubicaciones(datos: dict) -> None:
    escribir_json(UBICACIONES_FILE, normalizar_ubicaciones(datos))


def importar_pozos_bitacora(datos: dict | None = None) -> dict:
    datos = datos or cargar_ubicaciones()
    df = cargar_bitacora()
    if df.empty:
        return datos
    zonas = datos["zonas"]
    for valor in df["POZO"].tolist():
        nombre = str(valor or "").strip()
        if not nombre or nombre.lower() in {"nan", "none", "-"}:
            continue
        fid = clave_ubicacion(nombre)
        if fid in datos["fichas"] or ficha_infraestructura(nombre):
            continue
        ficha = ficha_vacia(nombre, inferir_zona_id(nombre, zonas))
        sector_fila = ""
        filas = df[df["POZO"].astype(str).str.strip() == nombre]
        if not filas.empty:
            sector_fila = str(filas.iloc[-1].get("SECTOR", "") or "").strip()
        if sector_fila:
            ficha["sector"] = sector_fila
        datos["fichas"][ficha["id"]] = ficha
    return datos


def resolver_ubicacion_historial(texto: str, datos: dict | None = None) -> str:
    """Cruza el texto de campo con el catálogo maestro (ficha_por_ubicacion)."""
    datos = datos if datos is not None else cargar_ubicaciones()
    ficha = ficha_por_ubicacion(datos, texto)
    if not ficha:
        return ""
    return str(ficha.get("nombre") or ficha.get("id") or "").strip()


def sincronizar_catalogo_historial(datos: dict | None = None) -> tuple[dict, int]:
    """Da de alta pozos, módulos y alias que aparecen en el histórico y no están en el JSON."""
    datos = datos or cargar_ubicaciones()
    ruta = historial_pjp4.ruta_archivo()
    if not ruta:
        return datos, 0
    zonas = datos["zonas"]
    cambios = 0
    for bruto in historial_pjp4.listar_ubicaciones(ruta):
        existente = ficha_por_ubicacion(datos, bruto)
        if existente:
            aliases = list(existente.get("aliases") or [])
            if bruto and clave_pozo_plana(bruto) not in _claves_ficha(existente):
                aliases.append(bruto)
                existente["aliases"] = aliases
                datos["fichas"][existente["id"]] = existente
                cambios += 1
                if getattr(_indice_catalogo, "_cache", None):
                    _indice_catalogo._cache = None
            continue
        nombre = canonizar_nombre_ubicacion(bruto) or " ".join(str(bruto).split())
        if not nombre:
            continue
        if ficha_por_ubicacion(datos, nombre):
            continue
        ficha = ficha_vacia(nombre, inferir_zona_id(nombre, zonas))
        if bruto and clave_pozo_plana(bruto) != clave_pozo_plana(nombre):
            ficha["aliases"] = [bruto]
        datos["fichas"][ficha["id"]] = ficha
        cambios += 1
        if getattr(_indice_catalogo, "_cache", None):
            _indice_catalogo._cache = None
    return datos, cambios


def catalogo_con_historial() -> dict:
    original = cargar_ubicaciones()
    n0 = len(original["fichas"])
    datos = importar_pozos_bitacora(original)
    datos, cambios = sincronizar_catalogo_historial(datos)
    if cambios or len(datos["fichas"]) > n0:
        guardar_ubicaciones(datos)
        return cargar_ubicaciones()
    return datos


def cargar_historial_campo() -> dict:
    ruta = historial_pjp4.ruta_archivo()
    mtime_hist = os.path.getmtime(ruta) if ruta and os.path.isfile(ruta) else 0
    mtime_cat = os.path.getmtime(UBICACIONES_FILE) if os.path.exists(UBICACIONES_FILE) else 0
    firma = (ruta, mtime_hist, mtime_cat)
    cache = getattr(cargar_historial_campo, "_cache", None)
    if cache and cache[0] == firma:
        return cache[1]
    datos = cargar_ubicaciones()
    resultado = historial_pjp4.procesar(
        ruta,
        lambda texto, catalogo=datos: resolver_ubicacion_historial(texto, catalogo),
    )
    cargar_historial_campo._cache = (firma, resultado)
    return resultado


def historial_de_pozo(ficha: dict) -> list[dict]:
    filas = []
    for fila in cargar_historial_campo().get("filas") or []:
        ubicacion = fila.get("UBICACIÓN") or fila.get("POZO") or ""
        if _texto_coincide_pozo(ubicacion, ficha):
            filas.append(fila)
    return filas


def eventos_historial_campo_pozo(ficha: dict) -> list[dict]:
    eventos: list[dict] = []
    for fila in historial_de_pozo(ficha):
        fch = fila.get("FECHA")
        if isinstance(fch, datetime):
            fch = fch.date()
        elif not isinstance(fch, date):
            fch = parsear_fecha_valor(fch)
        desc = str(
            fila.get("DESCRIPCIÓN DE LA ACTIVIDAD")
            or fila.get("DESCRIPCION DE LA ACTIVIDAD")
            or ""
        ).strip()
        eventos.append(
            {
                "idx": None,
                "fecha": fch,
                "fecha_txt": fch.strftime("%d/%m/%Y") if isinstance(fch, date) else "s/f",
                "sort": fch.isoformat() if isinstance(fch, date) else "0000-00-00",
                "tipo": "historial",
                "clase": "Histórico",
                "titulo": desc or "Actividad histórica",
                "cuadrilla": str(fila.get("CUADRILLA") or "").strip(),
                "descripcion": str(fila.get("COMENTARIOS") or "").strip(),
                "estatus": "",
                "equipo": str(fila.get("UBICACIÓN") or "").strip(),
                "insumos": insumos_de_fila(fila),
                "secuencia": "",
                "evidencia": "",
                "fotos": 0,
            }
        )
    return eventos


ZONAS_TITULO_EXCEL = {
    "ARAGON": "aragon",
    "PALO BLANCO": "palo-blanco",
    "COYOTES": "coyotes",
    "GALLO": "gallo",
    "SOLEDAD NORTE": "soledad-norte",
    "SOLEDAD": "soledad",
}


def _texto_celda(valor) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor).strip()


def _norm_titulo(valor: str) -> str:
    texto = _texto_celda(valor).upper()
    return (
        texto.replace("Á", "A")
        .replace("É", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Ú", "U")
        .replace("Ü", "U")
        .replace("Ñ", "N")
    )


def _es_encabezado_bloque(valor: str) -> bool:
    return _norm_titulo(valor) in {"UBICACION", "UBICACION:", "POZO"}


def _zona_id_desde_titulo(valor: str) -> str:
    titulo = _norm_titulo(valor)
    if titulo in ZONAS_TITULO_EXCEL:
        return ZONAS_TITULO_EXCEL[titulo]
    return ""


def normalizar_mci_excel(raw: str) -> str:
    texto = _norm_titulo(raw).replace("AROW", "ARROW")
    if not texto:
        return ""
    if "WEICHAI" in texto:
        return "MCI tipo Weichai"
    if "CUMMINS" in texto:
        return "MCI tipo Cummins"
    if "VORTEC" in texto and "V6" in texto:
        return "MCI tipo Vortec V6"
    if "VORTEC" in texto:
        return "MCI tipo Vortec V8"
    if "ARROW" in texto:
        return "MCI tipo Arrow"
    return raw.strip()


def normalizar_ub_excel(raw: str) -> str:
    texto = _norm_titulo(raw)
    if not texto:
        return ""
    if "HLU" in texto or "HIDRAUL" in texto:
        return "HLU (Unidad Hidráulica)"
    if "LUFKIN" in texto:
        return "UBM tipo Lufkin"
    if "WEATHER" in texto:
        return "UBM tipo Weatherford"
    return raw.strip()


def listar_excels_inventario_ubicaciones() -> list[str]:
    carpetas = [UBICACIONES_EXCEL_DIR, PROGRAMAS_EXCEL_DIR, "."]
    vistos = set()
    archivos = []
    for carpeta in carpetas:
        if not os.path.isdir(carpeta):
            continue
        try:
            nombres = os.listdir(carpeta)
        except OSError:
            continue
        for nombre in nombres:
            if nombre.startswith("~$"):
                continue
            if not nombre.lower().endswith((".xlsx", ".xls", ".xlsm")):
                continue
            clave = nombre.casefold()
            if carpeta != UBICACIONES_EXCEL_DIR and "ubicacion" not in clave:
                continue
            ruta = os.path.abspath(os.path.join(carpeta, nombre))
            if ruta in vistos:
                continue
            vistos.add(ruta)
            archivos.append(ruta)
    return sorted(archivos, key=lambda r: (0 if UBICACIONES_EXCEL_DIR in r else 1, r.casefold()))


def excel_inventario_predeterminado() -> str:
    archivos = listar_excels_inventario_ubicaciones()
    return archivos[0] if archivos else ""


def _copia_excel_legible(ruta: str) -> str:
    """Copia fresca por mtime para no leer un temporal viejo si OneDrive bloquea el original."""
    sello = str(int(_mtime_archivo(ruta) * 1000))
    destino = os.path.join(tempfile.gettempdir(), f"pjp4_{sello}_{os.path.basename(ruta)}")
    try:
        shutil.copy2(ruta, destino)
        return destino
    except OSError:
        return ruta


def _tipo_columna_catalogo_excel(valor: str) -> str:
    titulo = _norm_titulo(valor)
    if not titulo:
        return ""
    if "NOMBRE CORTO" in titulo:
        return "corto"
    if "NOMBRE LARGO" in titulo:
        return "largo"
    if re.search(r"\bMCI\b", titulo):
        return "mci"
    if titulo in {"UB", "UBM"} or re.search(r"\bUBM\b", titulo) or titulo.startswith("UNIDAD DE BOMBEO"):
        return "ub"
    if _es_encabezado_bloque(valor):
        return "corto"
    return ""


def _zona_bloque_catalogo_excel(valor: str) -> str:
    titulo = _norm_titulo(valor)
    if not titulo:
        return ""
    if titulo.startswith("UBICACION DE BATERI") or titulo.startswith("UBICACION DE ESTACION"):
        return "instalaciones"
    if titulo.startswith("UBICACION DE") or titulo in ZONAS_TITULO_EXCEL:
        for clave, zona_id in sorted(ZONAS_TITULO_EXCEL.items(), key=lambda item: -len(item[0])):
            if clave in titulo:
                return zona_id
    return _zona_id_desde_titulo(valor)


def _canon_corto_instalacion(corto: str) -> str:
    texto = " ".join(str(corto or "").split())
    if not texto:
        return ""
    sitio = sitios_operativos.clave_sitio(texto)
    if sitio:
        return sitio
    plano = clave_pozo_plana(texto)
    return _CANON_UBICACION_PLANO.get(plano) or texto


def extraer_catalogo_maestro_excel(ruta: str) -> list[dict]:
    """Lee las tablas Nombre Largo (Excel) / Nombre Corto (App) del inventario."""
    from openpyxl import load_workbook

    origen = _copia_excel_legible(ruta)
    wb = load_workbook(origen, data_only=True, read_only=True)
    try:
        hoja = wb[wb.sheetnames[0]]
        grid = [list(fila) for fila in hoja.iter_rows(values_only=True)]
    finally:
        wb.close()
    if not grid:
        return []

    bloques: list[tuple[int, int, int, str]] = []
    for r, fila in enumerate(grid):
        titulos = []
        for c, valor in enumerate(fila):
            zona_id = _zona_bloque_catalogo_excel(_texto_celda(valor))
            if zona_id:
                titulos.append((c, zona_id))
        for i, (c0, zona_id) in enumerate(titulos):
            c1 = titulos[i + 1][0] if i + 1 < len(titulos) else max(len(fila), c0 + 6)
            bloques.append((r, c0, c1, zona_id))

    filas: list[dict] = []
    vistos: set[str] = set()
    pares_maestro: list[tuple[str, str]] = []
    for r0, c0, c1, zona_id in bloques:
        cols = {"corto": None, "largo": None, "mci": None, "ub": None}
        inicio = r0 + 1
        fila_header = None
        for r in range(r0 + 1, min(r0 + 4, len(grid))):
            fila = grid[r]
            hallados = False
            for c in range(c0, min(c1, len(fila))):
                tipo = _tipo_columna_catalogo_excel(_texto_celda(fila[c]))
                if tipo:
                    cols[tipo] = c
                    hallados = True
            if hallados:
                fila_header = r
                inicio = r + 1
                break
        if fila_header is not None and (cols["mci"] is None or cols["ub"] is None):
            fila = grid[fila_header]
            origen = max((c for c in cols.values() if c is not None), default=c0) + 1
            for c in range(origen, min(len(fila), max(c1, origen) + 4)):
                if c >= c1:
                    titulo_zona = _texto_celda(grid[r0][c]) if c < len(grid[r0]) else ""
                    if _zona_bloque_catalogo_excel(titulo_zona):
                        break
                    tipo_siguiente = _tipo_columna_catalogo_excel(_texto_celda(fila[c]))
                    if tipo_siguiente in {"corto", "largo"}:
                        break
                tipo = _tipo_columna_catalogo_excel(_texto_celda(fila[c]))
                if tipo in {"mci", "ub"} and cols[tipo] is None:
                    cols[tipo] = c
        if cols["corto"] is None and cols["largo"] is None:
            continue
        zona = mapa_zonas().get(zona_id, {})
        es_instalacion = zona_id == "instalaciones"
        vacios = 0
        for r in range(inicio, len(grid)):
            fila = grid[r]
            corto = _texto_celda(fila[cols["corto"]]) if cols["corto"] is not None and cols["corto"] < len(fila) else ""
            largo = _texto_celda(fila[cols["largo"]]) if cols["largo"] is not None and cols["largo"] < len(fila) else ""
            if _zona_bloque_catalogo_excel(corto) or _zona_bloque_catalogo_excel(largo):
                break
            if _tipo_columna_catalogo_excel(corto) or _tipo_columna_catalogo_excel(largo):
                break
            if not corto and not largo:
                vacios += 1
                if vacios >= 3:
                    break
                continue
            vacios = 0
            mci_raw = _texto_celda(fila[cols["mci"]]) if cols["mci"] is not None and cols["mci"] < len(fila) else ""
            ub_raw = _texto_celda(fila[cols["ub"]]) if cols["ub"] is not None and cols["ub"] < len(fila) else ""
            nombre = _canon_corto_instalacion(corto or largo) if es_instalacion else (corto or largo)
            if not nombre or sitios_operativos.es_vacio_sitio(nombre):
                continue
            clave = clave_ubicacion(nombre)
            if clave in vistos:
                if largo and es_instalacion:
                    pares_maestro.append((largo, nombre))
                continue
            vistos.add(clave)
            aliases = []
            if largo and clave_pozo_plana(largo) != clave_pozo_plana(nombre):
                aliases.append(largo)
            if corto and corto != nombre:
                aliases.append(corto)
            if es_instalacion:
                sitios_operativos.registrar_par_catalogo(largo, nombre)
                pares_maestro.append((largo, nombre))
            filas.append(
                {
                    "nombre": nombre,
                    "zona_id": zona_id,
                    "sector": (
                        INSTALACIONES_SECTOR.get(nombre)
                        or sitios_operativos.nombre_oficial_sitio(nombre)
                        or zona.get("sector", "")
                    ),
                    "prefijo": "INS" if es_instalacion else zona.get("prefijo", ""),
                    "equipo_mci": "" if es_instalacion else normalizar_mci_excel(mci_raw),
                    "equipo_ub": "" if es_instalacion else normalizar_ub_excel(ub_raw),
                    "mci_raw": "" if es_instalacion else mci_raw,
                    "ub_raw": "" if es_instalacion else ub_raw,
                    "aliases": aliases,
                    "instalacion": es_instalacion,
                    "largo_excel": largo,
                    "corto_app": nombre,
                }
            )
    extraer_catalogo_maestro_excel._pares = pares_maestro
    return filas


def extraer_pozos_desde_excel(ruta: str) -> list[dict]:
    filas = extraer_catalogo_maestro_excel(ruta)
    if filas:
        return filas
    from openpyxl import load_workbook

    origen = _copia_excel_legible(ruta)
    wb = load_workbook(origen, data_only=True, read_only=True)
    try:
        hoja = wb[wb.sheetnames[0]]
        grid = [list(fila) for fila in hoja.iter_rows(values_only=True)]
    finally:
        wb.close()

    bloques: list[tuple[int, int, str]] = []
    for r, fila in enumerate(grid):
        for c, valor in enumerate(fila):
            zona_id = _zona_id_desde_titulo(_texto_celda(valor))
            if zona_id:
                bloques.append((r, c, zona_id))

    filas = []
    vistos: set[str] = set()
    for r0, c0, zona_id in bloques:
        inicio = r0 + 1
        for r in range(r0 + 1, min(r0 + 5, len(grid))):
            if c0 < len(grid[r]) and _es_encabezado_bloque(_texto_celda(grid[r][c0])):
                inicio = r + 1
                break
        for r in range(inicio, len(grid)):
            fila = grid[r]
            if c0 >= len(fila):
                break
            ubicacion = _texto_celda(fila[c0])
            if _zona_id_desde_titulo(ubicacion) or _es_encabezado_bloque(ubicacion):
                break
            if not ubicacion:
                continue
            if ubicacion.upper() in {"UBICACION Y EQUIPOS", "MCI", "UB", "UBM"}:
                continue
            mci_raw = _texto_celda(fila[c0 + 1]) if c0 + 1 < len(fila) else ""
            ub_raw = _texto_celda(fila[c0 + 2]) if c0 + 2 < len(fila) else ""
            clave = clave_ubicacion(ubicacion)
            if clave in vistos:
                continue
            vistos.add(clave)
            zona = mapa_zonas().get(zona_id, {})
            filas.append(
                {
                    "nombre": ubicacion,
                    "zona_id": zona_id,
                    "sector": zona.get("sector", ""),
                    "prefijo": zona.get("prefijo", ""),
                    "equipo_mci": normalizar_mci_excel(mci_raw),
                    "equipo_ub": normalizar_ub_excel(ub_raw),
                    "mci_raw": mci_raw,
                    "ub_raw": ub_raw,
                    "aliases": [],
                    "instalacion": zona_id == "instalaciones",
                }
            )
    return filas


def migracion_excel_pendiente(ruta: str = "") -> bool:
    ruta = ruta or excel_inventario_predeterminado()
    if not ruta or not os.path.exists(ruta):
        return False
    mig = cargar_ubicaciones().get("migracion") or {}
    try:
        mtime = os.path.getmtime(ruta)
    except OSError:
        return False
    if str(mig.get("archivo", "")) != os.path.abspath(ruta):
        return True
    return abs(float(mig.get("mtime") or 0) - mtime) > 1


def importar_excel_ubicaciones(ruta: str = "", *, forzar: bool = False) -> dict:
    ruta = ruta or excel_inventario_predeterminado()
    if not ruta:
        return {"ok": False, "error": "No hay un Excel de inventario en ubicacion_equipo/."}
    if not forzar and not migracion_excel_pendiente(ruta):
        return {"ok": True, "omitido": True, "nuevos": 0, "actualizados": 0, "total": 0}
    try:
        extraidos = extraer_pozos_desde_excel(ruta)
    except Exception as exc:
        return {"ok": False, "error": f"No se pudo leer el Excel: {exc}"}
    if not extraidos:
        return {"ok": False, "error": "El Excel no tiene tablas de UBICACIÓN / MCI / UB reconocibles."}

    datos = cargar_ubicaciones()
    nuevos = 0
    actualizados = 0
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    for item in extraidos:
        aliases_item = [a for a in (item.get("aliases") or []) if str(a).strip()]
        existente = encontrar_ficha(datos, item["nombre"]) or (
            encontrar_ficha(datos, aliases_item[0]) if aliases_item else None
        )
        if existente:
            existente["zona_id"] = item["zona_id"] or existente.get("zona_id", "")
            existente["prefijo"] = item["prefijo"] or existente.get("prefijo", "")
            existente["sector"] = item["sector"] or existente.get("sector", "")
            aliases = list(existente.get("aliases") or [])
            for alias in aliases_item:
                if alias not in aliases and clave_pozo_plana(alias) != clave_pozo_plana(existente.get("nombre")):
                    aliases.append(alias)
            existente["aliases"] = aliases
            if not item.get("instalacion"):
                if item.get("equipo_mci"):
                    existente["equipo_mci"] = item["equipo_mci"]
                if item.get("equipo_ub"):
                    existente["equipo_ub"] = item["equipo_ub"]
            existente["actualizado"] = ahora
            datos["fichas"][existente["id"]] = existente
            actualizados += 1
        else:
            ficha = ficha_vacia(item["nombre"], item["zona_id"], item["sector"])
            ficha["equipo_mci"] = "" if item.get("instalacion") else item.get("equipo_mci", "")
            ficha["equipo_ub"] = "" if item.get("instalacion") else item.get("equipo_ub", "")
            ficha["aliases"] = aliases_item
            ficha["actualizado"] = ahora
            datos["fichas"][ficha["id"]] = ficha
            nuevos += 1

    try:
        mtime = os.path.getmtime(ruta)
    except OSError:
        mtime = 0
    datos["migracion"] = {
        "archivo": os.path.abspath(ruta),
        "mtime": mtime,
        "fecha": ahora,
        "nuevos": nuevos,
        "actualizados": actualizados,
        "total_excel": len(extraidos),
    }
    guardar_ubicaciones(datos)
    return {
        "ok": True,
        "omitido": False,
        "nuevos": nuevos,
        "actualizados": actualizados,
        "total": len(extraidos),
        "archivo": os.path.basename(ruta),
        "preview": extraidos,
    }


def orden_natural_pozo(nombre: str) -> tuple:
    texto = str(nombre or "").strip()
    m = re.match(r"^([A-Za-z]+)\s*-?\s*(\d+)", texto)
    if m:
        return (m.group(1).upper(), int(m.group(2)), texto.upper())
    return (texto.upper(), 0, texto.upper())


def titulo_zona_historial(zona_id: str, zona: dict | None = None) -> str:
    for item in ZONAS_HISTORIAL_POZOS:
        if item["id"] == zona_id:
            return item["titulo"]
    zona = zona or {}
    prefijo = str(zona.get("prefijo") or "").strip()
    nombre = str(zona.get("nombre") or zona_id)
    return f"{prefijo} {nombre}".strip()


def fichas_de_zona(datos: dict, zona_id: str) -> list[dict]:
    fichas = [f for f in datos["fichas"].values() if f.get("zona_id") == zona_id]
    return sorted(fichas, key=lambda f: orden_natural_pozo(f.get("nombre", "")))


def obtener_ficha(datos: dict, ficha_id: str) -> dict | None:
    return datos["fichas"].get(ficha_id)


def _indice_catalogo(datos: dict) -> dict[str, dict]:
    fichas = datos.get("fichas") or {}
    marca = (id(datos), len(fichas))
    cache = getattr(_indice_catalogo, "_cache", None)
    if cache and cache[0] == marca:
        return cache[1]
    indice: dict[str, dict] = {}
    for ficha in fichas.values():
        for clave in _claves_ficha(ficha):
            indice.setdefault(clave, ficha)
    _indice_catalogo._cache = (marca, indice)
    return indice


def encontrar_ficha(datos: dict, nombre: str) -> dict | None:
    nombre = str(nombre or "").strip()
    if not nombre:
        return None
    candidatos = [nombre]
    canon = canonizar_nombre_ubicacion(nombre)
    if canon and canon.casefold() != nombre.casefold():
        candidatos.append(canon)
    indice = _indice_catalogo(datos)
    for candidato in candidatos:
        fid = clave_ubicacion(candidato)
        if fid in datos["fichas"]:
            return datos["fichas"][fid]
        for clave in _variantes_clave_ubicacion(candidato):
            hallada = indice.get(clave)
            if hallada:
                return hallada
        for ficha in datos["fichas"].values():
            if str(ficha.get("nombre", "")).casefold() == candidato.casefold():
                return ficha
    return ficha_infraestructura(nombre) or ficha_infraestructura(canon)


def reapuntar_cambios_equipo(
    datos: dict,
    old_id: str,
    old_nombre: str,
    new_id: str,
    new_nombre: str,
) -> None:
    old_clave = clave_ubicacion(old_nombre)
    for cambio in datos.get("cambios_equipo", []):
        if str(cambio.get("pozo_id", "")) == old_id or clave_ubicacion(cambio.get("pozo", "")) == old_clave:
            cambio["pozo_id"] = new_id
            cambio["pozo"] = new_nombre
        if str(cambio.get("origen_id", "")) == old_id or clave_ubicacion(cambio.get("origen", "")) == old_clave:
            cambio["origen_id"] = new_id
            cambio["origen"] = new_nombre


def alta_ubicacion(nombre: str, zona_id: str, sector: str = "") -> str | None:
    nombre = nombre.strip()
    if not nombre:
        return "Escribe el nombre o clave de la ubicación."
    datos = cargar_ubicaciones()
    if encontrar_ficha(datos, nombre):
        return (
            f"«{nombre}» ya está registrada. Activa «Actualizar si ya existe» "
            "para corregir zona o sector, o edítala desde su ficha."
        )
    if zona_id not in mapa_zonas(datos["zonas"]):
        zona_id = inferir_zona_id(nombre, datos["zonas"])
    ficha = ficha_vacia(nombre, zona_id, sector)
    datos["fichas"][ficha["id"]] = ficha
    guardar_ubicaciones(datos)
    return None


def actualizar_datos_ubicacion(
    ficha_id: str,
    nombre: str,
    zona_id: str,
    sector: str,
) -> str | None:
    nombre = nombre.strip()
    if not nombre:
        return "Escribe el nombre o clave de la ubicación."
    datos = cargar_ubicaciones()
    if ficha_id not in datos["fichas"]:
        return "No se encontró la ubicación a editar."
    zonas = mapa_zonas(datos["zonas"])
    if zona_id not in zonas:
        return "Selecciona una zona válida."
    nuevo_id = clave_ubicacion(nombre) or ficha_id
    if nuevo_id != ficha_id and nuevo_id in datos["fichas"]:
        return f"La clave «{nombre}» ya pertenece a otra ubicación."
    actual = dict(datos["fichas"][ficha_id])
    old_nombre = str(actual.get("nombre") or "")
    zona = zonas[zona_id]
    actual["nombre"] = nombre
    actual["zona_id"] = zona_id
    actual["prefijo"] = zona.get("prefijo", actual.get("prefijo", "OTR"))
    actual["sector"] = sector.strip()
    actual["id"] = nuevo_id
    actual["actualizado"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    if nuevo_id != ficha_id:
        del datos["fichas"][ficha_id]
        reapuntar_cambios_equipo(datos, ficha_id, old_nombre, nuevo_id, nombre)
    datos["fichas"][nuevo_id] = actual
    guardar_ubicaciones(datos)
    return None


def eliminar_ubicacion(ficha_id: str) -> str | None:
    datos = cargar_ubicaciones()
    if ficha_id not in datos["fichas"]:
        return "No se encontró la ubicación."
    del datos["fichas"][ficha_id]
    guardar_ubicaciones(datos)
    return None


def guardar_o_actualizar_ubicacion(
    nombre: str,
    zona_id: str,
    sector: str,
    actualizar: bool,
) -> tuple[str | None, str]:
    nombre = nombre.strip()
    datos = cargar_ubicaciones()
    existente = encontrar_ficha(datos, nombre)
    if existente:
        if not actualizar:
            return (
                f"«{nombre}» ya está registrada. Activa «Actualizar si ya existe» "
                "para corregir zona o sector, o edítala desde su ficha.",
                existente["id"],
            )
        error = actualizar_datos_ubicacion(existente["id"], nombre, zona_id, sector)
        nuevo = encontrar_ficha(cargar_ubicaciones(), nombre)
        return error, (nuevo["id"] if nuevo else existente["id"])
    error = alta_ubicacion(nombre, zona_id, sector)
    creado = encontrar_ficha(cargar_ubicaciones(), nombre)
    return error, (creado["id"] if creado else clave_ubicacion(nombre))


def actualizar_ficha(ficha_id: str, campos: dict) -> None:
    datos = cargar_ubicaciones()
    if ficha_id not in datos["fichas"]:
        return
    actual = datos["fichas"][ficha_id]
    actual.update({k: v for k, v in campos.items() if k != "id"})
    actual["actualizado"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    datos["fichas"][ficha_id] = actual
    guardar_ubicaciones(datos)


def registrar_cambio_equipo(
    ficha_id: str,
    tipo: str,
    equipo_nuevo: str,
    origen_id: str = "",
    motivo: str = "",
    fecha: date | None = None,
) -> str | None:
    if tipo not in {"MCI", "UB"}:
        return "Selecciona si el cambio es de motor MCI o de unidad de bombeo."
    datos = cargar_ubicaciones()
    destino = datos["fichas"].get(ficha_id)
    if not destino:
        return "No se encontró la ubicación destino."
    campo = "equipo_mci" if tipo == "MCI" else "equipo_ub"
    origen = datos["fichas"].get(origen_id) if origen_id else None
    movido = ""
    if origen:
        movido = str(origen.get(campo) or "").strip()
    equipo_nuevo = (equipo_nuevo or "").strip() or movido
    if not equipo_nuevo:
        return "Indica el equipo que se instala o selecciona el pozo de origen."
    anterior = str(destino.get(campo) or "").strip()
    destino[campo] = equipo_nuevo
    destino["actualizado"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    if origen:
        origen[campo] = ""
        origen["actualizado"] = destino["actualizado"]
        datos["fichas"][origen["id"]] = origen
    datos["fichas"][ficha_id] = destino
    datos["cambios_equipo"].append(
        {
            "id": uuid.uuid4().hex[:10],
            "fecha": (fecha or date.today()).strftime("%Y-%m-%d"),
            "tipo": tipo,
            "pozo": destino["nombre"],
            "pozo_id": ficha_id,
            "equipo_anterior": anterior,
            "equipo_nuevo": equipo_nuevo,
            "origen": origen["nombre"] if origen else "",
            "origen_id": origen_id,
            "motivo": motivo.strip(),
            "registrado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
    )
    guardar_ubicaciones(datos)
    return None


def _claves_pozo(*valores) -> set[str]:
    return {clave_pozo_plana(valor) for valor in valores if clave_pozo_plana(valor)}


def cambios_de_pozo(datos: dict, ficha: dict) -> list[dict]:
    nombres = _claves_pozo(ficha.get("nombre"), ficha.get("id"))
    historial = []
    for cambio in datos.get("cambios_equipo", []):
        ids = _claves_pozo(
            cambio.get("pozo"),
            cambio.get("pozo_id"),
            cambio.get("origen"),
            cambio.get("origen_id"),
        )
        if nombres & ids:
            historial.append(cambio)
    return sorted(historial, key=lambda c: str(c.get("fecha", "")), reverse=True)


def _texto_coincide_pozo(valor: str, ficha: dict) -> bool:
    texto = str(valor or "").strip()
    if not texto or texto.lower() in {"nan", "none"}:
        return False
    claves = _claves_ficha(ficha)
    if not claves:
        return False
    if clave_pozo_plana(texto) in claves:
        return True
    unido = re.sub(
        r"\b(PSN|PGA|PPB|PS|PA|PB|PC)[\s\-_.]+(\d+[A-Za-z]?)\b",
        r"\1\2",
        texto,
        flags=re.IGNORECASE,
    )
    for trozo in re.split(r"[^A-Za-z0-9]+", unido):
        if clave_pozo_plana(trozo) in claves:
            return True
    plana = clave_pozo_plana(texto)
    if plana.isdigit() and len(plana) >= 2:
        return any(clave.endswith(plana) and clave[: -len(plana)].isalpha() for clave in claves)
    return False


def bitacora_de_pozo(df: pd.DataFrame, ficha: dict) -> pd.DataFrame:
    if df.empty:
        return df
    columnas = [col for col in ("POZO", "UBICACIÓN", "EQUIPO/UBICACIÓN", "SECTOR") if col in df.columns]
    if columnas:
        mask_pozo = df.apply(
            lambda fila: any(_texto_coincide_pozo(fila.get(col), ficha) for col in columnas),
            axis=1,
        )
        if mask_pozo.any():
            return df.loc[mask_pozo].copy()
    sector = str(ficha.get("sector") or "").strip()
    if sector:
        mask_sec = df["SECTOR"].astype(str).str.strip() == sector
        vacio_pozo = df["POZO"].astype(str).str.strip().isin({"", "nan", "None"})
        return df.loc[mask_sec & vacio_pozo].copy()
    return df.iloc[0:0].copy()


def clasificar_mantenimiento(fila: pd.Series) -> str:
    tipo = str(fila.get("TIPO DE ACTIVIDAD", "") or "").strip()
    if tipo in TIPOS_MANTENIMIENTO_PROGRAMADO and not es_registro_no_programado(fila):
        return "Programado"
    if es_registro_no_programado(fila):
        return "No programado"
    if tipo in TIPOS_MANTENIMIENTO_PROGRAMADO:
        return "Programado"
    return "Bitácora"


def programados_de_pozo(ficha: dict) -> list[dict]:
    sector = str(ficha.get("sector") or "").strip()
    nombre = ficha.get("nombre", "")
    relacionados = []
    for item in cargar_programados_ops():
        ubic = str(item.get("ubicacion") or "").strip()
        extra = " ".join(
            [str(item.get("actividad", "")), str(item.get("observacion", "")), ubic]
        )
        por_nombre = _texto_coincide_pozo(extra, ficha) or _texto_coincide_pozo(ubic, ficha)
        por_sector = bool(sector) and (ubic == sector or sector.casefold() in ubic.casefold())
        if por_nombre or por_sector:
            relacionados.append(item)
    return sorted(relacionados, key=lambda i: str(i.get("mes", "")), reverse=True)


def parsear_fecha_valor(valor) -> date | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, pd.Timestamp):
        if pd.isna(valor):
            return None
        return valor.date()
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if 20000 <= float(valor) <= 80000:
            dt = pd.to_datetime(float(valor), unit="D", origin="1899-12-30", errors="coerce")
            if pd.isna(dt):
                return None
            return dt.date()
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nan", "none", "nat", "<na>"}:
        return None
    if re.match(r"^\d{4}-\d{2}-\d{2}", texto):
        dt = pd.to_datetime(texto, errors="coerce")
    else:
        dt = pd.to_datetime(texto, errors="coerce", dayfirst=True)
    if pd.isna(dt):
        return None
    return dt.date()


def insumos_de_fila(fila) -> list[str]:
    partes: list[str] = []
    for col in COLUMNAS_INSUMOS:
        try:
            val = float(fila.get(col) or 0)
        except (TypeError, ValueError):
            continue
        if val > 0:
            partes.append(f"{col}: {val:g}")
    try:
        bandas = float(fila.get("BANDAS CANTIDAD") or 0)
    except (TypeError, ValueError):
        bandas = 0
    if bandas > 0:
        desc = str(fila.get("BANDAS DESCRIPCIÓN") or "").strip()
        partes.append(f"Bandas: {bandas:g}" + (f" ({desc})" if desc else ""))
    try:
        alternador = float(fila.get("ALTERNADOR CANTIDAD") or 0)
    except (TypeError, ValueError):
        alternador = 0
    if alternador > 0:
        desc = str(fila.get("ALTERNADOR DESCRIPCIÓN") or "").strip()
        partes.append(f"Alternador: {alternador:g}" + (f" ({desc})" if desc else ""))
    otros = str(fila.get("OTROS INSUMOS") or "").strip()
    if otros and otros.lower() not in {"0", "nan", "none"}:
        partes.append(otros)
    piezas = str(fila.get("PIEZAS ROTATIVAS") or "").strip()
    if piezas and piezas.lower() not in {"0", "nan", "none"}:
        partes.append(piezas)
    return partes


def fecha_de_programado(item: dict) -> date | None:
    for clave in ("fecha", "fecha_programada", "dia"):
        extraida = parsear_fecha_valor(item.get(clave))
        if extraida:
            return extraida
    extra = " ".join(
        [
            str(item.get("observacion") or ""),
            str(item.get("actividad") or ""),
            str(item.get("ubicacion") or ""),
        ]
    )
    m = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", extra)
    if m:
        texto = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
        extraida = parsear_fecha_valor(texto)
        if extraida:
            return extraida
    mes = str(item.get("mes") or "").strip()
    m_dia = re.search(r"(?:d[ií]a|el)\s+(\d{1,2})\b", extra, flags=re.I)
    if m_dia and re.match(r"^\d{4}-\d{2}$", mes):
        try:
            return date(int(mes[:4]), int(mes[5:7]), int(m_dia.group(1)))
        except ValueError:
            return None
    return None


def fechas_bitacora_pozo(df_pozo: pd.DataFrame) -> set[date]:
    if df_pozo is None or df_pozo.empty:
        return set()
    fechas = set()
    for valor in df_pozo["FECHA"].tolist():
        extraida = parsear_fecha_valor(valor)
        if extraida:
            fechas.add(extraida)
    return fechas


def programados_ordenados_pozo(ficha: dict) -> list[dict]:
    def _clave(item: dict):
        estatus = str(item.get("estatus") or "")
        pendiente = 0 if estatus in {"No Cumplido", "Reprogramado"} else 1
        mes = str(item.get("mes") or "0000-00")
        try:
            peso_mes = -int(mes.replace("-", "") or 0)
        except ValueError:
            peso_mes = 0
        return (pendiente, peso_mes, str(item.get("actividad") or ""))

    return sorted(programados_de_pozo(ficha), key=_clave)


def eventos_historial_pozo(df_pozo: pd.DataFrame, cambios: list[dict]) -> list[dict]:
    eventos: list[dict] = []
    if df_pozo is not None and not df_pozo.empty:
        for idx, fila in df_pozo.iterrows():
            fch = parsear_fecha_valor(fila.get("FECHA"))
            clase = clasificar_mantenimiento(fila)
            eventos.append(
                {
                    "idx": idx,
                    "fecha": fch,
                    "fecha_txt": fch.strftime("%d/%m/%Y") if fch else str(fila.get("FECHA") or "s/f"),
                    "sort": fch.isoformat() if fch else "0000-00-00",
                    "tipo": "bitacora",
                    "clase": clase,
                    "titulo": str(fila.get("ACTIVIDAD CORTA") or fila.get("TIPO DE ACTIVIDAD") or "Actividad de bitácora").strip()
                    or "Actividad de bitácora",
                    "cuadrilla": str(fila.get("CUADRILLA") or "").strip(),
                    "descripcion": str(
                        fila.get("DESCRIPCIÓN DE LA ACTIVIDAD")
                        or fila.get("DESCRIPCION DE LA ACTIVIDAD")
                        or ""
                    ).strip(),
                    "estatus": str(fila.get("EVALUACIÓN DEL ESTATUS") or "").strip(),
                    "equipo": str(fila.get("EQUIPO/UBICACIÓN") or fila.get("UBICACIÓN") or "").strip(),
                    "insumos": insumos_de_fila(fila),
                    "secuencia": str(fila.get("SECUENCIA") or "").strip(),
                    "evidencia": estatus_evidencia_fila(fila),
                    "fotos": len(rutas_fotos_fila(fila)),
                }
            )
    for cambio in cambios or []:
        fch = parsear_fecha_valor(cambio.get("fecha"))
        tipo_eq = str(cambio.get("tipo") or "equipo")
        eventos.append(
            {
                "fecha": fch,
                "fecha_txt": fch.strftime("%d/%m/%Y") if fch else str(cambio.get("fecha") or "s/f"),
                "sort": fch.isoformat() if fch else "0000-00-00",
                "tipo": "cambio",
                "clase": "Cambio de equipo",
                "titulo": f"Cambio de {tipo_eq}: {cambio.get('equipo_nuevo') or 'equipo instalado'}",
                "cuadrilla": "",
                "descripcion": str(cambio.get("motivo") or "").strip(),
                "estatus": "",
                "equipo": str(cambio.get("equipo_anterior") or ""),
                    "idx": None,
                    "insumos": [],
                    "secuencia": "",
                    "origen": str(cambio.get("origen") or ""),
                    "evidencia": "",
                    "fotos": 0,
                }
        )
    return sorted(eventos, key=lambda e: e["sort"], reverse=True)


def html_calendario_pozo(
    anio: int,
    mes: int,
    dias_bitacora: set[date],
    dias_programados: set[date],
    dia_sel: date | None = None,
) -> str:
    cal = calendar.Calendar(firstweekday=calendar.MONDAY)
    hoy = date.today()
    encabezado = "".join(f"<th>{dia}</th>" for dia in DIAS_SEMANA_CAL)
    filas = []
    for semana in cal.monthdayscalendar(anio, mes):
        celdas = []
        for dia in semana:
            if dia == 0:
                celdas.append('<td><div class="cal-day empty"></div></td>')
                continue
            actual = date(anio, mes, dia)
            clases = ["cal-day"]
            if actual == hoy:
                clases.append("hoy")
            if actual in dias_bitacora:
                clases.append("bitacora")
            if actual in dias_programados:
                clases.append("programado")
            if dia_sel and actual == dia_sel:
                clases.append("sel")
            puntos = ""
            if actual in dias_bitacora:
                puntos += '<span class="dot bit"></span>'
            if actual in dias_programados:
                puntos += '<span class="dot prog"></span>'
            celdas.append(
                f'<td><div class="{" ".join(clases)}"><span class="num">{dia}</span>'
                f'<span class="dots">{puntos}</span></div></td>'
            )
        filas.append(f"<tr>{''.join(celdas)}</tr>")
    return f"""<!DOCTYPE html>
<html><head><style>
body {{ margin: 0; background: transparent; font-family: "Segoe UI", system-ui, sans-serif; }}
.cal-wrap {{
    background: #0f172a; border: 1px solid #334155; border-radius: 14px;
    padding: 1rem 1.1rem 0.9rem; color: #e2e8f0;
}}
.cal-head {{ color: #fff; font-weight: 700; font-size: 1.08rem; margin-bottom: 0.75rem; }}
.cal-table {{ width: 100%; border-collapse: separate; border-spacing: 4px; table-layout: fixed; }}
.cal-table th {{
    text-align: center; font-size: 0.7rem; letter-spacing: 0.06em;
    text-transform: uppercase; color: #93c5fd; font-weight: 700; padding: 0.15rem 0 0.3rem;
}}
.cal-day {{
    min-height: 48px; border-radius: 8px; background: #1e293b; border: 1px solid #334155;
    padding: 5px 6px 6px; color: #e2e8f0;
}}
.cal-day.empty {{ background: transparent; border-color: transparent; }}
.cal-day.hoy {{ border-color: #86efac; box-shadow: 0 0 0 1px rgba(134, 239, 172, 0.35); }}
.cal-day.bitacora {{ background: rgba(20, 83, 45, 0.45); border-color: #2d6a4f; }}
.cal-day.programado {{ background: rgba(30, 58, 95, 0.82); border-color: #3b82f6; }}
.cal-day.bitacora.programado {{
    background: linear-gradient(135deg, rgba(20, 83, 45, 0.5) 0%, rgba(30, 58, 95, 0.8) 100%);
    border-color: #86efac;
}}
.cal-day.sel {{ outline: 2px solid #fbbf24; outline-offset: 1px; }}
.cal-day .num {{ font-size: 0.9rem; font-weight: 700; display: block; }}
.cal-day .dots {{ display: flex; gap: 4px; margin-top: 8px; min-height: 7px; }}
.dot {{ width: 7px; height: 7px; border-radius: 50%; display: inline-block; }}
.dot.bit {{ background: #4ade80; }}
.dot.prog {{ background: #60a5fa; }}
.cal-legend {{
    display: flex; gap: 0.95rem; margin-top: 0.7rem; color: #94a3b8;
    font-size: 0.78rem; flex-wrap: wrap; align-items: center;
}}
</style></head><body>
<div class="cal-wrap">
  <div class="cal-head">{NOMBRES_MES[mes - 1]} {anio}</div>
  <table class="cal-table"><thead><tr>{encabezado}</tr></thead><tbody>{''.join(filas)}</tbody></table>
  <div class="cal-legend">
    <span><span class="dot bit"></span> Actividad ejecutada</span>
    <span><span class="dot prog"></span> Mantenimiento programado</span>
    <span>Borde verde: hoy</span>
  </div>
</div>
</body></html>"""


def render_calendario_pozo(
    anio: int,
    mes: int,
    dias_bitacora: set[date],
    dias_programados: set[date],
    dia_sel: date | None = None,
) -> None:
    components.html(
        html_calendario_pozo(anio, mes, dias_bitacora, dias_programados, dia_sel),
        height=430,
        scrolling=False,
    )


def bitacora_con_fechas(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=list(df.columns) + ["_fecha"] if df is not None else ["_fecha"])
    copia = df.copy()
    copia["_fecha"] = copia["FECHA"].map(parsear_fecha_valor)
    return copia


def fechas_disponibles_operacion(df: pd.DataFrame | None = None) -> list[date]:
    """Todas las jornadas de la bitácora y de las carpetas de evidencia, sin recorte."""
    base = df if df is not None and "_fecha" in getattr(df, "columns", []) else bitacora_con_fechas(
        df if df is not None else cargar_bitacora()
    )
    dias = {d for d in base.get("_fecha", pd.Series(dtype=object)).tolist() if isinstance(d, date)}
    dias.update(fechas_carpetas_evidencia())
    return sorted(dias, reverse=True)


def fechas_carpetas_evidencia() -> set[date]:
    dias: set[date] = set()
    patron = re.compile(r"(\d{4}-\d{2}-\d{2})|(\d{2}[-/]\d{2}[-/]\d{4})")
    for raiz in (EVIDENCIA_OCTUBRE_DIR, EVIDENCIAS_DIR):
        if not os.path.isdir(raiz):
            continue
        for dirpath, dirnames, _archivos in os.walk(raiz):
            for nombre in [os.path.basename(dirpath), *dirnames]:
                hallado = patron.search(str(nombre or ""))
                if not hallado:
                    continue
                dia = parsear_fecha_valor(hallado.group(0))
                if dia and dia.year >= 2020:
                    dias.add(dia)
    return dias


def bitacora_del_dia(df: pd.DataFrame, dia: date) -> pd.DataFrame:
    if df is None or df.empty or dia is None:
        return pd.DataFrame(columns=COLUMNAS_EXCEL)
    base = df if "_fecha" in df.columns else bitacora_con_fechas(df)
    return base.loc[base["_fecha"] == dia].copy()


def clasificar_estatus_equipo(estatus: str) -> str:
    texto = str(estatus or "").strip().casefold()
    if "fuera" in texto:
        return "paro"
    if "mantenimiento" in texto:
        return "mixto"
    if "operativ" in texto:
        return "ok"
    return "mixto" if texto else "ok"


def clasificar_bloque_cuadrilla(bloque: pd.DataFrame) -> str:
    marcas = {clasificar_estatus_equipo(v) for v in bloque["EVALUACIÓN DEL ESTATUS"].tolist()}
    if marcas == {"ok"} or not marcas:
        return "ok"
    if marcas == {"paro"}:
        return "paro"
    return "mixto"


def etiqueta_semaforo(marca: str) -> str:
    return {
        "ok": "Equipo operando",
        "paro": "Fuera de operación",
        "mixto": "Operando / con paro",
    }.get(marca, "Sin estatus")


def resumen_cuadrillas_dia(df_dia: pd.DataFrame) -> list[dict]:
    if df_dia is None or df_dia.empty:
        return []
    acumulado: dict[str, dict] = {}
    for idx, fila in df_dia.iterrows():
        bruto = str(fila.get("CUADRILLA") or "").strip() or "Sin cuadrilla"
        clave = clave_cuadrilla_oficial(bruto) or "__sin_cuadrilla__"
        grupo = acumulado.get(clave)
        if grupo is None:
            grupo = {
                "cuadrilla": etiqueta_cuadrilla_historial(bruto),
                "marca": "ok",
                "reportes": 0,
                "pozos": set(),
                "sectores": set(),
                "indices": [],
            }
            acumulado[clave] = grupo
        grupo["indices"].append(idx)
        grupo["reportes"] += 1
        pozo = str(fila.get("POZO") or "").strip()
        sector = str(fila.get("SECTOR") or "").strip()
        if pozo and pozo.lower() not in {"nan", "none"}:
            grupo["pozos"].add(pozo)
        if sector and sector.lower() not in {"nan", "none"}:
            grupo["sectores"].add(sector)
    grupos = []
    for grupo in acumulado.values():
        bloque = df_dia.loc[grupo["indices"]]
        grupo["marca"] = clasificar_bloque_cuadrilla(bloque)
        grupo["pozos"] = sorted(grupo["pozos"])
        grupo["sectores"] = sorted(grupo["sectores"])
        grupos.append(grupo)
    return sorted(grupos, key=lambda g: orden_cuadrilla_oficial(g["cuadrilla"]))


def mapa_dias_cuadrillas(df: pd.DataFrame, anio: int, mes: int) -> dict[date, str]:
    if df is None or df.empty:
        return {}
    base = df if "_fecha" in df.columns else bitacora_con_fechas(df)
    marcas: dict[date, str] = {}
    recorte = base.dropna(subset=["_fecha"])
    recorte = recorte[
        recorte["_fecha"].map(lambda d: isinstance(d, date) and d.year == anio and d.month == mes)
    ]
    for dia, bloque in recorte.groupby("_fecha"):
        marcas[dia] = clasificar_bloque_cuadrilla(bloque)
    return marcas


def materiales_de_filas(filas: list[pd.Series]) -> list[dict]:
    acumulado: dict[str, dict] = {}
    for fila in filas:
        for texto in insumos_de_fila(fila):
            if ":" in texto:
                nombre, _, cantidad = texto.partition(":")
                clave = nombre.strip()
                extra = cantidad.strip()
            else:
                clave = texto.strip()
                extra = ""
            item = acumulado.setdefault(clave, {"nombre": clave, "detalle": extra, "lineas": []})
            if extra and extra not in item["lineas"]:
                item["lineas"].append(extra)
            if extra:
                item["detalle"] = " + ".join(item["lineas"])
    amigable = {item.get("columna_excel"): item for item in catalogo_insumos_activos() if item.get("columna_excel")}
    salida = []
    for clave, item in acumulado.items():
        cat = amigable.get(clave)
        if cat:
            item["nombre"] = cat.get("nombre") or clave
            unidad = str(cat.get("unidad") or "").strip()
            if unidad and unidad not in item["detalle"]:
                item["detalle"] = f"{item['detalle']} {unidad}".strip()
        salida.append(item)
    return salida


def fotos_de_filas(filas: list[pd.Series]) -> list[dict]:
    evidencias = []
    for fila in filas:
        pozo = str(fila.get("POZO") or "").strip()
        tipo = str(fila.get("ACTIVIDAD CORTA") or fila.get("TIPO DE ACTIVIDAD") or "").strip()
        sec = str(fila.get("SECUENCIA") or "").strip()
        for i, ruta in enumerate(rutas_fotos_fila(fila), start=1):
            desc = " · ".join(p for p in [f"Evidencia {i}", tipo or None, pozo or None, f"OT {sec}" if sec else None] if p)
            evidencias.append({"ruta": ruta, "descripcion": desc})
    return evidencias


def personal_cuadrillas_default() -> dict:
    semillas = {
        "Electricos": ("", "", "", ""),
        "Instrumentos": (
            "1705 CARLOS HERNANDEZ GALLEGOS",
            "2107 ISAI HERNANDEZ GONZALEZ",
            "",
            "1705-2107",
        ),
        "Mecanico 1": (
            "2239 LUIS EMMANUEL GOMEZ MELO",
            "2603 DAMIAN HERNANDEZ RODRIGUEZ",
            "",
            "2239-2603",
        ),
        "Mecanico 2": (
            "1114 ANGEL PEREZ CASTAÑEDA",
            "2268 ROBERTO JUAREZ CUERVO",
            "",
            "1114-2268",
        ),
        "Mecanico 3": (
            "1424 SAUL LARA FLORES",
            "559 JAVIER VERA GARRIDO",
            "",
            "1424-559",
        ),
        "Mecanico 4": (
            "1498 HECTOR DANIEL ZALETA TELLEZ",
            "2647 JOSE LUIS DE LA CRUZ VICENCIO",
            "",
            "1498-2647",
        ),
        "Montaje": (
            "1360 VICTOR RUIZ R.",
            "1851 PASCUAL MARTINEZ",
            "631 ERNESTO ARELLANO",
            "1360-1851-631",
        ),
        "Soldadura": (
            "207 MIGUEL VALENCIA M.",
            "813 JAVIER RIVERA M.",
            "",
            "207-813",
        ),
        "Retro excavadora": ("Operador Retroexcavadora", "", "", ""),
    }
    posiciones = {}
    for cuadrilla in CUADRILLAS:
        mec, ayu, apoyo, codigo = semillas.get(cuadrilla, ("", "", "", ""))
        posiciones[cuadrilla] = {
            "mecanico": mec,
            "ayudante": ayu,
            "apoyo": apoyo,
            "codigo": codigo,
        }
    return {"posiciones": posiciones, "extras": []}


def _item_personal_limpio(item: dict | None) -> dict:
    dato = item if isinstance(item, dict) else {}
    return {
        "mecanico": str(dato.get("mecanico") or "").strip(),
        "ayudante": str(dato.get("ayudante") or "").strip(),
        "apoyo": str(dato.get("apoyo") or "").strip(),
        "codigo": str(dato.get("codigo") or "").strip(),
    }


def _texto_celda_asistencia(valor) -> str:
    if valor is None or isinstance(valor, bool):
        return "" if valor is None else str(valor)
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, int):
        return str(valor)
    texto = str(valor).strip()
    if texto.lower() in {"nan", "none", "nat", "<na>"}:
        return ""
    return texto


def _numero_celda_asistencia(valor) -> float | None:
    if isinstance(valor, bool) or valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = _texto_celda_asistencia(valor).replace("%", "").replace(",", ".")
    if not texto:
        return None
    try:
        return float(texto)
    except ValueError:
        return None


def _mtime_archivo(ruta: str) -> float:
    try:
        return os.path.getmtime(ruta)
    except OSError:
        return 0.0


def clave_asistencia_cuadrilla(nombre: str) -> str:
    clave = re.sub(r"\s+", " ", _sin_acentos(nombre).casefold()).strip()
    alias = {
        "electrico": "electricos",
        "electricos": "electricos",
        "retroexcavadora": "retro excavadora",
        "retro excavadora": "retro excavadora",
    }
    return alias.get(clave, clave)


def _cuadrilla_canonica_asistencia(nombre: str) -> str:
    objetivo = clave_asistencia_cuadrilla(nombre)
    if objetivo in {"de apoyo", "apoyo"}:
        return ""
    for oficial in CUADRILLAS:
        if clave_asistencia_cuadrilla(oficial) == objetivo:
            return oficial
    return ""


def _es_franco_asistencia(codigo: str) -> bool:
    clave = re.sub(r"[\s.\-/]+", "", _sin_acentos(codigo).casefold())
    return clave in {"fo", "franco", "descanso", "diafranco"}


def _codigo_asistencia_normalizado(codigo: str) -> str:
    return re.sub(r"\s+", "", str(codigo or "")).casefold()


def _hoja_asistencia(ruta: str, preferida: str, contiene: str) -> str | None:
    xl = pd.ExcelFile(ruta)
    if preferida in xl.sheet_names:
        return preferida
    objetivo = _sin_acentos(contiene).casefold()
    for nombre in xl.sheet_names:
        if objetivo in _sin_acentos(nombre).casefold():
            return nombre
    return xl.sheet_names[0] if xl.sheet_names else None


def _indice_columna_asistencia(encabezados: list[str], opciones: set[str]) -> int | None:
    for idx, texto in enumerate(encabezados):
        if texto in opciones:
            return idx
    return None


@st.cache_data(show_spinner=False)
def _leer_lista_asistencia_cache(ruta: str, mtime: float) -> dict:
    _ = mtime
    vacio = {"posiciones": {}, "extras": [], "disponible": False}
    if not ruta or not os.path.exists(ruta):
        return vacio
    try:
        hoja = _hoja_asistencia(ruta, "LISTA_PERSONAL", "lista")
        if not hoja:
            return vacio
        df = pd.read_excel(ruta, sheet_name=hoja, header=None, dtype=object)
    except Exception:
        return vacio
    fila_header = None
    encabezados: list[str] = []
    for i in range(len(df)):
        textos = [_sin_acentos(_texto_celda_asistencia(v)).casefold() for v in df.iloc[i].tolist()]
        if "personal" in textos and any(t.startswith("nombre de personal") for t in textos):
            fila_header = i
            encabezados = textos
            break
    if fila_header is None:
        return vacio
    i_personal = _indice_columna_asistencia(encabezados, {"personal"})
    i_p1 = _indice_columna_asistencia(encabezados, {"nombre de personal", "nombre de personal1", "personal 1"})
    i_p2 = _indice_columna_asistencia(encabezados, {"nombre de personal2", "nombre de personal 2", "personal 2"})
    i_p3 = _indice_columna_asistencia(encabezados, {"nombre de personal3", "nombre de personal 3", "personal 3"})
    i_num = next((i for i, h in enumerate(encabezados) if "numero de personal" in h), None)
    if i_personal is None:
        return vacio
    posiciones: dict[str, dict] = {}
    extras: list[str] = []
    vistos: set[str] = set()
    for _, fila in df.iloc[fila_header + 1 :].iterrows():
        etiqueta = _texto_celda_asistencia(fila.iloc[i_personal] if i_personal < len(fila) else "")
        if not etiqueta:
            continue
        p1 = _texto_celda_asistencia(fila.iloc[i_p1]) if i_p1 is not None and i_p1 < len(fila) else ""
        p2 = _texto_celda_asistencia(fila.iloc[i_p2]) if i_p2 is not None and i_p2 < len(fila) else ""
        p3 = _texto_celda_asistencia(fila.iloc[i_p3]) if i_p3 is not None and i_p3 < len(fila) else ""
        codigo = _texto_celda_asistencia(fila.iloc[i_num]) if i_num is not None and i_num < len(fila) else ""
        if clave_asistencia_cuadrilla(etiqueta) in {"de apoyo", "apoyo"}:
            for nombre in (p1, p2, p3):
                clave_n = nombre.casefold()
                if nombre and clave_n not in vistos:
                    extras.append(nombre)
                    vistos.add(clave_n)
            continue
        canon = _cuadrilla_canonica_asistencia(etiqueta)
        if not canon:
            continue
        item = {
            "mecanico": p1,
            "ayudante": p2,
            "apoyo": p3,
            "codigo": codigo,
        }
        if not item["codigo"]:
            item["codigo"] = codigo_dotacion_cuadrilla(item)
        posiciones[canon] = item
    return {"posiciones": posiciones, "extras": extras, "disponible": bool(posiciones or extras)}


def leer_lista_asistencia() -> dict:
    return _leer_lista_asistencia_cache(LISTA_ASISTENCIA_FILE, _mtime_archivo(LISTA_ASISTENCIA_FILE))


def _anio_mes_calendario_asistencia(df: pd.DataFrame, fila_header: int) -> tuple[int, int, str]:
    mapa_mes = {_sin_acentos(nombre).casefold(): i for i, nombre in enumerate(NOMBRES_MES, start=1)}
    anio = None
    mes = None
    titulo = ""
    for i in range(max(fila_header, 0)):
        for valor in df.iloc[i].tolist():
            texto = _texto_celda_asistencia(valor)
            if not texto:
                continue
            plano = _sin_acentos(texto).casefold()
            if "calendario" in plano and len(texto) > len(titulo):
                titulo = texto
            encontrado = re.search(r"(20\d{2})", texto)
            if encontrado and anio is None:
                anio = int(encontrado.group(1))
            for nombre, numero in mapa_mes.items():
                if nombre in plano:
                    mes = numero
                    break
    hoy = date.today()
    if mes is None:
        mes = hoy.month
    if anio is None:
        anio = hoy.year - 1 if mes > hoy.month else hoy.year
    if not titulo:
        titulo = f"Calendario de asistencia {NOMBRES_MES[mes - 1]}"
    return anio, mes, titulo


@st.cache_data(show_spinner=False)
def _leer_calendario_asistencia_cache(ruta: str, mtime: float) -> dict:
    _ = mtime
    vacio = {"cuadrillas": {}, "disponible": False, "anio": 0, "mes": 0, "titulo": "", "periodo": ""}
    if not ruta or not os.path.exists(ruta):
        return vacio
    try:
        hoja = _hoja_asistencia(ruta, "CALENDARIO_ASISTECNIA", "calendario")
        if not hoja:
            return vacio
        df = pd.read_excel(ruta, sheet_name=hoja, header=None, dtype=object)
    except Exception:
        return vacio
    fila_header = None
    encabezados: list[str] = []
    for i in range(len(df)):
        textos = [_sin_acentos(_texto_celda_asistencia(v)).casefold() for v in df.iloc[i].tolist()]
        if "cuadrillas" in textos and "realizadas" in textos:
            fila_header = i
            encabezados = textos
            break
    if fila_header is None or fila_header < 1:
        return vacio
    anio, mes, titulo = _anio_mes_calendario_asistencia(df, fila_header)
    i_cuadrilla = _indice_columna_asistencia(encabezados, {"cuadrillas", "cuadrilla"})
    i_real = _indice_columna_asistencia(encabezados, {"realizadas"})
    i_total = _indice_columna_asistencia(encabezados, {"total"})
    i_avance = _indice_columna_asistencia(encabezados, {"avance"})
    if i_cuadrilla is None:
        return vacio
    reservadas = {i_cuadrilla, i_real, i_total, i_avance}
    fila_dias = df.iloc[fila_header - 1]
    columnas_dia = []
    for idx, encabezado in enumerate(encabezados):
        if idx in reservadas or encabezado in {"semana 1", "semana 2", "semana 3", "semana 4", "semana 5", "semana 6"}:
            continue
        dia_num = None
        texto_dia = _texto_celda_asistencia(fila_dias.iloc[idx] if idx < len(fila_dias) else "")
        if re.fullmatch(r"\d{1,2}", texto_dia):
            numero = int(texto_dia)
            if 1 <= numero <= 31:
                dia_num = numero
        if dia_num is None:
            continue
        try:
            fecha = date(anio, mes, dia_num)
        except ValueError:
            continue
        columnas_dia.append((idx, fecha, _texto_celda_asistencia(df.iloc[fila_header, idx])))
    cuadrillas: dict[str, dict] = {}
    for _, fila in df.iloc[fila_header + 1 :].iterrows():
        etiqueta = _texto_celda_asistencia(fila.iloc[i_cuadrilla] if i_cuadrilla < len(fila) else "")
        canon = _cuadrilla_canonica_asistencia(etiqueta)
        if not canon:
            continue
        dias = []
        for idx, fecha, dia_semana in columnas_dia:
            codigo = _texto_celda_asistencia(fila.iloc[idx] if idx < len(fila) else "")
            dias.append(
                {
                    "fecha": fecha.isoformat(),
                    "codigo": codigo,
                    "laboral": bool(codigo) and not _es_franco_asistencia(codigo),
                    "dia_semana": dia_semana,
                }
            )
        total = _numero_celda_asistencia(fila.iloc[i_total]) if i_total is not None and i_total < len(fila) else None
        total_entero = int(round(total)) if total is not None else None
        realizadas_calc = sum(1 for dia in dias if dia["laboral"])
        avance_calc = (realizadas_calc / total_entero) if total_entero else None
        cuadrillas[clave_asistencia_cuadrilla(canon)] = {
            "nombre": canon,
            "realizadas": realizadas_calc,
            "total": total_entero,
            "avance": avance_calc,
            "dias": dias,
        }
    return {
        "cuadrillas": cuadrillas,
        "disponible": bool(cuadrillas),
        "anio": anio,
        "mes": mes,
        "titulo": titulo,
        "periodo": f"{NOMBRES_MES[mes - 1]} {anio}",
    }


def leer_calendario_asistencia() -> dict:
    return _leer_calendario_asistencia_cache(
        CALENDARIO_ASISTENCIA_FILE,
        _mtime_archivo(CALENDARIO_ASISTENCIA_FILE),
    )


@st.cache_data(show_spinner=False)
def _jornadas_bitacora_cache(ruta: str, mtime: float) -> dict:
    _ = ruta, mtime
    return jornadas_bitacora_por_cuadrilla()


def jornadas_bitacora_vigentes() -> dict[tuple[str, date], str]:
    return _jornadas_bitacora_cache(EXCEL_FILE, _mtime_archivo(EXCEL_FILE))


def _dias_con_ficha_de_bitacora(nombre: str, dias: list[dict]) -> list[dict]:
    """La ficha capturada en la bitácora manda sobre el código fijo del calendario."""
    try:
        jornadas = jornadas_bitacora_vigentes()
    except Exception:
        return dias
    clave = clave_cuadrilla_oficial(nombre)
    salida = []
    for dia in dias:
        copia = dict(dia)
        try:
            fecha = date.fromisoformat(str(copia.get("fecha") or ""))
        except ValueError:
            fecha = None
        if fecha is not None and (clave, fecha) in jornadas:
            copia["codigo"] = jornadas[(clave, fecha)] or "En campo"
            copia["laboral"] = True
            copia["desde_bitacora"] = True
        salida.append(copia)
    return salida


def resumen_jornadas_cuadrilla(cuadrilla: str) -> dict | None:
    cal = leer_calendario_asistencia()
    if not cal.get("disponible"):
        return None
    item = (cal.get("cuadrillas") or {}).get(clave_asistencia_cuadrilla(cuadrilla))
    if not item:
        return None
    salida = dict(item)
    salida["dias"] = _dias_con_ficha_de_bitacora(item.get("nombre") or cuadrilla, item.get("dias") or [])
    salida["realizadas"] = sum(1 for dia in salida["dias"] if dia.get("laboral"))
    if salida.get("total"):
        salida["avance"] = salida["realizadas"] / salida["total"]
    salida["anio"] = cal.get("anio")
    salida["mes"] = cal.get("mes")
    salida["periodo"] = cal.get("periodo") or ""
    salida["titulo"] = cal.get("titulo") or ""
    return salida


def detalle_jornada_asistencia(cuadrilla: str, dia: date | None) -> dict | None:
    item = resumen_jornadas_cuadrilla(cuadrilla)
    if not item or not isinstance(dia, date):
        return None
    if dia.year != item.get("anio") or dia.month != item.get("mes"):
        return None
    dia_info = next((d for d in item.get("dias") or [] if d.get("fecha") == dia.isoformat()), None)
    return {
        "periodo": item.get("periodo") or "",
        "realizadas": item.get("realizadas"),
        "total": item.get("total"),
        "avance": item.get("avance"),
        "codigo": str((dia_info or {}).get("codigo") or ""),
        "laboral": (dia_info or {}).get("laboral"),
        "dia_semana": str((dia_info or {}).get("dia_semana") or ""),
        "en_calendario": dia_info is not None,
    }


def _combinar_personal_con_lista(base: dict) -> dict:
    lista = leer_lista_asistencia()
    if not lista.get("disponible"):
        return base
    posiciones = base.setdefault("posiciones", {})
    for cuadrilla, item in (lista.get("posiciones") or {}).items():
        limpio = _item_personal_limpio(item)
        if any(limpio.get(k) for k in ("mecanico", "ayudante", "apoyo", "codigo")):
            posiciones[cuadrilla] = limpio
    apoyo = [str(n).strip() for n in (lista.get("extras") or []) if str(n).strip()]
    apoyo_cf = {n.casefold() for n in apoyo}
    manuales = []
    vistos = set(apoyo_cf)
    for nombre in base.get("extras") or []:
        texto = str(nombre or "").strip()
        clave = texto.casefold()
        if texto and clave not in vistos:
            manuales.append(texto)
            vistos.add(clave)
    base["extras"] = apoyo + manuales
    base["fuente_lista"] = LISTA_ASISTENCIA_FILE
    return base


def cargar_personal_cuadrillas() -> dict:
    datos = leer_json(PERSONAL_FILE, personal_cuadrillas_default())
    if not isinstance(datos, dict):
        return _combinar_personal_con_lista(personal_cuadrillas_default())
    base = personal_cuadrillas_default()
    crudo = datos.get("posiciones") if isinstance(datos.get("posiciones"), dict) else {}
    por_clave = {clave_hist_cuadrilla(n): _item_personal_limpio(v) for n, v in crudo.items()}
    for cuadrilla in CUADRILLAS:
        item = por_clave.get(clave_hist_cuadrilla(cuadrilla)) or _item_personal_limpio(crudo.get(cuadrilla))
        if any(item.get(k) for k in ("mecanico", "ayudante", "apoyo", "codigo")):
            base["posiciones"][cuadrilla] = item
        elif clave_hist_cuadrilla(cuadrilla) == "electricos":
            base["posiciones"][cuadrilla] = _item_personal_limpio({})
    extras = datos.get("extras") or []
    if isinstance(extras, list):
        limpios = []
        vistos = set()
        for nombre in extras:
            texto = str(nombre or "").strip()
            if texto and texto.casefold() not in vistos:
                limpios.append(texto)
                vistos.add(texto.casefold())
        base["extras"] = limpios
    return _combinar_personal_con_lista(base)


def guardar_personal_cuadrillas(datos: dict) -> None:
    escribir_json(PERSONAL_FILE, datos)


def posicion_personal_cuadrilla(cuadrilla: str) -> dict:
    datos = cargar_personal_cuadrillas()
    posiciones = datos.get("posiciones") or {}
    objetivo = clave_cuadrilla_oficial(cuadrilla)
    if cuadrilla in posiciones and isinstance(posiciones[cuadrilla], dict):
        return _item_personal_limpio(posiciones[cuadrilla])
    for nombre, item in posiciones.items():
        if clave_cuadrilla_oficial(nombre) == objetivo:
            return _item_personal_limpio(item)
    return _item_personal_limpio({})


def codigo_dotacion_cuadrilla(item: dict) -> str:
    if str(item.get("codigo") or "").strip():
        return str(item.get("codigo") or "").strip()
    fichas = []
    for clave in ("mecanico", "ayudante", "apoyo"):
        texto = str(item.get(clave) or "").strip()
        match = re.match(r"^(\d+)\b", texto)
        if match:
            fichas.append(match.group(1))
    return "-".join(fichas)


def titular_cuadrilla(cuadrilla: str) -> tuple[str, str]:
    item = posicion_personal_cuadrilla(cuadrilla)
    return item["mecanico"], item["ayudante"]


def nombres_de_codigo_personal(codigo: str) -> list[str]:
    if not codigo or _es_franco_asistencia(codigo):
        return []
    indice: dict[str, str] = {}
    datos = cargar_personal_cuadrillas()
    for item in (datos.get("posiciones") or {}).values():
        if not isinstance(item, dict):
            continue
        for clave in ("mecanico", "ayudante", "apoyo"):
            texto = str(item.get(clave) or "").strip()
            match = re.match(r"^(\d+)\b", texto)
            if match and texto:
                indice.setdefault(match.group(1), texto)
    for extra in datos.get("extras") or []:
        texto = str(extra or "").strip()
        match = re.match(r"^(\d+)\b", texto)
        if match and texto:
            indice.setdefault(match.group(1), texto)
    nombres = []
    for ficha in re.split(r"[-/]", str(codigo)):
        ficha = ficha.strip()
        if not ficha.isdigit():
            continue
        nombres.append(indice.get(ficha) or ficha)
    return nombres


def render_nota_asistencia(cuadrilla: str, dia: date | None = None, codigo_bitacora: str = "") -> None:
    item = resumen_jornadas_cuadrilla(cuadrilla)
    if not item:
        return
    realizadas = item.get("realizadas")
    total = item.get("total")
    avance = avance_texto(item.get("avance")) or "—"
    st.caption(
        f"Asistencia {item.get('periodo') or ''}: "
        f"{realizadas if realizadas is not None else '—'} realizadas de "
        f"{total if total is not None else '—'} · avance {avance}. "
        f"Fuente: `{CALENDARIO_ASISTENCIA_FILE}`."
    )
    if not isinstance(dia, date):
        return
    detalle = detalle_jornada_asistencia(cuadrilla, dia)
    if not detalle or not detalle.get("en_calendario"):
        return
    codigo_dia = str(detalle.get("codigo") or "").strip()
    codigo_reg = str(codigo_bitacora or "").strip()
    if detalle.get("laboral"):
        nombres = " · ".join(nombres_de_codigo_personal(codigo_dia))
        texto_dia = f"Jornada del {dia.strftime('%d/%m/%Y')}: {codigo_dia}"
        if nombres:
            texto_dia += f" · {nombres}"
    elif _es_franco_asistencia(codigo_dia):
        texto_dia = f"El calendario marca franco (F/O) el {dia.strftime('%d/%m/%Y')}."
    else:
        texto_dia = f"El calendario no tiene código de personal el {dia.strftime('%d/%m/%Y')}."
    fichas_dia = {_codigo_asistencia_normalizado(parte) for parte in codigo_dia.split("/")}
    coincide = _codigo_asistencia_normalizado(codigo_reg) in fichas_dia
    if codigo_reg and codigo_dia and not coincide:
        st.warning(f"{texto_dia} La bitácora registra {codigo_reg}.")
        return
    if codigo_reg and detalle.get("laboral") and coincide:
        st.caption(f"{texto_dia} Coincide con la bitácora.")
        return
    st.caption(texto_dia)


def bolsa_nombres_personal() -> list[str]:
    datos = cargar_personal_cuadrillas()
    vistos: set[str] = set()
    nombres: list[str] = []
    for item in (datos.get("posiciones") or {}).values():
        if not isinstance(item, dict):
            continue
        for clave in ("mecanico", "ayudante", "apoyo"):
            texto = str(item.get(clave) or "").strip()
            clave_n = texto.casefold()
            if texto and clave_n not in vistos:
                nombres.append(texto)
                vistos.add(clave_n)
    for extra in datos.get("extras") or []:
        texto = str(extra or "").strip()
        clave_n = texto.casefold()
        if texto and clave_n not in vistos:
            nombres.append(texto)
            vistos.add(clave_n)
    return nombres


def opciones_personal_turno(default: str) -> list[str]:
    nombres = bolsa_nombres_personal()
    if default and default not in nombres:
        nombres = [default] + nombres
    return nombres + [PERSONAL_OTRO]


def _selector_nombre_turno(etiqueta: str, default: str, key: str) -> str:
    opciones = opciones_personal_turno(default)
    indice = opciones.index(default) if default in opciones else 0
    if key in st.session_state and st.session_state[key] not in opciones:
        st.session_state.pop(key, None)
    if key in st.session_state:
        indice = None
    elegido = st.selectbox(
        etiqueta,
        opciones,
        index=indice,
        key=key,
        persist_state="session",
    )
    if elegido == PERSONAL_OTRO:
        return st.text_input(
            f"Nombre real · {etiqueta}",
            placeholder="Escribe el nombre de quien cubrió",
            key=f"{key}_otro",
            persist_state="session",
        ).strip()
    return str(elegido or "").strip()


def render_bloque_cobertura(form_id: int, cuadrilla: str) -> dict:
    item = posicion_personal_cuadrilla(cuadrilla)
    mec_tit = item["mecanico"]
    ayu_tit = item["ayudante"]
    apoyo_tit = item["apoyo"]
    codigo = codigo_dotacion_cuadrilla(item)
    integrantes = [p for p in [mec_tit, ayu_tit, apoyo_tit] if p]
    titular_txt = " · ".join(integrantes) or "Sin dotación titular capturada"
    lineas = "".join(f"<p>{html.escape(p)}</p>" for p in integrantes) or "<p>Sin dotación titular capturada</p>"
    extra_codigo = (
        f'<p><b>Número de personal:</b> {html.escape(codigo)}</p>' if codigo else ""
    )
    pendiente = ""
    if clave_hist_cuadrilla(cuadrilla) == "electricos" and not integrantes:
        pendiente = "<p>Dotación de eléctricos pendiente de captura.</p>"
    st.markdown(
        f"""
        <div class="cobertura-card">
            <div class="kicker">Posición oficial · {html.escape(etiqueta_cuadrilla_historial(cuadrilla))}</div>
            <h4>Personal titular</h4>
            {lineas}
            {extra_codigo}
            {pendiente}
            <p>OPS y facturación conservan esta cuadrilla. El personal en turno queda para auditoría interna.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    modo = st.radio(
        "Dotación del servicio",
        ["Personal titular", "Cobertura / sustitución"],
        horizontal=True,
        key=f"cobertura_modo_{form_id}_{cuadrilla}",
        persist_state="session",
        help="Si alguien cubre por falta o reasignación OPS, elige cobertura y cambia solo el nombre real.",
    )
    apoyo_real = apoyo_tit
    if modo == "Personal titular":
        mec_real, ayu_real, apoyo_real = mec_tit, ayu_tit, apoyo_tit
        if integrantes:
            st.caption(f"En turno (real): **{titular_txt}**" + (f" · **{codigo}**" if codigo else ""))
        elif clave_hist_cuadrilla(cuadrilla) == "electricos":
            st.caption("La categoría de eléctricos queda pendiente. Aún no hay personal titular asignado.")
        else:
            st.caption("Captura la dotación titular en Auditoría → Catálogos → Personal y coberturas.")
    else:
        cols = st.columns(3 if apoyo_tit else 2)
        with cols[0]:
            mec_real = _selector_nombre_turno(
                "Personal 1 en turno",
                mec_tit,
                f"mec_real_{form_id}_{cuadrilla}",
            )
        with cols[1]:
            ayu_real = _selector_nombre_turno(
                "Personal 2 en turno",
                ayu_tit,
                f"ayu_real_{form_id}_{cuadrilla}",
            )
        if apoyo_tit:
            with cols[2]:
                apoyo_real = _selector_nombre_turno(
                    "Personal 3 en turno",
                    apoyo_tit,
                    f"apoyo_real_{form_id}_{cuadrilla}",
                )
    hay_cobertura = bool(
        (mec_real and mec_real != mec_tit)
        or (ayu_real and ayu_real != ayu_tit)
        or (apoyo_real and apoyo_real != apoyo_tit)
    )
    return {
        "mecanico_titular": mec_tit,
        "ayudante_titular": " · ".join(p for p in [ayu_tit, apoyo_tit] if p),
        "mecanico_real": mec_real,
        "ayudante_real": " · ".join(p for p in [ayu_real, apoyo_real] if p),
        "cobertura": "Sí" if hay_cobertura else "No",
        "codigo": codigo,
    }


def primer_registro_jornada(fecha: date, cuadrilla: str):
    """Primer registro (consecutivo más bajo) de la cuadrilla en la fecha, o None."""
    objetivo = clave_cuadrilla_oficial(cuadrilla) if cuadrilla else ""
    if not objetivo:
        return None
    df = cargar_bitacora()
    if df is None or df.empty or "FECHA" not in df.columns or "CUADRILLA" not in df.columns:
        return None
    del_dia = bitacora_del_dia(bitacora_con_fechas(df), fecha)
    if del_dia.empty:
        return None
    del_dia = del_dia.loc[del_dia["CUADRILLA"].astype(str).map(clave_cuadrilla_oficial) == objetivo]
    if del_dia.empty:
        return None
    orden = del_dia["SECUENCIA"].map(lambda v: numero_secuencia_actividad(v) or 10**6)
    return del_dia.loc[orden.sort_values(kind="stable").index[0]]


def _id_jornada_personal(fecha: date, cuadrilla: str) -> str:
    clave = clave_cuadrilla_oficial(cuadrilla) or str(cuadrilla or "").strip().casefold()
    return f"{fecha.isoformat()}|{clave}"


def guardar_memoria_jornada(fecha: date, cuadrilla: str, datos: dict) -> None:
    """Conserva la dotación elegida para las siguientes actividades del mismo día."""
    if not str(cuadrilla or "").strip():
        return
    memoria = dict(st.session_state.get("_memoria_jornada_personal") or {})
    clave = _id_jornada_personal(fecha, cuadrilla)
    actual = dict(memoria.get(clave) or {})
    actual.update({k: v for k, v in datos.items() if v is not None})
    actual["cuadrilla"] = cuadrilla
    memoria[clave] = actual
    memoria["__ultima__"] = {"fecha": fecha.isoformat(), "cuadrilla": cuadrilla}
    st.session_state["_memoria_jornada_personal"] = memoria


def memoria_jornada(fecha: date, cuadrilla: str) -> dict:
    memoria = st.session_state.get("_memoria_jornada_personal") or {}
    return dict(memoria.get(_id_jornada_personal(fecha, cuadrilla)) or {})


def cuadrilla_recordada(fecha: date, opciones: list[str]) -> str:
    ultima = (st.session_state.get("_memoria_jornada_personal") or {}).get("__ultima__") or {}
    if ultima.get("fecha") != fecha.isoformat():
        return ""
    nombre = str(ultima.get("cuadrilla") or "")
    return nombre if nombre in opciones else ""


def sincronizar_memoria_dotacion(fecha: date, cuadrilla: str, form_id: int, secuencia: int | None = None) -> None:
    modo = st.session_state.get(f"cobertura_modo_{form_id}_{cuadrilla}")
    if modo not in {"Personal titular", "Cobertura / sustitución"}:
        return
    datos: dict = {"modo": modo}
    if secuencia:
        datos["secuencia"] = int(secuencia)
    if modo == "Cobertura / sustitución":
        for base in ("mec_real", "ayu_real", "apoyo_real"):
            clave = f"{base}_{form_id}_{cuadrilla}"
            if clave in st.session_state:
                datos[base] = st.session_state.get(clave) or ""
            otro = f"{clave}_otro"
            if otro in st.session_state:
                datos[f"{base}_otro"] = st.session_state.get(otro) or ""
    guardar_memoria_jornada(fecha, cuadrilla, datos)


def _volcar_dotacion_memoria(valores: dict, form_id: int, cuadrilla: str, mem: dict) -> None:
    modo = mem.get("modo")
    if modo not in {"Personal titular", "Cobertura / sustitución"}:
        return
    valores[f"cobertura_modo_{form_id}_{cuadrilla}"] = modo
    if modo != "Cobertura / sustitución":
        return
    item = posicion_personal_cuadrilla(cuadrilla)
    bases = (("mec_real", item["mecanico"]), ("ayu_real", item["ayudante"]), ("apoyo_real", item["apoyo"]))
    for base, titular in bases:
        if base == "apoyo_real" and not titular:
            continue
        elegido = str(mem.get(base) or "").strip()
        otro = str(mem.get(f"{base}_otro") or "").strip()
        if not elegido and not otro:
            continue
        clave = f"{base}_{form_id}_{cuadrilla}"
        opciones = opciones_personal_turno(titular)
        if elegido in opciones:
            valores[clave] = elegido
            if elegido == PERSONAL_OTRO and otro:
                valores[f"{clave}_otro"] = otro
        else:
            valores[clave] = PERSONAL_OTRO
            valores[f"{clave}_otro"] = otro or elegido


def _nombres_fila_personal(fila) -> list[str]:
    mecanico = str(fila.get("MECANICO EN TURNO") or "").strip()
    ayudantes = [p.strip() for p in str(fila.get("AYUDANTE EN TURNO") or "").split("·") if p.strip()]
    if mecanico or ayudantes:
        return [mecanico, *ayudantes]
    return [p.strip() for p in str(fila.get("PERSONAL") or "").split("·") if p.strip()]


def _dotacion_distinta(nombres: list[str], item: dict) -> bool:
    bases = [str(item.get("mecanico") or ""), str(item.get("ayudante") or ""), str(item.get("apoyo") or "")]
    return any(nombre and nombre != base for nombre, base in zip(nombres, bases))


def _sembrar_estado_jornada(valores: dict, *, solo_faltantes: bool) -> None:
    for clave, valor in valores.items():
        if str(clave).startswith("_"):
            continue
        if solo_faltantes and clave in st.session_state:
            continue
        st.session_state[clave] = valor


def aplicar_precarga_jornada(marca: str, firma: tuple, calcular) -> dict:
    """Siembra cuadrilla y personal una vez por jornada y los recupera si la sección se ocultó."""
    previa = st.session_state.get(marca) or {}
    if previa.get("firma") == firma:
        guardados = previa.get("valores") or {}
        if any(not str(clave).startswith("_") and clave not in st.session_state for clave in guardados):
            valores = calcular() or {}
            _sembrar_estado_jornada(valores, solo_faltantes=True)
            st.session_state[marca] = {"firma": firma, "valores": valores}
            return valores
        return guardados
    for clave, valor in (previa.get("valores") or {}).items():
        if not str(clave).startswith("_") and st.session_state.get(clave) == valor:
            st.session_state.pop(clave, None)
    valores = calcular() or {}
    _sembrar_estado_jornada(valores, solo_faltantes=False)
    st.session_state[marca] = {"firma": firma, "valores": valores}
    return valores


def valores_precarga_campo(form_id: int, fecha: date, cuadrilla: str) -> dict:
    valores: dict = {}
    mem = memoria_jornada(fecha, cuadrilla)
    if mem.get("modo"):
        _volcar_dotacion_memoria(valores, form_id, cuadrilla, mem)
    fila = primer_registro_jornada(fecha, cuadrilla)
    if fila is not None and f"cobertura_modo_{form_id}_{cuadrilla}" not in valores:
        cobertura_txt = str(fila.get("COBERTURA") or "").strip().casefold()
        item = posicion_personal_cuadrilla(cuadrilla)
        nombres = _nombres_fila_personal(fila)
        if cobertura_txt in {"sí", "si"} or _dotacion_distinta(nombres, item):
            _volcar_dotacion_memoria(
                valores,
                form_id,
                cuadrilla,
                {
                    "modo": "Cobertura / sustitución",
                    "mec_real": nombres[0] if nombres else "",
                    "ayu_real": nombres[1] if len(nombres) > 1 else "",
                    "apoyo_real": nombres[2] if len(nombres) > 2 else "",
                },
            )
        elif cobertura_txt or nombres:
            valores[f"cobertura_modo_{form_id}_{cuadrilla}"] = "Personal titular"
    if fila is not None and valores:
        valores["_secuencia"] = numero_secuencia_actividad(fila.get("SECUENCIA")) or int(mem.get("secuencia") or 1)
    elif valores:
        valores["_secuencia"] = int(mem.get("secuencia") or 1)
    return valores


def valores_precarga_ficha(epoch: int, fecha: date, cuadrilla: str) -> dict:
    mem = memoria_jornada(fecha, cuadrilla)
    ficha = str(mem.get("ficha") or "").strip()
    secuencia = int(mem.get("secuencia") or 0)
    if not ficha:
        fila = primer_registro_jornada(fecha, cuadrilla)
        if fila is None:
            return {}
        ficha = str(fila.get("PERSONAL") or fila.get("MECANICO EN TURNO") or "").strip()
        secuencia = numero_secuencia_actividad(fila.get("SECUENCIA")) or 1
    if not ficha:
        return {}
    return {
        f"bit_reg_personal_{epoch}": ficha,
        "_secuencia": secuencia or 1,
    }


def aviso_precarga_jornada(valores: dict) -> None:
    if valores:
        st.info(
            f"Jornada en curso: se precargaron los datos del registro "
            f"{etiqueta_secuencia_consecutiva(int(valores.get('_secuencia') or 1))} de esta cuadrilla. "
            "Cámbialos si hubo relevo de personal."
        )


def personal_de_registro(fila) -> dict:
    return {
        "cuadrilla": str(fila.get("CUADRILLA") or "").strip(),
        "mecanico_titular": str(fila.get("MECANICO TITULAR") or "").strip(),
        "ayudante_titular": str(fila.get("AYUDANTE TITULAR") or "").strip(),
        "mecanico_real": str(fila.get("MECANICO EN TURNO") or "").strip(),
        "ayudante_real": str(fila.get("AYUDANTE EN TURNO") or "").strip(),
        "cobertura": str(fila.get("COBERTURA") or "").strip(),
    }


def texto_par_personal(primero: str, segundo: str) -> str:
    return " · ".join(p for p in [primero, segundo] if p) or "—"


def personal_asignado_cuadrilla(dia: date, cuadrilla: str) -> str:
    asig = asignacion_del_dia(dia)
    if not hay_asignacion_publicada(asig):
        return ""
    directo = str(asig.get("cuadrillas", {}).get(cuadrilla) or "").strip()
    if directo:
        return directo
    objetivo = clave_ubicacion(cuadrilla)
    for nombre, texto in (asig.get("cuadrillas") or {}).items():
        if clave_ubicacion(str(nombre)) == objetivo or objetivo in clave_ubicacion(str(nombre)):
            return str(texto or "").strip()
    return ""


def html_calendario_cuadrillas(
    anio: int,
    mes: int,
    marcas: dict[date, str],
    dia_sel: date | None = None,
    dias_fotos: set[date] | None = None,
) -> str:
    cal = calendar.Calendar(firstweekday=calendar.MONDAY)
    hoy = date.today()
    dias_fotos = dias_fotos or set()
    encabezado = "".join(f"<th>{dia}</th>" for dia in DIAS_SEMANA_CAL)
    filas = []
    for semana in cal.monthdayscalendar(anio, mes):
        celdas = []
        for dia in semana:
            if dia == 0:
                celdas.append('<td><div class="cal-day empty"></div></td>')
                continue
            actual = date(anio, mes, dia)
            marca = marcas.get(actual, "")
            tiene_fotos = actual in dias_fotos
            clases = ["cal-day"]
            if actual == hoy:
                clases.append("hoy")
            if marca:
                clases.append(marca)
            if tiene_fotos:
                clases.append("fotos")
            if dia_sel and actual == dia_sel:
                clases.append("sel")
            punto = f'<span class="dot {marca}"></span>' if marca else ""
            if tiene_fotos:
                punto += '<span class="dot fotos"></span>'
            celdas.append(
                f'<td><div class="{" ".join(clases)}"><span class="num">{dia}</span>'
                f'<span class="dots">{punto}</span></div></td>'
            )
        filas.append(f"<tr>{''.join(celdas)}</tr>")
    return f"""<!DOCTYPE html>
<html><head><style>
body {{ margin: 0; background: transparent; font-family: "Segoe UI", system-ui, sans-serif; }}
.cal-wrap {{
    background: #0f172a; border: 1px solid #334155; border-radius: 14px;
    padding: 1rem 1.1rem 0.9rem; color: #e2e8f0;
}}
.cal-head {{ color: #fff; font-weight: 700; font-size: 1.08rem; margin-bottom: 0.75rem; }}
.cal-table {{ width: 100%; border-collapse: separate; border-spacing: 4px; table-layout: fixed; }}
.cal-table th {{
    text-align: center; font-size: 0.7rem; letter-spacing: 0.06em;
    text-transform: uppercase; color: #93c5fd; font-weight: 700; padding: 0.15rem 0 0.3rem;
}}
.cal-day {{
    min-height: 48px; border-radius: 8px; background: #1e293b; border: 1px solid #334155;
    padding: 5px 6px 6px; color: #e2e8f0;
}}
.cal-day.empty {{ background: transparent; border-color: transparent; }}
.cal-day.hoy {{ border-color: #93c5fd; box-shadow: 0 0 0 1px rgba(147, 197, 253, 0.35); }}
.cal-day.ok {{ background: rgba(20, 83, 45, 0.5); border-color: #22c55e; }}
.cal-day.paro {{ background: rgba(127, 29, 29, 0.45); border-color: #ef4444; }}
.cal-day.mixto {{ background: rgba(120, 53, 15, 0.4); border-color: #f59e0b; }}
.cal-day.fotos {{ border-color: #38bdf8; }}
.cal-day.sel {{ outline: 2px solid #fbbf24; outline-offset: 1px; }}
.cal-day .num {{ font-size: 0.9rem; font-weight: 700; display: block; }}
.cal-day .dots {{ display: flex; gap: 4px; margin-top: 8px; min-height: 7px; }}
.dot {{ width: 7px; height: 7px; border-radius: 50%; display: inline-block; }}
.dot.ok {{ background: #4ade80; }}
.dot.paro {{ background: #f87171; }}
.dot.mixto {{ background: #fbbf24; }}
.dot.fotos {{ background: #38bdf8; }}
.cal-legend {{
    display: flex; gap: 0.95rem; margin-top: 0.7rem; color: #94a3b8;
    font-size: 0.78rem; flex-wrap: wrap; align-items: center;
}}
</style></head><body>
<div class="cal-wrap">
  <div class="cal-head">{NOMBRES_MES[mes - 1]} {anio}</div>
  <table class="cal-table"><thead><tr>{encabezado}</tr></thead><tbody>{''.join(filas)}</tbody></table>
  <div class="cal-legend">
    <span><span class="dot ok"></span> Equipo operando</span>
    <span><span class="dot paro"></span> Fuera de operación</span>
    <span><span class="dot mixto"></span> Día mixto</span>
    <span><span class="dot fotos"></span> Evidencia fotográfica</span>
    <span>Borde azul: hoy</span>
  </div>
</div>
</body></html>"""


def render_calendario_cuadrillas(
    anio: int,
    mes: int,
    marcas: dict[date, str],
    dia_sel: date | None = None,
    dias_fotos: set[date] | None = None,
) -> None:
    components.html(
        html_calendario_cuadrillas(anio, mes, marcas, dia_sel, dias_fotos),
        height=460,
        scrolling=False,
    )


def personal_de_pozo(df_pozo: pd.DataFrame) -> list[dict]:
    if df_pozo.empty:
        return []
    grupos = []
    for cuadrilla, bloque in df_pozo.groupby(df_pozo["CUADRILLA"].astype(str)):
        nombre = str(cuadrilla or "").strip()
        if not nombre:
            continue
        fechas = [str(v) for v in bloque["FECHA"].tolist() if str(v).strip()]
        cortos = []
        for _, fila_b in bloque.iterrows():
            etiqueta = str(fila_b.get("ACTIVIDAD CORTA") or fila_b.get("TIPO DE ACTIVIDAD") or "").strip()
            if etiqueta:
                cortos.append(etiqueta)
        tipos = sorted(set(cortos))
        nombres_turno: list[str] = []
        vistos_n = set()
        for _, fila_b in bloque.iterrows():
            for col in ("MECANICO EN TURNO", "AYUDANTE EN TURNO"):
                persona = str(fila_b.get(col) or "").strip()
                clave_n = persona.casefold()
                if persona and clave_n not in vistos_n:
                    nombres_turno.append(persona)
                    vistos_n.add(clave_n)
        grupos.append(
            {
                "cuadrilla": nombre,
                "intervenciones": int(len(bloque)),
                "ultima": max(fechas) if fechas else "",
                "actividades": ", ".join(tipos) if tipos else "Sin tipo registrado",
                "nombres": ", ".join(nombres_turno) if nombres_turno else "",
            }
        )
    return sorted(grupos, key=lambda g: g["intervenciones"], reverse=True)


def asignacion_vacia() -> dict:
    return {
        "cuadrillas": {cuadrilla: "" for cuadrilla in CUADRILLAS_ASIGNACION},
        "salida": {cuadrilla: "" for cuadrilla in CUADRILLAS_ASIGNACION},
        "ausencias": "",
        "actualizado": "",
    }


def estatus_salida_normalizado(valor: str) -> str:
    texto = _sin_acentos(str(valor or "")).casefold().strip()
    if not texto:
        return ""
    if "en base" in texto or "no sale" in texto or texto in {"f/o", "fo", "franco"}:
        return ESTATUS_EN_BASE
    if "sale a campo" in texto or texto in {"campo", "sale"}:
        return ESTATUS_SALE_CAMPO
    if str(valor or "").strip() in ESTATUS_SALIDA_CUADRILLA:
        return str(valor).strip()
    return ""


def cargar_asignaciones() -> dict:
    datos = leer_json(ASIGNACIONES_FILE, {})
    return datos if isinstance(datos, dict) else {}


def asignacion_del_dia(fecha: date) -> dict:
    registro = cargar_asignaciones().get(fecha.strftime("%Y-%m-%d"), {})
    base = asignacion_vacia()
    if not isinstance(registro, dict):
        return base
    cuadrillas = registro.get("cuadrillas", {})
    if isinstance(cuadrillas, dict):
        for nombre in CUADRILLAS_ASIGNACION:
            base["cuadrillas"][nombre] = str(cuadrillas.get(nombre, "") or "")
    salida = registro.get("salida", {})
    if isinstance(salida, dict):
        for nombre in CUADRILLAS_ASIGNACION:
            base["salida"][nombre] = estatus_salida_normalizado(salida.get(nombre, ""))
    base["ausencias"] = str(registro.get("ausencias", "") or "")
    base["actualizado"] = str(registro.get("actualizado", "") or "")
    return base


def estatus_salida_sugerido(fecha: date, cuadrilla: str, guardado: str = "") -> str:
    elegido = estatus_salida_normalizado(guardado)
    if elegido:
        return elegido
    detalle = detalle_jornada_asistencia(cuadrilla, fecha)
    if detalle and detalle.get("en_calendario"):
        if _es_franco_asistencia(str(detalle.get("codigo") or "")):
            return ESTATUS_EN_BASE
        if detalle.get("laboral"):
            return ESTATUS_SALE_CAMPO
    return ESTATUS_SALE_CAMPO


def salida_visible_dia(fecha: date) -> dict[str, str]:
    asig = asignacion_del_dia(fecha)
    publicada = {
        nombre: estatus
        for nombre, estatus in (asig.get("salida") or {}).items()
        if estatus_salida_normalizado(estatus)
    }
    if publicada:
        return {nombre: estatus_salida_normalizado(estatus) for nombre, estatus in publicada.items()}
    inferida = {}
    for nombre in CUADRILLAS_ASIGNACION:
        detalle = detalle_jornada_asistencia(nombre, fecha)
        if not detalle or not detalle.get("en_calendario"):
            continue
        codigo = str(detalle.get("codigo") or "")
        if _es_franco_asistencia(codigo):
            inferida[nombre] = ESTATUS_EN_BASE
        elif detalle.get("laboral"):
            inferida[nombre] = ESTATUS_SALE_CAMPO
    return inferida


def guardar_asignacion_dia(
    fecha: date,
    cuadrillas: dict[str, str],
    ausencias: str,
    salida: dict[str, str] | None = None,
) -> None:
    datos = cargar_asignaciones()
    clave = fecha.strftime("%Y-%m-%d")
    datos[clave] = {
        "fecha": clave,
        "actualizado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "cuadrillas": {
            nombre: str(cuadrillas.get(nombre, "") or "").strip()
            for nombre in CUADRILLAS_ASIGNACION
        },
        "salida": {
            nombre: estatus_salida_normalizado(str((salida or {}).get(nombre, "") or "")) or ESTATUS_SALE_CAMPO
            for nombre in CUADRILLAS_ASIGNACION
        },
        "ausencias": ausencias.strip(),
    }
    escribir_json(ASIGNACIONES_FILE, datos)


def _ubicar_calendario_asistencia(ws) -> tuple[dict[int, int], dict[str, int]]:
    fila_header = None
    encabezados: list[str] = []
    for fila in range(1, (ws.max_row or 0) + 1):
        textos = [
            _sin_acentos(_texto_celda_asistencia(ws.cell(fila, col).value)).casefold()
            for col in range(1, (ws.max_column or 0) + 1)
        ]
        if "cuadrillas" in textos and "realizadas" in textos:
            fila_header = fila
            encabezados = textos
            break
    if not fila_header or fila_header < 2:
        return {}, {}
    columnas_dia: dict[int, int] = {}
    reservadas = {"cuadrillas", "cuadrilla", "realizadas", "total", "avance"}
    for col, encabezado in enumerate(encabezados, start=1):
        if encabezado in reservadas:
            continue
        dia_txt = _texto_celda_asistencia(ws.cell(fila_header - 1, col).value)
        if re.fullmatch(r"\d{1,2}", dia_txt):
            numero = int(dia_txt)
            if 1 <= numero <= 31:
                columnas_dia[numero] = col
    i_cuadrilla = encabezados.index("cuadrillas") + 1 if "cuadrillas" in encabezados else 2
    filas_cuadrilla: dict[str, int] = {}
    for fila in range(fila_header + 1, (ws.max_row or 0) + 1):
        etiqueta = _texto_celda_asistencia(ws.cell(fila, i_cuadrilla).value)
        canon = _cuadrilla_canonica_asistencia(etiqueta)
        if canon:
            filas_cuadrilla[clave_asistencia_cuadrilla(canon)] = fila
    return columnas_dia, filas_cuadrilla


def aplicar_salida_calendario(fecha: date, salida: dict[str, str]) -> dict:
    cal = leer_calendario_asistencia()
    if not cal.get("disponible"):
        return {
            "ok": False,
            "mensaje": f"No se encontró un calendario legible en {CALENDARIO_ASISTENCIA_FILE}.",
        }
    if fecha.year != cal.get("anio") or fecha.month != cal.get("mes"):
        return {
            "ok": False,
            "mensaje": (
                f"El estatus quedó publicado para campo. {CALENDARIO_ASISTENCIA_FILE} "
                f"corresponde a {cal.get('periodo') or 'otro mes'} y no se modificó."
            ),
        }
    try:
        import openpyxl

        wb = openpyxl.load_workbook(CALENDARIO_ASISTENCIA_FILE)
    except Exception as exc:
        return {"ok": False, "mensaje": f"No se pudo abrir el calendario de asistencia: {exc}"}
    hoja = None
    for nombre in wb.sheetnames:
        if "calendario" in _sin_acentos(nombre).casefold():
            hoja = nombre
            break
    ws = wb[hoja or wb.sheetnames[0]]
    columnas_dia, filas_cuadrilla = _ubicar_calendario_asistencia(ws)
    columna = columnas_dia.get(fecha.day)
    if not columna:
        wb.close()
        return {
            "ok": False,
            "mensaje": (
                f"El {fecha.strftime('%d/%m/%Y')} no tiene columna en el calendario "
                "(fin de semana o día fuera del plan). El estatus sí quedó publicado para campo."
            ),
        }
    try:
        jornadas = jornadas_bitacora_vigentes()
    except Exception:
        jornadas = {}
    actualizadas = []
    omitidas = []
    for nombre in CUADRILLAS_ASIGNACION:
        estatus = estatus_salida_normalizado(str(salida.get(nombre, "") or "")) or ESTATUS_SALE_CAMPO
        fila = filas_cuadrilla.get(clave_asistencia_cuadrilla(nombre))
        if not fila:
            omitidas.append(nombre)
            continue
        celda = ws.cell(fila, columna)
        clave_bit = (clave_cuadrilla_oficial(_cuadrilla_canonica_asistencia(nombre) or nombre), fecha)
        if clave_bit in jornadas:
            celda.value = jornadas[clave_bit] or "En campo"
        elif estatus == ESTATUS_EN_BASE:
            celda.value = "F/O"
        else:
            actual = _texto_celda_asistencia(celda.value)
            if not actual or _es_franco_asistencia(actual):
                canon = _cuadrilla_canonica_asistencia(nombre) or nombre
                codigo = codigo_dotacion_cuadrilla(posicion_personal_cuadrilla(canon))
                celda.value = codigo or None
        actualizadas.append(nombre)
    try:
        wb.save(CALENDARIO_ASISTENCIA_FILE)
    except Exception as exc:
        return {
            "ok": False,
            "mensaje": (
                "Las asignaciones se publicaron, pero el calendario no se pudo guardar. "
                f"Cierra {CALENDARIO_ASISTENCIA_FILE} si está abierto. Detalle: {exc}"
            ),
        }
    finally:
        wb.close()
    extra = ""
    if omitidas:
        extra = " Sin fila en el calendario: " + ", ".join(omitidas) + "."
    return {
        "ok": True,
        "mensaje": (
            f"Calendario actualizado en {CALENDARIO_ASISTENCIA_FILE}: "
            f"{len(actualizadas)} cuadrilla(s) del {fecha.strftime('%d/%m/%Y')}. "
            "Quien no sale queda como F/O y las fórmulas de realizadas y avance se recalculan."
            + extra
        ),
    }


def hay_asignacion_publicada(asig: dict) -> bool:
    if str(asig.get("ausencias", "")).strip():
        return True
    if any(str(texto).strip() for texto in (asig.get("cuadrillas") or {}).values()):
        return True
    return any(estatus_salida_normalizado(texto) for texto in (asig.get("salida") or {}).values())


def texto_html_asignacion(texto: str) -> str:
    limpio = html.escape(texto.strip()) if texto else ""
    return limpio.replace("\n", "<br>") if limpio else ""


def html_salida_dia(fecha: date) -> str:
    salida = salida_visible_dia(fecha)
    if not salida:
        return ""
    chips = []
    for cuadrilla in CUADRILLAS_ASIGNACION:
        estatus = salida.get(cuadrilla)
        if not estatus:
            continue
        clase = "base" if estatus == ESTATUS_EN_BASE else "campo"
        chips.append(
            f'<div class="salida-chip {clase}">'
            f"<span>{html.escape(cuadrilla)}</span>"
            f"<b>{html.escape(estatus)}</b></div>"
        )
    if not chips:
        return ""
    return (
        f'<div class="salida-banner">'
        f'<div class="salida-titulo">Salida del {fecha.strftime("%d/%m/%Y")}</div>'
        f'<div class="salida-chips">{"".join(chips)}</div>'
        f"</div>"
    )


def html_marca_salida(estatus: str) -> str:
    if estatus == ESTATUS_EN_BASE:
        return f'<div class="salida-mini base">{html.escape(estatus)}</div>'
    if estatus == ESTATUS_SALE_CAMPO:
        return f'<div class="salida-mini campo">{html.escape(estatus)}</div>'
    return ""


def clave_mes(anio: int, mes: int) -> str:
    return f"{int(anio):04d}-{int(mes):02d}"


def etiqueta_mes(mes_clave: str) -> str:
    try:
        anio, mes = mes_clave.split("-")
        return f"{NOMBRES_MES[int(mes) - 1]} {anio}"
    except Exception:
        return mes_clave


def cargar_programados_ops() -> list[dict]:
    datos = leer_json(PROGRAMADOS_FILE, [])
    if isinstance(datos, dict):
        plano = []
        for mes, filas in datos.items():
            if isinstance(filas, list):
                for fila in filas:
                    if isinstance(fila, dict):
                        item = dict(fila)
                        item.setdefault("mes", mes)
                        plano.append(item)
        return plano
    return [item for item in datos if isinstance(item, dict)] if isinstance(datos, list) else []


def programados_del_mes(mes_clave: str) -> list[dict]:
    return [item for item in cargar_programados_ops() if str(item.get("mes", "")) == mes_clave]


def guardar_programados_ops(registros: list[dict]) -> None:
    escribir_json(PROGRAMADOS_FILE, registros)


def guardar_programados_mes(mes_clave: str, filas: list[dict]) -> None:
    restantes = [item for item in cargar_programados_ops() if str(item.get("mes", "")) != mes_clave]
    for fila in filas:
        fila["mes"] = mes_clave
        if not str(fila.get("id", "")).strip():
            fila["id"] = uuid.uuid4().hex[:10]
    guardar_programados_ops(restantes + filas)


def porcentaje_cumplimiento(filas: list[dict]) -> float:
    if not filas:
        return 0.0
    cumplidos = sum(1 for fila in filas if str(fila.get("estatus", "")) == "Cumplido")
    return round(100.0 * cumplidos / len(filas), 1)


def _texto_programado(valor) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.casefold() in {"none", "nan", "<na>", "nat"}:
        return ""
    return texto


def dataframe_programados(filas: list[dict]) -> pd.DataFrame:
    columnas = ["id", "Actividad", "Ubicación", "Equipo", "Estatus", "Observación"]
    if not filas:
        return pd.DataFrame(columns=columnas)
    return pd.DataFrame(
        [
            {
                "id": _texto_programado(item.get("id")),
                "Actividad": _texto_programado(item.get("actividad")),
                "Ubicación": _texto_programado(item.get("ubicacion")),
                "Equipo": _texto_programado(item.get("equipo")),
                "Estatus": (
                    item.get("estatus")
                    if item.get("estatus") in ESTATUS_PROGRAMADO
                    else "No Cumplido"
                ),
                "Observación": _texto_programado(item.get("observacion")),
            }
            for item in filas
        ]
    )


def filas_desde_editor(editado: pd.DataFrame) -> list[dict]:
    registros = []
    for _, fila in editado.iterrows():
        actividad = str(fila.get("Actividad", "")).strip()
        if not actividad or actividad.lower() in {"nan", "none"}:
            continue
        estatus = str(fila.get("Estatus", "")).strip()
        if estatus not in ESTATUS_PROGRAMADO:
            estatus = "No Cumplido"
        registros.append(
            {
                "id": str(fila.get("id", "")).strip(),
                "actividad": actividad,
                "ubicacion": str(fila.get("Ubicación", "")).strip(),
                "equipo": str(fila.get("Equipo", "")).strip(),
                "estatus": estatus,
                "observacion": str(fila.get("Observación", "")).strip(),
            }
        )
    return registros


def mes_de_fecha(valor) -> str:
    fecha = pd.to_datetime(valor, errors="coerce")
    if pd.isna(fecha):
        return ""
    return fecha.strftime("%Y-%m")


def es_registro_no_programado(fila: pd.Series) -> bool:
    tipo = str(fila.get("TIPO DE ACTIVIDAD", "") or "").strip()
    descripcion = str(fila.get("DESCRIPCIÓN DE LA ACTIVIDAD", "") or "")
    estatus = str(fila.get("EVALUACIÓN DEL ESTATUS", "") or "").strip()
    if tipo in TIPOS_MANTENIMIENTO_PROGRAMADO:
        return bool(PATRON_NO_PROGRAMADO.search(descripcion))
    if tipo in TIPOS_MANTENIMIENTO_NO_PROGRAMADO:
        return True
    if estatus == "Fuera de servicio":
        return True
    return bool(PATRON_NO_PROGRAMADO.search(tipo) or PATRON_NO_PROGRAMADO.search(descripcion))


def bitacora_no_programada(df: pd.DataFrame, mes_clave: str) -> pd.DataFrame:
    if df.empty:
        return df
    trabajo = df.copy()
    trabajo["_mes"] = trabajo["FECHA"].map(mes_de_fecha)
    del_mes = trabajo[trabajo["_mes"] == mes_clave]
    if del_mes.empty:
        return del_mes.drop(columns=["_mes"], errors="ignore")
    mask = del_mes.apply(es_registro_no_programado, axis=1)
    return del_mes.loc[mask].drop(columns=["_mes"], errors="ignore")


# ── Funciones de datos ─────────────────────────────────────────────────────────
def sector_requiere_pozo(sector: str) -> bool:
    if sitios_operativos.es_instalacion(sector):
        return False
    return sector not in set(obtener_catalogos()["sectores_sin_pozo"]) and sector != SECTOR_OTRO


def cuadrilla_es_mecanica(cuadrilla: str) -> bool:
    return cuadrilla.startswith("Mecanico")


def _sin_acentos(texto: str) -> str:
    normal = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(c for c in normal if unicodedata.category(c) != "Mn")


def grupo_actividades_cuadrilla(cuadrilla: str) -> str | None:
    clave = re.sub(r"\s+", " ", _sin_acentos(cuadrilla).casefold()).strip()
    if clave.startswith("mecanico"):
        return "MECANICOS"
    if clave.startswith("electric"):
        return "ELECTRICO"
    if clave.startswith("instrument"):
        return "INSTRUMENTOS"
    if clave.startswith("montaje"):
        return "MONTAJE"
    if clave.startswith("soldad"):
        return "SOLDADURA"
    return None


def actividad_esta_activa(item: dict) -> bool:
    if not isinstance(item, dict):
        return False
    valor = item.get("activa", True)
    if isinstance(valor, str):
        return valor.strip().casefold() not in {"", "0", "false", "no", "inactiva", "inactivo"}
    return bool(valor)


def _item_actividad(nombre: str, corto: str = "", activa=True, id_item: str = "") -> dict:
    nombre = str(nombre or "").strip()
    corto = str(corto or "").strip()
    if not nombre or nombre.casefold() == "otra":
        return {}
    fila = {"nombre": nombre, "corto": corto or nombre, "activa": actividad_esta_activa({"activa": activa})}
    ident = str(id_item or "").strip()
    if ident:
        fila["id"] = ident
    return fila


def catalogo_actividades_default() -> dict[str, list[dict]]:
    return {
        grupo: [dict(item) for item in items]
        for grupo, items in CATALOGO_ACTIVIDADES_DEFAULT.items()
    }


def normalizar_catalogo_actividades(datos) -> dict[str, list[dict]]:
    base = catalogo_actividades_default()
    if not isinstance(datos, dict):
        return base
    for grupo in GRUPOS_ACTIVIDAD:
        origen = datos.get(grupo, base[grupo])
        if not isinstance(origen, list):
            continue
        limpios = []
        vistos = set()
        for item in origen:
            if isinstance(item, str):
                fila = _item_actividad(item)
            elif isinstance(item, dict):
                fila = _item_actividad(item.get("nombre"), item.get("corto"), item.get("activa", True), item.get("id"))
            else:
                continue
            clave = fila.get("nombre", "").casefold()
            if not clave or clave in vistos:
                continue
            vistos.add(clave)
            limpios.append(fila)
        if isinstance(datos, dict) and grupo in datos:
            base[grupo] = limpios
        elif limpios:
            base[grupo] = limpios
    return base


def cargar_catalogo_actividades() -> dict[str, list[dict]]:
    if "catalogo_actividades" not in st.session_state:
        datos = leer_json(ACTIVIDADES_FILE, catalogo_actividades_default())
        st.session_state.catalogo_actividades = normalizar_catalogo_actividades(datos)
        if not os.path.exists(ACTIVIDADES_FILE):
            escribir_json(ACTIVIDADES_FILE, st.session_state.catalogo_actividades)
    return st.session_state.catalogo_actividades


def guardar_catalogo_actividades(datos: dict) -> None:
    limpios = normalizar_catalogo_actividades(datos)
    st.session_state.catalogo_actividades = limpios
    escribir_json(ACTIVIDADES_FILE, limpios)


def actividades_de_cuadrilla(cuadrilla: str, *, incluir_inactivas: bool = False) -> list[dict]:
    grupo = grupo_actividades_cuadrilla(cuadrilla)
    if not grupo:
        return []
    items = list(cargar_catalogo_actividades().get(grupo) or [])
    if incluir_inactivas:
        return items
    return [item for item in items if actividad_esta_activa(item)]


def nombres_actividad_cuadrilla(cuadrilla: str) -> list[str]:
    if not grupo_actividades_cuadrilla(cuadrilla):
        return []
    nombres = [item["nombre"] for item in actividades_de_cuadrilla(cuadrilla)]
    if "Otra" not in nombres:
        nombres.append("Otra")
    return nombres


def buscar_actividad_catalogo(cuadrilla: str, nombre: str) -> dict | None:
    objetivo = str(nombre or "").strip().casefold()
    if not objetivo or objetivo == "otra":
        return None
    for item in actividades_de_cuadrilla(cuadrilla, incluir_inactivas=True):
        if item["nombre"].casefold() == objetivo and actividad_esta_activa(item):
            return item
    return None


def render_selector_actividad_cuadrilla(key: str, cuadrilla: str) -> tuple:
    """Selector de actividad oficial. Lee el catálogo activo de la cuadrilla."""
    if not str(cuadrilla or "").strip() or not grupo_actividades_cuadrilla(cuadrilla):
        if str(cuadrilla or "").strip():
            st.caption("Esta cuadrilla no tiene catálogo oficial de actividades.")
        return None, "", ""
    opciones = nombres_actividad_cuadrilla(cuadrilla)
    if st.session_state.get(key) not in opciones:
        st.session_state.pop(key, None)
    grupo = grupo_actividades_cuadrilla(cuadrilla) or ""
    etiqueta = ETIQUETAS_GRUPO_ACTIVIDAD.get(grupo, grupo)
    tipo = st.selectbox(
        "Actividad oficial",
        opciones,
        index=None,
        placeholder="Seleccionar...",
        key=key,
        persist_state="session",
        help=f"Catálogo maestro {etiqueta}. Solo se listan las actividades activas de esta cuadrilla.",
    )
    otro = ""
    corto = ""
    if tipo == "Otra":
        otro = st.text_input(
            "Nombre de la actividad",
            placeholder="Escribe el nombre oficial completo",
            key=f"{key}_otra",
            persist_state="session",
        )
        corto = st.text_input(
            "Nombre corto",
            placeholder="Ej. MTTO TIERRA FISICA SKID ICI",
            key=f"{key}_corta",
            persist_state="session",
        )
    elif tipo:
        oficial = buscar_actividad_catalogo(cuadrilla, tipo)
        if oficial:
            st.caption(f"Nombre corto: **{oficial['corto']}**")
    return tipo, otro, corto


def resolver_actividad_catalogo(
    cuadrilla: str,
    tipo_sel: str | None,
    tipo_otro: str = "",
    corto_otro: str = "",
) -> tuple[str, str]:
    if not tipo_sel or tipo_sel == "Seleccionar...":
        return "", ""
    if tipo_sel == "Otra":
        nombre = tipo_otro.strip()
        corto = corto_otro.strip() or nombre
        return nombre, corto
    item = buscar_actividad_catalogo(cuadrilla, tipo_sel)
    if item:
        return item["nombre"], item["corto"]
    return tipo_sel.strip(), tipo_sel.strip()


def resolver_equipo(equipo_sel: str | None) -> str:
    if not equipo_sel or equipo_sel == "Seleccionar...":
        return ""
    if sitios_operativos.es_vacio_sitio(equipo_sel):
        return ""
    return equipo_sel


def resolver_tipo_actividad(tipo_sel: str | None, tipo_otro: str) -> str:
    if not tipo_sel or tipo_sel == "Seleccionar...":
        return ""
    if tipo_sel == "Otra":
        return tipo_otro.strip()
    return tipo_sel


def resolver_sector(sector_sel: str, sector_otro: str) -> str:
    if sector_sel == SECTOR_OTRO:
        return sector_otro.strip()
    return sector_sel


def formatear_secuencia(numero: int) -> str:
    return f"Act. {int(numero)}"


def etiqueta_secuencia_consecutiva(numero: int) -> str:
    return f"#{int(numero):02d}"


def _siguiente_numero_secuencia(bloque: pd.DataFrame) -> int:
    if bloque is None or bloque.empty:
        return 1
    numeros = []
    if "SECUENCIA" in bloque.columns:
        for valor in bloque["SECUENCIA"].astype(str):
            encontrado = re.search(r"(\d+)", valor)
            if encontrado:
                numeros.append(int(encontrado.group(1)))
    if numeros:
        return max(numeros) + 1
    return len(bloque) + 1


def siguiente_secuencia_del_dia(fecha: date) -> int:
    df = cargar_bitacora()
    if df is None or df.empty or "FECHA" not in df.columns:
        return 1
    return _siguiente_numero_secuencia(bitacora_del_dia(bitacora_con_fechas(df), fecha))


def siguiente_secuencia_turno(fecha: date, cuadrilla: str = "") -> int:
    df = cargar_bitacora()
    if df is None or df.empty or "FECHA" not in df.columns:
        return 1
    del_dia = bitacora_del_dia(bitacora_con_fechas(df), fecha)
    if cuadrilla and not del_dia.empty and "CUADRILLA" in del_dia.columns:
        objetivo = clave_cuadrilla_oficial(cuadrilla)
        if objetivo:
            del_dia = del_dia.loc[
                del_dia["CUADRILLA"].astype(str).map(clave_cuadrilla_oficial) == objetivo
            ]
    return _siguiente_numero_secuencia(del_dia)


def insumos_vacios() -> dict[str, float | int | str]:
    valores: dict[str, float | int | str] = {col: 0 for col in COLUMNAS_INSUMOS}
    valores.update(
        {
            "BANDAS CANTIDAD": 0,
            "BANDAS DESCRIPCIÓN": "",
            "BANDAS PROPORCIONADO POR": "",
            "ALTERNADOR CANTIDAD": 0,
            "ALTERNADOR ESTADO/ORIGEN": "",
            "ALTERNADOR DESCRIPCIÓN": "",
            "OTROS INSUMOS": "",
            "PIEZAS ROTATIVAS": "",
        }
    )
    return valores


def _texto_cantidad_insumo(cantidad) -> str:
    try:
        numero = float(cantidad or 0)
    except (TypeError, ValueError):
        return "0"
    if numero.is_integer():
        return str(int(numero))
    return f"{numero:g}"


def texto_pieza_rotativa(consumo: dict) -> str:
    nombre = str(consumo.get("nombre") or "Componente").strip()
    cantidad = _texto_cantidad_insumo(consumo.get("cantidad"))
    condicion = str(consumo.get("condicion_instalada") or "").strip()
    destino = str(consumo.get("destino_retirada") or "").strip()
    return f"{nombre} x{cantidad}: instalada {condicion}; retirada {destino}"


def volcar_detalle_insumos(valores: dict, consumos: list[dict]) -> None:
    bandas_cant = 0.0
    bandas_desc: list[str] = []
    bandas_origen: list[str] = []
    alt_cant = 0.0
    alt_cond: list[str] = []
    alt_dest: list[str] = []
    bomba_cant = 0.0
    piezas: list[str] = []
    otros_extra: list[str] = []
    for consumo in consumos:
        nombre = str(consumo.get("nombre") or "").strip()
        clave = _sin_acentos(nombre).casefold()
        try:
            cantidad = float(consumo.get("cantidad") or 0)
        except (TypeError, ValueError):
            cantidad = 0.0
        if str(consumo.get("tipo") or "") == TIPO_INSUMO_ROTATIVO:
            piezas.append(texto_pieza_rotativa(consumo))
            if "alternador" in clave:
                alt_cant += cantidad
                if consumo.get("condicion_instalada"):
                    alt_cond.append(str(consumo["condicion_instalada"]))
                if consumo.get("destino_retirada"):
                    alt_dest.append(str(consumo["destino_retirada"]))
            elif "bomba de agua" in clave:
                bomba_cant += cantidad
            else:
                otros_extra.append(f"{nombre} {_texto_cantidad_insumo(cantidad)}")
            continue
        if "banda" in clave:
            bandas_cant += cantidad
            bandas_desc.append(
                nombre if cantidad == 1 else f"{nombre} x{_texto_cantidad_insumo(cantidad)}"
            )
            origen = str(consumo.get("origen") or "").strip()
            if origen and origen not in bandas_origen:
                bandas_origen.append(origen)
    if bandas_cant:
        valores["BANDAS CANTIDAD"] = bandas_cant
        valores["BANDAS DESCRIPCIÓN"] = ", ".join(bandas_desc)
        valores["BANDAS PROPORCIONADO POR"] = ", ".join(bandas_origen)
    if alt_cant:
        valores["ALTERNADOR CANTIDAD"] = alt_cant
        valores["ALTERNADOR ESTADO/ORIGEN"] = " | ".join(dict.fromkeys(alt_cond))
        valores["ALTERNADOR DESCRIPCIÓN"] = " | ".join(dict.fromkeys(alt_dest))
    if bomba_cant:
        valores["BOMBA DE AGUA REPARADA PJP4"] = bomba_cant
    if piezas:
        valores["PIEZAS ROTATIVAS"] = " | ".join(piezas)
    if otros_extra:
        previo = str(valores.get("OTROS INSUMOS") or "").strip()
        extra = ", ".join(otros_extra)
        valores["OTROS INSUMOS"] = f"{previo}, {extra}" if previo else extra


def renderizar_insumo_catalogo(form_id: int, item: dict) -> dict | None:
    slug = item.get("id") or slug_insumo(item["nombre"])
    rotativo = es_insumo_rotativo(item)
    usa = st.checkbox(item["nombre"], key=f"usa_cat_{slug}_{form_id}")
    if not usa:
        return None
    unidad = item.get("unidad") or "Pzas"
    if unidad_es_entera(unidad):
        cantidad = st.number_input(
            f"Cantidad ({unidad})",
            min_value=0,
            step=1,
            value=1,
            key=f"cant_cat_{slug}_{form_id}",
        )
    else:
        paso = 0.1 if unidad.lower() == "kg" else 0.5
        cantidad = st.number_input(
            f"Cantidad ({unidad})",
            min_value=0.0,
            step=float(paso),
            value=float(paso),
            key=f"cant_cat_{slug}_{form_id}",
        )
    origen = st.radio(
        "Origen del insumo",
        ORIGENES_INSUMO,
        index=None,
        horizontal=True,
        key=f"origen_cat_{slug}_{form_id}",
    )
    condicion = ""
    destino = ""
    if rotativo:
        condicion = st.selectbox(
            "Condición de la pieza instalada",
            CONDICION_PIEZA_INSTALADA,
            index=None,
            placeholder="Seleccionar...",
            key=f"cond_cat_{slug}_{form_id}",
        ) or ""
        destino = st.selectbox(
            "Destino de la pieza retirada",
            DESTINO_PIEZA_RETIRADA,
            index=None,
            placeholder="Seleccionar...",
            key=f"dest_cat_{slug}_{form_id}",
        ) or ""
    return {
        "insumo_id": slug,
        "nombre": item["nombre"],
        "cantidad": cantidad,
        "unidad": unidad,
        "origen": origen or "",
        "columna_excel": item.get("columna_excel", ""),
        "tipo": TIPO_INSUMO_ROTATIVO if rotativo else TIPO_INSUMO_CONSUMIBLE,
        "condicion_instalada": condicion,
        "destino_retirada": destino,
    }


def _render_grupo_insumos(form_id: int, titulo: str, ayuda: str, items: list[dict]) -> list[dict]:
    if not items:
        return []
    st.markdown(f"**{titulo}**")
    st.caption(ayuda)
    consumos: list[dict] = []
    columnas = st.columns(3, gap="medium")
    for indice, item in enumerate(items):
        with columnas[indice % 3]:
            consumo = renderizar_insumo_catalogo(form_id, item)
            if consumo:
                consumos.append(consumo)
    return consumos


def renderizar_insumos(form_id: int) -> dict[str, float | int | str]:
    valores = insumos_vacios()
    catalogo = catalogo_insumos_activos()
    consumibles = [item for item in catalogo if not es_insumo_rotativo(item)]
    rotativos = [item for item in catalogo if es_insumo_rotativo(item)]
    consumos = _render_grupo_insumos(
        form_id,
        "Consumibles",
        "Aceite, filtros, bandas (C-115, C-112, C-240, Micro V) y el resto de material que se gasta.",
        consumibles,
    )
    consumos.extend(
        _render_grupo_insumos(
            form_id,
            "Componentes rotativos / reparables",
            "Bomba de agua, alternador y marcha. Al marcar uno se pide la condición de la pieza instalada y el destino de la retirada.",
            rotativos,
        )
    )
    for consumo in consumos:
        col_excel = consumo.get("columna_excel") or ""
        if col_excel in COLUMNAS_INSUMOS:
            valores[col_excel] = consumo["cantidad"]

    usa_otros = st.checkbox("Otros insumos", key=f"usa_otros_{form_id}")
    if usa_otros:
        valores["OTROS INSUMOS"] = st.text_area(
            "Detalle de otros insumos",
            height=120,
            placeholder="Escribe cualquier otro insumo utilizado…",
            key=f"otros_insumos_{form_id}",
        ).strip()
        origen_otros = st.radio(
            "Origen del insumo",
            ORIGENES_INSUMO,
            index=None,
            horizontal=True,
            key=f"otros_origen_{form_id}",
        )
        if valores["OTROS INSUMOS"]:
            consumos.append(
                {
                    "insumo_id": "otros",
                    "nombre": valores["OTROS INSUMOS"],
                    "cantidad": 1,
                    "unidad": "Pzas",
                    "origen": origen_otros or "",
                    "tipo": TIPO_INSUMO_CONSUMIBLE,
                }
            )

    volcar_detalle_insumos(valores, consumos)
    valores["_consumos"] = consumos
    return valores


def fotos_vacias() -> list[str]:
    return [""] * MAX_FOTOS


def limitar_fotos(archivos: list | None) -> list:
    if not archivos:
        return []
    return list(archivos)[:MAX_FOTOS]


def renderizar_galeria(archivos: list) -> None:
    fotos = limitar_fotos(archivos)
    for fila in range(3):
        columnas = st.columns(3, gap="small")
        for col_i, col in enumerate(columnas):
            idx = fila * 3 + col_i
            with col:
                if idx < len(fotos):
                    st.image(fotos[idx], use_container_width=True)
                    st.caption(f"{idx + 1}. {fotos[idx].name}")
                else:
                    st.markdown(
                        f'<div class="foto-vacia">Foto {idx + 1}</div>',
                        unsafe_allow_html=True,
                    )


def renderizar_galeria_rutas(rutas: list[str]) -> None:
    fotos = list(rutas) + [""] * MAX_FOTOS
    for fila in range(3):
        columnas = st.columns(3, gap="small")
        for col_i, col in enumerate(columnas):
            idx = fila * 3 + col_i
            with col:
                ruta = fotos[idx]
                if ruta and os.path.exists(ruta):
                    st.image(ruta, use_container_width=True)
                    st.caption(f"{idx + 1}. {os.path.basename(ruta)}")
                elif ruta:
                    st.markdown(
                        f'<div class="foto-vacia">No encontrada<br>{os.path.basename(ruta)}</div>',
                        unsafe_allow_html=True,
                    )
                else:
                    st.markdown(
                        f'<div class="foto-vacia">Foto {idx + 1}</div>',
                        unsafe_allow_html=True,
                    )


def rutas_fotos_fila(fila) -> list[str]:
    rutas = []
    vistos: set[str] = set()
    for col in COLUMNAS_FOTOS:
        ruta = str(fila.get(col) or "").strip()
        if ruta and ruta.lower() not in {"nan", "none"} and ruta not in vistos:
            rutas.append(ruta)
            vistos.add(ruta)
    extra = str(fila.get("FOTOS") or "").strip()
    if extra:
        for parte in extra.split("|"):
            ruta = parte.strip()
            if ruta and ruta.lower() not in {"nan", "none"} and ruta not in vistos:
                rutas.append(ruta)
                vistos.add(ruta)
    dia = parsear_fecha_valor(fila.get("FECHA"))
    if dia is not None:
        for ruta in fotos_evidencia_octubre_secuencia(
            str(fila.get("CUADRILLA") or ""),
            dia,
            fila.get("SECUENCIA"),
        ):
            if ruta not in vistos:
                rutas.append(ruta)
                vistos.add(ruta)
    return rutas


def id_registro_fila(fila) -> str:
    existente = str(fila.get("ID_REGISTRO") or "").strip()
    if existente and existente.lower() not in {"nan", "none"}:
        return existente
    fecha = parsear_fecha_valor(fila.get("FECHA"))
    base = "|".join(
        [
            fecha.isoformat() if fecha else "",
            str(fila.get("CUADRILLA") or ""),
            str(fila.get("PERSONAL") or ""),
            str(fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("POZO") or ""),
            str(fila.get("SECUENCIA") or ""),
            str(fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD") or "")[:80],
        ]
    )
    digest = hashlib.sha1(_sin_acentos(base).encode("utf-8", errors="ignore")).hexdigest()[:8].upper()
    prefijo = fecha.strftime("%Y%m%d") if fecha else "SINFECHA"
    return f"{prefijo}-{digest}"


def carpeta_evidencia_registro(fila) -> str:
    fecha = parsear_fecha_valor(fila.get("FECHA")) or date.today()
    return os.path.join(EVIDENCIAS_DIR, fecha.strftime("%Y-%m-%d"), id_registro_fila(fila))


def estatus_evidencia_fila(fila) -> str:
    actual = str(fila.get("EVIDENCIA") or "").strip()
    if len(rutas_fotos_fila(fila)) >= MIN_FOTOS_COMPLETO:
        return EVIDENCIA_COMPLETA
    if actual == EVIDENCIA_COMPLETA:
        return EVIDENCIA_PENDIENTE
    return actual or EVIDENCIA_PENDIENTE


def campos_faltantes_registro(fila) -> list[str]:
    faltan = []
    fotos = rutas_fotos_fila(fila)
    huecos = MAX_FOTOS - len(fotos)
    if not fotos:
        faltan.append(f"Evidencia fotográfica (0 de {MAX_FOTOS})")
    elif huecos:
        faltan.append(f"Espacios fotográficos libres ({huecos} de {MAX_FOTOS})")
    desc = str(fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD") or "").strip()
    if not desc:
        faltan.append("Descripción de la actividad")
    personal = str(fila.get("PERSONAL") or "").strip() or texto_par_personal(
        str(fila.get("MECANICO EN TURNO") or fila.get("MECANICO TITULAR") or ""),
        str(fila.get("AYUDANTE EN TURNO") or fila.get("AYUDANTE TITULAR") or ""),
    )
    if not personal or personal == "—":
        faltan.append("Personal")
    ubic = str(fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("POZO") or fila.get("SECTOR") or "").strip()
    if not ubic:
        faltan.append("Ubicación")
    return faltan


def anexar_fotos_a_registro(df: pd.DataFrame, idx, archivos: list) -> int:
    fila = df.loc[idx]
    rid = id_registro_fila(fila)
    existentes = rutas_fotos_fila(fila)
    disponibles = max(0, MAX_FOTOS - len(existentes))
    nuevos = limitar_fotos(archivos)[:disponibles]
    if not nuevos:
        return 0
    carpeta = carpeta_evidencia_registro(fila)
    os.makedirs(carpeta, exist_ok=True)
    rutas = list(existentes)
    for archivo in nuevos:
        n = len(rutas) + 1
        extension = os.path.splitext(getattr(archivo, "name", "") or "")[1].lower() or ".jpg"
        destino = os.path.join(carpeta, f"foto_{n:02d}{extension}").replace("\\", "/")
        with open(destino, "wb") as salida:
            salida.write(archivo.getvalue())
        rutas.append(destino)
    for i, col in enumerate(COLUMNAS_FOTOS):
        df.at[idx, col] = rutas[i] if i < len(rutas) else ""
    df.at[idx, "FOTOS"] = " | ".join(rutas)
    df.at[idx, "ID_REGISTRO"] = rid
    df.at[idx, "EVIDENCIA"] = EVIDENCIA_COMPLETA if len(rutas) >= MIN_FOTOS_COMPLETO else EVIDENCIA_PENDIENTE
    return len(nuevos)


def render_panel_evidencia_posterior(df: pd.DataFrame, idx, *, clave: str) -> None:
    fila = df.loc[idx]
    fotos = rutas_fotos_fila(fila)
    faltan = campos_faltantes_registro(fila)
    rid = id_registro_fila(fila)
    estatus = estatus_evidencia_fila(fila)
    carpeta = carpeta_evidencia_registro(fila).replace("\\", "/")
    if estatus == EVIDENCIA_COMPLETA:
        st.success(f"Evidencia **completa** · {len(fotos)} foto(s) · ID `{rid}`")
    else:
        st.warning(f"Evidencia **pendiente** · {len(fotos)} foto(s) de {MAX_FOTOS} · ID `{rid}`")
    st.caption(f"Las imágenes se guardan en `{carpeta}` y la ruta queda en el registro de `{EXCEL_FILE}`.")
    if faltan:
        st.markdown("**Campos o fotos que faltan**")
        for item in faltan:
            st.write(f"- {item}")
    huecos = MAX_FOTOS - len(fotos)
    if huecos <= 0:
        st.caption("Este registro ya tiene los 9 espacios fotográficos ocupados.")
        return
    archivos = st.file_uploader(
        "Anexar fotografías posteriores",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key=f"ev_post_{clave}",
        help="Puedes cargar la evidencia días después. Se agregan a los espacios vacíos sin borrar las fotos ya guardadas.",
    )
    if archivos and len(archivos) > huecos:
        st.caption(f"Solo se usarán {huecos} archivo(s); el resto de espacios ya está lleno.")
    if st.button("📷 Guardar evidencia y marcar completo", type="primary", key=f"ev_guardar_{clave}"):
        if not archivos:
            st.warning("Selecciona al menos una fotografía para anexar.")
            return
        agregadas = anexar_fotos_a_registro(df, idx, archivos)
        if not agregadas:
            st.warning("No había espacios libres para más fotografías.")
            return
        guardar_en_excel(df)
        st.success(
            f"Se anexaron {agregadas} foto(s). El registro quedó **{df.at[idx, 'EVIDENCIA']}**."
        )
        st.rerun()


def guardar_evidencias(fecha: date, secuencia: str, archivos: list) -> list[str]:
    rutas = fotos_vacias()
    fotos = limitar_fotos(archivos)
    if not fotos:
        return rutas

    stamp = datetime.now().strftime("%H%M%S")
    carpeta_nombre = (
        f"{fecha.strftime('%Y-%m-%d')}_"
        f"{secuencia.replace(' ', '').replace('.', '')}_"
        f"{stamp}"
    )
    carpeta = os.path.join(EVIDENCIAS_DIR, carpeta_nombre)
    os.makedirs(carpeta, exist_ok=True)

    for i, archivo in enumerate(fotos):
        extension = os.path.splitext(archivo.name)[1].lower() or ".jpg"
        nombre = f"foto_{i + 1}{extension}"
        destino = os.path.join(carpeta, nombre)
        with open(destino, "wb") as salida:
            salida.write(archivo.getvalue())
        rutas[i] = os.path.join(EVIDENCIAS_DIR, carpeta_nombre, nombre).replace("\\", "/")
    persistencia.marcar_evidencia_nueva([ruta for ruta in rutas if ruta])
    return rutas


def _clave_columna_bitacora(nombre: str) -> str:
    texto = _sin_acentos(str(nombre or "")).casefold()
    return re.sub(r"[^a-z0-9%]+", " ", texto).strip()


def _alias_columna_bitacora(nombre: str) -> str:
    clave = _clave_columna_bitacora(nombre)
    aliases = {
        "fecha": "FECHA",
        "cuadrilla": "CUADRILLA",
        "personal": "PERSONAL",
        "secuencia": "SECUENCIA",
        "secuencia de actividades": "SECUENCIA",
        "secuencia de actividad": "SECUENCIA",
        "ubicacion": "UBICACIÓN",
        "descripcion de la actividad": "DESCRIPCION DE LA ACTIVIDAD",
        "descripcion de actividad": "DESCRIPCION DE LA ACTIVIDAD",
        "aceite sae 50 pjp4": "ACEITE SAE 50 PJP4",
        "aceite lts": "ACEITE (LTS)",
        "bandas c144 pjp4": "BANDAS C144 PJP4",
        "bomba de agua reparada pjp4": "BOMBA DE AGUA REPARADA PJP4",
        "grasa ops": "GRASA OPS",
        "refrigerante ops": "REFRIGERANTE OPS",
        "comentarios": "COMENTARIOS",
        "avance": "AVANCE %",
        "avance %": "AVANCE %",
        "revision": "REVISIÓN",
        "observacion revision": "OBSERVACIÓN REVISIÓN",
        "programado": COLUMNA_PROGRAMADO,
        "programado si": COLUMNA_PROGRAMADO,
        "es programado": COLUMNA_PROGRAMADO,
    }
    return aliases.get(clave, str(nombre or "").strip())


def es_formato_soledad(df: pd.DataFrame) -> bool:
    if df is None or df.empty:
        return False
    claves = {_clave_columna_bitacora(c) for c in df.columns}
    return "personal" in claves or "avance %" in claves or "avance" in claves or "comentarios" in claves


def _renombrar_columnas_bitacora(df: pd.DataFrame) -> pd.DataFrame:
    copia = df.copy()
    grupos: dict[str, list[int]] = {}
    orden: list[str] = []
    for pos, col in enumerate(copia.columns):
        destino = _alias_columna_bitacora(col)
        if destino not in grupos:
            grupos[destino] = []
            orden.append(destino)
        grupos[destino].append(pos)
    salida = pd.DataFrame(index=copia.index)
    for destino in orden:
        serie = None
        for pos in grupos[destino]:
            otra = copia.iloc[:, pos]
            if serie is None:
                serie = otra
                continue
            vacia = serie.isna() | serie.astype(str).str.strip().isin(["", "nan", "None", "<NA>"])
            serie = serie.where(~vacia, otra)
        salida[destino] = serie
    return salida


_CLAVES_NO_INSUMO = {
    "fecha",
    "secuencia",
    "secuencia de actividades",
    "secuencia de actividad",
    "sector",
    "pozo",
    "cuadrilla",
    "personal",
    "mecanico titular",
    "ayudante titular",
    "mecanico en turno",
    "ayudante en turno",
    "cobertura",
    "ubicacion",
    "equipo ubicacion",
    "tipo de actividad",
    "actividad corta",
    "descripcion de la actividad",
    "evaluacion del estatus",
    "fotos",
    "revision",
    "observacion revision",
    "id registro",
    "evidencia",
    "comentarios",
    "avance",
    "avance %",
    "campo",
    "ot",
}


def es_columna_insumo_excel(nombre: str) -> bool:
    clave = _clave_columna_bitacora(nombre)
    if not clave or clave in _CLAVES_NO_INSUMO or clave.startswith("foto"):
        return False
    if nombre in COLUMNAS_INSUMOS or nombre in COLUMNAS_INSUMOS_NUEVOS:
        return False
    return True


def _texto_insumo_celda(columna: str, valor) -> str:
    if texto_visible(valor, "") == "":
        return ""
    texto = str(valor).strip()
    try:
        numero = float(texto.replace(",", "."))
    except ValueError:
        if texto.lower() in {"0", "0.0"}:
            return ""
        return f"{columna}: {texto}"
    if numero <= 0:
        return ""
    return f"{columna}: {numero:g}"


def consumibles_de_fila(fila) -> str:
    partes: list[str] = []
    for col in getattr(fila, "index", []):
        if not es_columna_insumo_excel(str(col)):
            continue
        texto = _texto_insumo_celda(str(col), fila.get(col))
        if texto:
            partes.append(texto)
    if not partes:
        partes = list(insumos_de_fila(fila))
    vistos: set[str] = set()
    unicos = []
    for item in partes:
        clave = item.casefold()
        if clave not in vistos:
            unicos.append(item)
            vistos.add(clave)
    return " · ".join(unicos)


def texto_visible(valor, defecto: str = "—") -> str:
    if valor is None:
        return defecto
    try:
        if pd.isna(valor):
            return defecto
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nan", "none", "<na>", "nat"}:
        return defecto
    return texto


def texto_secuencia_limpio(valor) -> str:
    texto = texto_visible(valor, "")
    if not texto:
        return ""
    try:
        numero = float(texto.replace(",", ".").replace("%", ""))
    except ValueError:
        return texto
    if numero.is_integer():
        return str(int(numero))
    return texto


def _porcentaje_avance_registro(valor) -> int:
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        numero = 100
    if 0 < numero < 1:
        numero *= 100
    return max(0, min(100, int(round(numero))))


def avance_texto(valor) -> str:
    if texto_visible(valor, "") == "":
        return ""
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nan", "none", "<na>"}:
        return ""
    ya_tiene_porcentaje = "%" in texto
    try:
        num = float(texto.replace("%", "").replace(",", "."))
    except ValueError:
        return texto
    if not ya_tiene_porcentaje and 0 < num <= 1:
        return f"{int(round(num * 100))}%"
    return f"{int(round(num))}%"


def mapear_soledad_a_interno(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    aplicar_catalogo_maestro_excel()
    copia = _renombrar_columnas_bitacora(df)
    if "DESCRIPCION DE LA ACTIVIDAD" in copia.columns and "DESCRIPCIÓN DE LA ACTIVIDAD" not in copia.columns:
        copia["DESCRIPCIÓN DE LA ACTIVIDAD"] = copia["DESCRIPCION DE LA ACTIVIDAD"]
    if "UBICACIÓN" in copia.columns:
        if "EQUIPO/UBICACIÓN" not in copia.columns:
            copia["EQUIPO/UBICACIÓN"] = copia["UBICACIÓN"]
        if "SECTOR" not in copia.columns:
            copia["SECTOR"] = copia["UBICACIÓN"]
        def _pozo_desde_ubicacion(valor):
            sitio = sitios_operativos.clave_sitio(valor)
            if sitio:
                return sitio
            texto = str(valor or "").strip()
            return "" if sitios_operativos.es_vacio_sitio(texto) else texto

        if "POZO" not in copia.columns:
            copia["POZO"] = copia["UBICACIÓN"].map(_pozo_desde_ubicacion)
        else:
            vacio_pozo = copia["POZO"].isna() | copia["POZO"].astype(str).str.strip().str.casefold().isin(
                ["", "nan", "none", "<na>", "ninguno"]
            )
            copia.loc[vacio_pozo, "POZO"] = copia.loc[vacio_pozo, "UBICACIÓN"].map(_pozo_desde_ubicacion)
    if "PERSONAL" in copia.columns:
        personal = copia["PERSONAL"].fillna("").astype(str)
        if "MECANICO TITULAR" not in copia.columns:
            copia["MECANICO TITULAR"] = personal
        if "MECANICO EN TURNO" not in copia.columns:
            copia["MECANICO EN TURNO"] = personal
    if "ACEITE SAE 50 PJP4" in copia.columns and "ACEITE (LTS)" not in copia.columns:
        copia["ACEITE (LTS)"] = pd.to_numeric(copia["ACEITE SAE 50 PJP4"], errors="coerce").fillna(0)
    if "BANDAS C144 PJP4" in copia.columns and "BANDAS CANTIDAD" not in copia.columns:
        copia["BANDAS CANTIDAD"] = pd.to_numeric(copia["BANDAS C144 PJP4"], errors="coerce").fillna(0)
    if "GRASA OPS" in copia.columns and "GRASA (KG)" not in copia.columns:
        copia["GRASA (KG)"] = pd.to_numeric(copia["GRASA OPS"], errors="coerce").fillna(0)
    if "REFRIGERANTE OPS" in copia.columns and "REFRIGERANTE (LTS)" not in copia.columns:
        copia["REFRIGERANTE (LTS)"] = pd.to_numeric(copia["REFRIGERANTE OPS"], errors="coerce").fillna(0)
    if "BOMBA DE AGUA REPARADA PJP4" in copia.columns:
        bomba = pd.to_numeric(copia["BOMBA DE AGUA REPARADA PJP4"], errors="coerce").fillna(0)
        extra = copia["BOMBA DE AGUA REPARADA PJP4"].where(bomba > 0, "")
        if "OTROS INSUMOS" in copia.columns:
            copia["OTROS INSUMOS"] = copia["OTROS INSUMOS"].astype(str)
        else:
            copia["OTROS INSUMOS"] = extra.astype(str).replace({"nan": ""})
    return copia


def vista_revision_soledad(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(
            columns=[
                "Fecha",
                "Cuadrilla",
                "Personal",
                "Ubicación",
                "Descripción de Actividad",
                "Consumibles",
                "Comentarios",
                "Avance",
                "Evidencia",
            ]
        )
    filas = []
    for _, fila in df.iterrows():
        ubicacion = etiqueta_ubicacion_fila(fila) or str(
            fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("POZO") or fila.get("SECTOR") or ""
        ).strip()
        descripcion = str(
            fila.get("DESCRIPCION DE LA ACTIVIDAD")
            or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD")
            or ""
        ).strip()
        personal = str(fila.get("PERSONAL") or "").strip()
        if not personal:
            personal = texto_par_personal(
                str(fila.get("MECANICO EN TURNO") or fila.get("MECANICO TITULAR") or ""),
                str(fila.get("AYUDANTE EN TURNO") or fila.get("AYUDANTE TITULAR") or ""),
            )
        fecha = parsear_fecha_valor(fila.get("FECHA"))
        filas.append(
            {
                "Fecha": fecha.strftime("%d/%m/%Y") if fecha else str(fila.get("FECHA") or ""),
                "Cuadrilla": str(fila.get("CUADRILLA") or "").strip(),
                "Personal": personal,
                "Ubicación": ubicacion,
                "Descripción de Actividad": descripcion,
                "Consumibles": consumibles_de_fila(fila),
                "Comentarios": texto_visible(
                    fila["COMENTARIOS"] if "COMENTARIOS" in fila.index else fila.get("OBSERVACIÓN REVISIÓN"),
                    "",
                ),
                "Avance": avance_texto(fila["AVANCE %"] if "AVANCE %" in fila.index else fila.get("AVANCE")),
                "Evidencia": estatus_evidencia_fila(fila),
            }
        )
    return pd.DataFrame(filas)


def _valor_fecha_exportacion(valor):
    dia = parsear_fecha_valor(valor)
    return datetime(dia.year, dia.month, dia.day) if dia else ""


def columnas_oficiales_exportacion_bitacora(df: pd.DataFrame) -> list[str]:
    columnas = list(COLUMNAS_BITACORA_SOLEDAD)
    for extra in ("REVISIÓN", "OBSERVACIÓN REVISIÓN", "EVIDENCIA"):
        if extra in df.columns and extra not in columnas:
            columnas.append(extra)
    return columnas


def excel_bitacora_filtrada(df: pd.DataFrame) -> bytes:
    """Excel con las columnas oficiales y solo las filas que se ven en pantalla."""
    origen = pd.DataFrame() if df is None else df.drop(columns=["_fecha"], errors="ignore")
    columnas = columnas_oficiales_exportacion_bitacora(origen)
    salida = pd.DataFrame(index=origen.index)
    for col in columnas:
        if col in origen.columns:
            salida[col] = origen[col]
        else:
            salida[col] = ""
    if "FECHA" in salida.columns:
        salida["FECHA"] = salida["FECHA"].map(_valor_fecha_exportacion)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        salida.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    return buffer.getvalue()


def nombre_excel_bitacora_filtrada(filtro_fecha: str, df: pd.DataFrame) -> str:
    dia = parsear_fecha_valor(filtro_fecha) if filtro_fecha and filtro_fecha != "Todas" else None
    if dia is None and df is not None and not df.empty:
        origen_fechas = df["_fecha"] if "_fecha" in df.columns else df.get("FECHA")
        unicas = sorted(
            {
                parsear_fecha_valor(valor)
                for valor in (origen_fechas.tolist() if origen_fechas is not None else [])
            }
            - {None}
        )
        if len(unicas) == 1:
            dia = unicas[0]
        elif len(unicas) >= 2:
            return (
                f"bitacora_filtrada_{unicas[0].strftime('%Y-%m-%d')}"
                f"_{unicas[-1].strftime('%Y-%m-%d')}.xlsx"
            )
    if dia is None:
        dia = date.today()
    return f"bitacora_filtrada_{dia.strftime('%Y-%m-%d')}.xlsx"


def _hoja_parece_bitacora(columnas) -> bool:
    claves = {_clave_columna_bitacora(c) for c in columnas}
    return "fecha" in claves and ("cuadrilla" in claves or "personal" in claves)


def _leer_excel_crudo(ruta: str) -> pd.DataFrame:
    origen = _copia_excel_legible(ruta)
    frames: list[pd.DataFrame] = []
    ultimo = None
    try:
        libro = pd.ExcelFile(origen)
        preferidas = [h for h in HOJAS_BITACORA if h in libro.sheet_names]
        hojas = preferidas or list(libro.sheet_names)
        extras = [h for h in libro.sheet_names if h not in hojas]
        hojas = hojas + extras
    except Exception:
        hojas = list(HOJAS_BITACORA)
        libro = None
    for hoja in hojas:
        for header in (0, 1, 2):
            try:
                candidato = pd.read_excel(origen, sheet_name=hoja, header=header)
            except Exception:
                continue
            if candidato is None or candidato.empty:
                continue
            ultimo = candidato
            if _hoja_parece_bitacora(candidato.columns):
                frames.append(candidato)
                break
    if frames:
        return pd.concat(frames, ignore_index=True, sort=False)
    if ultimo is not None:
        return ultimo
    raise FileNotFoundError(ruta)


def asegurar_archivo_bitacora_soledad() -> str:
    os.makedirs(os.path.dirname(EXCEL_FILE) or DATABASE_DIR, exist_ok=True)
    if os.path.exists(EXCEL_FILE):
        return EXCEL_FILE
    for origen in FUENTES_BITACORA_SOLEDAD[1:]:
        if not os.path.exists(origen):
            continue
        try:
            crudo = _leer_excel_crudo(origen)
            interno = normalizar_dataframe(mapear_soledad_a_interno(crudo))
            guardar_en_excel(interno)
            return EXCEL_FILE
        except Exception:
            continue
    vacio = pd.DataFrame(columns=COLUMNAS_EXCEL)
    guardar_en_excel(vacio, sincronizar=False)
    return EXCEL_FILE


def normalizar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame(columns=COLUMNAS_EXCEL)
    df = _renombrar_columnas_bitacora(df)
    if es_formato_soledad(df):
        df = mapear_soledad_a_interno(df)
    df = df.drop(columns=["CAMPO", "OT"], errors="ignore")
    extras = [c for c in df.columns if c not in COLUMNAS_EXCEL]
    for col in COLUMNAS_EXCEL:
        if col not in df.columns:
            df[col] = 0 if col in COLUMNAS_NUMERICAS else ""
    for col in COLUMNAS_BITACORA_SOLEDAD:
        if col not in df.columns:
            df[col] = ""
    for col in COLUMNAS_EXCEL:
        if col in COLUMNAS_NUMERICAS:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = df[col].fillna("").astype(str)
            df.loc[df[col].isin(["nan", "None", "<NA>"]), col] = ""
            if col == COLUMNA_PROGRAMADO:
                df[col] = df[col].map(_normalizar_marca_programado)
    if "SECUENCIA" in df.columns:
        df["SECUENCIA"] = df["SECUENCIA"].map(texto_secuencia_limpio)
    for col in extras + list(COLUMNAS_BITACORA_SOLEDAD):
        if col not in df.columns:
            continue
        if col in COLUMNAS_CONSUMIBLE_SOLEDAD or col == "AVANCE %":
            df[col] = pd.to_numeric(df[col], errors="coerce")
        else:
            df[col] = df[col].fillna("").astype(str)
            df.loc[df[col].isin(["nan", "None", "<NA>"]), col] = ""
    vacio_revision = df["REVISIÓN"].astype(str).str.strip() == ""
    df.loc[vacio_revision, "REVISIÓN"] = "Pendiente"
    if "ID_REGISTRO" not in df.columns:
        df["ID_REGISTRO"] = ""
    if "EVIDENCIA" not in df.columns:
        df["EVIDENCIA"] = EVIDENCIA_PENDIENTE
    df["ID_REGISTRO"] = [id_registro_fila(df.loc[i]) for i in df.index]
    df["EVIDENCIA"] = [estatus_evidencia_fila(df.loc[i]) for i in df.index]
    orden = list(COLUMNAS_EXCEL)
    for col in list(COLUMNAS_BITACORA_SOLEDAD) + extras:
        if col not in orden and col in df.columns:
            orden.append(col)
    return df[orden].copy()


def cargar_bitacora() -> pd.DataFrame:
    aplicar_catalogo_maestro_excel()
    asegurar_archivo_bitacora_soledad()
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=COLUMNAS_EXCEL)
    try:
        df = normalizar_dataframe(_leer_excel_crudo(EXCEL_FILE))
        if df is not None and not df.empty and "ID_REGISTRO" in df.columns:
            df = df.drop_duplicates(subset=["ID_REGISTRO"], keep="last")
        return df
    except Exception:
        if os.path.exists(EXCEL_FILE_LEGACY):
            try:
                return normalizar_dataframe(_leer_excel_crudo(EXCEL_FILE_LEGACY))
            except Exception:
                return pd.DataFrame(columns=COLUMNAS_EXCEL)
        return pd.DataFrame(columns=COLUMNAS_EXCEL)


def guardar_en_excel(df_final: pd.DataFrame, *, sincronizar: bool = True) -> None:
    os.makedirs(os.path.dirname(EXCEL_FILE) or DATABASE_DIR, exist_ok=True)
    df_final = df_final.drop(columns=["_fecha"], errors="ignore")
    df_final = normalizar_dataframe(df_final)
    columnas_salida = []
    for col in list(COLUMNAS_BITACORA_SOLEDAD) + ["REVISIÓN", "OBSERVACIÓN REVISIÓN", "ID_REGISTRO", "EVIDENCIA"] + list(COLUMNAS_EXCEL):
        if col in df_final.columns and col not in columnas_salida:
            columnas_salida.append(col)
    salida = df_final[columnas_salida].copy()
    actuales = persistencia.contar_filas_bitacora(EXCEL_FILE) if os.path.exists(EXCEL_FILE) else 0
    if persistencia.debe_bloquear_vaciado(actuales, len(salida)):
        return
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        if os.path.exists(EXCEL_FILE):
            from openpyxl import load_workbook
            from openpyxl.utils.dataframe import dataframe_to_rows

            wb = load_workbook(_copia_excel_legible(EXCEL_FILE))
            if SHEET_NAME in wb.sheetnames:
                del wb[SHEET_NAME]
            ws = wb.create_sheet(SHEET_NAME, 0)
            for fila in dataframe_to_rows(salida, index=False, header=True):
                ws.append(list(fila))
            wb.save(tmp)
            wb.close()
        else:
            with pd.ExcelWriter(tmp, engine="openpyxl", mode="w") as writer:
                salida.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        try:
            persistencia.copiar_seguro(tmp, EXCEL_FILE)
        except PermissionError:
            alt = os.path.join(os.path.dirname(EXCEL_FILE) or DATABASE_DIR, "~pjp4_bitacora_write.xlsx")
            persistencia.copiar_seguro(tmp, alt)
            persistencia.reemplazar_seguro(alt, EXCEL_FILE)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    if sincronizar and len(salida) > 0:
        persistencia.respaldar_cambios("Actualización de bitácora")
        try:
            sincronizar_calendario_desde_bitacora()
        except Exception:
            pass


def _encabezados_hoja_bitacora(ws) -> tuple[int, list[tuple[int, str]]]:
    limite = min(int(ws.max_row or 1), 12)
    for fila in range(1, limite + 1):
        celdas = []
        for col in range(1, int(ws.max_column or 1) + 1):
            valor = ws.cell(fila, col).value
            if valor is None or str(valor).strip() == "":
                continue
            celdas.append((col, str(valor).strip()))
        claves = {_clave_columna_bitacora(nombre) for _, nombre in celdas}
        if "fecha" in claves and ("cuadrilla" in claves or "personal" in claves):
            return fila, celdas
    raise ValueError("La bitácora no tiene una fila de encabezados reconocible.")


def _hoja_bitacora_editable(wb):
    nombres = [nombre for nombre in HOJAS_BITACORA if nombre in wb.sheetnames] or list(wb.sheetnames)
    for nombre in nombres:
        try:
            fila, encabezados = _encabezados_hoja_bitacora(wb[nombre])
        except ValueError:
            continue
        return wb[nombre], fila, encabezados
    raise ValueError("No se encontró la hoja de la bitácora operativa.")


def _columna_encabezado(encabezados: list[tuple[int, str]], *buscados: str) -> int | None:
    objetivos = {_clave_columna_bitacora(item) for item in buscados}
    for col, nombre in encabezados:
        if _alias_columna_bitacora(nombre) in buscados or _clave_columna_bitacora(nombre) in objetivos:
            return col
    return None


def _siguiente_fila_bitacora(ws, fila_header: int, col_fecha: int) -> int:
    ultima = fila_header
    for fila in range(fila_header + 1, int(ws.max_row or fila_header) + 1):
        valor = ws.cell(fila, col_fecha).value
        if valor is not None and str(valor).strip() != "":
            ultima = fila
    return ultima + 1


def _siguiente_secuencia_en_hoja(ws, fila_header: int, encabezados: list[tuple[int, str]], dia: date, cuadrilla: str) -> int:
    col_fecha = _columna_encabezado(encabezados, "FECHA")
    col_cuadrilla = _columna_encabezado(encabezados, "CUADRILLA")
    col_secuencia = _columna_encabezado(encabezados, "SECUENCIA", "SECUENCIA DE ACTIVIDADES")
    if not col_fecha or not col_cuadrilla:
        return 1
    objetivo = clave_cuadrilla_oficial(cuadrilla)
    numeros = []
    for fila in range(fila_header + 1, int(ws.max_row or fila_header) + 1):
        if parsear_fecha_valor(ws.cell(fila, col_fecha).value) != dia:
            continue
        if clave_cuadrilla_oficial(str(ws.cell(fila, col_cuadrilla).value or "")) != objetivo:
            continue
        numero = numero_secuencia_actividad(ws.cell(fila, col_secuencia).value) if col_secuencia else None
        if numero:
            numeros.append(numero)
    return max(numeros) + 1 if numeros else 1


def _mejor_columna_consumible(token: str, encabezados: list[tuple[int, str]]) -> tuple[int, str] | None:
    clave_token = _clave_columna_bitacora(token)
    if not clave_token:
        return None
    mejor = None
    mejor_puntaje = -1.0
    for col, nombre in encabezados:
        if not es_columna_insumo_excel(nombre):
            continue
        clave_col = _clave_columna_bitacora(nombre)
        if not clave_col:
            continue
        if clave_token == clave_col:
            puntaje = 1000 + len(clave_col)
        elif clave_token in clave_col or clave_col in clave_token:
            puntaje = float(min(len(clave_token), len(clave_col))) - abs(len(clave_col) - len(clave_token)) * 0.05
        else:
            continue
        if puntaje > mejor_puntaje:
            mejor = (col, nombre)
            mejor_puntaje = puntaje
    return mejor


def _cantidad_consumible(token: str, columna: str) -> float:
    numeros_columna = set(re.findall(r"\d+(?:[.,]\d+)?", columna))
    propios = [numero for numero in re.findall(r"\d+(?:[.,]\d+)?", token) if numero not in numeros_columna]
    if not propios:
        return 1
    try:
        return float(propios[-1].replace(",", "."))
    except ValueError:
        return 1


def _asignar_consumibles(texto: str, encabezados: list[tuple[int, str]]) -> tuple[dict[int, float], list[str]]:
    partes = [parte.strip() for parte in re.split(r"[,;\n]+", str(texto or "")) if parte.strip()]
    usados: dict[int, float] = {}
    sueltos: list[str] = []
    for parte in partes:
        elegido = _mejor_columna_consumible(parte, encabezados)
        if elegido is None:
            sueltos.append(parte)
            continue
        col, nombre = elegido
        usados[col] = _cantidad_consumible(parte, nombre)
    return usados, sueltos


def _nombre_foto_bitacora(indice: int, original: str) -> str:
    base = os.path.basename(str(original or "foto.jpg"))
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", base).strip("._") or "foto.jpg"
    raiz, extension = os.path.splitext(base)
    if extension.lower() not in {".jpg", ".jpeg", ".png"}:
        extension = ".jpg"
    return f"{int(indice):02d}_{raiz}{extension.lower()}"


def guardar_fotos_bitacora_octubre(
    cuadrilla: str,
    dia: date,
    secuencia,
    archivos: list | None,
    raiz: str | None = None,
) -> list[str]:
    if not archivos:
        return []
    carpeta = carpeta_evidencia_octubre_secuencia(cuadrilla, dia, secuencia)
    if raiz:
        carpeta = os.path.join(raiz, os.path.relpath(carpeta, EVIDENCIA_OCTUBRE_DIR))
    os.makedirs(carpeta, exist_ok=True)
    rutas = []
    for indice, archivo in enumerate(archivos, start=1):
        destino = os.path.join(carpeta, _nombre_foto_bitacora(indice, getattr(archivo, "name", "foto.jpg")))
        raiz_nombre, extension = os.path.splitext(destino)
        copia = 2
        while os.path.exists(destino):
            destino = f"{raiz_nombre}_{copia}{extension}"
            copia += 1
        contenido = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo.read()
        with open(destino, "wb") as salida:
            salida.write(contenido)
        rutas.append(destino.replace("\\", "/"))
    persistencia.marcar_evidencia_nueva(rutas)
    return rutas


def _guardar_libro_bitacora(wb, destino: str) -> None:
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb.save(tmp)
        try:
            persistencia.copiar_seguro(tmp, destino)
        except PermissionError:
            alt = os.path.join(os.path.dirname(destino) or DATABASE_DIR, "~pjp4_bitacora_write.xlsx")
            persistencia.copiar_seguro(tmp, alt)
            persistencia.reemplazar_seguro(alt, destino)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        wb.close()


def _escribir_texto_catalogo(ws, fila_header: int, encabezados: list, fila: int, valor, titulo: str) -> None:
    texto = str(valor or "").strip()
    if not texto:
        return
    col = _columna_encabezado(encabezados, titulo)
    if not col:
        col = max((columna for columna, _nombre in encabezados), default=0) + 1
        ws.cell(fila_header, col).value = titulo
        encabezados.append((col, titulo))
    ws.cell(fila, col).value = texto


def agregar_registro_bitacora_soledad(
    fecha: date,
    cuadrilla: str,
    personal: str,
    ubicacion: str,
    descripcion: str,
    consumibles: str = "",
    comentarios: str = "",
    fotos: list | None = None,
    ruta: str | None = None,
    raiz_fotos: str | None = None,
    secuencia: int | None = None,
    avance: int | float | None = None,
    equipo: str = "",
    tipo_actividad: str = "",
    actividad_corta: str = "",
) -> dict:
    """Agrega una fila a la hoja original de la bitácora y conserva el resto del Excel."""
    destino = ruta or EXCEL_FILE
    if not os.path.exists(destino):
        raise FileNotFoundError(destino)
    from openpyxl import load_workbook

    wb = load_workbook(destino)
    ws, fila_header, encabezados = _hoja_bitacora_editable(wb)
    col_fecha = _columna_encabezado(encabezados, "FECHA")
    col_cuadrilla = _columna_encabezado(encabezados, "CUADRILLA")
    col_personal = _columna_encabezado(encabezados, "PERSONAL")
    col_secuencia = _columna_encabezado(encabezados, "SECUENCIA", "SECUENCIA DE ACTIVIDADES")
    col_ubicacion = _columna_encabezado(encabezados, "UBICACIÓN", "UBICACION")
    col_desc = _columna_encabezado(encabezados, "DESCRIPCION DE LA ACTIVIDAD")
    col_coment = _columna_encabezado(encabezados, "COMENTARIOS")
    if not col_fecha or not col_cuadrilla or not col_desc:
        wb.close()
        raise ValueError("La bitácora no tiene fecha, cuadrilla o descripción.")
    if secuencia is None:
        secuencia = _siguiente_secuencia_en_hoja(ws, fila_header, encabezados, fecha, cuadrilla)
    else:
        secuencia = max(1, int(secuencia))
    fila = _siguiente_fila_bitacora(ws, fila_header, col_fecha)
    rutas = guardar_fotos_bitacora_octubre(cuadrilla, fecha, secuencia, fotos or [], raiz_fotos)
    try:
        ws.cell(fila, col_fecha).value = datetime(fecha.year, fecha.month, fecha.day)
        ws.cell(fila, col_cuadrilla).value = cuadrilla.strip()
        if col_personal:
            ws.cell(fila, col_personal).value = personal.strip()
        if col_secuencia:
            ws.cell(fila, col_secuencia).value = int(secuencia)
        if col_ubicacion:
            ws.cell(fila, col_ubicacion).value = ubicacion.strip()
        ws.cell(fila, col_desc).value = descripcion.strip()
        cantidades, sueltos = _asignar_consumibles(consumibles, encabezados)
        for col, cantidad in cantidades.items():
            ws.cell(fila, col).value = int(cantidad) if float(cantidad).is_integer() else float(cantidad)
        texto_coment = comentarios.strip()
        if sueltos:
            extra = "Consumibles: " + ", ".join(sueltos)
            texto_coment = f"{texto_coment} · {extra}" if texto_coment else extra
        if col_coment and texto_coment:
            ws.cell(fila, col_coment).value = texto_coment
        if avance is not None:
            col_avance = _columna_encabezado(encabezados, "AVANCE %", "AVANCE")
            if not col_avance:
                col_avance = max((col for col, _nombre in encabezados), default=0) + 1
                ws.cell(fila_header, col_avance).value = "AVANCE %"
                encabezados.append((col_avance, "AVANCE %"))
            ws.cell(fila, col_avance).value = _porcentaje_avance_registro(avance)
        equipo_final = nombre_equipo_catalogo(equipo)
        if equipo_final:
            col_equipo = _columna_encabezado(encabezados, "EQUIPO")
            if not col_equipo:
                col_equipo = max((col for col, _nombre in encabezados), default=0) + 1
                ws.cell(fila_header, col_equipo).value = "EQUIPO"
            ws.cell(fila, col_equipo).value = equipo_final
        _escribir_texto_catalogo(ws, fila_header, encabezados, fila, tipo_actividad, "TIPO DE ACTIVIDAD")
        _escribir_texto_catalogo(
            ws, fila_header, encabezados, fila, actividad_corta, "ACTIVIDAD CORTA"
        )
        _guardar_libro_bitacora(wb, destino)
    except Exception:
        for ruta_foto in rutas:
            try:
                os.remove(ruta_foto)
            except OSError:
                pass
        raise
    try:
        df_bitacora = cargar_bitacora() if destino == EXCEL_FILE else pd.DataFrame()
    except Exception:
        df_bitacora = pd.DataFrame()
    if destino == EXCEL_FILE:
        persistencia.respaldar_cambios("Registro de actividad en bitácora")
        try:
            sincronizar_calendario_desde_bitacora()
        except Exception:
            pass
    return {
        "secuencia": int(secuencia),
        "fila": fila,
        "fotos": rutas,
        "carpeta": (
            os.path.dirname(rutas[0]).replace("\\", "/")
            if rutas
            else carpeta_evidencia_octubre_secuencia(cuadrilla, fecha, secuencia).replace("\\", "/")
        ),
        "df_bitacora": df_bitacora,
    }


def _texto_comparable_bitacora(valor) -> str:
    texto = texto_visible(valor, "")
    return re.sub(r"\s+", " ", _sin_acentos(texto)).strip().casefold()


def _columnas_registro_hoja(encabezados: list[tuple[int, str]]) -> dict[str, int | None]:
    return {
        "fecha": _columna_encabezado(encabezados, "FECHA"),
        "cuadrilla": _columna_encabezado(encabezados, "CUADRILLA"),
        "personal": _columna_encabezado(encabezados, "PERSONAL"),
        "secuencia": _columna_encabezado(encabezados, "SECUENCIA", "SECUENCIA DE ACTIVIDADES"),
        "ubicacion": _columna_encabezado(encabezados, "UBICACIÓN", "UBICACION"),
        "descripcion": _columna_encabezado(encabezados, "DESCRIPCION DE LA ACTIVIDAD"),
        "comentarios": _columna_encabezado(encabezados, "COMENTARIOS"),
    }


def _fila_hoja_coincide(ws, fila: int, cols: dict, registro) -> bool:
    if parsear_fecha_valor(ws.cell(fila, cols["fecha"]).value) != parsear_fecha_valor(registro.get("FECHA")):
        return False
    if clave_cuadrilla_oficial(str(ws.cell(fila, cols["cuadrilla"]).value or "")) != clave_cuadrilla_oficial(
        str(registro.get("CUADRILLA") or "")
    ):
        return False
    if cols["secuencia"] and numero_secuencia_actividad(ws.cell(fila, cols["secuencia"]).value) != numero_secuencia_actividad(
        registro.get("SECUENCIA")
    ):
        return False
    if cols["descripcion"]:
        desc_df = registro.get("DESCRIPCION DE LA ACTIVIDAD") or registro.get("DESCRIPCIÓN DE LA ACTIVIDAD")
        if _texto_comparable_bitacora(ws.cell(fila, cols["descripcion"]).value) != _texto_comparable_bitacora(desc_df):
            return False
    return True


def _localizar_fila_hoja(ws, fila_header: int, cols: dict, registro, idx) -> int:
    """Ubica en la hoja la fila del registro elegido, sin depender solo del índice."""
    if not cols["fecha"] or not cols["cuadrilla"]:
        raise ValueError("La bitácora no tiene columnas de fecha y cuadrilla.")
    candidata = None
    try:
        candidata = fila_header + 1 + int(idx)
    except (TypeError, ValueError):
        candidata = None
    if candidata and candidata <= ws.max_row and _fila_hoja_coincide(ws, candidata, cols, registro):
        return candidata
    coincidencias = [
        fila
        for fila in range(fila_header + 1, int(ws.max_row or fila_header) + 1)
        if _fila_hoja_coincide(ws, fila, cols, registro)
    ]
    if not coincidencias:
        raise ValueError("No se encontró ese registro en la bitácora. Recarga la vista e inténtalo de nuevo.")
    if candidata:
        coincidencias.sort(key=lambda fila: abs(fila - candidata))
    return coincidencias[0]


def _secuencia_ocupada_hoja(ws, fila_header: int, cols: dict, dia: date, cuadrilla: str, secuencia: int, excluir: int) -> bool:
    if not cols["secuencia"]:
        return False
    objetivo = clave_cuadrilla_oficial(cuadrilla)
    for fila in range(fila_header + 1, int(ws.max_row or fila_header) + 1):
        if fila == excluir:
            continue
        if parsear_fecha_valor(ws.cell(fila, cols["fecha"]).value) != dia:
            continue
        if clave_cuadrilla_oficial(str(ws.cell(fila, cols["cuadrilla"]).value or "")) != objetivo:
            continue
        if numero_secuencia_actividad(ws.cell(fila, cols["secuencia"]).value) == secuencia:
            return True
    return False


def _abrir_libro_bitacora(destino: str):
    from openpyxl import load_workbook

    if not os.path.exists(destino):
        raise FileNotFoundError(destino)
    wb = load_workbook(_copia_excel_legible(destino))
    ws, fila_header, encabezados = _hoja_bitacora_editable(wb)
    return wb, ws, fila_header, _columnas_registro_hoja(encabezados)


def es_ruta_evidencia_editable(ruta: str) -> bool:
    """Solo se borran archivos dentro de las carpetas de evidencia del proyecto."""
    if not ruta:
        return False
    absoluta = os.path.realpath(ruta if os.path.isabs(ruta) else os.path.join(persistencia.RAIZ, ruta))
    for carpeta in persistencia.CARPETAS_EVIDENCIA:
        base = os.path.realpath(os.path.join(persistencia.RAIZ, carpeta))
        try:
            if os.path.commonpath([absoluta, base]) == base and absoluta != base:
                return True
        except ValueError:
            continue
    return False


def _quitar_carpetas_vacias(carpeta: str) -> None:
    raiz = os.path.realpath(os.path.join(persistencia.RAIZ, EVIDENCIA_OCTUBRE_DIR))
    actual = os.path.realpath(carpeta)
    while actual and actual != raiz and es_ruta_evidencia_editable(actual):
        try:
            if os.listdir(actual):
                return
            persistencia.borrar_carpeta_vacia(actual)
        except OSError:
            return
        actual = os.path.dirname(actual)


def carpeta_fotos_registro_bitacora(fila) -> str:
    dia = parsear_fecha_valor(fila.get("FECHA")) or date.today()
    existentes = fotos_evidencia_octubre_secuencia(str(fila.get("CUADRILLA") or ""), dia, fila.get("SECUENCIA"))
    if existentes:
        return os.path.dirname(existentes[0])
    return carpeta_evidencia_octubre_secuencia(str(fila.get("CUADRILLA") or ""), dia, fila.get("SECUENCIA"))


def _mover_fotos_registro(origen: str, destino: str) -> tuple[list[str], list[str]]:
    if not origen or not os.path.isdir(origen):
        return [], []
    if os.path.realpath(origen) == os.path.realpath(destino):
        return [], []
    os.makedirs(destino, exist_ok=True)
    nuevas, viejas = [], []
    for nombre in sorted(os.listdir(origen)):
        fuente = os.path.join(origen, nombre)
        if not os.path.isfile(fuente):
            continue
        objetivo = os.path.join(destino, nombre)
        raiz, extension = os.path.splitext(objetivo)
        copia = 2
        while os.path.exists(objetivo):
            objetivo = f"{raiz}_{copia}{extension}"
            copia += 1
        shutil.move(fuente, objetivo)
        viejas.append(fuente.replace("\\", "/"))
        nuevas.append(objetivo.replace("\\", "/"))
    _quitar_carpetas_vacias(origen)
    return nuevas, viejas


def _ajustar_tablas_tras_borrar(ws, fila: int) -> None:
    from openpyxl.utils.cell import get_column_letter, range_boundaries

    for tabla in list(ws.tables.values()):
        try:
            min_col, min_row, max_col, max_row = range_boundaries(tabla.ref)
        except (TypeError, ValueError):
            continue
        if not (min_row < fila <= max_row) or max_row - 1 <= min_row:
            continue
        nuevo = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row - 1}"
        tabla.ref = nuevo
        if getattr(tabla, "autoFilter", None) is not None:
            tabla.autoFilter.ref = nuevo


def despues_de_cambio_bitacora(motivo: str) -> dict:
    """Respalda la bitácora y vuelve a cruzar el calendario de asistencia."""
    try:
        persistencia.respaldar_cambios(motivo)
    except Exception:
        pass
    for clave in ("_cal_sync_firma", "_cal_sync_fallo"):
        st.session_state.pop(clave, None)
    try:
        return sincronizar_calendario_desde_bitacora()
    except Exception as exc:
        return {"ok": False, "cambios": 0, "mensaje": str(exc)}


def actualizar_registro_bitacora(idx, registro, cambios: dict, ruta: str | None = None) -> dict:
    """Reescribe en su misma fila los campos editados de un registro de la bitácora."""
    destino = ruta or EXCEL_FILE
    wb, ws, fila_header, cols = _abrir_libro_bitacora(destino)
    try:
        fila = _localizar_fila_hoja(ws, fila_header, cols, registro, idx)
        dia_viejo = parsear_fecha_valor(registro.get("FECHA"))
        dia = cambios.get("fecha") or dia_viejo
        cuadrilla = str(cambios.get("cuadrilla") or registro.get("CUADRILLA") or "").strip()
        secuencia = cambios.get("secuencia")
        if secuencia is None:
            secuencia = numero_secuencia_actividad(registro.get("SECUENCIA"))
        if not isinstance(dia, date):
            raise ValueError("La fecha del registro no es válida.")
        if not cuadrilla:
            raise ValueError("Selecciona la cuadrilla.")
        descripcion = str(cambios.get("descripcion", "")).strip()
        if not descripcion:
            raise ValueError("La descripción de la actividad no puede quedar vacía.")
        if secuencia and _secuencia_ocupada_hoja(ws, fila_header, cols, dia, cuadrilla, int(secuencia), fila):
            raise ValueError(
                f"La secuencia {int(secuencia)} ya existe para {etiqueta_cuadrilla_historial(cuadrilla)} "
                f"el {dia.strftime('%d/%m/%Y')}. Elige otra."
            )
        ws.cell(fila, cols["fecha"]).value = datetime(dia.year, dia.month, dia.day)
        ws.cell(fila, cols["cuadrilla"]).value = cuadrilla
        if cols["personal"] and "personal" in cambios:
            ws.cell(fila, cols["personal"]).value = str(cambios["personal"]).strip() or None
        if cols["secuencia"] and secuencia:
            ws.cell(fila, cols["secuencia"]).value = int(secuencia)
        if cols["ubicacion"] and "ubicacion" in cambios:
            ws.cell(fila, cols["ubicacion"]).value = str(cambios["ubicacion"]).strip() or None
        if cols["descripcion"]:
            ws.cell(fila, cols["descripcion"]).value = descripcion
        if cols["comentarios"] and "comentarios" in cambios:
            ws.cell(fila, cols["comentarios"]).value = str(cambios["comentarios"]).strip() or None
        origen_fotos = carpeta_fotos_registro_bitacora(registro)
        _guardar_libro_bitacora(wb, destino)
    except Exception:
        wb.close()
        raise
    destino_fotos = carpeta_evidencia_octubre_secuencia(cuadrilla, dia, secuencia)
    movidas, viejas = [], []
    aviso = ""
    try:
        movidas, viejas = _mover_fotos_registro(origen_fotos, destino_fotos)
    except OSError as exc:
        aviso = f"Las fotos se quedaron en {origen_fotos.replace(os.sep, '/')}: {exc}"
    if movidas:
        persistencia.marcar_evidencia_nueva(movidas)
        persistencia.marcar_evidencia_eliminada(viejas)
    return {"fila": fila, "fotos_movidas": len(movidas), "aviso": aviso}


def eliminar_registro_bitacora(idx, registro, ruta: str | None = None) -> dict:
    """Quita la fila de la bitácora y borra la carpeta de fotos de esa actividad."""
    destino = ruta or EXCEL_FILE
    wb, ws, fila_header, cols = _abrir_libro_bitacora(destino)
    try:
        fila = _localizar_fila_hoja(ws, fila_header, cols, registro, idx)
        ws.delete_rows(fila, 1)
        _ajustar_tablas_tras_borrar(ws, fila)
        _guardar_libro_bitacora(wb, destino)
    except Exception:
        wb.close()
        raise
    borradas = []
    carpetas = set()
    for ruta_foto in rutas_fotos_fila(registro):
        if not es_ruta_evidencia_editable(ruta_foto) or not os.path.isfile(ruta_foto):
            continue
        try:
            persistencia.borrar_archivo(ruta_foto)
        except OSError:
            continue
        borradas.append(ruta_foto.replace("\\", "/"))
        carpetas.add(os.path.dirname(ruta_foto))
    for carpeta in carpetas:
        _quitar_carpetas_vacias(carpeta)
    persistencia.registrar_baja_bitacora(1)
    persistencia.marcar_evidencia_eliminada(borradas)
    return {"fila": fila, "fotos_borradas": len(borradas)}


def eliminar_foto_registro(ruta_foto: str) -> bool:
    if not es_ruta_evidencia_editable(ruta_foto):
        raise ValueError("Esa ruta no pertenece a las carpetas de evidencia.")
    if not os.path.isfile(ruta_foto):
        return False
    persistencia.borrar_archivo(ruta_foto)
    persistencia.marcar_evidencia_eliminada([ruta_foto])
    _quitar_carpetas_vacias(os.path.dirname(ruta_foto))
    return True


def anexar_fotos_registro_bitacora(registro, archivos: list) -> list[str]:
    dia = parsear_fecha_valor(registro.get("FECHA")) or date.today()
    carpeta = carpeta_fotos_registro_bitacora(registro)
    os.makedirs(carpeta, exist_ok=True)
    inicio = len(fotos_evidencia_octubre_secuencia(str(registro.get("CUADRILLA") or ""), dia, registro.get("SECUENCIA")))
    rutas = []
    for numero, archivo in enumerate(archivos or [], start=inicio + 1):
        destino = os.path.join(carpeta, _nombre_foto_bitacora(numero, getattr(archivo, "name", "foto.jpg")))
        raiz, extension = os.path.splitext(destino)
        copia = 2
        while os.path.exists(destino):
            destino = f"{raiz}_{copia}{extension}"
            copia += 1
        contenido = archivo.getvalue() if hasattr(archivo, "getvalue") else archivo.read()
        with open(destino, "wb") as salida:
            salida.write(contenido)
        rutas.append(destino.replace("\\", "/"))
    persistencia.marcar_evidencia_nueva(rutas)
    return rutas


_MESES_PROGRAMA = {
    "enero": 1,
    "ene": 1,
    "febrero": 2,
    "feb": 2,
    "marzo": 3,
    "mar": 3,
    "abril": 4,
    "abr": 4,
    "mayo": 5,
    "may": 5,
    "junio": 6,
    "jun": 6,
    "julio": 7,
    "jul": 7,
    "agosto": 8,
    "ago": 8,
    "septiembre": 9,
    "setiembre": 9,
    "sept": 9,
    "sep": 9,
    "octubre": 10,
    "oct": 10,
    "noviembre": 11,
    "nov": 11,
    "diciembre": 12,
    "dic": 12,
}
_TOKENS_ACTIVIDAD_RUIDO = {
    "de",
    "del",
    "la",
    "el",
    "los",
    "las",
    "a",
    "al",
    "y",
    "en",
    "por",
    "para",
    "con",
    "mantenimiento",
    "mtto",
    "mantto",
    "mant",
}
_TOKENS_EQUIPO_PROGRAMA = {"mci", "ubm", "ici", "ub"}
_COLUMNAS_PROGRAMA_ARCHIVO = [
    "ID",
    "ZONA",
    "UBICACION",
    "CODIGO",
    "TIPO DE MANTENIMIENTO",
    "FECHA PROGRAMADA",
    "MES",
]
_PREFIJOS_CODIGO_PROGRAMA = (
    ("SN", "PSN"),
    ("CY", "PC"),
    ("AR", "PA"),
    ("PPB", "PPB"),
    ("PB", "PB"),
    ("PGA", "PGA"),
    ("PSN", "PSN"),
    ("PA", "PA"),
    ("PS", "PS"),
    ("PC", "PC"),
    ("G", "PGA"),
    ("S", "PS"),
)


def _normalizar_marca_programado(valor) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    texto = _sin_acentos(str(valor)).strip()
    if not texto or texto.casefold() in {"nan", "none", "nat", "<na>"}:
        return ""
    plano = texto.casefold()
    if plano in {"si", "s", "yes", "1", "true", "cumplido", "programado"} or plano.startswith("si "):
        return "Sí"
    if plano in {"no", "n", "0", "false"}:
        return "No"
    return texto


def _marca_programado_si(valor) -> bool:
    return _normalizar_marca_programado(valor) == "Sí"


def _anio_en_texto(valor, defecto: int) -> int:
    coincidencia = re.search(r"(20\d{2})", str(valor or ""))
    return int(coincidencia.group(1)) if coincidencia else int(defecto)


def _mes_desde_texto(valor) -> int | None:
    plano = _norm_encabezado(valor)
    if not plano:
        return None
    for nombre, numero in sorted(_MESES_PROGRAMA.items(), key=lambda item: len(item[0]), reverse=True):
        if re.search(rf"\b{re.escape(nombre)}\b", plano):
            return numero
    return None


def _fecha_programada(valor, anio_defecto: int | None = None) -> date | None:
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor).strip()
    if not texto or texto.casefold() in {"nan", "none", "nat", "<na>"}:
        return None
    mes_anio = re.fullmatch(r"(\d{1,2})[/-](\d{4})", texto)
    if mes_anio:
        mes, anio = int(mes_anio.group(1)), int(mes_anio.group(2))
        if 1 <= mes <= 12:
            return date(anio, mes, 1)
    anio_mes = re.fullmatch(r"(\d{4})[/-](\d{1,2})", texto)
    if anio_mes:
        anio, mes = int(anio_mes.group(1)), int(anio_mes.group(2))
        if 1 <= mes <= 12:
            return date(anio, mes, 1)
    mes = _mes_desde_texto(texto)
    anio_texto = re.search(r"(20\d{2})", texto)
    if mes and anio_texto:
        return date(int(anio_texto.group(1)), mes, 1)
    if mes and anio_defecto:
        return date(int(anio_defecto), mes, 1)
    if _norm_encabezado(texto) in {"mes", "fecha"}:
        return None
    return parsear_fecha_valor(valor)


def _etiqueta_mes_programa(cuando: date | None) -> str:
    if not cuando:
        return ""
    return f"{NOMBRES_MES[cuando.month - 1]} {cuando.year}"


def _rol_columna_programa(normal: str) -> str | None:
    if not normal or normal in {"id", "nro", "no"}:
        return None
    if any(token in normal for token in ("ubicacion", "pozo", "localizacion")) and "tipo" not in normal:
        return "ubicacion"
    if ("tipo" in normal and "mant" in normal) or normal in {
        "tipo",
        "actividad",
        "mantenimiento",
        "trabajo",
    }:
        return "tipo"
    if normal.startswith("actividad") or normal.startswith("tipo "):
        return "tipo"
    if "fecha" in normal or normal in {"mes", "periodo"} or normal.startswith("mes "):
        return "fecha"
    return None


def _fecha_marca_mes(valor, anio: int, mes: int) -> date | None:
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        dia = int(valor)
        if dia == 0:
            return None
        if 1 <= dia <= 31:
            try:
                return date(anio, mes, dia)
            except ValueError:
                return date(anio, mes, 1)
        return None
    texto = _texto_actividad(valor)
    plano = _sin_acentos(texto).casefold()
    if not plano or plano in {"-", "0", "no", "n", "nan", "none"}:
        return None
    fecha = _fecha_programada(texto, anio)
    if fecha:
        return fecha
    return date(anio, mes, 1)


def _pozo_desde_codigo_programa(codigo: str) -> tuple[str, str]:
    """AR-67 (UBM) queda en PA67 y UBM; SN-46 (Motor) queda en PSN46 y MCI."""
    texto = " ".join(str(codigo or "").split())
    if not texto:
        return "", ""
    plano = _sin_acentos(texto)
    equipo = ""
    if re.search(r"\bUBM\b", plano, re.IGNORECASE):
        equipo = "UBM"
    elif re.search(r"\bMOTOR\b", plano, re.IGNORECASE):
        equipo = "MCI"
    base = re.split(r"\s*\(", texto, maxsplit=1)[0].strip()
    compacto = clave_pozo_plana(base)
    for prefijo, destino in _PREFIJOS_CODIGO_PROGRAMA:
        if compacto.startswith(prefijo) and len(compacto) > len(prefijo) and compacto[len(prefijo)].isdigit():
            return f"{destino}{compacto[len(prefijo):]}", equipo
    return base or texto, equipo


def _tipo_programa_operativo(tipo: str, actividad: str, equipo: str) -> str:
    base = " ".join(str(tipo or actividad or "").split())
    if not base:
        return ""
    if equipo and equipo.casefold() not in _sin_acentos(base).casefold():
        return f"{base} {equipo}"
    return base


def _id_programa_mantenimiento(ubicacion: str, tipo: str, cuando: date | None, zona: str = "") -> str:
    base = "|".join(
        [
            _sin_acentos(zona).upper(),
            clave_pozo_plana(ubicacion),
            " ".join(sorted(_tokens_actividad_programa(tipo))),
            cuando.isoformat() if cuando else "",
        ]
    )
    return hashlib.sha1(base.encode("utf-8", errors="ignore")).hexdigest()[:12].upper()


def _mapa_historial_programa(fila) -> dict[str, int] | None:
    roles: dict[str, int] = {}
    for indice, valor in enumerate(list(fila)):
        normal = _norm_encabezado(valor)
        if not normal:
            continue
        if normal in {"cod", "codigo"} or normal.startswith("codigo") or normal.startswith("cod "):
            roles.setdefault("codigo", indice)
        elif normal == "equipo" or normal.startswith("equipo "):
            roles.setdefault("equipo", indice)
        elif "actividad" in normal:
            roles.setdefault("actividad", indice)
        elif "tipo" in normal:
            roles.setdefault("tipo", indice)
        elif "fecha programada" in normal or normal == "fecha":
            roles.setdefault("fecha", indice)
    if ("codigo" in roles or "equipo" in roles) and "fecha" in roles and ("actividad" in roles or "tipo" in roles):
        return roles
    return None


def _filas_historial_operativo(tabla: pd.DataFrame, anio_defecto: int, zona: str) -> list[dict]:
    if tabla is None or tabla.empty:
        return []
    mejor = None
    for indice in range(min(len(tabla), 25)):
        mapa = _mapa_historial_programa(tabla.iloc[indice].tolist())
        if mapa:
            mejor = (indice, mapa)
            break
    if not mejor:
        return []
    fila_header, roles = mejor
    filas = []
    vistos = set()
    for indice in range(fila_header + 1, len(tabla)):
        renglon = tabla.iloc[indice].tolist()

        def celda(campo: str) -> str:
            columna = roles.get(campo)
            if columna is None or columna >= len(renglon):
                return ""
            return _texto_actividad(renglon[columna])

        codigo = celda("codigo") or celda("equipo")
        if not codigo or _norm_encabezado(codigo) in {"cod", "codigo", "equipo", "total"}:
            continue
        cuando = _fecha_programada(
            renglon[roles["fecha"]] if roles["fecha"] < len(renglon) else None,
            anio_defecto,
        )
        if cuando is None:
            continue
        ubicacion, equipo = _pozo_desde_codigo_programa(codigo)
        tipo = _tipo_programa_operativo(celda("tipo"), celda("actividad"), equipo)
        if not ubicacion or not tipo:
            continue
        item = {
            "ID": _id_programa_mantenimiento(ubicacion, tipo, cuando, zona),
            "ZONA": zona,
            "UBICACION": ubicacion,
            "CODIGO": codigo,
            "TIPO DE MANTENIMIENTO": tipo,
            "FECHA PROGRAMADA": cuando.isoformat(),
            "MES": _etiqueta_mes_programa(cuando),
        }
        if item["ID"] in vistos:
            continue
        vistos.add(item["ID"])
        filas.append(item)
    return filas


def _filas_tabla_programa(tabla: pd.DataFrame, anio_defecto: int, zona: str = "") -> list[dict]:
    if tabla is None or tabla.empty:
        return []
    mejor = None
    mejor_puntaje = 0
    for indice in range(min(len(tabla), 25)):
        roles: dict[str, int] = {}
        meses: dict[int, tuple[int, int]] = {}
        for columna, valor in enumerate(tabla.iloc[indice].tolist()):
            normal = _norm_encabezado(valor)
            rol = _rol_columna_programa(normal)
            if rol and rol not in roles:
                roles[rol] = columna
            mes = _mes_desde_texto(valor)
            if mes and rol not in {"ubicacion", "tipo", "fecha"}:
                meses[columna] = (_anio_en_texto(valor, anio_defecto), mes)
        puntaje = 0
        if "ubicacion" in roles:
            puntaje += 2
        if "tipo" in roles:
            puntaje += 2
        if "fecha" in roles:
            puntaje += 2
        if len(meses) >= 3:
            puntaje += 2
        if (
            puntaje > mejor_puntaje
            and "ubicacion" in roles
            and "tipo" in roles
            and ("fecha" in roles or len(meses) >= 3)
        ):
            mejor = (indice, roles, meses)
            mejor_puntaje = puntaje
    if not mejor:
        return []
    fila_header, roles, meses = mejor
    filas = []
    vistos = set()
    for indice in range(fila_header + 1, len(tabla)):
        renglon = tabla.iloc[indice].tolist()
        ubicacion = _texto_actividad(renglon[roles["ubicacion"]] if roles["ubicacion"] < len(renglon) else "")
        tipo = _texto_actividad(renglon[roles["tipo"]] if roles["tipo"] < len(renglon) else "")
        if not ubicacion or not tipo:
            continue
        if _norm_encabezado(ubicacion) in {"ubicacion", "pozo", "ubicacion pozo", "total"}:
            continue
        fechas: list[date | None] = []
        if "fecha" in roles and len(meses) < 3:
            fechas.append(
                _fecha_programada(
                    renglon[roles["fecha"]] if roles["fecha"] < len(renglon) else None,
                    anio_defecto,
                )
            )
        elif meses:
            for columna, (anio, mes) in meses.items():
                marca = _fecha_marca_mes(renglon[columna] if columna < len(renglon) else None, anio, mes)
                if marca:
                    fechas.append(marca)
        for cuando in fechas:
            if cuando is None:
                continue
            item = {
                "ID": _id_programa_mantenimiento(ubicacion, tipo, cuando, zona),
                "ZONA": zona,
                "UBICACION": ubicacion,
                "CODIGO": "",
                "TIPO DE MANTENIMIENTO": tipo,
                "FECHA PROGRAMADA": cuando.isoformat(),
                "MES": _etiqueta_mes_programa(cuando),
            }
            if item["ID"] in vistos:
                continue
            vistos.add(item["ID"])
            filas.append(item)
    return filas


def interpretar_programa_mantenimiento(contenido: bytes, nombre: str, zona: str = "") -> list[dict]:
    """Lee un Excel o CSV. Los libros UBM usan la hoja Historial; una tabla simple también sirve."""
    if not contenido:
        raise ValueError("El archivo está vacío.")
    if len(contenido) > 20 * 1024 * 1024:
        raise ValueError("El archivo supera 20 MB.")
    anio = _anio_en_texto(nombre, date.today().year)
    tablas: list[tuple[str, pd.DataFrame]] = []
    if str(nombre or "").lower().endswith(".csv"):
        ultimo = None
        for codec in ("utf-8-sig", "latin-1"):
            try:
                tablas = [("csv", pd.read_csv(io.BytesIO(contenido), header=None, dtype=object, encoding=codec))]
                break
            except Exception as exc:
                ultimo = exc
        if not tablas:
            raise ValueError(f"No se pudo leer el CSV: {ultimo}")
    else:
        try:
            libro = pd.ExcelFile(io.BytesIO(contenido))
        except Exception as exc:
            raise ValueError(f"No se pudo leer el Excel: {exc}") from exc
        nombres = list(libro.sheet_names)
        tiene_historial = any(_norm_encabezado(hoja) == "historial" for hoja in nombres)
        for hoja in nombres:
            if tiene_historial and _norm_encabezado(hoja) != "historial":
                continue
            try:
                tabla = pd.read_excel(libro, sheet_name=hoja, header=None, dtype=object)
            except Exception:
                continue
            if isinstance(tabla, pd.DataFrame):
                tablas.append((hoja, tabla))
    operativas: list[dict] = []
    simples: list[dict] = []
    for _nombre, tabla in tablas:
        encontradas = _filas_historial_operativo(tabla, anio, zona)
        if encontradas:
            operativas.extend(encontradas)
        else:
            simples.extend(_filas_tabla_programa(tabla, anio, zona))
    origen = operativas or simples
    filas: list[dict] = []
    vistos = set()
    for item in origen:
        if item["ID"] in vistos:
            continue
        vistos.add(item["ID"])
        filas.append(item)
        if len(filas) > 20000:
            raise ValueError("El programa tiene más de 20,000 actividades. Revisa el archivo.")
    if not filas:
        raise ValueError(
            "No encontré actividades con pozo o código, tipo de mantenimiento y fecha programada."
        )
    return filas


def guardar_programa_mantenimiento(filas: list[dict], ruta: str) -> int:
    df = pd.DataFrame(filas, columns=_COLUMNAS_PROGRAMA_ARCHIVO)
    os.makedirs(os.path.dirname(ruta) or DATABASE_DIR, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Programa", index=False)
        try:
            persistencia.copiar_seguro(tmp, ruta)
        except PermissionError:
            alt = os.path.join(os.path.dirname(ruta) or DATABASE_DIR, "~pjp4_programa_write.xlsx")
            persistencia.copiar_seguro(tmp, alt)
            persistencia.reemplazar_seguro(alt, ruta)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    cargar_programa_mantenimiento._cache = None
    return len(df)


def _destino_zona_programa(zona: str) -> tuple[str, str]:
    clave = str(zona or "").casefold()
    if clave == "norte":
        return "Norte", PROGRAMA_NORTE_FILE
    if clave == "sur":
        return "Sur", PROGRAMA_SUR_FILE
    raise ValueError("La zona tiene que ser Norte o Sur.")


def importar_programa_zona(zona: str, contenido: bytes, nombre: str) -> int:
    etiqueta, ruta = _destino_zona_programa(zona)
    filas = interpretar_programa_mantenimiento(contenido, nombre, etiqueta)
    return guardar_programa_mantenimiento(filas, ruta)


def _programa_vacio() -> pd.DataFrame:
    return pd.DataFrame(columns=_COLUMNAS_PROGRAMA_ARCHIVO + ["FECHA"])


def _leer_programa_guardado(ruta: str, zona_defecto: str) -> pd.DataFrame:
    try:
        crudo = pd.read_excel(ruta, sheet_name=0, dtype=object)
    except Exception:
        return _programa_vacio()
    if crudo is None or crudo.empty:
        return _programa_vacio()
    renombre = {}
    for col in crudo.columns:
        normal = _norm_encabezado(col)
        rol = _rol_columna_programa(normal)
        if rol == "ubicacion":
            renombre[col] = "UBICACION"
        elif rol == "tipo":
            renombre[col] = "TIPO DE MANTENIMIENTO"
        elif rol == "fecha" and "FECHA PROGRAMADA" not in renombre.values():
            renombre[col] = "FECHA PROGRAMADA"
        elif normal == "id":
            renombre[col] = "ID"
        elif normal == "mes":
            renombre[col] = "MES"
        elif normal == "zona":
            renombre[col] = "ZONA"
        elif normal in {"codigo", "cod"}:
            renombre[col] = "CODIGO"
    df = crudo.rename(columns=renombre)
    for col in _COLUMNAS_PROGRAMA_ARCHIVO:
        if col not in df.columns:
            df[col] = zona_defecto if col == "ZONA" else ""
    df["ZONA"] = df["ZONA"].map(_texto_actividad)
    if zona_defecto:
        df.loc[df["ZONA"] == "", "ZONA"] = zona_defecto
    df["UBICACION"] = df["UBICACION"].map(_texto_actividad)
    df["CODIGO"] = df["CODIGO"].map(_texto_actividad)
    df["TIPO DE MANTENIMIENTO"] = df["TIPO DE MANTENIMIENTO"].map(_texto_actividad)
    df["FECHA"] = df["FECHA PROGRAMADA"].map(lambda valor: _fecha_programada(valor, date.today().year))
    df["MES"] = [
        _texto_actividad(mes) or _etiqueta_mes_programa(cuando)
        for mes, cuando in zip(df["MES"], df["FECHA"])
    ]
    df["ID"] = [
        _texto_actividad(identificador)
        or _id_programa_mantenimiento(ubicacion, tipo, cuando, zona)
        for identificador, ubicacion, tipo, cuando, zona in zip(
            df["ID"], df["UBICACION"], df["TIPO DE MANTENIMIENTO"], df["FECHA"], df["ZONA"]
        )
    ]
    df = df[df["UBICACION"].astype(bool) & df["TIPO DE MANTENIMIENTO"].astype(bool)].copy()
    return df[_COLUMNAS_PROGRAMA_ARCHIVO + ["FECHA"]]


def _firmas_programa_maestro() -> tuple:
    firmas = []
    for _zona, ruta in _rutas_programa_maestro():
        if os.path.isfile(ruta):
            firmas.append((ruta, os.path.getmtime(ruta), os.path.getsize(ruta)))
    return tuple(firmas)


def _rutas_programa_maestro() -> list[tuple[str, str]]:
    zonas = [("Norte", PROGRAMA_NORTE_FILE), ("Sur", PROGRAMA_SUR_FILE)]
    if any(os.path.isfile(ruta) for _, ruta in zonas):
        return zonas
    return [("", PROGRAMA_MANTENIMIENTO_FILE)]


def cargar_programa_mantenimiento() -> pd.DataFrame:
    """Une Zona Norte y Zona Sur en un solo programa para las alertas y el avance."""
    firmas = _firmas_programa_maestro()
    if not firmas:
        return _programa_vacio()
    cache = getattr(cargar_programa_mantenimiento, "_cache", None)
    if cache and cache[0] == firmas:
        return cache[1].copy()
    partes = []
    for zona, ruta in _rutas_programa_maestro():
        if not os.path.isfile(ruta):
            continue
        parte = _leer_programa_guardado(ruta, zona)
        if not parte.empty:
            partes.append(parte)
    if not partes:
        df = _programa_vacio()
    else:
        df = pd.concat(partes, ignore_index=True)
        df = df.drop_duplicates(subset=["ID"], keep="first")
    cargar_programa_mantenimiento._cache = (firmas, df)
    return df.copy()


def _tokens_actividad_programa(texto: str) -> set[str]:
    plano = re.sub(r"[^a-z0-9]+", " ", _sin_acentos(texto).casefold())
    return {token for token in plano.split() if len(token) > 1 and token not in _TOKENS_ACTIVIDAD_RUIDO}


def actividades_programadas_coinciden(izquierda: str, derecha: str) -> bool:
    tokens_a = _tokens_actividad_programa(izquierda)
    tokens_b = _tokens_actividad_programa(derecha)
    if not tokens_a or not tokens_b:
        return False
    equipos_a = tokens_a & _TOKENS_EQUIPO_PROGRAMA
    equipos_b = tokens_b & _TOKENS_EQUIPO_PROGRAMA
    if equipos_a and equipos_b and not (equipos_a & equipos_b):
        return False
    if not (tokens_a <= tokens_b or tokens_b <= tokens_a):
        return False
    corto = tokens_a if len(tokens_a) <= len(tokens_b) else tokens_b
    return any(len(token) >= 4 for token in corto) or bool(corto & _TOKENS_EQUIPO_PROGRAMA)


def _claves_pozo_programa(*textos: str) -> set[str]:
    claves = set()
    for texto in textos:
        bruto = str(texto or "").strip()
        if not bruto or bruto.casefold() in {"nan", "none"}:
            continue
        for trozo in re.split(r"[·|;,/]+", bruto):
            trozo = trozo.strip()
            if not trozo:
                continue
            plana = clave_pozo_plana(trozo)
            if plana:
                claves.add(plana)
            try:
                largo = nombre_largo_excel(trozo)
            except Exception:
                largo = ""
            if largo:
                claves.add(clave_pozo_plana(largo))
            corto, _equipo = _pozo_desde_codigo_programa(trozo)
            if corto and corto != trozo:
                claves.add(clave_pozo_plana(corto))
                try:
                    largo_corto = nombre_largo_excel(corto)
                except Exception:
                    largo_corto = ""
                if largo_corto:
                    claves.add(clave_pozo_plana(largo_corto))
    return {clave for clave in claves if clave}


def _bitacora_para_programa() -> pd.DataFrame:
    ruta = EXCEL_FILE
    if not os.path.isfile(ruta):
        return cargar_bitacora()
    firma = (os.path.getmtime(ruta), os.path.getsize(ruta))
    cache = getattr(_bitacora_para_programa, "_cache", None)
    if cache and cache[0] == firma:
        return cache[1]
    df = cargar_bitacora()
    _bitacora_para_programa._cache = (firma, df)
    return df


def cruce_cumplimiento_programa(programa: pd.DataFrame, bitacora: pd.DataFrame) -> pd.DataFrame:
    """Cada actividad del programa queda Realizada cuando una bitácora con Programado = Sí la cubre."""
    columnas = [
        "ID",
        "ZONA",
        "UBICACION",
        "CODIGO",
        "TIPO DE MANTENIMIENTO",
        "FECHA",
        "MES",
        "ESTATUS",
        "FECHA REALIZADA",
    ]
    if programa is None or programa.empty:
        return pd.DataFrame(columns=columnas)
    confirmados = []
    if bitacora is not None and not bitacora.empty and COLUMNA_PROGRAMADO in bitacora.columns:
        for idx, fila in bitacora.iterrows():
            if not _marca_programado_si(fila.get(COLUMNA_PROGRAMADO)):
                continue
            textos = [
                fila.get("POZO"),
                fila.get("UBICACIÓN"),
                fila.get("EQUIPO/UBICACIÓN"),
                fila.get("SECTOR"),
            ]
            confirmados.append(
                {
                    "idx": idx,
                    "claves": _claves_pozo_programa(*[str(texto or "") for texto in textos]),
                    "actividad": str(fila.get("TIPO DE ACTIVIDAD") or ""),
                    "corta": str(fila.get("ACTIVIDAD CORTA") or ""),
                    "fecha": parsear_fecha_valor(fila.get("FECHA")),
                }
            )
    usados: set = set()
    ordenados = programa.copy()
    if "FECHA" not in ordenados.columns:
        ordenados["FECHA"] = ordenados["FECHA PROGRAMADA"].map(
            lambda valor: _fecha_programada(valor, date.today().year)
        )
    ordenados = ordenados.sort_values(by="FECHA", key=lambda serie: serie.map(lambda valor: valor or date.max))
    salida = []
    for _, fila in ordenados.iterrows():
        cuando = fila.get("FECHA")
        if not isinstance(cuando, date):
            cuando = _fecha_programada(fila.get("FECHA PROGRAMADA"), date.today().year)
        claves = _claves_pozo_programa(str(fila.get("UBICACION") or ""), str(fila.get("CODIGO") or ""))
        tipo = str(fila.get("TIPO DE MANTENIMIENTO") or "")
        candidatos = []
        for registro in confirmados:
            if registro["idx"] in usados:
                continue
            if not claves or not registro["claves"] or not (claves & registro["claves"]):
                continue
            if not (
                actividades_programadas_coinciden(tipo, registro["actividad"])
                or actividades_programadas_coinciden(tipo, registro["corta"])
            ):
                continue
            mismo_mes = bool(
                cuando
                and registro["fecha"]
                and cuando.year == registro["fecha"].year
                and cuando.month == registro["fecha"].month
            )
            candidatos.append((0 if mismo_mes else 1, registro["fecha"] or date.max, registro["idx"], registro["fecha"]))
        estatus = "Pendiente"
        fecha_realizada = ""
        if candidatos:
            candidatos.sort(key=lambda item: (item[0], item[1], str(item[2])))
            _, _, idx, fecha_real = candidatos[0]
            usados.add(idx)
            estatus = "Realizado"
            fecha_realizada = fecha_real.strftime("%d/%m/%Y") if fecha_real else ""
        salida.append(
            {
                "ID": fila.get("ID") or "",
                "ZONA": fila.get("ZONA") or "",
                "UBICACION": fila.get("UBICACION") or "",
                "CODIGO": fila.get("CODIGO") or "",
                "TIPO DE MANTENIMIENTO": tipo,
                "FECHA": cuando,
                "MES": fila.get("MES") or _etiqueta_mes_programa(cuando),
                "ESTATUS": estatus,
                "FECHA REALIZADA": fecha_realizada,
            }
        )
    return pd.DataFrame(salida, columns=columnas)


def coincidencias_pendientes_programa(
    ubicacion: str,
    actividad: str,
    dia: date,
    programa: pd.DataFrame,
    bitacora: pd.DataFrame,
    actividad_corta: str = "",
) -> list[dict]:
    if programa is None or programa.empty or not str(ubicacion or "").strip() or not str(actividad or "").strip():
        return []
    claves = _claves_pozo_programa(ubicacion)
    if not claves:
        return []
    cruce = cruce_cumplimiento_programa(programa, bitacora)
    pendientes = []
    for _, fila in cruce.iterrows():
        if fila.get("ESTATUS") != "Pendiente":
            continue
        if not (claves & _claves_pozo_programa(str(fila.get("UBICACION") or ""), str(fila.get("CODIGO") or ""))):
            continue
        tipo = str(fila.get("TIPO DE MANTENIMIENTO") or "")
        if not (
            actividades_programadas_coinciden(actividad, tipo)
            or actividades_programadas_coinciden(actividad_corta, tipo)
        ):
            continue
        cuando = fila.get("FECHA") if isinstance(fila.get("FECHA"), date) else None
        mismo_mes = bool(cuando and dia and cuando.year == dia.year and cuando.month == dia.month)
        pendientes.append(
            {
                "ubicacion": str(fila.get("UBICACION") or ""),
                "zona": str(fila.get("ZONA") or ""),
                "tipo": tipo,
                "mes": str(fila.get("MES") or _etiqueta_mes_programa(cuando)),
                "fecha": cuando,
                "mismo_mes": mismo_mes,
            }
        )
    pendientes.sort(key=lambda item: (0 if item["mismo_mes"] else 1, item["fecha"] or date.max))
    return pendientes


def pendientes_programa_campo(
    ubicacion: str,
    actividad: str,
    dia: date,
    actividad_corta: str = "",
) -> list[dict]:
    return coincidencias_pendientes_programa(
        ubicacion,
        actividad,
        dia,
        cargar_programa_mantenimiento(),
        _bitacora_para_programa(),
        actividad_corta,
    )


def construir_registro(
    fecha: date,
    secuencia: str,
    sector: str,
    pozo: str,
    cuadrilla: str,
    equipo: str,
    tipo_actividad: str,
    actividad_corta: str = "",
    insumos: dict[str, float | int | str] | None = None,
    descripcion: str = "",
    estatus: str = "",
    fotos: list[str] | None = None,
    cobertura: dict[str, str] | None = None,
    programado: str = "",
    porcentaje_avance: int | float = 100,
) -> dict:
    insumos = insumos or {}
    fotos = fotos or []
    rutas = list(fotos) + [""] * MAX_FOTOS
    pozo = "" if sitios_operativos.es_vacio_sitio(pozo) else str(pozo or "").strip()
    equipo = "" if sitios_operativos.es_vacio_sitio(equipo) else str(equipo or "").strip()
    equipo_catalogo = nombre_equipo_catalogo(equipo)
    sector = "" if sitios_operativos.es_vacio_sitio(sector) else str(sector or "").strip()
    tag_pozo = componer_tag_pozo(sector, pozo)
    if not tag_pozo and _PATRON_CODIGO_POZO.match(clave_pozo_plana(pozo)):
        tag_pozo = etiqueta_codigo_ubicacion(pozo)
    if es_codigo_infraestructura(pozo):
        partes_ubicacion = [pozo]
        sitio = pozo
    elif tag_pozo:
        pozo = tag_pozo
        zona, _n = descomponer_tag_pozo(tag_pozo)
        sector = zona or sector or tag_pozo
        partes_ubicacion = [tag_pozo]
        sitio = tag_pozo
    else:
        sitio = sitios_operativos.clave_sitio_campos(pozo, sector, equipo, descripcion)
        if sitio:
            pozo = sitio
            sector = sitio
            partes_ubicacion = [sitio]
            equipo = ""
        else:
            partes_ubicacion = []
            for parte in (sector, pozo):
                texto = str(parte or "").strip()
                if not texto or sitios_operativos.es_vacio_sitio(texto):
                    continue
                if texto not in partes_ubicacion:
                    partes_ubicacion.append(texto)
    registro = {
        "FECHA": fecha.strftime("%Y-%m-%d"),
        "SECUENCIA": secuencia,
        "SECTOR": sector,
        "POZO": pozo,
        "CUADRILLA": cuadrilla,
        "MECANICO TITULAR": str((cobertura or {}).get("mecanico_titular") or ""),
        "AYUDANTE TITULAR": str((cobertura or {}).get("ayudante_titular") or ""),
        "MECANICO EN TURNO": str((cobertura or {}).get("mecanico_real") or ""),
        "AYUDANTE EN TURNO": str((cobertura or {}).get("ayudante_real") or ""),
        "COBERTURA": str((cobertura or {}).get("cobertura") or "No"),
        "EQUIPO": equipo_catalogo,
        "EQUIPO/UBICACIÓN": equipo,
        "TIPO DE ACTIVIDAD": tipo_actividad.strip(),
        "ACTIVIDAD CORTA": actividad_corta.strip() or tipo_actividad.strip(),
        "DESCRIPCIÓN DE LA ACTIVIDAD": descripcion.strip(),
        "DESCRIPCION DE LA ACTIVIDAD": descripcion.strip(),
        "EVALUACIÓN DEL ESTATUS": estatus,
        "FOTOS": " | ".join(p for p in rutas[:MAX_FOTOS] if p),
        "REVISIÓN": "Pendiente",
        "OBSERVACIÓN REVISIÓN": "",
        "PERSONAL": str((cobertura or {}).get("codigo") or "").strip()
        or " · ".join(
            p
            for p in [
                str((cobertura or {}).get("mecanico_real") or (cobertura or {}).get("mecanico_titular") or ""),
                str((cobertura or {}).get("ayudante_real") or (cobertura or {}).get("ayudante_titular") or ""),
            ]
            if p
        ),
        "UBICACIÓN": sitio or sitios_operativos.texto_ubicacion_final(*partes_ubicacion),
        "COMENTARIOS": "",
        "AVANCE %": _porcentaje_avance_registro(porcentaje_avance),
        "ID_REGISTRO": "",
        "EVIDENCIA": EVIDENCIA_PENDIENTE,
        COLUMNA_PROGRAMADO: _normalizar_marca_programado(programado),
    }
    for col in COLUMNAS_INSUMOS + COLUMNAS_INSUMOS_NUEVOS:
        registro[col] = insumos.get(col, 0 if col in COLUMNAS_NUMERICAS else "")
    for origen, destino in MAPA_CONSUMIBLE_SOLEDAD.items():
        registro[destino] = insumos.get(origen, 0)
    bomba = insumos.get("BOMBA DE AGUA REPARADA PJP4")
    if bomba:
        registro["BOMBA DE AGUA REPARADA PJP4"] = bomba
    elif insumos.get("OTROS INSUMOS"):
        registro["BOMBA DE AGUA REPARADA PJP4"] = insumos.get("OTROS INSUMOS")
    for i, col in enumerate(COLUMNAS_FOTOS):
        registro[col] = rutas[i]
    registro["FOTOS"] = " | ".join(p for p in rutas[:MAX_FOTOS] if p)
    registro["ID_REGISTRO"] = id_registro_fila(pd.Series(registro))
    registro["EVIDENCIA"] = estatus_evidencia_fila(pd.Series(registro))
    return registro


def es_registro_demo(fila) -> bool:
    secuencia = str(fila.get("SECUENCIA") or "").strip()
    observacion = str(fila.get("OBSERVACIÓN REVISIÓN") or "")
    return secuencia.startswith(PREFIJO_SECUENCIA_DEMO) or MARCA_DEMO in observacion


def es_registro_prueba_manual(fila) -> bool:
    if es_registro_demo(fila):
        return False
    descripcion = str(fila.get("DESCRIPCIÓN DE LA ACTIVIDAD") or "")
    tipo = str(fila.get("TIPO DE ACTIVIDAD") or "").strip()
    cuadrilla = str(fila.get("CUADRILLA") or "").strip()
    sector = str(fila.get("SECTOR") or "").strip()
    if PATRON_REGISTRO_PRUEBA.search(descripcion):
        return True
    if cuadrilla in {"MECA-1", "MECA-2", "MECA-3", "MECA-4"}:
        return True
    if sector in {"Otro", "Zona piloto", "Sector San Andrés"}:
        return True
    if not tipo and cuadrilla in {"Electricos", "Mecanico 1", "Mecanico 2", "Mecanico 3", "Mecanico 4"}:
        return True
    return False


def depurar_bitacora(
    df: pd.DataFrame,
    *,
    quitar_pruebas: bool = True,
    quitar_demo: bool = False,
) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_EXCEL)
    vista = normalizar_dataframe(df)
    mask = pd.Series(True, index=vista.index)
    if quitar_pruebas:
        mask &= ~vista.apply(es_registro_prueba_manual, axis=1)
    if quitar_demo:
        mask &= ~vista.apply(es_registro_demo, axis=1)
    return normalizar_dataframe(vista.loc[mask].reset_index(drop=True))


def _sitios_demo() -> list[tuple[str, str, str]]:
    datos = leer_json(UBICACIONES_FILE, {})
    fichas = datos.get("fichas") if isinstance(datos, dict) else {}
    sitios: list[tuple[str, str, str]] = []
    omitir = {"aops", "base pjp4", "almacen ops (aops)"}
    for ficha in (fichas or {}).values():
        if not isinstance(ficha, dict):
            continue
        nombre = str(ficha.get("nombre") or "").strip()
        sector = str(ficha.get("sector") or "").strip()
        if not nombre or not sector or nombre.casefold() in omitir:
            continue
        equipo = str(ficha.get("equipo_mci") or ficha.get("equipo_ub") or "").strip()
        sitios.append((sector, nombre, equipo))
    if sitios:
        return sitios
    return [
        ("Soledad (S)", "PS-1", "MCI tipo Cummins"),
        ("Soledad Norte (SN)", "PSN-1", "MCI tipo Weichai"),
        ("Aragon (A)", "PA-3", "UBM tipo Lufkin"),
        ("Gallo (GA)", "PGA-1", "MCI tipo Arrow"),
        ("Coyotes (Coy)", "PC-1", "UBM tipo Weatherford"),
    ]


def _actividades_grupo_demo(cuadrilla: str) -> list[dict]:
    grupo = grupo_actividades_cuadrilla(cuadrilla)
    catalogo = leer_json(ACTIVIDADES_FILE, CATALOGO_ACTIVIDADES_DEFAULT)
    if not isinstance(catalogo, dict):
        catalogo = CATALOGO_ACTIVIDADES_DEFAULT
    items = catalogo.get(grupo) if grupo else None
    if not isinstance(items, list) or not items:
        items = CATALOGO_ACTIVIDADES_DEFAULT.get(grupo or "MECANICOS", [])
    limpios = []
    for item in items:
        if isinstance(item, dict) and item.get("nombre") and actividad_esta_activa(item):
            limpios.append({"nombre": str(item["nombre"]), "corto": str(item.get("corto") or item["nombre"])})
    return limpios or [{"nombre": "Revisión a MCI", "corto": "REV MCI"}]


def _personal_demo(cuadrilla: str) -> dict[str, str]:
    datos = leer_json(PERSONAL_FILE, {})
    posiciones = datos.get("posiciones") if isinstance(datos, dict) else {}
    item = posicion_personal_cuadrilla(cuadrilla)
    titular_m = item["mecanico"] or f"Titular {cuadrilla}"
    titular_a = " · ".join(p for p in [item["ayudante"], item["apoyo"]] if p) or f"Ayudante {cuadrilla}"
    return {
        "mecanico_titular": titular_m,
        "ayudante_titular": titular_a,
        "mecanico_real": titular_m,
        "ayudante_real": titular_a,
        "cobertura": "No",
        "codigo": codigo_dotacion_cuadrilla(item),
    }


def _equipo_demo(sitio: tuple[str, str, str], actividad: str, rng: random.Random) -> str:
    _, pozo, equipo_ficha = sitio
    if equipo_ficha:
        return equipo_ficha
    if "UBM" in actividad.upper() or "UBM" in actividad:
        return rng.choice(EQUIPOS_UB_DEFAULT)
    if "ICI" in actividad.upper():
        return "ICI GR-20"
    return rng.choice(EQUIPOS_MCI_DEFAULT + [f"MCI {pozo}"])


def _descripcion_demo(actividad: str, pozo: str, estatus: str, rng: random.Random) -> str:
    cierres = {
        "Operativo": "El equipo queda operativo y se libera el pozo.",
        "En mantenimiento": "Queda pendiente una segunda intervención para concluir el servicio.",
        "Fuera de servicio": "Se deja el equipo fuera de servicio a la espera de refacción.",
    }
    plantillas = [
        f"Se ejecutó {actividad} en {pozo or 'la instalación'}. {cierres[estatus]}",
        f"Cuadrilla interviene {pozo or 'el sitio'}: {actividad}. {cierres[estatus]}",
        f"{actividad} conforme a programa OPS. Ubicación {pozo or 'campo'}. {cierres[estatus]}",
    ]
    return rng.choice(plantillas)


def _insumos_demo(actividad: str, rng: random.Random) -> dict:
    texto = actividad.casefold()
    insumos: dict = {}
    if "bimestral" in texto or "semestral" in texto or "revisión a mci" in texto:
        insumos["ACEITE (LTS)"] = rng.choice([8, 12, 16, 20])
        insumos["FILTROS (PZAS)"] = rng.choice([1, 2])
        if rng.random() < 0.4:
            insumos["GRASA (KG)"] = rng.choice([1, 2])
    if "habilitado" in texto or "restablecimiento" in texto:
        insumos["ACEITE (LTS)"] = rng.choice([4, 8])
        insumos["REFRIGERANTE (LTS)"] = rng.choice([2, 4])
    if "soldadura" in texto or "corte" in texto:
        insumos["SOLDADURA/ELECTRODOS (KG)"] = rng.choice([1, 2, 3])
        insumos["DISCOS DE CORTE/DESBASTE (PZAS)"] = rng.choice([2, 4])
    if "eléctrico" in texto or "tablero" in texto or "cable" in texto:
        insumos["CINTA AISLANTE (ROLLOS)"] = 1
        insumos["CABLE (MTS)"] = rng.choice([5, 10, 15])
    return insumos


def _ots_por_dia_demo(dia: date) -> int:
    if dia.weekday() == 6:
        return 0
    if dia.weekday() == 5:
        return 1 if dia.day % 2 == 0 else 0
    semana = min((dia.day - 1) // 7, 4)
    return [2, 2, 3, 2, 2][semana]


def _revision_demo(dia: date, rng: random.Random) -> str:
    if dia >= date(dia.year, 9, 22):
        pesos = [0.52, 0.28, 0.20]
    elif dia >= date(dia.year, 9, 12):
        pesos = [0.22, 0.58, 0.20]
    else:
        pesos = [0.10, 0.75, 0.15]
    return rng.choices(["Pendiente", "Validado", "Con observaciones"], pesos)[0]


def construir_ots_demo_septiembre(anio: int = 2026, mes: int = 9) -> list[dict]:
    rng = random.Random(anio * 100 + mes)
    sitios = _sitios_demo()
    cuadrillas_mec = ["Mecanico 1", "Mecanico 2", "Mecanico 3", "Mecanico 4"]
    cuadrillas_apoyo = ["Electricos", "Instrumentos", "Montaje", "Soldadura"]
    ultimo = date(anio, mes, calendar.monthrange(anio, mes)[1])
    registros: list[dict] = []
    folio = 1
    turno_mec = 0
    turno_apoyo = 0
    cursor = date(anio, mes, 1)
    while cursor <= ultimo:
        for _ in range(_ots_por_dia_demo(cursor)):
            if rng.random() < 0.74:
                cuadrilla = cuadrillas_mec[turno_mec % len(cuadrillas_mec)]
                turno_mec += 1
            else:
                cuadrilla = cuadrillas_apoyo[turno_apoyo % len(cuadrillas_apoyo)]
                turno_apoyo += 1
            actividad = rng.choice(_actividades_grupo_demo(cuadrilla))
            sitio = rng.choice(sitios)
            sector, pozo, _equipo_f = sitio
            revision = _revision_demo(cursor, rng)
            if revision == "Con observaciones":
                estatus = rng.choice(["En mantenimiento", "Fuera de servicio", "Operativo"])
            elif revision == "Pendiente":
                estatus = rng.choice(["Operativo", "En mantenimiento"])
            else:
                estatus = "Operativo"
            cobertura = _personal_demo(cuadrilla)
            if rng.random() < 0.16:
                relevo = rng.choice(cuadrillas_mec)
                otro = _personal_demo(relevo)
                cobertura["mecanico_real"] = otro["mecanico_titular"] or cobertura["mecanico_titular"]
                cobertura["ayudante_real"] = otro["ayudante_titular"] or cobertura["ayudante_titular"]
                cobertura["cobertura"] = "Sí"
            registro = construir_registro(
                fecha=cursor,
                secuencia=f"{PREFIJO_SECUENCIA_DEMO}{folio:03d}",
                sector=sector,
                pozo=pozo,
                cuadrilla=cuadrilla,
                equipo=_equipo_demo(sitio, actividad["nombre"], rng),
                tipo_actividad=actividad["nombre"],
                actividad_corta=actividad["corto"],
                insumos=_insumos_demo(actividad["nombre"], rng),
                descripcion=_descripcion_demo(actividad["nombre"], pozo, estatus, rng),
                estatus=estatus,
                cobertura=cobertura,
            )
            registro["REVISIÓN"] = revision
            if revision == "Validado":
                registro["OBSERVACIÓN REVISIÓN"] = f"{MARCA_DEMO} · Cerrada / validada"
            elif revision == "Con observaciones":
                registro["OBSERVACIÓN REVISIÓN"] = f"{MARCA_DEMO} · Retraso: evidencia o retrabajo pendiente"
            else:
                registro["OBSERVACIÓN REVISIÓN"] = f"{MARCA_DEMO} · En proceso / pendiente de cierre"
            registros.append(registro)
            folio += 1
        cursor += timedelta(days=1)
    return registros


def construir_programados_demo(mes_clave: str = "2026-09") -> list[dict]:
    filas = [
        ("Mantenimiento bimestral a MCI", "Soledad (S)", "MCI tipo Cummins", "Cumplido", ""),
        ("Mantenimiento bimestral a UBM", "Gallo (GA)", "UBM tipo Lufkin", "Cumplido", ""),
        ("Mantenimiento semestral a MCI", "Aragon (A)", "MCI tipo Weichai", "Cumplido", ""),
        ("Mantenimiento semestral a UBM", "Palo Blanco (PB)", "UBM tipo Weatherford", "Cumplido", ""),
        ("Revisión a MCI", "Coyotes (Coy)", "MCI tipo Arrow", "Cumplido", ""),
        ("Revisión a ICI- GR-20", "Soledad Norte (SN)", "ICI GR-20", "Cumplido", ""),
        ("Habilitado de MCI y UBM", "Soledad (S)", "MCI tipo Cummins", "Cumplido", ""),
        ("Mantenimiento bimestral a MCI", "Gallo (GA)", "MCI tipo Vortec V8", "Cumplido", ""),
        ("Restablecimiento a MCI", "Aragon (A)", "MCI tipo Arrow", "No Cumplido", "Espera de refacción"),
        ("Mantenimiento bimestral a UBM", "Coyotes (Coy)", "UBM tipo Lufkin", "Reprogramado", "Condición climática"),
    ]
    return [
        {
            "actividad": act,
            "ubicacion": ubi,
            "equipo": eq,
            "estatus": est,
            "observacion": obs,
            "mes": mes_clave,
            "id": f"demo-{idx:03d}",
            "demo": True,
        }
        for idx, (act, ubi, eq, est, obs) in enumerate(filas, start=1)
    ]


def construir_ots_flota_demo() -> list[dict]:
    return [
        {
            "id": "DEMO-VEH-001",
            "tipo": "preventivo",
            "unidad": "U-560",
            "estatus": ESTATUS_OT_CERRADA,
            "fecha_creacion": "2026-09-04",
            "fecha_estimada": "2026-09-08",
            "fecha_cierre": "2026-09-08",
            "descripcion": "Mantenimiento preventivo de 10,000 km cerrado en taller.",
            "origen": "demo",
        },
        {
            "id": "DEMO-VEH-002",
            "tipo": "correctivo",
            "unidad": "U-614",
            "estatus": ESTATUS_OT_CERRADA,
            "fecha_creacion": "2026-09-12",
            "fecha_estimada": "2026-09-16",
            "fecha_cierre": "2026-09-15",
            "descripcion": "Correctivo por falla eléctrica; unidad liberada.",
            "origen": "demo",
        },
        {
            "id": "DEMO-VEH-003",
            "tipo": "preventivo",
            "unidad": "U-501",
            "estatus": ESTATUS_OT_ABIERTA,
            "fecha_creacion": "2026-09-10",
            "fecha_estimada": "2026-09-18",
            "fecha_cierre": "",
            "descripcion": "Preventivo vencido: pendiente de ingreso a taller.",
            "origen": "demo",
        },
        {
            "id": "DEMO-VEH-004",
            "tipo": "correctivo",
            "unidad": "U-503",
            "estatus": ESTATUS_OT_PROCESO,
            "fecha_creacion": "2026-09-22",
            "fecha_estimada": "2026-09-30",
            "fecha_cierre": "",
            "descripcion": "Correctivo en proceso: espera de refacción de suspensión.",
            "origen": "demo",
        },
        {
            "id": "DEMO-VEH-005",
            "tipo": "preventivo",
            "unidad": "U-612",
            "estatus": ESTATUS_OT_ABIERTA,
            "fecha_creacion": "2026-09-24",
            "fecha_estimada": "2026-10-04",
            "fecha_cierre": "",
            "descripcion": "Alerta de kilometraje: OT preventiva abierta.",
            "origen": "demo",
        },
    ]


def construir_ingresos_taller_demo() -> list[dict]:
    return [
        {
            "id": "DEMO-ING-001",
            "unidad": "U-503",
            "placas": "S-503-PJP",
            "fecha_entrada": "2026-09-22",
            "fecha_salida": "",
            "km_entrada": 86420,
            "km_salida": 0,
            "horometro_entrada": 0,
            "horometro_salida": 0,
            "responsable": "Mecanico 2",
            "falla": "Ruido en suspensión delantera y fuga en amortiguador. Unidad inhabilitada en base.",
            "tipo": TIPO_TALLER_CORRECTIVO,
            "semaforo": SEMAFORO_DETENIDA,
            "fecha_programacion": "2026-09-30",
            "horas_estimadas": 24,
            "mecanico": "Mecanico 2",
            "estatus": ESTATUS_INGRESO_ESPERA,
            "ot_id": "DEMO-VEH-004",
            "refacciones": [{"nombre": "Amortiguador delantero", "cantidad": 2, "en_almacen": False}],
            "solicitud": {
                "id": "DEMO-SOL-001",
                "estatus": SOLICITUD_PENDIENTE,
                "fecha": "2026-09-22",
                "piezas": [{"nombre": "Amortiguador delantero", "cantidad": 2}],
            },
            "externo": {},
            "evidencias": {"antes": [], "durante": [], "despues": []},
            "origen": "demo",
        },
        {
            "id": "DEMO-ING-002",
            "unidad": "U-501",
            "placas": "S-501-PJP",
            "fecha_entrada": "2026-09-26",
            "fecha_salida": "",
            "km_entrada": 120350,
            "km_salida": 0,
            "horometro_entrada": 0,
            "horometro_salida": 0,
            "responsable": "Mecanico 1",
            "falla": "Preventivo de 10,000 km. Filtros y aceite disponibles en almacén.",
            "tipo": TIPO_TALLER_PREVENTIVO,
            "semaforo": SEMAFORO_PRECAUCION,
            "fecha_programacion": "2026-09-29",
            "horas_estimadas": 8,
            "mecanico": "Mecanico 1",
            "estatus": ESTATUS_INGRESO_TALLER,
            "ot_id": "DEMO-VEH-003",
            "refacciones": [{"nombre": "Aceite SAE 50", "cantidad": 8, "en_almacen": True}],
            "solicitud": {},
            "externo": {},
            "evidencias": {"antes": [], "durante": [], "despues": []},
            "origen": "demo",
        },
        {
            "id": "DEMO-ING-003",
            "unidad": "U-612",
            "placas": "S-612-PJP",
            "fecha_entrada": "2026-09-18",
            "fecha_salida": "",
            "km_entrada": 97880,
            "km_salida": 0,
            "horometro_entrada": 4120,
            "horometro_salida": 0,
            "responsable": "Mecanico 3",
            "falla": "Falla de transmisión que rebasa capacidad del taller local.",
            "tipo": TIPO_TALLER_EXTERNO,
            "semaforo": SEMAFORO_EXTERNO,
            "fecha_programacion": "2026-09-20",
            "horas_estimadas": 72,
            "mecanico": "Mecanico 3",
            "estatus": ESTATUS_INGRESO_EXTERNO,
            "ot_id": "DEMO-VEH-005",
            "refacciones": [],
            "solicitud": {},
            "externo": {
                "motivo": "Requiere banco de transmisiones y equipo de diagnóstico especializado.",
                "taller": "Taller Diesel del Golfo",
                "cotizacion": "COT-612-2026-04",
                "autorizacion": "Pendiente de autorización",
                "fecha_envio": "2026-09-19",
                "fecha_retorno": "",
            },
            "evidencias": {"antes": [], "durante": [], "despues": []},
            "origen": "demo",
        },
    ]


def construir_asignaciones_demo() -> list[dict]:
    return [
        {
            "id": "DEMO-ASG-001",
            "unidad": "U-560",
            "proposito": "Unidad de relevo (principal en taller)",
            "cuadrilla": "Mecanico 1",
            "responsable": "Mecanico 1",
            "fecha_salida": "2026-09-27",
            "km_salida": 278980,
            "combustible_salida": "Lleno",
            "fluidos_salida": "Correctos",
            "obs_salida": "Unidad de relevo entregada en base para cubrir jornada.",
            "fotos_salida": [],
            "fecha_cierre": "2026-09-27",
            "km_cierre": 279315,
            "combustible_cierre": "1/2",
            "fluidos_cierre": "Correctos",
            "obs_cierre": "Devuelta sin novedad. 335 km de jornada.",
            "fotos_cierre": [],
            "km_recorridos": 335,
            "estatus": ESTATUS_ASIG_CERRADA,
            "origen": "demo",
        },
        {
            "id": "DEMO-ASG-002",
            "unidad": "U-614",
            "proposito": "Jornada de campo",
            "cuadrilla": "Mecanico 4",
            "responsable": "Mecanico 4",
            "fecha_salida": "2026-09-28",
            "km_salida": 182930,
            "combustible_salida": "3/4",
            "fluidos_salida": "Correctos",
            "obs_salida": "Salida a sector Soledad. Check-in de mañana.",
            "fotos_salida": [],
            "fecha_cierre": "",
            "km_cierre": 0,
            "combustible_cierre": "",
            "fluidos_cierre": "",
            "obs_cierre": "",
            "fotos_cierre": [],
            "km_recorridos": 0,
            "estatus": ESTATUS_ASIG_RUTA,
            "origen": "demo",
        },
    ]


def bitacora_requiere_demo(df: pd.DataFrame | None = None) -> bool:
    df = cargar_bitacora() if df is None else df
    limpia = depurar_bitacora(df, quitar_pruebas=True, quitar_demo=False)
    return len(limpia) < UMBRAL_REGISTROS_DEMO


def aplicar_datos_demo() -> dict[str, int]:
    bitacora = depurar_bitacora(cargar_bitacora(), quitar_pruebas=True, quitar_demo=True)
    ots = construir_ots_demo_septiembre()
    combinada = pd.concat([bitacora, pd.DataFrame(ots)], ignore_index=True) if ots else bitacora
    guardar_en_excel(combinada)

    programados = [item for item in cargar_programados_ops() if not item.get("demo")]
    demo_prog = construir_programados_demo()
    guardar_programados_ops(programados + demo_prog)

    veh = cargar_vehiculos()
    veh["ots"] = [ot for ot in (veh.get("ots") or []) if str(ot.get("origen") or "") != "demo"]
    veh["ots"].extend(construir_ots_flota_demo())
    veh["ingresos"] = [ing for ing in (veh.get("ingresos") or []) if str(ing.get("origen") or "") != "demo"]
    veh["ingresos"].extend(construir_ingresos_taller_demo())
    veh["asignaciones"] = [a for a in (veh.get("asignaciones") or []) if str(a.get("origen") or "") != "demo"]
    veh["asignaciones"].extend(construir_asignaciones_demo())
    guardar_vehiculos(veh)

    return {
        "ots": len(ots),
        "programados": len(demo_prog),
        "flota": len(construir_ots_flota_demo()),
        "reales": int(len(bitacora)),
    }


def quitar_datos_demo() -> dict[str, int]:
    bitacora = depurar_bitacora(cargar_bitacora(), quitar_pruebas=True, quitar_demo=True)
    guardar_en_excel(bitacora)
    programados = [item for item in cargar_programados_ops() if not item.get("demo")]
    guardar_programados_ops(programados)
    veh = cargar_vehiculos()
    veh["ots"] = [ot for ot in (veh.get("ots") or []) if str(ot.get("origen") or "") != "demo"]
    veh["ingresos"] = [ing for ing in (veh.get("ingresos") or []) if str(ing.get("origen") or "") != "demo"]
    veh["asignaciones"] = [a for a in (veh.get("asignaciones") or []) if str(a.get("origen") or "") != "demo"]
    guardar_vehiculos(veh)
    return {"ots": int(len(bitacora))}


def render_bloque_modo_demo() -> None:
    st.markdown("**Modo demo**")
    st.caption("OT simuladas de septiembre 2026 para ver el dashboard con métricas reales de prueba.")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Cargar demo", key="btn_demo_cargar", use_container_width=True):
            info = aplicar_datos_demo()
            st.session_state.demo_retirado = False
            st.session_state.demo_mensaje = (
                f"Demo lista: {info['ots']} OT de campo, {info['programados']} programados y {info['flota']} OT de flota."
            )
            st.rerun()
    with c2:
        if st.button("Quitar demo", key="btn_demo_quitar", use_container_width=True):
            quitar_datos_demo()
            st.session_state.demo_retirado = True
            st.session_state.demo_mensaje = "Se retiraron los registros de demostración."
            st.rerun()
    if st.session_state.get("demo_mensaje"):
        st.caption(st.session_state.demo_mensaje)


def asegurar_datos_demo_si_faltan() -> bool:
    if st.session_state.get("demo_retirado"):
        return False
    if not bitacora_requiere_demo():
        return False
    aplicar_datos_demo()
    st.session_state.demo_auto = True
    return True


def render_encabezado(titulo: str, subtitulo: str) -> None:
    st.markdown(
        f"""
        <div class="panel-header">
            <h1>{titulo}</h1>
            <p>{subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ── Selector de rol ────────────────────────────────────────────────────────────
CLAVES_POR_ROL = {
    ROL_OPERATIVO: "OPERATIVO_PASSWORD",
    ROL_ADMIN: "ADMIN_PASSWORD",
}


def aplicar_rol_pendiente() -> None:
    if "rol_pendiente" in st.session_state:
        st.session_state.rol_seleccionado = st.session_state.pop("rol_pendiente")


def rol_esta_desbloqueado(rol: str) -> bool:
    desbloqueados = st.session_state.get("roles_desbloqueados") or {}
    return bool(desbloqueados.get(rol))


def marcar_rol_desbloqueado(rol: str) -> None:
    desbloqueados = dict(st.session_state.get("roles_desbloqueados") or {})
    desbloqueados[rol] = True
    st.session_state.roles_desbloqueados = desbloqueados


def aplicar_cierre_de_rol() -> None:
    cerrar = st.session_state.pop("cerrar_sesion_rol", False)
    volver = st.session_state.pop("volver_inicio", False)
    if not cerrar and not volver:
        return
    if cerrar:
        st.session_state.roles_desbloqueados = {}
    st.session_state.rol_seleccionado = None
    st.session_state.pop("rol_pendiente", None)


def leer_secreto(nombre: str) -> str | None:
    try:
        valor = str(st.secrets[nombre]).strip()
    except Exception:
        return None
    return valor or None


def render_acceso_rol(rol: str) -> None:
    nombre_secreto = CLAVES_POR_ROL[rol]
    render_encabezado(rol, "Ingresa la clave de este módulo para continuar.")
    _, centro, _ = st.columns([1, 1.15, 1])
    with centro:
        with st.form(f"acceso_rol_{nombre_secreto}"):
            clave = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Entrar")
        if entrar:
            esperada = leer_secreto(nombre_secreto)
            if esperada is None:
                st.error(
                    f"Falta el secreto {nombre_secreto}. "
                    "Configúralo en Streamlit Cloud o en .streamlit/secrets.toml."
                )
            elif hmac.compare_digest(clave, esperada):
                marcar_rol_desbloqueado(rol)
                st.rerun()
            else:
                st.error("Contraseña incorrecta.")
        if st.button("Volver", key=f"volver_{nombre_secreto}"):
            st.session_state.volver_inicio = True
            st.rerun()


def render_selector_rol() -> str | None:
    aplicar_cierre_de_rol()
    aplicar_rol_pendiente()
    with st.sidebar:
        st.markdown("### Bitácora PJP4")
        st.caption("Selecciona el módulo o rol para continuar.")
        rol = st.radio(
            "Módulo / rol",
            DESTINOS,
            index=None,
            key="rol_seleccionado",
        )
        if rol in CLAVES_POR_ROL and not rol_esta_desbloqueado(rol):
            st.caption(f"Pendiente de clave: **{rol}**")
        elif rol:
            st.caption(f"Sesión: **{rol}**")
        if any(rol_esta_desbloqueado(item) for item in CLAVES_POR_ROL):
            if st.button("Cerrar sesión de rol", use_container_width=True):
                st.session_state.cerrar_sesion_rol = True
                st.rerun()
        st.markdown("---")
        render_bloque_modo_demo()
        st.markdown("---")
        st.caption("Operativo y Administrativo piden su propia clave.")
    return rol


def render_bienvenida() -> None:
    render_encabezado(
        "Bitácora de mantenimiento — PJP4",
        "Selecciona un módulo para cargar la interfaz correspondiente.",
    )
    col_op, col_ad = st.columns(2, gap="medium")
    with col_op:
        st.markdown(
            """
            <div class="rol-card">
                <h3>🚜 Operativo (Campo)</h3>
                <p>Captura de bitácora, ubicación, insumos críticos y evidencia fotográfica.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Entrar como Operativo (Campo)", type="primary", use_container_width=True):
            st.session_state.rol_pendiente = ROL_OPERATIVO
            st.rerun()
    with col_ad:
        st.markdown(
            """
            <div class="rol-card">
                <h3>🗂️ Administrativo</h3>
                <p>Revisión de bitácoras, catálogos de equipos y control de permisos.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Entrar como Administrativo", use_container_width=True):
            st.session_state.rol_pendiente = ROL_ADMIN
            st.rerun()
    col_ub, col_hc = st.columns(2, gap="medium")
    with col_ub:
        st.markdown(
            """
            <div class="rol-card">
                <h3>📍 Ubicación y Equipos</h3>
                <p>Ficha técnica por pozo: MCI, UB, historial, programados y personal que intervino.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Abrir Ubicación y Equipos", use_container_width=True):
            st.session_state.rol_pendiente = MODULO_UBICACIONES
            st.rerun()
    with col_hc:
        st.markdown(
            """
            <div class="rol-card">
                <h3>👷 Historial de Cuadrillas</h3>
                <p>Carpeta por cuadrilla: fechas, OT, personal, materiales y evidencia fotográfica.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Abrir Historial de Cuadrillas", use_container_width=True):
            st.session_state.rol_pendiente = MODULO_CUADRILLAS
            st.rerun()


def render_tarjeta_asignaciones_dia(fecha: date) -> None:
    asig = asignacion_del_dia(fecha)
    fecha_txt = fecha.strftime("%d/%m/%Y")
    actualizado = asig.get("actualizado") or ""
    meta = f"Actividades de mantenimiento para el {fecha_txt}."
    if actualizado:
        meta += f" Última actualización: {actualizado}."

    banner = html_salida_dia(fecha)
    if banner:
        st.markdown(banner, unsafe_allow_html=True)

    if not hay_asignacion_publicada(asig):
        st.markdown(
            f"""
            <div class="asignacion-board">
                <h2>📋 Actividades de mantenimiento diarias</h2>
                <p class="asignacion-meta">{html.escape(meta)}</p>
                <p class="asignacion-vacia">Aún no hay asignaciones publicadas para este día.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    salida = salida_visible_dia(fecha)
    bloques = []
    for cuadrilla in CUADRILLAS_ASIGNACION:
        pendiente = str(asig["cuadrillas"].get(cuadrilla, "")).strip()
        estatus = salida.get(cuadrilla, "")
        if pendiente:
            cuerpo = texto_html_asignacion(pendiente)
            extra = ""
        else:
            cuerpo = "Sin pendientes publicados."
            extra = " asignacion-vacia"
        tono = " en-base" if estatus == ESTATUS_EN_BASE else " en-campo" if estatus == ESTATUS_SALE_CAMPO else ""
        bloques.append(
            f'<div class="asignacion-item{tono}">'
            f'<div class="cuadrilla">{html.escape(cuadrilla)}</div>'
            f"{html_marca_salida(estatus)}"
            f'<p class="{extra.strip()}">{cuerpo}</p>'
            f"</div>"
        )

    ausencias = str(asig.get("ausencias", "")).strip()
    if ausencias:
        bloques.append(
            f'<div class="asignacion-ausencias">'
            f'<div class="cuadrilla">Ausencias / permisos</div>'
            f"<p>{texto_html_asignacion(ausencias)}</p>"
            f"</div>"
        )

    st.markdown(
        f"""
        <div class="asignacion-board">
            <h2>📋 Actividades de mantenimiento diarias</h2>
            <p class="asignacion-meta">{html.escape(meta)}</p>
            {''.join(bloques)}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_editor_asignaciones_diarias() -> None:
    st.markdown(
        '<div class="section-title">Editor de asignaciones diarias</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Publica los pendientes por cuadrilla y si ese día sale a campo o se queda en base. "
        "Al guardar, el calendario de asistencia marca F/O cuando no sale, y el personal de campo "
        "ve el estatus al consultar el día."
    )
    fecha_asig = st.date_input(
        "Fecha de la asignación",
        value=date.today(),
        format="DD/MM/YYYY",
        key="fecha_asignacion_admin",
    )
    actual = asignacion_del_dia(fecha_asig)
    clave_fecha = fecha_asig.strftime("%Y-%m-%d")

    if st.session_state.pop("asignacion_publicada", False):
        st.success(
            f"Asignaciones del {fecha_asig.strftime('%d/%m/%Y')} guardadas en {ASIGNACIONES_FILE}."
        )
    aviso_cal = st.session_state.pop("asignacion_calendario_aviso", "")
    ok_cal = st.session_state.pop("asignacion_calendario_ok", "")
    if ok_cal:
        st.success(ok_cal)
    if aviso_cal:
        st.warning(aviso_cal)

    pendientes: dict[str, str] = {}
    salidas: dict[str, str] = {}
    col_izq, col_der = st.columns(2, gap="medium")
    mitad = (len(CUADRILLAS_ASIGNACION) + 1) // 2
    for i, cuadrilla in enumerate(CUADRILLAS_ASIGNACION):
        key = f"asig_{clave_fecha}_{cuadrilla}"
        key_salida = f"salida_{clave_fecha}_{cuadrilla}"
        if key not in st.session_state:
            st.session_state[key] = actual["cuadrillas"].get(cuadrilla, "")
        if key_salida not in st.session_state:
            st.session_state[key_salida] = estatus_salida_sugerido(
                fecha_asig,
                cuadrilla,
                str((actual.get("salida") or {}).get(cuadrilla, "") or ""),
            )
        destino = col_izq if i < mitad else col_der
        etiqueta_asig = etiqueta_cuadrilla_historial(cuadrilla)
        with destino:
            st.markdown(f"**{etiqueta_asig}**")
            salidas[cuadrilla] = st.radio(
                f"Estatus de salida — {etiqueta_asig}",
                ESTATUS_SALIDA_CUADRILLA,
                horizontal=True,
                key=key_salida,
            )
            pendientes[cuadrilla] = st.text_area(
                f"Pendientes — {etiqueta_asig}",
                height=120,
                placeholder="Escribe las actividades asignadas a esta cuadrilla…",
                key=key,
            )

    key_aus = f"asig_{clave_fecha}_ausencias"
    if key_aus not in st.session_state:
        st.session_state[key_aus] = actual.get("ausencias", "")
    ausencias = st.text_area(
        "Notas de ausencias o permisos",
        height=110,
        placeholder="Ej. Mecánico 4: permiso. Instrumentos: incapacidad.",
        key=key_aus,
    )

    if actual.get("actualizado"):
        st.caption(f"Última publicación de este día: {actual['actualizado']}")

    _, col_btn, _ = st.columns([1, 2, 1])
    with col_btn:
        guardar = st.button(
            "💾 Publicar asignaciones del día",
            type="primary",
            use_container_width=True,
        )
    if guardar:
        guardar_asignacion_dia(fecha_asig, pendientes, ausencias, salidas)
        resultado_cal = aplicar_salida_calendario(fecha_asig, salidas)
        st.session_state.asignacion_publicada = True
        if resultado_cal.get("ok"):
            st.session_state.asignacion_calendario_ok = resultado_cal.get("mensaje") or ""
        else:
            st.session_state.asignacion_calendario_aviso = resultado_cal.get("mensaje") or ""
        st.rerun()


# ── Panel operativo ────────────────────────────────────────────────────────────
MODULOS_OPERATIVO = [
    {
        "id": "captura",
        "icono": "📝",
        "titulo": "Captura de Actividades de Campo",
        "descripcion": "Llena órdenes de trabajo, desglose de materiales y evidencias fotográficas.",
    },
    {
        "id": "evidencia_cuadrilla",
        "icono": "📷",
        "titulo": "Evidencia fotográfica de cuadrilla",
        "descripcion": "Sube hasta 9 fotos del día. Quedan en Evidencia Fotografica y el calendario las marca.",
    },
    {
        "id": "asignaciones",
        "icono": "📋",
        "titulo": "Asignaciones Diarias",
        "descripcion": "Consulta los pendientes e instrucciones enviadas por la administración.",
    },
    {
        "id": "historial",
        "icono": "🗂️",
        "titulo": "Historial y Estado de Pozos / Cuadrillas",
        "descripcion": "Revisa el historial de pozos y los reportes diarios de las cuadrillas.",
    },
    {
        "id": "unidades",
        "icono": "🔑",
        "titulo": "Check-in / Check-out de Unidades",
        "descripcion": "Recibe o entrega una unidad de relevo: km, combustible, fluidos y fotos.",
    },
    {
        "id": "combustible",
        "icono": "⛽",
        "titulo": "Carga de Combustible",
        "descripcion": "Registra litros, kilometraje y tipo de combustible de la unidad.",
    },
]
SUBMENUS_OP_HISTORIAL = [
    {
        "id": "pozos",
        "icono": "🛢️",
        "titulo": "Historial y Estado de Pozos",
        "descripcion": "Explora zonas, abre un pozo y consulta calendario, bitácora y programados.",
    },
    {
        "id": "cuadrillas",
        "icono": "👷",
        "titulo": "Historial de Cuadrillas",
        "descripcion": "Carpeta por cuadrilla: fechas, OT, personal e evidencia fotográfica.",
    },
    {
        "id": "ot_diario",
        "icono": "📑",
        "titulo": "Historial Diario de Órdenes de Trabajo",
        "descripcion": "Abre los PDF diarios de OT (DDMMYYYY) y el resumen de cuadrillas de esa jornada.",
    },
]


def reset_submodulos_operativo() -> None:
    st.session_state.op_modulo = None
    st.session_state.op_hist_seccion = None
    st.session_state.op_origen = None
    _reset_explorador_pozos()
    _reset_historial_cuadrillas()


def abrir_modulo_operativo(modulo_id: str, origen: str = "", proposito: str = "") -> None:
    st.session_state.op_modulo = modulo_id
    st.session_state.op_origen = origen or None
    if proposito:
        st.session_state.asig_proposito_pref = proposito
    if modulo_id == "combustible":
        st.session_state.comb_form_abierto = True
    st.rerun()


def salir_perfil_campo() -> None:
    reset_submodulos_operativo()
    st.session_state.pop("rol_seleccionado", None)
    st.session_state.pop("rol_pendiente", None)
    st.rerun()


PATRON_NO_UNIDAD_MOVIL = re.compile(
    r"bateria|batería|pozo|separador|skid|mci|ubm|ici|estacion",
    re.IGNORECASE,
)
PATRON_UNIDAD_ECONOMICO = re.compile(r"^u[-\s]?\d+", re.IGNORECASE)
NOMBRES_NO_UNIDAD = {"otros", "otro", "celaje", "na", "n/a", "none", "sin unidad"}


def parece_unidad_movil(nombre: str) -> bool:
    texto = str(nombre or "").strip()
    if not texto or texto.casefold() in NOMBRES_NO_UNIDAD:
        return False
    if PATRON_NO_UNIDAD_MOVIL.search(texto):
        return False
    if PATRON_UNIDAD_ECONOMICO.match(texto):
        return True
    return bool(re.search(r"\d", texto)) and len(texto) >= 3


def unidades_panorama_campo(diario: pd.DataFrame, parque: pd.DataFrame, datos: dict) -> list[str]:
    todas = unidades_taller(diario, parque, datos)
    abiertas = {str(a.get("unidad") or "").strip() for a in asignaciones_abiertas(datos)}
    criticas: list[str] = []
    moviles: list[str] = []
    for unidad in todas:
        estado = semaforo_efectivo_unidad(datos, unidad)
        if unidad in abiertas or estado["semaforo"] != SEMAFORO_OPERATIVA:
            criticas.append(unidad)
        elif parece_unidad_movil(unidad):
            moviles.append(unidad)
    vistos: set[str] = set()
    salida: list[str] = []
    for unidad in criticas + moviles:
        clave = unidad.casefold()
        if clave in vistos:
            continue
        vistos.add(clave)
        salida.append(unidad)
        if len(salida) >= 10:
            break
    return salida


def fila_hangar_unidad(datos: dict, diario: pd.DataFrame, unidad: str) -> dict:
    estado = semaforo_efectivo_unidad(datos, unidad)
    semaforo = estado["semaforo"]
    asgs = asignaciones_de_unidad(datos, unidad)
    abierta = next((a for a in asgs if str(a.get("estatus")) != ESTATUS_ASIG_CERRADA), None)
    ultima = asgs[0] if asgs else None
    if abierta:
        detalle = (
            f"En ruta · {abierta.get('proposito') or 'Asignación'} · "
            f"{html_fecha(abierta.get('fecha_salida'))}"
        )
    elif estado.get("ingreso"):
        ing = estado["ingreso"]
        detalle = str(ing.get("falla") or etiqueta_tipo_taller(str(ing.get("tipo") or "")) or "En taller")
    elif ultima:
        detalle = f"Último check-in: {html_fecha(ultima.get('fecha_salida'))}"
    else:
        detalle = "Sin check-in reciente"
    tono = TONO_ALERTA_SEMAFORO.get(semaforo, "ok")
    if semaforo == SEMAFORO_OPERATIVA:
        foco = "🟢"
    elif semaforo == SEMAFORO_PRECAUCION:
        foco = "🟡"
    else:
        foco = "🔴"
    return {
        "unidad": unidad,
        "semaforo": semaforo,
        "etiqueta": ETIQUETA_SEMAFORO_UNIDAD.get(semaforo, semaforo),
        "tono": tono,
        "foco": foco,
        "odometro": ultimo_odometro_unidad(diario, unidad),
        "detalle": detalle,
        "abierta": abierta,
    }


def metricas_hangar_campo(datos: dict, diario: pd.DataFrame) -> dict:
    hoy = date.today()
    asignaciones = [a for a in (datos.get("asignaciones") or []) if isinstance(a, dict)]
    checkins_hoy = sum(1 for a in asignaciones if parsear_fecha_valor(a.get("fecha_salida")) == hoy)
    litros = 0.0
    if diario is not None and not diario.empty and "Fecha" in diario.columns and "Litros" in diario.columns:
        fechas = pd.to_datetime(diario["Fecha"], errors="coerce")
        hoy_mask = fechas.dt.date == hoy
        litros = float(pd.to_numeric(diario.loc[hoy_mask, "Litros"], errors="coerce").fillna(0).sum())
    return {
        "checkins": checkins_hoy,
        "litros": litros,
        "en_ruta": len(asignaciones_abiertas(datos)),
    }


def render_tarjeta_hangar_unidad(fila: dict, seleccionada: str) -> None:
    sel = " sel" if seleccionada and seleccionada == fila["unidad"] else ""
    km = fila["odometro"]
    km_txt = f"{km:,.0f} km" if km else "Sin odómetro"
    st.markdown(
        f"""
        <div class="hangar-card {html.escape(fila['tono'])}{sel}">
            <div class="kicker">{fila['foco']} {html.escape(fila['etiqueta'])}</div>
            <h4>{html.escape(fila['unidad'])}</h4>
            <p><b>Km / horas:</b> {html.escape(km_txt)}</p>
            <p><i>{html.escape(fila['detalle'])}</i></p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button(
        f"Seleccionar {fila['unidad']}",
        key=f"hangar_sel_{fila['unidad']}",
        use_container_width=True,
        type="primary" if sel else "secondary",
    ):
        st.session_state.hangar_unidad = fila["unidad"]
        st.rerun()


def render_hangar_campo() -> None:
    diario, parque = cargar_datos_flota()
    datos = cargar_vehiculos()
    metricas = metricas_hangar_campo(datos, diario)
    unidad_sel = str(st.session_state.get("hangar_unidad") or "")

    st.markdown(
        """
        <div class="hangar-brand">
            <div class="prefijo">Servicios PJP4 · Base Álamo</div>
            <h1>Soledad · Centro de comando de campo</h1>
            <p>Hangar perimetral: acciones a los lados y panorama de unidades al centro.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    banner_salida = html_salida_dia(date.today())
    if banner_salida:
        st.markdown(banner_salida, unsafe_allow_html=True)
    top1, top2, top3 = st.columns([2, 2, 2])
    with top1:
        st.markdown(
            '<div class="hangar-top-meta">👤 <b>Operador:</b> Personal de campo · Base Álamo</div>',
            unsafe_allow_html=True,
        )
    with top2:
        pendientes = metricas["en_ruta"]
        st.markdown(
            f'<div class="hangar-online">🟢 <b>Sistema Online</b> | {pendientes} en ruta</div>',
            unsafe_allow_html=True,
        )
    with top3:
        if st.button("🚪 Cambiar perfil", use_container_width=True, key="hangar_cambiar_perfil"):
            salir_perfil_campo()

    st.divider()
    col_left, col_center, col_right = st.columns([1.15, 2.7, 1.15], gap="medium")

    with col_left:
        st.markdown("### ⚡ Acciones")
        if st.button("🏁 Check-in / Check-out", use_container_width=True, key="hangar_act_checkin"):
            abrir_modulo_operativo("unidades", origen="checkin")
        if st.button("⛽ Carga de combustible", use_container_width=True, key="hangar_act_comb"):
            abrir_modulo_operativo("combustible", origen="combustible")
        if st.button("📷 Subir evidencia", use_container_width=True, key="hangar_act_evid"):
            abrir_modulo_operativo("captura", origen="evidencia")
        if st.button("📷 Evidencia de cuadrilla", use_container_width=True, key="hangar_act_ev_cuadrilla"):
            abrir_modulo_operativo("evidencia_cuadrilla")
        st.markdown("---")
        if unidad_sel:
            st.success(f"{unidad_sel} lista para check-in, carga o evidencia.")
        else:
            st.markdown(
                '<div class="hangar-hint">💡 <b>Consejo de campo:</b> toca una unidad '
                "en el centro para dejarla seleccionada antes de una acción rápida.</div>",
                unsafe_allow_html=True,
            )

    with col_center:
        st.markdown("### 🚜 Panorama de unidades asignadas (Base Álamo)")
        st.caption("Estado en tiempo real de la flota: semáforo, último check-in y kilometraje.")
        unidades = unidades_panorama_campo(diario, parque, datos)
        if not unidades:
            st.info("No hay unidades de flota para mostrar. Revisa el control de combustible o un check-in.")
        else:
            filas = [fila_hangar_unidad(datos, diario, u) for u in unidades]
            orden = {SEMAFORO_DETENIDA: 0, SEMAFORO_EXTERNO: 1, SEMAFORO_PRECAUCION: 2, SEMAFORO_OPERATIVA: 3}
            filas.sort(key=lambda f: (orden.get(f["semaforo"], 9), f["unidad"]))
            for inicio in range(0, len(filas), 2):
                cols = st.columns(2, gap="medium")
                for col, fila in zip(cols, filas[inicio : inicio + 2]):
                    with col:
                        render_tarjeta_hangar_unidad(fila, unidad_sel)

    with col_right:
        st.markdown("### 📊 Control y alertas")
        if st.button("⚠️ Reportar avería", use_container_width=True, key="hangar_act_averia"):
            abrir_modulo_operativo("captura", origen="averia")
        if st.button("📋 Bitácora del día", use_container_width=True, key="hangar_act_bit"):
            abrir_modulo_operativo("historial")
        if st.button("📌 Asignaciones del día", use_container_width=True, key="hangar_act_asig"):
            abrir_modulo_operativo("asignaciones")
        st.markdown("---")
        st.markdown("#### 📌 Métricas rápidas")
        st.metric("Check-ins hoy", f"{metricas['checkins']} registros")
        st.metric("Combustible cargado", f"{metricas['litros']:,.0f} litros")
        st.metric("Unidades en ruta", str(metricas["en_ruta"]))


def render_consulta_asignaciones_operativo() -> None:
    st.markdown('<div class="section-title">Asignaciones del día</div>', unsafe_allow_html=True)
    st.caption("Pendientes e instrucciones publicadas por administración. Solo consulta de campo.")
    fecha = st.date_input(
        "Fecha a consultar",
        value=date.today(),
        format="DD/MM/YYYY",
        key="op_asig_fecha",
    )
    render_tarjeta_asignaciones_dia(fecha)


def render_historial_operativo() -> None:
    seccion = st.session_state.get("op_hist_seccion")
    if not seccion:
        render_cuadricula_explorador(
            SUBMENUS_OP_HISTORIAL,
            "op_hist_seccion",
            "Elige si quieres consultar pozos, cuadrillas o el historial diario de OT.",
        )
        return
    render_barra_volver(
        f'Historial de campo / <b>{html.escape(titulo_submenu(SUBMENUS_OP_HISTORIAL, seccion))}</b>',
        "op_hist_seccion",
        "← Volver al historial",
    )
    if seccion == "pozos":
        if migracion_excel_pendiente() and "ub_excel_autohecho" not in st.session_state:
            auto = importar_excel_ubicaciones(forzar=True)
            st.session_state.ub_excel_autohecho = True
            if auto.get("ok") and not auto.get("omitido"):
                st.success(
                    f"Inventario Excel importado: {auto['nuevos']} pozos nuevos y "
                    f"{auto['actualizados']} actualizados."
                )
        datos = catalogo_con_historial()
        render_flujo_historial_pozos(datos)
        return
    if seccion == "ot_diario":
        render_historial_diario_ot()
        return
    render_historial_cuadrillas(encabezado=False)


def render_captura_actividades_campo() -> None:
    form_id = st.session_state.form_id
    origen = str(st.session_state.get("op_origen") or "")
    unidad_sel = str(st.session_state.get("hangar_unidad") or "")
    if origen == "evidencia":
        extra = f" de {unidad_sel}" if unidad_sel else ""
        st.info(f"Sube la evidencia fotográfica{extra}. Queda en la bitácora de campo.")
    elif origen == "averia":
        extra = f" en {unidad_sel}" if unidad_sel else ""
        st.warning(f"Describe la avería{extra} y adjunta fotos. El registro queda en la bitácora del día.")

    if st.session_state.pop("just_saved", False):
        st.success(
            f"Registro guardado en {EXCEL_FILE}. "
            "Los consumos quedaron en almacen_insumos.json."
        )
        st.balloons()

    with st.container():
        col_u, col_o = st.columns(2, gap="medium")

        with col_u:
            st.markdown('<div class="section-title">📍 Ubicación</div>', unsafe_allow_html=True)
            fecha = st.date_input(
                "Fecha",
                value=date.today(),
                format="DD/MM/YYYY",
                key=f"fecha_{form_id}",
                persist_state="session",
            )
            ubicacion_sel, ubicacion_nueva, tag_instalacion = render_selector_ubicacion_captura(
                f"captura_campo_{form_id}"
            )

        with col_o:
            st.markdown('<div class="section-title">🔧 Operación</div>', unsafe_allow_html=True)
            clave_cuadrilla = f"cuadrilla_{form_id}"
            if clave_cuadrilla not in st.session_state:
                recordada = cuadrilla_recordada(fecha, list(CUADRILLAS))
                if recordada:
                    st.session_state[clave_cuadrilla] = recordada
            cuadrilla = st.selectbox(
                "Cuadrilla / posición oficial",
                CUADRILLAS,
                format_func=etiqueta_cuadrilla_historial,
                key=clave_cuadrilla,
                persist_state="session",
                help="Etiqueta que audita OPS (Mecánico 1, 2, 3, 4…). No cambia aunque haya cobertura.",
            )
            siguiente = siguiente_secuencia_turno(fecha, cuadrilla)
            secuencia_n = st.number_input(
                "Secuencia del día",
                min_value=1,
                step=1,
                value=int(siguiente),
                key=f"secuencia_{form_id}_{fecha.isoformat()}_{clave_cuadrilla_oficial(cuadrilla)}",
            )
            st.caption(
                f"Siguiente consecutivo de {etiqueta_cuadrilla_historial(cuadrilla)} "
                f"en esta fecha: **{etiqueta_secuencia_consecutiva(secuencia_n)}**."
            )
            precarga = aplicar_precarga_jornada(
                f"_precarga_campo_{form_id}",
                (fecha.isoformat(), cuadrilla),
                lambda: valores_precarga_campo(form_id, fecha, cuadrilla),
            )
            aviso_precarga_jornada(precarga)
            cobertura = render_bloque_cobertura(form_id, cuadrilla)
            sincronizar_memoria_dotacion(fecha, cuadrilla, form_id, int(secuencia_n))
            if tag_instalacion:
                st.caption(f"Ubicación del registro: **{tag_instalacion}**.")
            equipo_sel = render_selector_equipo_ubicacion(f"equipo_{form_id}", tag_instalacion)

            tipo_actividad, tipo_actividad_otro, actividad_corta_otro = render_selector_actividad_cuadrilla(
                f"tipo_actividad_{form_id}_{cuadrilla}",
                cuadrilla,
            )
            opciones_act = nombres_actividad_cuadrilla(cuadrilla)

    confirma_programa = False
    ubicacion_programa = str(
        tag_instalacion or resolver_tag_macro(ubicacion_sel, ubicacion_nueva) or ""
    ).strip()
    if tipo_actividad and tipo_actividad != "Otra":
        actividad_programa = str(tipo_actividad)
        oficial_programa = buscar_actividad_catalogo(cuadrilla, tipo_actividad)
        actividad_corta_programa = str((oficial_programa or {}).get("corto") or "")
    else:
        actividad_programa = str(tipo_actividad_otro or "").strip()
        actividad_corta_programa = str(actividad_corta_otro or "").strip()
    if ubicacion_programa and actividad_programa:
        try:
            coincidencias_programa = pendientes_programa_campo(
                ubicacion_programa,
                actividad_programa,
                fecha,
                actividad_corta_programa,
            )
        except Exception:
            coincidencias_programa = []
            st.caption("No se pudo consultar el programa de mantenimiento.")
        if coincidencias_programa:
            detalle = []
            for item in coincidencias_programa[:6]:
                cuando = f" · {item['mes']}" if item.get("mes") else ""
                zona = f" · {item['zona']}" if item.get("zona") else ""
                detalle.append(f"- **{item['tipo']}** en {item['ubicacion']}{zona}{cuando}")
            extras = len(coincidencias_programa) - len(detalle)
            if extras > 0:
                detalle.append(f"- y {extras} más")
            st.warning(
                "Hay una actividad pendiente en el programa de mantenimiento para esta ubicación.\n\n"
                + "\n".join(detalle)
            )
            confirma_programa = st.checkbox(
                "¿Es esta una actividad del Programa de Mantenimiento?",
                key=(
                    f"prog_confirma_{form_id}_"
                    f"{clave_pozo_plana(ubicacion_programa)}_"
                    f"{clave_pozo_plana(actividad_programa)}"
                ),
            )

    st.markdown("")
    st.markdown('<div class="section-title">🛢️ Insumos utilizados</div>', unsafe_allow_html=True)
    st.caption(
        "Activa solo los insumos que aplican. El origen — PJP4 (Propio) u OPS (Cliente) — es obligatorio. "
        "En un componente rotativo también se registra la condición de la pieza instalada y el destino de la retirada."
    )
    insumos = renderizar_insumos(form_id)

    st.markdown("---")
    st.markdown('<div class="section-title">📷 Evidencia fotográfica</div>', unsafe_allow_html=True)
    st.caption("Carga directa de la actividad. Hasta 9 imágenes se guardan en el registro al pulsar Guardar.")
    archivos_foto = st.file_uploader(
        "Cargar fotografías",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key=f"fotos_{form_id}",
    )
    if archivos_foto and len(archivos_foto) > MAX_FOTOS:
        st.warning("Solo se usarán las primeras 9 imágenes.")
    renderizar_galeria(archivos_foto)

    st.markdown("---")
    st.markdown("#### 📝 Estatus y actividad")

    col_desc, col_est, col_avance = st.columns([2.6, 1.15, 1.15], gap="medium")
    with col_desc:
        descripcion = st.text_area(
            "Descripción de la actividad",
            height=130,
            placeholder="Describe el trabajo realizado, hallazgos y acciones tomadas…",
            key=f"descripcion_{form_id}",
        )
    with col_est:
        estatus = st.selectbox(
            "Estatus del equipo",
            ESTATUS_OPCIONES,
            key=f"estatus_{form_id}",
        )
    with col_avance:
        porcentaje_avance = st.number_input(
            "Porcentaje de avance (%)",
            min_value=0,
            max_value=100,
            value=100,
            step=1,
            key=f"avance_{form_id}",
        )

    st.markdown("")
    _, col_btn_center, _ = st.columns([1, 2, 1])
    with col_btn_center:
        guardar = st.button("💾 Guardar Registro", type="primary", use_container_width=True)

    if not guardar:
        return

    equipo_final = resolver_equipo(equipo_sel)
    captura = resolver_captura_ubicacion(ubicacion_sel, ubicacion_nueva, tag_instalacion)
    sector_final = captura["sector"]
    pozo_final = captura["pozo"]
    tipo_actividad_final, actividad_corta_final = resolver_actividad_catalogo(
        cuadrilla, tipo_actividad, tipo_actividad_otro, actividad_corta_otro
    )
    secuencia_final = formatear_secuencia(secuencia_n)
    consumos = list(insumos.get("_consumos") or [])
    error_consumo = validar_consumos_campo(consumos)

    if not descripcion.strip():
        st.warning("Por favor escribe una descripción de la actividad antes de guardar.")
    elif not captura.get("ubicacion"):
        st.warning("Completa la ubicación: elige la zona y el pozo del catálogo, o la instalación especial.")
    elif opciones_act and tipo_actividad == "Otra" and not tipo_actividad_final:
        st.warning("Por favor escribe el nombre de la actividad.")
    elif opciones_act and not tipo_actividad_final:
        st.warning("Por favor selecciona una actividad oficial de la cuadrilla.")
    elif cobertura.get("cobertura") == "Sí" and not cobertura.get("mecanico_real"):
        st.warning("Captura el nombre real de quien ejecutó el servicio en esta cobertura.")
    elif error_consumo:
        st.warning(error_consumo)
    else:
        registro = construir_registro(
            fecha=fecha,
            secuencia=secuencia_final,
            sector=sector_final,
            pozo=pozo_final,
            cuadrilla=cuadrilla,
            equipo=equipo_final,
            tipo_actividad=tipo_actividad_final or "",
            actividad_corta=actividad_corta_final or "",
            insumos=insumos,
            descripcion=descripcion,
            estatus=estatus,
            fotos=guardar_evidencias(fecha, secuencia_final, archivos_foto),
            cobertura=cobertura,
            programado="Sí" if confirma_programa else "No",
            porcentaje_avance=porcentaje_avance,
        )

        try:
            personal = " · ".join(
                parte.strip()
                for parte in [
                    str((cobertura or {}).get("mecanico_real") or ""),
                    str((cobertura or {}).get("ayudante_real") or ""),
                ]
                if parte.strip()
            )
            nombres_consumo = [
                str(item.get("nombre") or "").strip()
                for item in consumos
                if str(item.get("nombre") or "").strip()
            ]
            try:
                agregar_registro_bitacora_soledad(
                    fecha,
                    cuadrilla,
                    personal,
                    captura.get("ubicacion") or pozo_final or sector_final,
                    descripcion,
                    ", ".join(nombres_consumo),
                    "",
                    limitar_fotos(list(archivos_foto or [])),
                    secuencia=int(secuencia_n),
                    avance=porcentaje_avance,
                    equipo=equipo_final,
                    tipo_actividad=tipo_actividad_final,
                    actividad_corta=actividad_corta_final,
                )
            except Exception:
                df_existente = cargar_bitacora()
                if df_existente.empty:
                    df_final = pd.DataFrame([registro])
                else:
                    df_final = pd.concat(
                        [df_existente, pd.DataFrame([registro])],
                        ignore_index=True,
                    )
                guardar_en_excel(df_final)
            registrar_consumos_campo(
                fecha, secuencia_final, cuadrilla, sector_final, consumos
            )
            st.session_state.just_saved = True
            st.session_state.form_id += 1
            st.rerun()
        except Exception as e:
            st.error(f"No se pudo guardar el registro: {e}")


def render_carga_combustible_campo() -> None:
    st.markdown('<div class="section-title">Carga de combustible</div>', unsafe_allow_html=True)
    st.caption("Registra litros, km y tipo. El renglón queda en el control de combustible de la unidad.")
    diario, parque = cargar_datos_flota()
    if diario.empty and parque.empty:
        st.warning("No hay control de combustible disponible. Puedes registrar la salida a estación con un check-in.")
        if st.button("Registrar salida a estación (check-in)", type="primary", key="comb_campo_asig"):
            abrir_modulo_operativo("unidades", origen="combustible", proposito=PROPOSITO_CARGA_COMBUSTIBLE)
        return
    st.session_state.comb_form_abierto = True
    pref = str(st.session_state.get("hangar_unidad") or "")
    fid = st.session_state.get("comb_form_id", 0)
    key = f"comb_cap_unidad_{fid}"
    unidades = unidades_combustible(diario, parque)
    if pref and pref in unidades and key not in st.session_state:
        st.session_state[key] = pref
    render_formulario_carga_combustible(diario, parque)


def render_panel_operativo() -> None:
    modulo = st.session_state.get("op_modulo")
    if not modulo:
        st.session_state.op_hist_seccion = None
        render_hangar_campo()
        return

    render_encabezado(
        "🚜 Panel del Personal Operativo",
        "Centro de comando de campo — Soledad · Base Álamo",
    )
    render_barra_volver(
        f'Hangar de campo / <b>{html.escape(titulo_submenu(MODULOS_OPERATIVO, modulo))}</b>',
        "op_modulo",
        "← Volver al hangar",
    )
    if modulo == "captura":
        render_captura_actividades_campo()
    elif modulo == "asignaciones":
        render_consulta_asignaciones_operativo()
    elif modulo == "historial":
        render_historial_operativo()
    elif modulo == "unidades":
        render_control_asignacion_unidades()
    elif modulo == "combustible":
        render_carga_combustible_campo()
    elif modulo == "evidencia_cuadrilla":
        render_subida_evidencia_cuadrilla(key_prefix="op_ev")
    else:
        st.session_state.op_modulo = None
        render_hangar_campo()


# ── Panel administrativo ───────────────────────────────────────────────────────
def etiqueta_registro(df: pd.DataFrame, idx: int) -> str:
    fila = df.loc[idx]
    fecha = parsear_fecha_valor(fila.get("FECHA"))
    fecha_txt = fecha.strftime("%d/%m/%Y") if fecha else str(fila.get("FECHA") or "")
    ubic = etiqueta_ubicacion_fila(fila) or str(
        fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("SECTOR") or ""
    ).strip()
    evid = estatus_evidencia_fila(fila)
    secuencia = numero_secuencia_actividad(fila.get("SECUENCIA"))
    actividad = texto_visible(fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD"), "")
    if len(actividad) > 45:
        actividad = actividad[:44].rstrip() + "…"
    # Streamlit identifica la opción elegida por su texto: dos etiquetas iguales saltan al otro registro.
    return (
        f"{fecha_txt} · {fila.get('CUADRILLA', '')} · Sec. {secuencia or '—'} · {ubic or 'Sin ubicación'} · "
        f"{actividad or 'Sin descripción'} · {evid} · {len(rutas_fotos_fila(fila))} foto(s) · reg. {idx + 1}"
    )


MENSAJE_BITACORA_EN_USO = (
    "La bitácora está abierta en Excel o OneDrive la está sincronizando. "
    "Ciérrala y vuelve a intentar; no se guardó ningún cambio."
)


def _avisar_revision(tipo: str, texto: str) -> None:
    st.session_state["_rev_aviso"] = (tipo, texto)
    st.session_state["_rev_ver"] = int(st.session_state.get("_rev_ver", 0)) + 1


def render_aviso_revision() -> None:
    aviso = st.session_state.pop("_rev_aviso", None)
    if not aviso:
        return
    tipo, texto = aviso
    {"success": st.success, "warning": st.warning, "error": st.error}.get(tipo, st.info)(texto)


def _texto_sync_calendario(resultado: dict) -> str:
    if resultado.get("ok"):
        return " El calendario de asistencia quedó recalculado."
    mensaje = str(resultado.get("mensaje") or "").strip()
    return f" El calendario se recalcula en pantalla; el archivo se actualizará cuando quede libre ({mensaje})." if mensaje else ""


def _opciones_cuadrilla_edicion(df: pd.DataFrame, actual: str) -> list[str]:
    por_clave: dict[str, str] = {}
    for valor in [actual, *df.get("CUADRILLA", pd.Series(dtype=object)).astype(str).tolist()]:
        texto = str(valor or "").strip()
        clave = clave_cuadrilla_oficial(texto)
        if texto and clave and clave not in por_clave:
            por_clave[clave] = texto
    for oficial in ORDEN_CUADRILLAS_OFICIAL:
        clave = clave_cuadrilla_oficial(oficial)
        if clave not in por_clave:
            por_clave[clave] = _sin_acentos(oficial).upper()
    return sorted(por_clave.values(), key=orden_cuadrilla_oficial)


def render_edicion_registro_bitacora(df: pd.DataFrame, idx) -> None:
    fila = df.loc[idx]
    ver = int(st.session_state.get("_rev_ver", 0))
    dia = parsear_fecha_valor(fila.get("FECHA")) or date.today()
    cuadrilla_actual = str(fila.get("CUADRILLA") or "").strip()
    opciones_cuadrilla = _opciones_cuadrilla_edicion(df, cuadrilla_actual)
    clave_actual = clave_cuadrilla_oficial(cuadrilla_actual)
    indice_cuadrilla = next(
        (i for i, op in enumerate(opciones_cuadrilla) if clave_cuadrilla_oficial(op) == clave_actual),
        0,
    )
    ubic_actual = str(fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("POZO") or "").strip()
    secuencia_actual = numero_secuencia_actividad(fila.get("SECUENCIA")) or 1
    descripcion_actual = texto_visible(
        fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD"), ""
    )
    comentarios_actual = texto_visible(fila.get("COMENTARIOS"), "")
    personal_actual = texto_visible(fila.get("PERSONAL"), "")

    st.markdown("##### ✏️ Editar datos del registro")
    with st.form(f"rev_editar_{idx}_{ver}"):
        c1, c2, c3 = st.columns([1, 1.2, 0.8])
        with c1:
            nueva_fecha = st.date_input("Fecha", value=dia, format="DD/MM/YYYY")
        with c2:
            clave_cuad_edit = f"rev_edit_cuadrilla_{idx}"
            if clave_cuad_edit in st.session_state and st.session_state[clave_cuad_edit] not in opciones_cuadrilla:
                st.session_state.pop(clave_cuad_edit, None)
            if clave_cuad_edit not in st.session_state:
                coincidencia = next(
                    (op for op in opciones_cuadrilla if clave_cuadrilla_oficial(op) == clave_actual),
                    None,
                )
                if coincidencia:
                    st.session_state[clave_cuad_edit] = coincidencia
            nueva_cuadrilla = st.selectbox(
                "Cuadrilla",
                opciones_cuadrilla,
                index=None if clave_cuad_edit in st.session_state else indice_cuadrilla,
                format_func=etiqueta_cuadrilla_historial,
                key=clave_cuad_edit,
                persist_state="session",
            )
        with c3:
            nueva_secuencia = st.number_input("Secuencia", min_value=1, step=1, value=int(secuencia_actual))
        c4, c5 = st.columns([2, 1])
        with c4:
            nueva_ubic, nuevo_numero, _tag_edit = render_selector_ubicacion_captura(
                f"rev_edit_{idx}_{ver}",
                valor_actual=ubic_actual,
            )
        with c5:
            clave_personal_edit = f"rev_edit_personal_{idx}"
            if clave_personal_edit not in st.session_state:
                st.session_state[clave_personal_edit] = personal_actual
            nuevo_personal = st.text_input(
                "Ficha / Personal",
                key=clave_personal_edit,
                persist_state="session",
            )
        nueva_desc = st.text_area("Descripción de la actividad", value=descripcion_actual, height=110)
        nuevos_coment = st.text_area("Comentarios", value=comentarios_actual, height=80)
        guardar = st.form_submit_button("💾 Guardar Cambios en Bitácora", type="primary")
    if not guardar:
        return
    captura = resolver_captura_ubicacion(nueva_ubic, nuevo_numero, _tag_edit)
    ubicacion_guardada = captura["ubicacion"]
    cambios = {
        "fecha": nueva_fecha,
        "cuadrilla": nueva_cuadrilla,
        "secuencia": int(nueva_secuencia),
        "ubicacion": ubicacion_guardada,
        "personal": nuevo_personal,
        "descripcion": nueva_desc,
        "comentarios": nuevos_coment,
    }
    sin_cambios = (
        nueva_fecha == dia
        and clave_cuadrilla_oficial(nueva_cuadrilla) == clave_actual
        and int(nueva_secuencia) == int(secuencia_actual)
        and cambios["ubicacion"] == ubic_actual
        and nuevo_personal.strip() == personal_actual
        and nueva_desc.strip() == descripcion_actual
        and nuevos_coment.strip() == comentarios_actual
    )
    if sin_cambios:
        st.info("No hay cambios por guardar.")
        return
    try:
        resultado = actualizar_registro_bitacora(idx, fila, cambios)
    except PermissionError:
        st.error(MENSAJE_BITACORA_EN_USO)
        return
    except (ValueError, FileNotFoundError) as exc:
        st.error(str(exc))
        return
    sync = despues_de_cambio_bitacora("Edición de registro en bitácora")
    texto = f"Registro actualizado en la fila {resultado['fila']} de la bitácora."
    if resultado["fotos_movidas"]:
        texto += f" Se movieron {resultado['fotos_movidas']} foto(s) a la carpeta de la nueva fecha, cuadrilla o secuencia."
    texto += _texto_sync_calendario(sync)
    if resultado.get("aviso"):
        _avisar_revision("warning", f"{texto} {resultado['aviso']}")
    else:
        _avisar_revision("success", texto)
    st.session_state["_rev_seleccion"] = idx
    st.rerun()


def render_fotos_registro_editables(df: pd.DataFrame, idx) -> None:
    fila = df.loc[idx]
    ver = int(st.session_state.get("_rev_ver", 0))
    rutas = rutas_fotos_fila(fila)
    en_disco = [ruta for ruta in rutas if os.path.isfile(ruta)]
    carpeta = carpeta_fotos_registro_bitacora(fila).replace("\\", "/")
    st.markdown("##### 📷 Fotografías del registro")
    st.caption(f"{len(en_disco)} foto(s) en `{carpeta}`.")
    if not rutas:
        st.info("Este registro todavía no tiene fotografías.")
    for inicio in range(0, len(rutas), 3):
        columnas = st.columns(3, gap="small")
        for col, ruta in zip(columnas, rutas[inicio : inicio + 3]):
            with col:
                nombre = os.path.basename(ruta)
                if not os.path.isfile(ruta):
                    st.markdown(f'<div class="foto-vacia">No encontrada<br>{html.escape(nombre)}</div>', unsafe_allow_html=True)
                    continue
                st.image(ruta, use_container_width=True)
                st.caption(nombre)
                firma = hashlib.sha1(ruta.encode("utf-8", errors="ignore")).hexdigest()[:10]
                if st.button("🗑️ Eliminar foto", key=f"rev_borrar_foto_{idx}_{firma}", use_container_width=True):
                    try:
                        eliminar_foto_registro(ruta)
                    except (OSError, ValueError) as exc:
                        st.error(f"No se pudo borrar {nombre}: {exc}")
                        continue
                    try:
                        persistencia.respaldar_cambios("Foto eliminada de un registro")
                    except Exception:
                        pass
                    _avisar_revision("success", f"Se eliminó la foto {nombre}.")
                    st.session_state["_rev_seleccion"] = idx
                    st.rerun()
    huecos = MAX_FOTOS - len(en_disco)
    if huecos <= 0:
        st.caption(f"El registro ya tiene {MAX_FOTOS} fotografías. Elimina una para adjuntar otra.")
        return
    archivos = st.file_uploader(
        f"Adjuntar fotografías a este registro (quedan {huecos} espacio(s))",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        key=f"rev_fotos_nuevas_{idx}_{ver}",
    )
    if archivos and len(archivos) > huecos:
        st.caption(f"Solo se guardarán las primeras {huecos} foto(s).")
    if st.button("📷 Adjuntar fotos al registro", key=f"rev_adjuntar_{idx}_{ver}", disabled=not archivos):
        nuevas = anexar_fotos_registro_bitacora(fila, list(archivos)[:huecos])
        try:
            persistencia.respaldar_cambios("Evidencia anexada a un registro")
        except Exception:
            pass
        _avisar_revision("success", f"Se adjuntaron {len(nuevas)} foto(s) al registro.")
        st.session_state["_rev_seleccion"] = idx
        st.rerun()


def render_eliminar_registro_bitacora(df: pd.DataFrame, idx) -> None:
    fila = df.loc[idx]
    ver = int(st.session_state.get("_rev_ver", 0))
    fotos = [ruta for ruta in rutas_fotos_fila(fila) if os.path.isfile(ruta)]
    st.markdown("##### ⚠️ Eliminar actividad")
    with st.container(border=True):
        st.caption(
            "Quita la fila de la bitácora operativa y borra la carpeta de fotografías de esta actividad. "
            "El calendario de asistencia se recalcula al terminar."
        )
        confirmar = st.checkbox(
            f"Confirmo que quiero eliminar esta actividad y sus {len(fotos)} fotografía(s). No se puede deshacer.",
            key=f"rev_confirmar_borrado_{idx}_{ver}",
        )
        if not st.button(
            "🗑️ Eliminar Actividad Completa",
            type="primary",
            disabled=not confirmar,
            key=f"rev_borrar_registro_{idx}_{ver}",
        ):
            return
    dia_txt = (parsear_fecha_valor(fila.get("FECHA")) or date.today()).strftime("%d/%m/%Y")
    etiqueta = f"{dia_txt} · {fila.get('CUADRILLA', '')} · Sec. {numero_secuencia_actividad(fila.get('SECUENCIA')) or '—'}"
    try:
        resultado = eliminar_registro_bitacora(idx, fila)
    except PermissionError:
        st.error(MENSAJE_BITACORA_EN_USO)
        return
    except (ValueError, FileNotFoundError) as exc:
        st.error(str(exc))
        return
    sync = despues_de_cambio_bitacora("Registro eliminado de la bitácora")
    _avisar_revision(
        "success",
        f"Se eliminó la actividad {etiqueta} y {resultado['fotos_borradas']} foto(s).{_texto_sync_calendario(sync)}",
    )
    st.session_state["_rev_seleccion"] = idx
    st.rerun()


def render_revision_bitacoras() -> None:
    render_captura_bitacora_admin()
    st.markdown('<div class="section-title">Revisión de bitácoras</div>', unsafe_allow_html=True)
    marca = datetime.fromtimestamp(_mtime_archivo(EXCEL_FILE)).strftime("%d/%m/%Y %H:%M") if os.path.exists(EXCEL_FILE) else "sin archivo"
    df = bitacora_con_fechas(cargar_bitacora())
    disponibles = fechas_disponibles_operacion(df)
    rango = (
        f"{disponibles[-1].strftime('%d/%m/%Y')} a {disponibles[0].strftime('%d/%m/%Y')}"
        if disponibles
        else "sin fechas"
    )
    st.caption(
        f"Fuente operativa: `{EXCEL_FILE}` (actualizado {marca}). "
        f"Se listan todas las jornadas disponibles ({rango})."
    )
    if df.empty:
        st.info(f"Aún no hay registros en {EXCEL_FILE}.")
        return

    base = df.copy()

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    fechas = [d.strftime("%d/%m/%Y") for d in fechas_disponibles_operacion(base)]
    ubicaciones = sorted(
        {
            str(v).strip()
            for v in base.get("UBICACIÓN", base.get("EQUIPO/UBICACIÓN", pd.Series(dtype=object)))
            if str(v).strip() and str(v).strip().lower() not in {"nan", "none"}
        }
    )
    cuadrillas = sorted(
        {etiqueta_cuadrilla_historial(str(v)) for v in base["CUADRILLA"] if str(v).strip()},
        key=orden_cuadrilla_oficial,
    )
    with col_f1:
        filtro_fecha = st.selectbox("Fecha", ["Todas"] + fechas)
    with col_f2:
        filtro_ubi = st.selectbox("Ubicación", ["Todas"] + ubicaciones)
    with col_f3:
        opciones_filtro_cuadrilla = ["Todas"] + cuadrillas
        if st.session_state.get("rev_filtro_cuadrilla") not in opciones_filtro_cuadrilla:
            st.session_state.pop("rev_filtro_cuadrilla", None)
        filtro_cuadrilla = st.selectbox(
            "Cuadrilla",
            opciones_filtro_cuadrilla,
            key="rev_filtro_cuadrilla",
            persist_state="session",
        )
    with col_f4:
        filtro_revision = st.selectbox("Revisión", ["Todas"] + ESTADOS_REVISION)
    filtro_evidencia = st.selectbox(
        "Evidencia fotográfica",
        ["Todas", EVIDENCIA_PENDIENTE, EVIDENCIA_COMPLETA],
        key="rev_filtro_evidencia",
    )

    filtrado = base.copy()
    if filtro_fecha != "Todas":
        dia = parsear_fecha_valor(filtro_fecha)
        if dia is not None and "_fecha" in filtrado.columns:
            filtrado = filtrado[filtrado["_fecha"] == dia]
    if filtro_ubi != "Todas":
        col_ubi = "UBICACIÓN" if "UBICACIÓN" in filtrado.columns else "EQUIPO/UBICACIÓN"
        filtrado = filtrado[filtrado[col_ubi].astype(str).str.strip() == filtro_ubi]
    if filtro_cuadrilla != "Todas":
        objetivo = clave_cuadrilla_oficial(filtro_cuadrilla)
        filtrado = filtrado[filtrado["CUADRILLA"].astype(str).map(clave_cuadrilla_oficial) == objetivo]
    if filtro_revision != "Todas" and "REVISIÓN" in filtrado.columns:
        filtrado = filtrado[filtrado["REVISIÓN"].astype(str) == filtro_revision]
    if filtro_evidencia != "Todas":
        mask_ev = filtrado.apply(lambda r: estatus_evidencia_fila(r) == filtro_evidencia, axis=1)
        filtrado = filtrado.loc[mask_ev]

    firma_export = hashlib.md5(
        "|".join(
            [
                str(filtro_fecha),
                str(filtro_ubi),
                str(filtro_cuadrilla),
                str(filtro_revision),
                str(filtro_evidencia),
                str(len(filtrado)),
            ]
        ).encode()
    ).hexdigest()[:8]
    col_exp, col_exp_info = st.columns([1, 2])
    with col_exp:
        st.download_button(
            "⬇️ Exportar Excel filtrado",
            data=excel_bitacora_filtrada(filtrado),
            file_name=nombre_excel_bitacora_filtrada(filtro_fecha, filtrado),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"rev_descargar_excel_{firma_export}",
            disabled=filtrado.empty,
        )
    with col_exp_info:
        st.caption(
            f"{len(filtrado)} registro(s) visibles. El archivo usa las columnas de `{EXCEL_FILE}`."
        )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Registros", len(filtrado))
    m2.metric("Pendientes", int((filtrado["REVISIÓN"] == "Pendiente").sum()) if "REVISIÓN" in filtrado.columns else 0)
    m3.metric("Validados", int((filtrado["REVISIÓN"] == "Validado").sum()) if "REVISIÓN" in filtrado.columns else 0)
    m4.metric("Con observaciones", int((filtrado["REVISIÓN"] == "Con observaciones").sum()) if "REVISIÓN" in filtrado.columns else 0)
    pend_ev = int(filtrado.apply(lambda r: estatus_evidencia_fila(r) != EVIDENCIA_COMPLETA, axis=1).sum()) if not filtrado.empty else 0
    st.caption(f"{pend_ev} registro(s) con evidencia fotográfica pendiente.")

    st.dataframe(
        vista_revision_soledad(filtrado),
        use_container_width=True,
        hide_index=True,
    )

    if filtrado.empty:
        st.caption("No hay registros con esos filtros.")
        return

    indices = list(filtrado.index)
    render_aviso_revision()
    # Al editar, borrar o filtrar cambian las opciones; un key nuevo evita que el navegador conserve el texto viejo.
    firma_opciones = hashlib.md5(",".join(map(str, indices)).encode()).hexdigest()[:8]
    clave_registro = f"rev_registro_{int(st.session_state.get('_rev_ver', 0))}_{firma_opciones}"
    pendiente = st.session_state.pop("_rev_seleccion", None)
    if pendiente not in indices and clave_registro not in st.session_state:
        pendiente = st.session_state.get("_rev_ultimo")
    if pendiente in indices:
        st.session_state[clave_registro] = pendiente
    elif st.session_state.get(clave_registro) not in indices:
        st.session_state.pop(clave_registro, None)
    seleccionado = st.selectbox(
        "Registro a auditar",
        indices,
        format_func=lambda i: etiqueta_registro(df, i),
        key=clave_registro,
    )
    st.session_state["_rev_ultimo"] = seleccionado
    fila = df.loc[seleccionado]
    gente = personal_de_registro(fila)
    st.caption(
        f"Posición oficial: **{gente['cuadrilla'] or '—'}** · "
        f"Titular: **{texto_par_personal(gente['mecanico_titular'], gente['ayudante_titular'])}** · "
        f"En turno: **{texto_par_personal(gente['mecanico_real'], gente['ayudante_real'])}** · "
        f"Cobertura: **{gente['cobertura'] or 'No'}**"
    )
    dia_revision = fila.get("_fecha")
    if isinstance(dia_revision, datetime):
        dia_revision = dia_revision.date()
    if not isinstance(dia_revision, date):
        dia_revision = parsear_fecha_valor(fila.get("FECHA"))
    render_nota_asistencia(
        gente["cuadrilla"],
        dia_revision,
        str(fila.get("PERSONAL") or ""),
    )
    render_edicion_registro_bitacora(df, seleccionado)
    render_fotos_registro_editables(df, seleccionado)

    st.markdown("##### Estatus de revisión")
    col_est, col_obs = st.columns([1, 2])
    with col_est:
        estado_actual = str(fila.get("REVISIÓN") or "Pendiente")
        idx_estado = ESTADOS_REVISION.index(estado_actual) if estado_actual in ESTADOS_REVISION else 0
        nuevo_estado = st.selectbox("Estatus de revisión", ESTADOS_REVISION, index=idx_estado)
    with col_obs:
        observacion = st.text_area(
            "Observación de revisión",
            value=str(fila.get("OBSERVACIÓN REVISIÓN") or ""),
            height=90,
        )
    if st.button("Guardar revisión", type="primary"):
        df.loc[seleccionado, "REVISIÓN"] = nuevo_estado
        df.loc[seleccionado, "OBSERVACIÓN REVISIÓN"] = observacion.strip()
        guardar_en_excel(df)
        st.success("Revisión actualizada.")
        st.rerun()

    render_eliminar_registro_bitacora(df, seleccionado)


def render_editor_lista(titulo: str, clave: str, ayuda: str) -> None:
    st.markdown(f"**{titulo}**")
    st.caption(ayuda)
    catalogos = obtener_catalogos()
    df = pd.DataFrame({titulo: catalogos[clave]})
    editado = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"editor_{clave}",
    )
    if st.button(f"Guardar {titulo.lower()}", key=f"btn_guardar_{clave}"):
        valores = [
            str(v).strip()
            for v in editado[titulo].tolist()
            if str(v).strip() and str(v).strip().lower() not in {"nan", "none"}
        ]
        catalogos[clave] = valores
        guardar_catalogos(catalogos)
        st.success(f"{titulo} actualizado. El panel operativo usará esta lista.")
        st.rerun()


def render_gestion_catalogos() -> None:
    st.markdown('<div class="section-title">Gestión de catálogos</div>', unsafe_allow_html=True)
    st.caption("Los cambios se guardan en catalogos.json y el formulario de campo los toma de inmediato.")
    sub = st.session_state.get("admin_sub_catalogos")
    if not sub:
        render_cuadricula_explorador(
            SUBMENUS_CATALOGOS,
            "admin_sub_catalogos",
            "Elige el catálogo que quieres editar.",
        )
        return
    render_barra_volver(
        f'Catálogos / <b>{html.escape(titulo_submenu(SUBMENUS_CATALOGOS, sub))}</b>',
        "admin_sub_catalogos",
        "← Volver a Catálogos",
    )
    if sub == "ubicaciones":
        catalogos = obtener_catalogos()
        sin_pozo = set(catalogos["sectores_sin_pozo"])
        filas = []
        for sector in catalogos["sectores"]:
            if sector == SECTOR_OTRO:
                continue
            filas.append(
                {
                    "Ubicación": sector,
                    "Requiere pozo": sector not in sin_pozo,
                }
            )
        editado = st.data_editor(
            pd.DataFrame(filas),
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="editor_ubicaciones",
            column_config={
                "Ubicación": st.column_config.TextColumn(required=True),
                "Requiere pozo": st.column_config.CheckboxColumn(),
            },
        )
        if st.button("Guardar ubicaciones", type="primary"):
            sectores = []
            sin_pozo_nuevos = []
            for _, fila in editado.iterrows():
                nombre = str(fila.get("Ubicación", "")).strip()
                if not nombre or nombre == SECTOR_OTRO:
                    continue
                if nombre in sectores:
                    continue
                sectores.append(nombre)
                if not bool(fila.get("Requiere pozo", True)):
                    sin_pozo_nuevos.append(nombre)
            sectores.append(SECTOR_OTRO)
            catalogos["sectores"] = sectores
            catalogos["sectores_sin_pozo"] = sin_pozo_nuevos
            guardar_catalogos(catalogos)
            st.success("Ubicaciones actualizadas.")
            st.rerun()
    elif sub == "mci":
        render_editor_lista(
            "Equipos MCI",
            "equipos_mci",
            "Motores de combustión interna que aparecen en el selector de Equipo.",
        )
    elif sub == "ub":
        render_editor_lista(
            "Equipos UB",
            "equipos_ub",
            "Unidades de bombeo que aparecen en el selector de Equipo.",
        )
    elif sub == "actividades":
        render_editor_actividades_cuadrilla()
    elif sub == "personal":
        render_editor_personal_cuadrillas()


def _ids_actividades(items: list[dict]) -> list[dict]:
    cambio = False
    for item in items:
        if not str(item.get("id") or "").strip():
            item["id"] = uuid.uuid4().hex[:10]
            cambio = True
    return items if cambio else items


def _guardar_grupo_actividades(grupo: str, items: list[dict]) -> None:
    datos = cargar_catalogo_actividades()
    datos[grupo] = items
    guardar_catalogo_actividades(datos)


def _nombre_actividad_repetido(items: list[dict], nombre: str, id_item: str = "") -> bool:
    clave = str(nombre or "").strip().casefold()
    return any(
        str(item.get("nombre") or "").casefold() == clave and str(item.get("id") or "") != id_item
        for item in items
    )


def render_gestionar_actividades() -> None:
    st.markdown('<div class="section-title">Gestionar Actividades</div>', unsafe_allow_html=True)
    st.caption(
        f"Catálogo de mantenimiento por especialidad. Cada alta, edición o baja se guarda en "
        f"`{ACTIVIDADES_FILE}` y el selector de la bitácora usa las actividades activas de inmediato."
    )
    datos = cargar_catalogo_actividades()
    grupo = st.selectbox(
        "Especialidad",
        GRUPOS_ACTIVIDAD,
        format_func=lambda clave: ETIQUETAS_GRUPO_ACTIVIDAD.get(clave, clave),
        key="gact_grupo",
        persist_state="session",
    )
    items = [dict(item) for item in (datos.get(grupo) or [])]
    if _ids_actividades(items) and any(not str(origen.get("id") or "") for origen in (datos.get(grupo) or [])):
        _guardar_grupo_actividades(grupo, items)
        items = [dict(item) for item in (cargar_catalogo_actividades().get(grupo) or [])]
    activas = sum(1 for item in items if actividad_esta_activa(item))
    st.caption(f"{activas} activa(s) · {len(items) - activas} inactiva(s) en {ETIQUETAS_GRUPO_ACTIVIDAD.get(grupo, grupo)}.")

    with st.form("gact_alta", clear_on_submit=True):
        st.markdown("**Nueva actividad**")
        c1, c2 = st.columns(2)
        nombre_nuevo = c1.text_input("Nombre de la actividad", placeholder="Nombre oficial completo")
        corto_nuevo = c2.text_input("Nombre corto", placeholder="Ej. MTTO BIM MCI")
        agregar = st.form_submit_button("Agregar actividad", type="primary")
    if agregar:
        nueva = _item_actividad(nombre_nuevo, corto_nuevo, True, uuid.uuid4().hex[:10])
        if not nueva:
            st.warning("Escribe el nombre de la actividad. «Otra» queda reservada para captura libre.")
        elif _nombre_actividad_repetido(items, nueva["nombre"]):
            st.warning("Esa actividad ya está en esta especialidad.")
        else:
            items.append(nueva)
            _guardar_grupo_actividades(grupo, items)
            st.success("Actividad agregada. Ya aparece en el selector de la bitácora.")
            st.rerun()

    def _alternar_actividad(grupo_cb: str, ident_cb: str) -> None:
        clave = f"gact_{grupo_cb}_{ident_cb}_on"
        vigentes = [dict(actual) for actual in (cargar_catalogo_actividades().get(grupo_cb) or [])]
        for actual in vigentes:
            if str(actual.get("id") or "") == ident_cb:
                actual["activa"] = bool(st.session_state.get(clave))
        _guardar_grupo_actividades(grupo_cb, vigentes)

    st.markdown("**Catálogo vigente**")
    if not items:
        st.info("No hay actividades en esta especialidad. Agrega la primera arriba.")
    for item in items:
        ident = str(item.get("id") or "")
        with st.container(border=True):
            c_nom, c_cor = st.columns([1.4, 1])
            clave_nom = f"gact_{grupo}_{ident}_nom"
            clave_cor = f"gact_{grupo}_{ident}_cor"
            clave_on = f"gact_{grupo}_{ident}_on"
            if clave_nom not in st.session_state:
                st.session_state[clave_nom] = item.get("nombre") or ""
            if clave_cor not in st.session_state:
                st.session_state[clave_cor] = item.get("corto") or ""
            if clave_on not in st.session_state:
                st.session_state[clave_on] = actividad_esta_activa(item)
            nombre_edit = c_nom.text_input("Nombre de la actividad", key=clave_nom)
            corto_edit = c_cor.text_input("Nombre corto", key=clave_cor)
            c_on, c_save, c_del = st.columns([1, 1, 1])
            activa = c_on.checkbox(
                "Activa en bitácora",
                key=clave_on,
                on_change=_alternar_actividad,
                args=(grupo, ident),
            )
            if not activa:
                c_on.caption("No se lista en el registro.")
            if c_save.button("Guardar cambios", key=f"gact_{grupo}_{ident}_save", use_container_width=True):
                editada = _item_actividad(nombre_edit, corto_edit, activa, ident)
                if not editada:
                    st.warning("El nombre no puede quedar vacío ni llamarse «Otra».")
                elif _nombre_actividad_repetido(items, editada["nombre"], ident):
                    st.warning("Ya hay otra actividad con ese nombre en esta especialidad.")
                else:
                    reemplazo = [editada if str(actual.get("id") or "") == ident else actual for actual in items]
                    _guardar_grupo_actividades(grupo, reemplazo)
                    st.success("Cambios guardados en el catálogo.")
                    st.rerun()
            if c_del.button("Eliminar", key=f"gact_{grupo}_{ident}_del", use_container_width=True):
                restantes = [actual for actual in items if str(actual.get("id") or "") != ident]
                _guardar_grupo_actividades(grupo, restantes)
                st.session_state.pop(clave_nom, None)
                st.session_state.pop(clave_cor, None)
                st.session_state.pop(clave_on, None)
                st.success("Actividad eliminada del catálogo.")
                st.rerun()


def render_editor_actividades_cuadrilla() -> None:
    render_gestionar_actividades()


def render_editor_personal_cuadrillas() -> None:
    st.markdown("**Dotación titular por posición oficial**")
    lista = leer_lista_asistencia()
    cal = leer_calendario_asistencia()
    nota_lista = (
        f" Personal, Personal 2, Personal 3 y DE APOYO se leen de `{LISTA_ASISTENCIA_FILE}`."
        if lista.get("disponible")
        else ""
    )
    nota_cal = (
        f" Las jornadas (realizadas, total y avance) salen de `{CALENDARIO_ASISTENCIA_FILE}`."
        if cal.get("disponible")
        else ""
    )
    st.caption(
        f"La cuadrilla (Mecánico 1, 2, 3, 4…) es la etiqueta de OPS. "
        f"Al elegirla en captura se rellenan ficha, nombre y número combinado. "
        f"Retroexcavadora y extras manuales se guardan en {PERSONAL_FILE}."
        f"{nota_lista}{nota_cal}"
    )
    datos = cargar_personal_cuadrillas()
    filas = []
    for cuadrilla in CUADRILLAS:
        item = posicion_personal_cuadrilla(cuadrilla)
        filas.append(
            {
                "Posición oficial": cuadrilla,
                "Personal 1": item["mecanico"],
                "Personal 2": item["ayudante"],
                "Personal 3": item["apoyo"],
                "Número combinado": codigo_dotacion_cuadrilla(item),
            }
        )
    editado = st.data_editor(
        pd.DataFrame(filas),
        hide_index=True,
        use_container_width=True,
        disabled=["Posición oficial"],
        key="editor_personal_posiciones",
        column_config={
            "Personal 1": st.column_config.TextColumn(),
            "Personal 2": st.column_config.TextColumn(),
            "Personal 3": st.column_config.TextColumn(),
            "Número combinado": st.column_config.TextColumn(),
        },
    )
    apoyos = [str(n).strip() for n in (lista.get("extras") or []) if str(n).strip()]
    apoyo_cf = {n.casefold() for n in apoyos}
    if apoyos:
        st.markdown("**Personal de apoyo (DE APOYO)**")
        st.caption(f"Leído de `{LISTA_ASISTENCIA_FILE}`. Entra a la bolsa de nombres para coberturas.")
        st.dataframe(
            pd.DataFrame({"Nombre": apoyos}),
            hide_index=True,
            use_container_width=True,
        )
    extras_manuales = [n for n in (datos.get("extras") or []) if str(n).casefold() not in apoyo_cf]
    extras_df = pd.DataFrame({"Nombre": extras_manuales})
    if extras_df.empty:
        extras_df = pd.DataFrame({"Nombre": [""]})
    extras_ed = st.data_editor(
        extras_df,
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        key="editor_personal_extras",
        column_config={"Nombre": st.column_config.TextColumn("Personal extra (puede cubrir)")},
    )
    if st.button("💾 Guardar personal y coberturas", type="primary", key="btn_guardar_personal"):
        posiciones = {}
        for _, fila in editado.iterrows():
            posicion = str(fila.get("Posición oficial") or "").strip()
            if not posicion:
                continue
            item = {
                "mecanico": str(fila.get("Personal 1") or "").strip(),
                "ayudante": str(fila.get("Personal 2") or "").strip(),
                "apoyo": str(fila.get("Personal 3") or "").strip(),
                "codigo": str(fila.get("Número combinado") or "").strip(),
            }
            if not item["codigo"]:
                item["codigo"] = codigo_dotacion_cuadrilla(item)
            posiciones[posicion] = item
        if lista.get("disponible"):
            for cuadrilla, item_lista in (lista.get("posiciones") or {}).items():
                posiciones[cuadrilla] = _item_personal_limpio(item_lista)
        extras = []
        vistos = set()
        for nombre in apoyos:
            clave = nombre.casefold()
            if clave not in vistos:
                extras.append(nombre)
                vistos.add(clave)
        for _, fila in extras_ed.iterrows():
            nombre = str(fila.get("Nombre") or "").strip()
            clave = nombre.casefold()
            if nombre and clave not in vistos:
                extras.append(nombre)
                vistos.add(clave)
        guardar_personal_cuadrillas({"posiciones": posiciones, "extras": extras})
        st.success(f"Dotación actualizada en {PERSONAL_FILE}. La captura usará estos titulares.")
        st.rerun()
    if cal.get("disponible"):
        st.markdown("**Registro de jornadas por cuadrilla**")
        st.caption(f"{cal.get('titulo') or 'Calendario de asistencia'} · `{CALENDARIO_ASISTENCIA_FILE}`.")
        filas_cal = []
        for item in (cal.get("cuadrillas") or {}).values():
            filas_cal.append(
                {
                    "Cuadrilla": etiqueta_cuadrilla_historial(item.get("nombre") or ""),
                    "Realizadas": item.get("realizadas") if item.get("realizadas") is not None else "—",
                    "Total": item.get("total") if item.get("total") is not None else "—",
                    "Avance": avance_texto(item.get("avance")) or "—",
                }
            )
        st.dataframe(pd.DataFrame(filas_cal), hide_index=True, use_container_width=True)


def render_control_permisos() -> None:
    st.markdown('<div class="section-title">Control de permisos</div>', unsafe_allow_html=True)
    st.caption(
        "Registro local de accesos (permisos.json). Define quién puede operar como campo o como administración. "
        "Aún no hay inicio de sesión; el selector de rol de la barra lateral es la puerta de entrada actual."
    )
    df = pd.DataFrame(st.session_state.permisos)
    if df.empty:
        df = pd.DataFrame(PERMISOS_DEFAULT)
    for col in ["usuario", "nombre", "rol", "activo"]:
        if col not in df.columns:
            df[col] = False if col == "activo" else ""
    df = df[["usuario", "nombre", "rol", "activo"]]
    df["activo"] = df["activo"].astype(bool)

    editado = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="editor_permisos",
        column_config={
            "usuario": st.column_config.TextColumn("Usuario", required=True),
            "nombre": st.column_config.TextColumn("Nombre"),
            "rol": st.column_config.SelectboxColumn("Rol", options=ROLES),
            "activo": st.column_config.CheckboxColumn("Activo"),
        },
    )
    if st.button("Guardar permisos", type="primary"):
        registros = []
        vistos = set()
        for _, fila in editado.iterrows():
            usuario = str(fila.get("usuario", "")).strip()
            if not usuario or usuario.lower() in {"nan", "none"}:
                continue
            if usuario in vistos:
                continue
            vistos.add(usuario)
            rol = str(fila.get("rol", "")).strip()
            if rol not in ROLES:
                rol = ROL_OPERATIVO
            registros.append(
                {
                    "usuario": usuario,
                    "nombre": str(fila.get("nombre", "")).strip(),
                    "rol": rol,
                    "activo": bool(fila.get("activo", False)),
                }
            )
        if not registros:
            st.warning("Deja al menos un usuario activo.")
            return
        guardar_permisos(registros)
        st.success("Permisos actualizados.")
        st.rerun()


ESTATUS_OPS_EXCEL = ("Cumplido", "No Cumplido")


def estatus_guardado_actividad(guardado: dict) -> str:
    estatus = guardado.get("estatus")
    if estatus in ESTATUS_OPS_EXCEL:
        return estatus
    if guardado.get("realizado"):
        return "Cumplido"
    return "No Cumplido"


def _texto_observacion_editor(valor) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor).strip()


def _registro_control_vacio(registro: dict) -> bool:
    if isinstance(registro.get("evidencia_campo"), dict) and registro.get("evidencia_campo"):
        return False
    return (
        registro.get("estatus", "No Cumplido") != "Cumplido"
        and not registro.get("observaciones")
        and not registro.get("realizado")
    )


_RE_CODIGO_OPS = re.compile(r"[A-Za-z][A-Za-z0-9]*(?:-[A-Za-z0-9]+){1,6}")
EXTENSIONES_FOTO_OPS = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
COLUMNAS_TEXTO_BITACORA_OPS = (
    "CUADRILLA",
    "SECUENCIA",
    "UBICACIÓN",
    "EQUIPO/UBICACIÓN",
    "POZO",
    "SECTOR",
    "DESCRIPCION DE LA ACTIVIDAD",
    "DESCRIPCIÓN DE LA ACTIVIDAD",
    "ACTIVIDAD CORTA",
    "TIPO DE ACTIVIDAD",
    "COMENTARIOS",
)


def _compacto_busqueda(texto: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", _sin_acentos(texto).casefold())


def agujas_evidencia_ops(codigo: str, equipo: str) -> list[str]:
    agujas: list[str] = []

    def agregar(texto: str) -> None:
        compacto = _compacto_busqueda(texto)
        if len(compacto) < 4 or compacto in agujas:
            return
        if not re.search(r"[a-z]", compacto) or not re.search(r"\d", compacto):
            return
        agujas.append(compacto)

    agregar(codigo)
    for texto in (codigo, equipo):
        for coincidencia in _RE_CODIGO_OPS.findall(str(texto or "")):
            agregar(coincidencia)
    return agujas


def texto_contiene_agujas(texto: str, agujas: list[str]) -> bool:
    if not agujas:
        return False
    plano = re.sub(r"[^a-z0-9]+", " ", _sin_acentos(texto).casefold())
    partes = [parte for parte in plano.split() if parte]
    if not partes:
        return False
    unidos = set(partes)
    for largo in (2, 3, 4):
        for inicio in range(len(partes) - largo + 1):
            unidos.add("".join(partes[inicio : inicio + largo]))
    for aguja in agujas:
        if aguja in unidos:
            return True
        if len(aguja) < 5:
            continue
        for parte in partes:
            if parte == aguja or not parte.endswith(aguja):
                continue
            prefijo = parte[: -len(aguja)]
            if prefijo.isalpha() and len(prefijo) <= 2:
                return True
    return False


def _casilla_marcada(valor) -> bool:
    if valor is None:
        return False
    try:
        if pd.isna(valor):
            return False
    except (TypeError, ValueError):
        pass
    return bool(valor)


def firma_evidencia_octubre() -> tuple:
    carpeta = EVIDENCIA_OCTUBRE_DIR
    if not os.path.isdir(carpeta):
        return (0, 0.0)
    total = 0
    reciente = 0.0
    for raiz, dirs, archivos in os.walk(carpeta):
        dirs[:] = [nombre for nombre in dirs if not nombre.startswith(".")]
        for nombre in archivos:
            if nombre.startswith("~$"):
                continue
            total += 1
            try:
                reciente = max(reciente, os.path.getmtime(os.path.join(raiz, nombre)))
            except OSError:
                continue
    return (total, reciente)


@st.cache_data(show_spinner=False)
def indice_fotos_octubre(firma: tuple) -> list[dict]:
    _ = firma
    fotos = []
    carpeta = EVIDENCIA_OCTUBRE_DIR
    if not os.path.isdir(carpeta):
        return fotos
    for raiz, dirs, archivos in os.walk(carpeta):
        dirs[:] = [nombre for nombre in dirs if not nombre.startswith(".")]
        for nombre in archivos:
            if nombre.startswith("~$"):
                continue
            if os.path.splitext(nombre)[1].lower() not in EXTENSIONES_FOTO_OPS:
                continue
            ruta = os.path.join(raiz, nombre)
            if not os.path.isfile(ruta):
                continue
            rel = os.path.relpath(ruta, carpeta).replace("\\", "/")
            fotos.append({"ruta": ruta.replace("\\", "/"), "rel": rel, "nombre": nombre})
    return fotos


def _texto_fila_bitacora(fila) -> str:
    return " ".join(str(fila.get(col) or "") for col in COLUMNAS_TEXTO_BITACORA_OPS)


def buscar_evidencias_ops(
    codigo: str,
    equipo: str,
    df_bitacora: pd.DataFrame | None,
    anio: int,
    mes: int,
) -> dict:
    agujas = agujas_evidencia_ops(codigo, equipo)
    fotos_directas = []
    if agujas:
        for foto in indice_fotos_octubre(firma_evidencia_octubre()):
            if texto_contiene_agujas(foto["rel"], agujas):
                fotos_directas.append(foto["ruta"])
    bitacoras = []
    if agujas and df_bitacora is not None and not df_bitacora.empty:
        for _, fila in df_bitacora.iterrows():
            dia = parsear_fecha_valor(fila.get("FECHA"))
            if dia is None or dia.year != int(anio) or dia.month != int(mes):
                continue
            if not texto_contiene_agujas(_texto_fila_bitacora(fila), agujas):
                continue
            fotos = [ruta for ruta in rutas_fotos_fila(fila) if ruta and os.path.isfile(ruta)]
            ubicacion = str(fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("POZO") or "").strip()
            descripcion = str(
                fila.get("DESCRIPCION DE LA ACTIVIDAD")
                or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD")
                or fila.get("ACTIVIDAD CORTA")
                or ""
            ).strip()
            bitacoras.append(
                {
                    "fecha": dia,
                    "cuadrilla": str(fila.get("CUADRILLA") or "").strip(),
                    "secuencia": str(fila.get("SECUENCIA") or "").strip(),
                    "ubicacion": ubicacion,
                    "descripcion": descripcion,
                    "fotos": fotos,
                }
            )
    bitacoras.sort(
        key=lambda item: (item["fecha"], orden_cuadrilla_oficial(item["cuadrilla"]), item["secuencia"])
    )
    fotos = []
    vistos: set[str] = set()
    for registro in bitacoras:
        for ruta in registro["fotos"]:
            if ruta not in vistos:
                fotos.append(ruta)
                vistos.add(ruta)
    for ruta in fotos_directas:
        if ruta not in vistos:
            fotos.append(ruta)
            vistos.add(ruta)
    pdfs = []
    if bitacoras:
        por_fecha = {reporte["fecha"]: reporte for reporte in listar_reportes_ot_diarios()}
        vistos_pdf: set[date] = set()
        for registro in bitacoras:
            dia = registro["fecha"]
            reporte = por_fecha.get(dia)
            if reporte and dia not in vistos_pdf:
                pdfs.append(reporte)
                vistos_pdf.add(dia)
    return {"agujas": agujas, "bitacoras": bitacoras, "fotos": fotos, "pdfs": pdfs}


def _hora_foto_exif(ruta: str) -> str:
    try:
        from PIL import Image
        from PIL.ExifTags import TAGS
    except Exception:
        return ""
    try:
        with Image.open(ruta) as imagen:
            exif = imagen.getexif()
            if not exif:
                return ""
            bloques = [exif]
            try:
                sub = exif.get_ifd(0x8769)
                if sub:
                    bloques.append(sub)
            except Exception:
                pass
            preferida = ""
            respaldo = ""
            for bloque in bloques:
                for clave, valor in bloque.items():
                    nombre = TAGS.get(clave, clave)
                    if nombre not in {"DateTimeOriginal", "DateTimeDigitized", "DateTime"}:
                        continue
                    coincidencia = re.search(r"(?:^|\s)(\d{1,2}):(\d{2})(?::\d{2})?", str(valor or ""))
                    if not coincidencia:
                        continue
                    hora = int(coincidencia.group(1))
                    minuto = int(coincidencia.group(2))
                    if not (0 <= hora <= 23 and 0 <= minuto <= 59):
                        continue
                    texto = f"{hora:02d}:{minuto:02d}"
                    if nombre == "DateTimeOriginal":
                        preferida = texto
                    elif not respaldo:
                        respaldo = texto
            return preferida or respaldo
    except Exception:
        return ""


@st.cache_data(show_spinner=False)
def _hora_foto_exif_cache(ruta: str, mtime_ms: int) -> str:
    _ = mtime_ms
    return _hora_foto_exif(ruta)


def hora_reporte_campo(fotos: list[str]) -> str:
    for ruta in fotos[:3]:
        if not ruta or not os.path.isfile(ruta):
            continue
        try:
            marca = _hora_foto_exif_cache(ruta, int(os.path.getmtime(ruta) * 1000))
        except OSError:
            continue
        if marca:
            return marca
    return ""


def texto_vinculo_evidencia(evidencia) -> str:
    if not isinstance(evidencia, dict) or not evidencia:
        return ""
    fecha = parsear_fecha_valor(evidencia.get("fecha"))
    fecha_txt = fecha.strftime("%d/%m/%Y") if fecha else str(evidencia.get("fecha") or "").strip()
    cuadrilla = etiqueta_cuadrilla_historial(str(evidencia.get("cuadrilla") or ""))
    hora = str(evidencia.get("hora") or "").strip()
    secuencia = str(evidencia.get("secuencia") or "").strip()
    partes = [parte for parte in [fecha_txt, cuadrilla, hora, f"Sec. {secuencia}" if secuencia else ""] if parte and parte != "Sin cuadrilla"]
    if cuadrilla == "Sin cuadrilla" and not any(partes):
        return ""
    if cuadrilla == "Sin cuadrilla":
        partes.insert(1, cuadrilla)
    return " · ".join(partes)


def cuadrillas_activas_en_fecha(df: pd.DataFrame | None, dia: date) -> list[dict]:
    if df is None or df.empty or dia is None:
        return []
    del_dia = bitacora_del_dia(bitacora_con_fechas(df), dia)
    if del_dia.empty or "CUADRILLA" not in del_dia.columns:
        return []
    vistos: dict[str, dict] = {}
    for _, fila in del_dia.iterrows():
        nombre = str(fila.get("CUADRILLA") or "").strip()
        clave = clave_cuadrilla_oficial(nombre) or "__sin_cuadrilla__"
        item = vistos.get(clave)
        if item is None:
            vistos[clave] = {
                "cuadrilla": etiqueta_cuadrilla_historial(nombre),
                "reportes": 1,
                "clave": clave,
            }
        else:
            item["reportes"] += 1
    return sorted(vistos.values(), key=lambda item: orden_cuadrilla_oficial(item["cuadrilla"]))


def reportes_cuadrilla_en_fecha(df: pd.DataFrame | None, dia: date, cuadrilla: str) -> list[dict]:
    if df is None or df.empty or dia is None or cuadrilla is None:
        return []
    del_dia = bitacora_del_dia(bitacora_con_fechas(df), dia)
    bloque = bloque_historial_cuadrilla(del_dia, cuadrilla)
    if bloque.empty:
        return []
    reportes = []
    for fila in filas_secuencia_historial(bloque):
        nombre = str(fila.get("CUADRILLA") or cuadrilla).strip() or cuadrilla
        fotos = [
            str(ruta).replace("\\", "/")
            for ruta in fotos_actividad_ot(fila, dia, nombre)
            if ruta and os.path.isfile(str(ruta))
        ]
        descripcion = texto_visible(
            fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD"),
            "Sin descripción",
        )
        ubicacion = etiqueta_ubicacion_fila(fila) or texto_visible(
            fila.get("UBICACIÓN") or fila.get("POZO") or fila.get("SECTOR") or fila.get("EQUIPO/UBICACIÓN"),
            "",
        )
        secuencia = texto_secuencia_limpio(fila.get("SECUENCIA"))
        reportes.append(
            {
                "id_registro": id_registro_fila(fila),
                "secuencia": secuencia,
                "hora": hora_reporte_campo(fotos),
                "descripcion": descripcion,
                "ubicacion": ubicacion,
                "cuadrilla": nombre,
                "fotos": fotos,
            }
        )
    return reportes


def guardar_cumplimiento_con_evidencia(clave: str, evidencia: dict) -> None:
    control = leer_control_programa_octubre()
    previo = control.get(clave) if isinstance(control.get(clave), dict) else {}
    nuevo = dict(previo or {})
    nuevo["estatus"] = "Cumplido"
    nuevo["realizado"] = True
    nuevo["evidencia_campo"] = evidencia
    control[clave] = nuevo
    guardar_control_programa_octubre(control)


def render_vinculo_evidencia_guardado(vinculo: dict | None) -> None:
    if not isinstance(vinculo, dict) or not vinculo:
        return
    st.markdown("**Evidencia de campo vinculada**")
    resumen = texto_vinculo_evidencia(vinculo)
    if resumen:
        st.caption(resumen)
    descripcion = str(vinculo.get("descripcion") or "").strip()
    if descripcion:
        st.write(descripcion)
    ubicacion = str(vinculo.get("ubicacion") or "").strip()
    if ubicacion:
        st.caption(ubicacion)
    fotos = [str(ruta) for ruta in (vinculo.get("fotos") or []) if ruta and os.path.isfile(str(ruta))]
    if fotos:
        render_grid_evidencia_actividad(fotos)
    else:
        st.caption("Este vínculo no tiene fotografías disponibles en disco.")


def _cerrar_dialogo_validar_ops() -> None:
    if st.session_state.pop("ops_val_confirmado", False):
        st.session_state.ops_validar_abierta = None
        return
    abierto = st.session_state.get("ops_validar_abierta") or {}
    st.session_state.ops_validar_cancelada = abierto.get("id")
    st.session_state.ops_editor_epoch = int(st.session_state.get("ops_editor_epoch") or 0) + 1
    st.session_state.ops_val_paso = 1
    st.session_state.ops_val_id = None
    st.session_state.pop("ops_val_cuadrilla", None)
    st.session_state.ops_validar_abierta = None


def _bitacora_dialogo_ops() -> pd.DataFrame:
    bitacora = st.session_state.get("ops_df_bitacora")
    if isinstance(bitacora, pd.DataFrame):
        return bitacora
    return cargar_bitacora()


def _refrescar_paso_dialogo() -> None:
    try:
        st.rerun(scope="fragment")
    except Exception as exc:
        if exc.__class__.__name__ != "StreamlitAPIException":
            raise
        st.rerun()


@st.dialog("Vincular evidencia y marcar cumplido", width="large", on_dismiss=_cerrar_dialogo_validar_ops)
def dialogo_vincular_cumplido_ops() -> None:
    datos = st.session_state.get("ops_validar_abierta") or {}
    if not datos:
        return
    if st.session_state.get("ops_val_id") != datos.get("id"):
        st.session_state.ops_val_id = datos.get("id")
        st.session_state.ops_val_paso = 1
    paso = int(st.session_state.get("ops_val_paso") or 1)
    actividad_id = str(datos.get("id") or "actividad")
    with st.container(border=True):
        nro = str(datos.get("nro") or "—")
        codigo = str(datos.get("codigo") or "Sin código")
        st.markdown(f"**Nro {html.escape(nro)} · {html.escape(codigo)}**")
        equipo = str(datos.get("equipo") or "").strip()
        fecha_prog = str(datos.get("fecha") or "").strip()
        st.caption(html.escape(" · ".join(parte for parte in [equipo, f"programada el {fecha_prog}" if fecha_prog else ""] if parte)))
        actividad = str(datos.get("actividad") or "").strip()
        if actividad:
            st.write(actividad)
    st.caption("El estatus pasa a Cumplido solo al confirmar el reporte de campo.")

    if paso <= 1:
        st.markdown("**Paso 1 de 3 · Fecha de ejecución**")
        fecha_default = parsear_fecha_valor(datos.get("fecha_iso") or datos.get("fecha")) or date.today()
        fecha = st.date_input(
            "Fecha en que se ejecutó el trabajo en campo",
            value=fecha_default,
            min_value=date(2024, 1, 1),
            max_value=date(2035, 12, 31),
            format="DD/MM/YYYY",
            key=f"ops_val_fecha_input_{actividad_id}",
        )
        if st.button("Continuar con las cuadrillas", type="primary", key=f"ops_val_sig_fecha_{actividad_id}"):
            elegida = fecha if isinstance(fecha, date) else fecha_default
            st.session_state.ops_val_fecha = elegida.isoformat()
            st.session_state.ops_val_paso = 2
            st.session_state.pop("ops_val_cuadrilla", None)
            _refrescar_paso_dialogo()
        return

    fecha = parsear_fecha_valor(st.session_state.get("ops_val_fecha"))
    if fecha is None:
        st.session_state.ops_val_paso = 1
        _refrescar_paso_dialogo()

    if paso == 2:
        st.markdown(f"**Paso 2 de 3 · Cuadrilla del {fecha.strftime('%d/%m/%Y')}**")
        if st.button("← Cambiar fecha", key=f"ops_val_atras_fecha_{actividad_id}"):
            st.session_state.ops_val_paso = 1
            st.session_state.pop("ops_val_cuadrilla", None)
            _refrescar_paso_dialogo()
        cuadrillas = cuadrillas_activas_en_fecha(_bitacora_dialogo_ops(), fecha)
        if not cuadrillas:
            st.info("Ninguna cuadrilla registró actividad en la bitácora de campo en esa fecha.")
            return
        conteo = {item["cuadrilla"]: item["reportes"] for item in cuadrillas}
        cuadrilla = st.selectbox(
            "Cuadrilla con reporte ese día",
            list(conteo),
            format_func=lambda nombre: (
                f"{etiqueta_cuadrilla_historial(nombre) or 'Sin cuadrilla'} · "
                f"{conteo[nombre]} {'reporte' if conteo[nombre] == 1 else 'reportes'}"
            ),
            key=f"ops_val_cuadrilla_input_{actividad_id}_{fecha.isoformat()}",
        )
        if st.button("Ver actividades de la cuadrilla", type="primary", key=f"ops_val_sig_cuadrilla_{actividad_id}"):
            st.session_state.ops_val_cuadrilla = cuadrilla
            st.session_state.ops_val_paso = 3
            _refrescar_paso_dialogo()
        return

    if "ops_val_cuadrilla" not in st.session_state:
        st.session_state.ops_val_paso = 2
        _refrescar_paso_dialogo()
    cuadrilla = str(st.session_state.get("ops_val_cuadrilla") or "").strip()
    st.markdown(
        f"**Paso 3 de 3 · Actividad de {html.escape(etiqueta_cuadrilla_historial(cuadrilla))}**"
    )
    st.caption(f"Reportes del {fecha.strftime('%d/%m/%Y')}. La hora sale de la fotografía de campo.")
    volver_fecha, volver_cuadrilla = st.columns(2)
    with volver_fecha:
        if st.button("← Cambiar fecha", key=f"ops_val_atras_fecha_3_{actividad_id}", width="stretch"):
            st.session_state.ops_val_paso = 1
            st.session_state.pop("ops_val_cuadrilla", None)
            _refrescar_paso_dialogo()
    with volver_cuadrilla:
        if st.button("← Cambiar cuadrilla", key=f"ops_val_atras_cuadrilla_{actividad_id}", width="stretch"):
            st.session_state.ops_val_paso = 2
            _refrescar_paso_dialogo()

    reportes = reportes_cuadrilla_en_fecha(_bitacora_dialogo_ops(), fecha, cuadrilla)
    if not reportes:
        st.info("Esa cuadrilla no tiene actividades registradas en la fecha elegida.")
        return
    etiquetas = []
    for item in reportes:
        hora = item["hora"] or "Sin hora"
        secuencia = item["secuencia"] or "—"
        descripcion = item["descripcion"]
        if len(descripcion) > 72:
            descripcion = descripcion[:69] + "…"
        etiquetas.append(f"{hora} · Sec. {secuencia} · {descripcion} · {len(item['fotos'])} foto(s)")
    indice = st.selectbox(
        "Actividad / evidencia de la bitácora",
        list(range(len(reportes))),
        format_func=lambda i: etiquetas[i],
        key=f"ops_val_reporte_{actividad_id}_{fecha.isoformat()}_{clave_hist_cuadrilla(cuadrilla)}",
    )
    elegido = reportes[int(indice)]
    with st.container(border=True):
        st.markdown(
            f"**{html.escape(elegido['hora'] or 'Sin hora registrada')}** · "
            f"Secuencia {html.escape(elegido['secuencia'] or '—')}"
        )
        st.write(elegido["descripcion"])
        if elegido.get("ubicacion"):
            st.caption(elegido["ubicacion"])
        if elegido["fotos"]:
            render_grid_evidencia_actividad(elegido["fotos"])
        else:
            st.caption("Este reporte no tiene fotografías adjuntas.")
    if st.button("Confirmar vínculo y marcar Cumplido", type="primary", key=f"ops_val_confirmar_{actividad_id}"):
        evidencia = {
            "fecha": fecha.isoformat(),
            "cuadrilla": elegido.get("cuadrilla") or cuadrilla,
            "id_registro": elegido.get("id_registro") or "",
            "secuencia": elegido.get("secuencia") or "",
            "hora": elegido.get("hora") or "",
            "descripcion": elegido.get("descripcion") or "",
            "ubicacion": elegido.get("ubicacion") or "",
            "fotos": list(elegido.get("fotos") or []),
            "vinculado_en": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        guardar_cumplimiento_con_evidencia(actividad_id, evidencia)
        st.session_state.ops_validar_mensaje = (
            f"{codigo} quedó en Cumplido, vinculada a "
            f"{etiqueta_cuadrilla_historial(evidencia['cuadrilla'])} del {fecha.strftime('%d/%m/%Y')}."
        )
        st.session_state.ops_val_confirmado = True
        st.session_state.ops_editor_epoch = int(st.session_state.get("ops_editor_epoch") or 0) + 1
        st.session_state.ops_val_paso = 1
        st.session_state.ops_val_id = None
        st.session_state.ops_validar_abierta = None
        st.rerun()


def render_panel_evidencias_ops(
    codigo: str,
    equipo: str,
    actividad: str,
    fecha_programada: str,
    df_bitacora: pd.DataFrame | None,
    anio: int,
    mes: int,
    vinculo: dict | None = None,
) -> None:
    etiqueta = etiqueta_mes(clave_mes(anio, mes))
    titulo = codigo or equipo or "esta actividad"
    st.markdown(f"##### Evidencias · {titulo}")
    if isinstance(vinculo, dict) and vinculo:
        render_vinculo_evidencia_guardado(vinculo)
        st.divider()
    st.caption(
        f"{equipo} · programada el {fecha_programada}. "
        f"Se buscó el código en `{EVIDENCIA_OCTUBRE_DIR}` y en la bitácora de {etiqueta}."
    )
    if actividad:
        st.caption(actividad)
    resultado = buscar_evidencias_ops(codigo, equipo, df_bitacora, anio, mes)
    tiene_vinculo = isinstance(vinculo, dict) and bool(vinculo)
    if not resultado["agujas"]:
        if not tiene_vinculo:
            st.info("Esta fila no tiene un código o equipo con el cual buscar evidencias.")
        return
    if not resultado["bitacoras"] and not resultado["fotos"]:
        if not tiene_vinculo:
            st.info(
                f"No hay bitácoras ni fotografías que coincidan con {titulo} "
                f"en {etiqueta} ni en {EVIDENCIA_OCTUBRE_DIR}."
            )
        return
    if resultado["bitacoras"]:
        st.markdown("**Bitácora coincidente**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Fecha": item["fecha"].strftime("%d/%m/%Y"),
                        "Cuadrilla": item["cuadrilla"],
                        "Secuencia": item["secuencia"],
                        "Ubicación": item["ubicacion"],
                        "Descripción": item["descripcion"],
                    }
                    for item in resultado["bitacoras"]
                ]
            ),
            hide_index=True,
            width="stretch",
        )
    if resultado["pdfs"]:
        st.markdown("**Orden de trabajo del día**")
        columnas_pdf = st.columns(min(len(resultado["pdfs"]), 3))
        for indice, reporte in enumerate(resultado["pdfs"]):
            with columnas_pdf[indice % len(columnas_pdf)]:
                ruta_ot = str(reporte.get("ruta") or "")
                if not ruta_ot or not os.path.isfile(ruta_ot):
                    st.caption(
                        f"OT · {reporte['fecha'].strftime('%d/%m/%Y')} ya se descargó y se retiró del servidor."
                    )
                    continue
                with open(ruta_ot, "rb") as archivo:
                    datos_ot = archivo.read()
                extra = {}
                if persistencia.es_archivo_pesado(ruta_ot):
                    extra = {
                        "on_click": _al_descargar_pdf_ot,
                        "args": (ruta_ot, reporte["fecha"].isoformat(), reporte["nombre"]),
                    }
                st.download_button(
                    f"OT · {reporte['fecha'].strftime('%d/%m/%Y')}",
                    data=datos_ot,
                    file_name=reporte["nombre"],
                    mime="application/pdf",
                    key=f"ot_ops_{anio}_{mes}_{_compacto_busqueda(titulo)}_{reporte['fecha'].isoformat()}",
                    width="stretch",
                    **extra,
                )
    fotos = resultado["fotos"]
    if not fotos:
        st.caption("La bitácora coincide, y esa secuencia no tiene fotografías en la carpeta de octubre.")
        return
    mostradas = fotos[:18]
    st.markdown(f"**Fotografías · {len(fotos)}**")
    if len(fotos) > len(mostradas):
        st.caption(f"Se muestran {len(mostradas)} de {len(fotos)}.")
    for inicio in range(0, len(mostradas), 3):
        columnas = st.columns(3, gap="small")
        for columna, ruta in zip(columnas, mostradas[inicio : inicio + 3]):
            with columna:
                st.image(ruta, width="stretch")
                st.caption(os.path.relpath(ruta, EVIDENCIA_OCTUBRE_DIR).replace("\\", "/"))


def _cerrar_dialogo_evidencias_ops() -> None:
    abierto = st.session_state.get("ops_evidencia_abierta") or {}
    st.session_state.ops_evidencia_cerrada = abierto.get("id")
    st.session_state.ops_evidencia_abierta = None


@st.dialog("Evidencias de la actividad", width="large", on_dismiss=_cerrar_dialogo_evidencias_ops)
def dialogo_evidencias_ops() -> None:
    datos = st.session_state.get("ops_evidencia_abierta") or {}
    if not datos:
        return
    bitacora = st.session_state.get("ops_df_bitacora")
    if not isinstance(bitacora, pd.DataFrame):
        bitacora = cargar_bitacora()
    vinculo = datos.get("vinculo") if isinstance(datos.get("vinculo"), dict) else None
    render_panel_evidencias_ops(
        str(datos.get("codigo") or ""),
        str(datos.get("equipo") or ""),
        str(datos.get("actividad") or ""),
        str(datos.get("fecha") or ""),
        bitacora,
        int(datos.get("anio") or date.today().year),
        int(datos.get("mes") or date.today().month),
        vinculo=vinculo,
    )


def render_interruptor_edicion_ops(*, contexto: str = "ops") -> bool:
    with st.container(border=True):
        col_ctrl, col_ayuda = st.columns([1.25, 2.3], vertical_alignment="center")
        with col_ctrl:
            activo = bool(
                st.toggle(
                    "Modo Edición Administrador",
                    key="ops_modo_edicion_admin",
                    help=(
                        "Actívalo para registrar una actividad y subir sus fotos."
                        if contexto == "bitacora"
                        else "Actívalo para marcar una actividad como Cumplido y vincular el reporte de campo."
                    ),
                )
            )
        with col_ayuda:
            if contexto == "bitacora" and activo:
                ayuda = (
                    "<b>Edición activa.</b> Completa el formulario de abajo para guardar "
                    "la actividad en la bitácora y adjuntar las fotografías."
                )
            elif contexto == "bitacora":
                ayuda = (
                    "Activa el modo para abrir el formulario de registro. "
                    "Ahí se captura la actividad del día y la evidencia fotográfica."
                )
            elif activo:
                ayuda = (
                    "<b>Edición activa.</b> Marca <b>Vincular</b> en la fila, o cambia el estatus a "
                    "<b>Cumplido</b>, para abrir el asistente: fecha, cuadrilla y evidencia."
                )
            else:
                ayuda = (
                    "La tabla está en consulta. "
                    "Activa el modo para habilitar el estatus y vincular la evidencia de campo."
                )
            st.markdown(f'<p class="ops-modo-ayuda">{ayuda}</p>', unsafe_allow_html=True)
    return activo


def render_programados_ops_excel(
    anio: int,
    mes: int,
    df_bitacora: pd.DataFrame | None = None,
    modo_edicion: bool = False,
) -> None:
    carpeta = directorio_programas_mtto()
    etiqueta = etiqueta_mes(clave_mes(anio, mes))
    st.caption(
        f"Actividades de {etiqueta} leídas de los Excel en {carpeta}. "
        "El estatus y las observaciones se guardan en este equipo."
    )
    if not os.path.isdir(carpeta):
        st.info("Coloca los archivos .xlsx en la carpeta PROGRAMAS MTTO OCTUBRE.")
        return
    archivos = listar_archivos_programa_octubre()
    if not archivos:
        st.info(f"No hay archivos Excel en {carpeta}.")
        return

    actividades = cargar_actividades_octubre(firmas_programa_octubre(), anio, mes)
    if not actividades:
        st.info(f"No hay actividades de {etiqueta} en los Excel de {carpeta}.")
        return

    especialidades = [
        nombre
        for nombre in ESPECIALIDADES_OCTUBRE
        if any(item["especialidad"] == nombre for item in actividades)
    ]
    extra = sorted({item["especialidad"] for item in actividades} - set(especialidades))
    especialidades.extend(extra)
    pestanas_nombre = ["Todos"] + especialidades
    aviso = st.session_state.pop("ops_validar_mensaje", None)
    if aviso:
        st.success(aviso)
    modo_edicion = bool(modo_edicion)

    control = leer_control_programa_octubre()
    evidencia_elegida = None
    pendiente_validar = None
    epoch = int(st.session_state.get("ops_editor_epoch") or 0)
    pestanas = st.tabs(pestanas_nombre)
    for pestana, nombre_pestana in zip(pestanas, pestanas_nombre):
        with pestana:
            filas = (
                actividades
                if nombre_pestana == "Todos"
                else [item for item in actividades if item["especialidad"] == nombre_pestana]
            )
            origenes = sorted({item["programa"] for item in filas})
            origen = "Todos"
            if len(origenes) > 1:
                origen = st.selectbox(
                    "Archivo o especialidad",
                    ["Todos"] + origenes,
                    key=f"origen_ops_{anio}_{mes}_{nombre_pestana}",
                )
                if origen != "Todos":
                    filas = [item for item in filas if item["programa"] == origen]
            registros = []
            for item in filas:
                clave = clave_control_actividad(item)
                guardado = control.get(clave) or {}
                estatus_fila = estatus_guardado_actividad(guardado)
                fila_tabla = {
                        "_id": clave,
                        "Ver Evidencias": False,
                        "Estatus": estatus_fila,
                        "Vínculo": texto_vinculo_evidencia(guardado.get("evidencia_campo"))
                        if estatus_fila == "Cumplido"
                        else "",
                }
                if modo_edicion:
                    fila_tabla["Vincular"] = False
                registros.append(
                    fila_tabla | {
                        "Observaciones": str(guardado.get("observaciones") or ""),
                        "Nro": item["nro"],
                        "Cód": item["codigo"],
                        "Equipo": item["equipo"],
                        "Actividad": item["actividad"],
                        "Tipo Mantto.": item["tipo"],
                        "Fecha Programada": datetime.fromisoformat(item["fecha"]).strftime("%d/%m/%Y"),
                    }
                )
            tabla = pd.DataFrame(registros)
            cumplidos = int((tabla["Estatus"] == "Cumplido").sum()) if not tabla.empty else 0
            if modo_edicion:
                ayuda_tabla = "Marca Vincular, o cambia Estatus a Cumplido, para abrir el asistente de evidencia."
            else:
                ayuda_tabla = "Marca Ver Evidencias en una fila para abrir sus fotos y la bitácora."
            st.caption(
                f"{len(tabla)} actividades · {cumplidos} cumplidas · {len(tabla) - cumplidos} no cumplidas. "
                f"{ayuda_tabla}"
            )
            columnas_bloqueadas = ["Vínculo", "Nro", "Cód", "Equipo", "Actividad", "Tipo Mantto.", "Fecha Programada"]
            if not modo_edicion:
                columnas_bloqueadas.append("Estatus")
            orden_columnas = [
                "Ver Evidencias",
                "Estatus",
                "Vínculo",
                "Observaciones",
                "Nro",
                "Cód",
                "Equipo",
                "Actividad",
                "Tipo Mantto.",
                "Fecha Programada",
            ]
            if modo_edicion:
                orden_columnas.insert(0, "Vincular")
            config_columnas = {
                "Ver Evidencias": st.column_config.CheckboxColumn("Ver Evidencias", default=False),
                "Estatus": st.column_config.SelectboxColumn(
                    "Estatus",
                    options=list(ESTATUS_OPS_EXCEL),
                    required=True,
                    help="En modo edición, Cumplido abre el asistente de evidencia antes de guardarse.",
                ),
                "Vínculo": st.column_config.TextColumn(
                    "Vínculo",
                    help="Fecha, cuadrilla y reporte de campo ligados al cumplimiento.",
                ),
                "Observaciones": st.column_config.TextColumn("Observaciones"),
                "Actividad": st.column_config.TextColumn("Actividad", width="large"),
            }
            if modo_edicion:
                config_columnas["Vincular"] = st.column_config.CheckboxColumn(
                    "Vincular",
                    default=False,
                    help="Abre el asistente para marcar la actividad como Cumplido y ligar la evidencia.",
                )
            editada = st.data_editor(
                tabla,
                column_order=orden_columnas,
                column_config=config_columnas,
                disabled=columnas_bloqueadas,
                hide_index=True,
                width="stretch",
                key=f"editor_ops_{anio}_{mes}_{nombre_pestana}_{origen}_{int(modo_edicion)}_{epoch}",
            )
            cambio = False
            for _, fila in editada.iterrows():
                clave = str(fila["_id"])
                previo = control.get(clave) if isinstance(control.get(clave), dict) else {}
                estatus_editor = fila["Estatus"] if fila["Estatus"] in ESTATUS_OPS_EXCEL else "No Cumplido"
                estatus_previo = estatus_guardado_actividad(previo)
                marcar_cumplido = bool(modo_edicion) and estatus_editor == "Cumplido" and estatus_previo != "Cumplido"
                vincular_fila = bool(modo_edicion) and _casilla_marcada(fila.get("Vincular"))
                estatus = estatus_previo if marcar_cumplido else estatus_editor
                if (marcar_cumplido or vincular_fila) and pendiente_validar is None:
                    fecha_prog = parsear_fecha_valor(fila.get("Fecha Programada"))
                    pendiente_validar = {
                        "id": clave,
                        "codigo": str(fila.get("Cód") or ""),
                        "equipo": str(fila.get("Equipo") or ""),
                        "actividad": str(fila.get("Actividad") or ""),
                        "nro": str(fila.get("Nro") or ""),
                        "fecha": str(fila.get("Fecha Programada") or ""),
                        "fecha_iso": fecha_prog.isoformat() if fecha_prog else "",
                        "anio": int(anio),
                        "mes": int(mes),
                    }
                nuevo = dict(previo)
                nuevo["estatus"] = estatus
                nuevo["observaciones"] = _texto_observacion_editor(fila["Observaciones"])
                nuevo["realizado"] = estatus == "Cumplido"
                if not marcar_cumplido and estatus != "Cumplido":
                    nuevo.pop("evidencia_campo", None)
                if control.get(clave) == nuevo:
                    continue
                if _registro_control_vacio(nuevo):
                    if clave in control:
                        control.pop(clave, None)
                        cambio = True
                    continue
                control[clave] = nuevo
                cambio = True
            if cambio:
                guardar_control_programa_octubre(control)
            marcadas = [fila for _, fila in editada.iterrows() if _casilla_marcada(fila.get("Ver Evidencias"))]
            if marcadas:
                fila_evidencia = marcadas[-1]
                guardado_ev = control.get(str(fila_evidencia.get("_id") or ""))
                vinculo = guardado_ev.get("evidencia_campo") if isinstance(guardado_ev, dict) else None
                evidencia_elegida = {
                    "id": str(fila_evidencia.get("_id") or ""),
                    "codigo": str(fila_evidencia.get("Cód") or ""),
                    "equipo": str(fila_evidencia.get("Equipo") or ""),
                    "actividad": str(fila_evidencia.get("Actividad") or ""),
                    "fecha": str(fila_evidencia.get("Fecha Programada") or ""),
                    "anio": int(anio),
                    "mes": int(mes),
                    "vinculo": vinculo if isinstance(vinculo, dict) else None,
                }
    cancelada = st.session_state.get("ops_validar_cancelada")
    if pendiente_validar and cancelada == pendiente_validar["id"]:
        pass
    elif pendiente_validar:
        st.session_state.ops_validar_cancelada = None
        st.session_state.ops_validar_abierta = pendiente_validar
        st.session_state.ops_df_bitacora = df_bitacora if isinstance(df_bitacora, pd.DataFrame) else cargar_bitacora()
        dialogo_vincular_cumplido_ops()
    elif cancelada:
        st.session_state.ops_validar_cancelada = None
    elif evidencia_elegida and st.session_state.get("ops_evidencia_cerrada") != evidencia_elegida["id"]:
        st.session_state.ops_evidencia_abierta = evidencia_elegida
        st.session_state.ops_df_bitacora = df_bitacora if isinstance(df_bitacora, pd.DataFrame) else cargar_bitacora()
        dialogo_evidencias_ops()
    elif not evidencia_elegida:
        st.session_state.ops_evidencia_cerrada = None
        st.session_state.ops_evidencia_abierta = None


def _items_insumos_registro() -> list[dict]:
    items: list[dict] = []
    vistos: set[str] = set()
    for item in catalogo_insumos_activos():
        nombre = str(item.get("nombre") or "").strip()
        if not nombre or nombre.casefold() in vistos:
            continue
        vistos.add(nombre.casefold())
        items.append(
            {
                "nombre": nombre,
                "tipo": tipo_insumo_item(item),
                "id": str(item.get("id") or slug_insumo(nombre)),
            }
        )
    for nombre in INSUMOS_REGISTRO_COMUNES:
        if nombre.casefold() in vistos:
            continue
        vistos.add(nombre.casefold())
        items.append({"nombre": nombre, "tipo": TIPO_INSUMO_CONSUMIBLE, "id": slug_insumo(nombre)})
    return items


def render_insumos_registro_bitacora(epoch: int) -> dict:
    st.caption(
        "Marca solo lo que se instaló o se gastó. "
        "Un componente rotativo pide la condición de la pieza nueva y el destino de la retirada."
    )
    items = _items_insumos_registro()
    nombres: list[str] = []
    rotativos: list[dict] = []
    grupos = [
        (
            TIPO_INSUMO_CONSUMIBLE,
            "Consumibles",
            "Bandas C-115, C-112, C-240, Micro V y el resto de material que se gasta.",
        ),
        (
            TIPO_INSUMO_ROTATIVO,
            "Componentes rotativos / reparables",
            "Bomba de agua, alternador y marcha / motor de arranque.",
        ),
    ]
    for tipo, titulo, ayuda in grupos:
        del_tipo = [item for item in items if item["tipo"] == tipo]
        if not del_tipo:
            continue
        st.markdown(f"**{titulo}**")
        st.caption(ayuda)
        columnas = st.columns(3, gap="medium")
        for indice, item in enumerate(del_tipo):
            with columnas[indice % 3]:
                slug = item["id"]
                if not st.checkbox(item["nombre"], key=f"bit_reg_insumo_{slug}_{epoch}"):
                    continue
                nombres.append(item["nombre"])
                if tipo != TIPO_INSUMO_ROTATIVO:
                    continue
                condicion = st.selectbox(
                    "Condición de la pieza instalada",
                    CONDICION_PIEZA_INSTALADA,
                    index=None,
                    placeholder="Seleccionar...",
                    key=f"bit_reg_cond_{slug}_{epoch}",
                )
                destino = st.selectbox(
                    "Destino de la pieza retirada",
                    DESTINO_PIEZA_RETIRADA,
                    index=None,
                    placeholder="Seleccionar...",
                    key=f"bit_reg_dest_{slug}_{epoch}",
                )
                rotativos.append(
                    {
                        "nombre": item["nombre"],
                        "tipo": TIPO_INSUMO_ROTATIVO,
                        "condicion_instalada": condicion or "",
                        "destino_retirada": destino or "",
                    }
                )
    usa_otros = st.checkbox("Otros insumos", key=f"bit_reg_insumo_otros_{epoch}")
    if usa_otros:
        detalle = st.text_input(
            "Detalle de otros insumos",
            placeholder="Escribe cualquier otro insumo utilizado",
            key=f"bit_reg_insumo_otros_txt_{epoch}",
        ).strip()
        nombres.append(detalle or "Otros insumos")
    if nombres or rotativos:
        resumen = list(nombres)
        resumen.extend(texto_pieza_rotativa(pieza) for pieza in rotativos)
        st.caption("Se guardará: **" + " · ".join(resumen) + "**")
    return {"nombres": nombres, "rotativos": rotativos}


def render_registro_actividad_bitacora() -> None:
    aviso = st.session_state.pop("bitacora_registro_ok", None)
    if aviso:
        st.success(aviso)
    epoch = int(st.session_state.get("bitacora_form_epoch") or 0)
    df_bitacora = bitacora_con_fechas(cargar_bitacora())
    opciones = catalogo_cuadrillas_historial(df_bitacora)
    otra = "Otra (escribir)"
    st.markdown('<div class="section-title">➕ Registrar Actividad en Bitácora</div>', unsafe_allow_html=True)
    st.caption(
        "El registro se agrega a la bitácora operativa. "
        "Las fotografías se cargan directo y quedan en la carpeta de la cuadrilla, la fecha y la secuencia."
    )

    col_u, col_o = st.columns(2, gap="medium")
    with col_u:
        with st.container(border=True):
            st.markdown('<div class="section-title">📍 Ubicación</div>', unsafe_allow_html=True)
            fecha = st.date_input(
                "Fecha",
                value=date.today(),
                format="DD/MM/YYYY",
                key=f"bit_reg_fecha_{epoch}",
                persist_state="session",
            )
            ubicacion_sel, ubicacion_nueva, tag_instalacion = render_selector_ubicacion_captura(
                f"bit_reg_{epoch}"
            )
            equipo_sel = render_selector_equipo_ubicacion(f"bit_reg_equipo_{epoch}", tag_instalacion)

    with col_o:
        with st.container(border=True):
            st.markdown('<div class="section-title">🔧 Operación</div>', unsafe_allow_html=True)
            clave_cuadrilla_reg = f"bit_reg_cuadrilla_{epoch}"
            if clave_cuadrilla_reg not in st.session_state:
                recordada = cuadrilla_recordada(fecha, opciones)
                if recordada:
                    st.session_state[clave_cuadrilla_reg] = recordada
            cuadrilla_sel = st.selectbox(
                "Cuadrilla",
                opciones + [otra],
                format_func=lambda nombre: nombre if nombre == otra else etiqueta_cuadrilla_historial(nombre),
                key=clave_cuadrilla_reg,
                persist_state="session",
            )
            cuadrilla_libre = ""
            if cuadrilla_sel == otra:
                cuadrilla_libre = st.text_input(
                    "Nombre de la cuadrilla",
                    placeholder="Escribe la cuadrilla",
                    key=f"bit_reg_cuadrilla_libre_{epoch}",
                    persist_state="session",
                )
            cuadrilla = cuadrilla_libre.strip() if cuadrilla_sel == otra else str(cuadrilla_sel or "").strip()
            siguiente = siguiente_secuencia_turno(fecha, cuadrilla) if cuadrilla else 1
            secuencia_n = st.number_input(
                "Secuencia del día",
                min_value=1,
                step=1,
                value=int(siguiente),
                key=(
                    f"bit_reg_secuencia_{epoch}_{fecha.isoformat()}_"
                    f"{clave_cuadrilla_oficial(cuadrilla) or 'sin'}"
                ),
            )
            st.caption(
                f"Siguiente consecutivo de esta cuadrilla en la fecha: "
                f"**{etiqueta_secuencia_consecutiva(int(secuencia_n))}**."
            )
            precarga = aplicar_precarga_jornada(
                f"_precarga_ficha_{epoch}",
                (fecha.isoformat(), cuadrilla),
                lambda: valores_precarga_ficha(epoch, fecha, cuadrilla),
            )
            aviso_precarga_jornada(precarga)
            clave_personal_reg = f"bit_reg_personal_{epoch}"
            if clave_personal_reg not in st.session_state and cuadrilla:
                ficha_recordada = str(memoria_jornada(fecha, cuadrilla).get("ficha") or "")
                if ficha_recordada:
                    st.session_state[clave_personal_reg] = ficha_recordada
            personal = st.text_input(
                "Ficha / Personal a cargo",
                placeholder="Ej. 1705-2107",
                key=clave_personal_reg,
                persist_state="session",
            )
            if cuadrilla:
                guardar_memoria_jornada(
                    fecha,
                    cuadrilla,
                    {"ficha": personal, "secuencia": int(secuencia_n)},
                )
            tipo_actividad, tipo_actividad_otro, actividad_corta_otro = render_selector_actividad_cuadrilla(
                f"bit_reg_act_{epoch}_{clave_cuadrilla_oficial(cuadrilla) or 'sin'}",
                cuadrilla,
            )
            descripcion = st.text_area(
                "Descripción de la actividad realizada",
                height=120,
                placeholder="Describe el trabajo realizado en campo",
                key=f"bit_reg_desc_{epoch}",
            )
            comentarios = st.text_area(
                "Comentarios / Observaciones",
                height=90,
                placeholder="Observaciones del trabajo",
                key=f"bit_reg_comentarios_{epoch}",
            )

    with st.container(border=True):
        st.markdown('<div class="section-title">🧪 Insumos utilizados</div>', unsafe_allow_html=True)
        seleccion_insumos = render_insumos_registro_bitacora(epoch)

    with st.container(border=True):
        st.markdown(
            '<div class="section-title">📷 Evidencia fotográfica</div>',
            unsafe_allow_html=True,
        )
        st.caption("Carga directa de la actividad. Hasta 9 imágenes se guardan en el registro al pulsar Guardar.")
        fotos = st.file_uploader(
            "Cargar fotografías",
            type=["jpg", "jpeg", "png"],
            accept_multiple_files=True,
            key=f"bit_reg_fotos_{epoch}",
        )
        if fotos and len(fotos) > MAX_FOTOS:
            st.warning("Solo se usarán las primeras 9 imágenes.")
        renderizar_galeria(fotos)

    st.markdown("")
    _, col_btn, _ = st.columns([1, 2, 1])
    with col_btn:
        guardar = st.button(
            "💾 Guardar Registro",
            type="primary",
            use_container_width=True,
            key=f"bit_reg_guardar_{epoch}",
        )
    if not guardar:
        return

    captura = resolver_captura_ubicacion(ubicacion_sel, ubicacion_nueva, tag_instalacion)
    ubicacion = captura["ubicacion"]
    tipo_actividad_final, actividad_corta_final = resolver_actividad_catalogo(
        cuadrilla, tipo_actividad, tipo_actividad_otro, actividad_corta_otro
    )
    if not cuadrilla:
        st.error("Escribe la cuadrilla que realizó la actividad.")
    elif not ubicacion:
        st.error("Completa la ubicación: elige la zona y el pozo del catálogo, o la instalación especial.")
    elif not str(descripcion or "").strip():
        st.error("Escribe la descripción de la actividad realizada.")
    elif nombres_actividad_cuadrilla(cuadrilla) and not tipo_actividad_final:
        st.error("Selecciona una actividad oficial de la cuadrilla.")
    else:
        error_rotativo = validar_piezas_rotativas(seleccion_insumos["rotativos"])
        if error_rotativo:
            st.error(error_rotativo)
            return
        consumibles = ", ".join(seleccion_insumos["nombres"])
        notas_rotativas = [texto_pieza_rotativa(pieza) for pieza in seleccion_insumos["rotativos"]]
        if notas_rotativas:
            bloque = "Piezas rotativas: " + " | ".join(notas_rotativas)
            comentarios = f"{str(comentarios or '').strip()} · {bloque}".strip(" ·") if str(comentarios or "").strip() else bloque
        try:
            resultado = agregar_registro_bitacora_soledad(
                fecha,
                cuadrilla,
                personal,
                ubicacion,
                descripcion,
                consumibles,
                comentarios,
                limitar_fotos(list(fotos or [])),
                secuencia=int(secuencia_n),
                equipo=equipo_sel,
                tipo_actividad=tipo_actividad_final,
                actividad_corta=actividad_corta_final,
            )
        except PermissionError:
            st.error("No se pudo guardar la bitácora. Cierra el Excel si lo tienes abierto e inténtalo de nuevo.")
        except Exception as exc:
            st.error(f"No se pudo guardar el registro: {exc}")
        else:
            st.session_state.ops_df_bitacora = resultado.get("df_bitacora")
            st.session_state.bitacora_form_epoch = epoch + 1
            if resultado.get("fotos"):
                st.session_state.bitacora_registro_ok = "Registro guardado y evidencia adjuntada correctamente"
            else:
                st.session_state.bitacora_registro_ok = "Registro guardado correctamente"
            st.rerun()


def render_captura_bitacora_admin() -> None:
    render_respaldo_base_datos("bitacora")
    activo = render_interruptor_edicion_ops(contexto="bitacora")
    if activo:
        render_registro_actividad_bitacora()
    else:
        st.info("Activa Modo Edición Administrador para registrar una actividad y subir las fotos de evidencia.")


def render_control_programados() -> None:
    st.markdown(
        '<div class="section-title">Control de programados vs no programados</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Compara el plan mensual de OPS (facturación) contra los paros imprevistos y correctivos "
        "capturados en la bitácora de campo."
    )

    hoy = date.today()
    col_anio, col_mes = st.columns(2)
    with col_anio:
        anio = st.number_input(
            "Año",
            min_value=2024,
            max_value=2035,
            value=hoy.year,
            step=1,
            key="prog_anio",
        )
    with col_mes:
        mes = st.selectbox(
            "Mes",
            list(range(1, 13)),
            index=hoy.month - 1,
            format_func=lambda n: NOMBRES_MES[n - 1],
            key="prog_mes",
        )
    mes_clave = clave_mes(anio, mes)
    filas_mes = programados_del_mes(mes_clave)
    df_bitacora = cargar_bitacora()
    df_imprevistos = bitacora_no_programada(df_bitacora, mes_clave)

    actividades_excel = []
    control_excel: dict = {}
    if listar_archivos_programa_octubre():
        actividades_excel = cargar_actividades_octubre(firmas_programa_octubre(), int(anio), int(mes))
        control_excel = leer_control_programa_octubre()
    if actividades_excel:
        cumplidos = sum(
            1
            for item in actividades_excel
            if estatus_guardado_actividad(control_excel.get(clave_control_actividad(item)) or {}) == "Cumplido"
        )
        no_cumplidos = len(actividades_excel) - cumplidos
        reprogramados = 0
        pct = round(100 * cumplidos / len(actividades_excel))
        total_ops = len(actividades_excel)
    else:
        cumplidos = sum(1 for fila in filas_mes if fila.get("estatus") == "Cumplido")
        no_cumplidos = sum(1 for fila in filas_mes if fila.get("estatus") == "No Cumplido")
        reprogramados = sum(1 for fila in filas_mes if fila.get("estatus") == "Reprogramado")
        pct = porcentaje_cumplimiento(filas_mes)
        total_ops = len(filas_mes)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Cumplimiento del mes", f"{pct}%")
    m2.metric("Programados OPS", total_ops)
    m3.metric("Cumplidos", cumplidos)
    m4.metric("No programados / emergencias", len(df_imprevistos))
    st.caption(
        f"{etiqueta_mes(mes_clave)} · No cumplidos: {no_cumplidos} · Reprogramados: {reprogramados}. "
        "El porcentaje es Cumplidos / total de actividades programadas OPS."
    )

    sub = st.session_state.get("admin_sub_programados")
    if not sub:
        render_cuadricula_explorador(
            SUBMENUS_PROGRAMADOS,
            "admin_sub_programados",
            "Elige si vas a cargar el plan OPS o a revisar imprevistos del mes.",
        )
        return
    ruta_programados = f'Programados / <b>{html.escape(titulo_submenu(SUBMENUS_PROGRAMADOS, sub))}</b>'
    if sub == "ops":
        st.markdown(f'<div class="explorer-crumb">{ruta_programados}</div>', unsafe_allow_html=True)
        modo_edicion_ops = render_interruptor_edicion_ops()
        if st.button("← Volver a Programados", key="volver_admin_sub_programados"):
            st.session_state.admin_sub_programados = None
            st.rerun()
        render_programados_ops_excel(int(anio), int(mes), df_bitacora, modo_edicion=modo_edicion_ops)
    else:
        render_barra_volver(
            ruta_programados,
            "admin_sub_programados",
            "← Volver a Programados",
        )

    if sub == "imprevistos":
        st.caption(
            "Registros de campo del mes que corresponden a paros imprevistos, correctivos, "
            "restablecimientos o equipos fuera de servicio."
        )
        if df_imprevistos.empty:
            st.info(f"No hay mantenimientos no programados en {etiqueta_mes(mes_clave)}.")
        else:
            col_a, col_b, col_c = st.columns(3)
            sectores_imp = sorted({str(v) for v in df_imprevistos["SECTOR"] if str(v).strip()})
            cuadrillas_imp = sorted(
                {etiqueta_cuadrilla_historial(str(v)) for v in df_imprevistos["CUADRILLA"] if str(v).strip()},
                key=orden_cuadrilla_oficial,
            )
            with col_a:
                filtro_sector = st.selectbox(
                    "Ubicación",
                    ["Todas"] + sectores_imp,
                    key="imp_sector",
                )
            with col_b:
                filtro_cuadrilla = st.selectbox(
                    "Cuadrilla",
                    ["Todas"] + cuadrillas_imp,
                    key="imp_cuadrilla",
                )
            with col_c:
                filtro_criterio = st.selectbox(
                    "Criterio",
                    ["Todos", "Fuera de servicio", "Restablecimientos / correctivos"],
                    key="imp_criterio",
                )

            vista = df_imprevistos.copy()
            if filtro_sector != "Todas":
                vista = vista[vista["SECTOR"].astype(str) == filtro_sector]
            if filtro_cuadrilla != "Todas":
                objetivo = clave_cuadrilla_oficial(filtro_cuadrilla)
                vista = vista[vista["CUADRILLA"].astype(str).map(clave_cuadrilla_oficial) == objetivo]
            if filtro_criterio == "Fuera de servicio":
                vista = vista[vista["EVALUACIÓN DEL ESTATUS"].astype(str) == "Fuera de servicio"]
            elif filtro_criterio == "Restablecimientos / correctivos":
                vista = vista[
                    vista["TIPO DE ACTIVIDAD"].astype(str).isin(TIPOS_MANTENIMIENTO_NO_PROGRAMADO)
                    | vista["TIPO DE ACTIVIDAD"].astype(str).str.contains(
                        PATRON_NO_PROGRAMADO, na=False
                    )
                    | vista["DESCRIPCIÓN DE LA ACTIVIDAD"].astype(str).str.contains(
                        PATRON_NO_PROGRAMADO, na=False
                    )
                ]

            st.metric("Emergencias visibles", len(vista))
            cols_imp = [c for c in COLUMNAS_REVISION_TABLA if c in vista.columns]
            st.dataframe(
                vista[cols_imp] if cols_imp else vista_revision_soledad(vista),
                use_container_width=True,
                hide_index=True,
            )


def letra_columna_excel(indice: int) -> str:
    letras = ""
    n = indice
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def listar_programas_excel() -> list[str]:
    if not os.path.isdir(PROGRAMAS_EXCEL_DIR):
        return []
    archivos = []
    for nombre in os.listdir(PROGRAMAS_EXCEL_DIR):
        if nombre.startswith("~$"):
            continue
        if nombre.lower().endswith(EXTENSIONES_PROGRAMA_EXCEL):
            archivos.append(nombre)
    return sorted(archivos, key=str.casefold)


def ruta_programa_excel(nombre: str) -> str:
    return os.path.join(PROGRAMAS_EXCEL_DIR, nombre)


def hojas_programa_excel(ruta: str) -> list[str]:
    with pd.ExcelFile(ruta) as libro:
        return list(libro.sheet_names)


def formatear_celda_programa(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, pd.Timestamp):
        if valor.hour == 0 and valor.minute == 0 and valor.second == 0:
            return valor.strftime("%d/%m/%Y")
        return valor.strftime("%d/%m/%Y %H:%M")
    if isinstance(valor, datetime):
        if valor.hour == 0 and valor.minute == 0 and valor.second == 0:
            return valor.strftime("%d/%m/%Y")
        return valor.strftime("%d/%m/%Y %H:%M")
    if isinstance(valor, date) and not isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, bool):
        return str(valor)
    if isinstance(valor, float):
        if valor.is_integer():
            return str(int(valor))
        return str(valor)
    return str(valor)


@st.cache_data(show_spinner="Leyendo el programa de Excel…")
def leer_hoja_programa(ruta: str, hoja: str, mtime: float) -> pd.DataFrame:
    _ = mtime
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(ruta, sheet_name=hoja, header=None, dtype=object)
    while df.shape[1] > 0 and df.iloc[:, -1].isna().all():
        df = df.iloc[:, :-1]
    while df.shape[0] > 0 and df.iloc[-1].isna().all():
        df = df.iloc[:-1]
    if df.empty:
        return pd.DataFrame()
    df = df.map(formatear_celda_programa)
    df.columns = [letra_columna_excel(i + 1) for i in range(df.shape[1])]
    df.index = range(1, len(df) + 1)
    return df


def render_consulta_programas_mensuales() -> None:
    st.markdown(
        '<div class="section-title">Consulta de programas mensuales</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Programas de mantenimiento de la empresa en la carpeta programas_excel. "
        "Se muestran tal como están en cada hoja de cálculo."
    )
    archivos = listar_programas_excel()
    if not archivos:
        st.info(
            "No hay archivos Excel en programas_excel/. "
            "Coloca ahí los programas de eléctricos, instrumentación, UBM u otras áreas."
        )
        return

    archivo = st.selectbox(
        "Archivo de programa",
        archivos,
        key="consulta_programa_archivo",
    )
    ruta = ruta_programa_excel(archivo)
    try:
        hojas = hojas_programa_excel(ruta)
    except Exception as e:
        st.error(f"No se pudieron leer las hojas de {archivo}: {e}")
        return

    if not hojas:
        st.warning("El archivo no tiene hojas visibles.")
        return

    if len(hojas) == 1:
        hoja = hojas[0]
        st.caption(f"Hoja: **{hoja}**")
    else:
        hoja = st.selectbox(
            "Hoja del libro",
            hojas,
            key=f"consulta_programa_hoja_{archivo}",
        )

    try:
        df = leer_hoja_programa(ruta, hoja, os.path.getmtime(ruta))
    except Exception as e:
        st.error(f"No se pudo leer la hoja «{hoja}»: {e}")
        return

    if df.empty:
        st.info(f"La hoja «{hoja}» no tiene datos.")
        return

    st.caption(
        f"**{archivo}** · hoja **{hoja}** · {len(df):,} filas × {df.shape[1]} columnas "
        f"(se conserva la disposición original del Excel)."
    )
    st.dataframe(df, use_container_width=True, hide_index=False)


def _texto_actividad(valor) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", " ", str(valor)).strip()


def _norm_encabezado(valor) -> str:
    texto = unicodedata.normalize("NFKD", _texto_actividad(valor))
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return texto.casefold()


def especialidad_programa(texto: str) -> str | None:
    normal = _norm_encabezado(texto)
    if "elect" in normal:
        return "Eléctrico"
    if "inst" in normal:
        return "Instrumentación"
    if "ubm" in normal or "mecan" in normal:
        return "UBM"
    return None


def _fecha_actividad(valor) -> date | None:
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    parsed = pd.to_datetime(valor, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def _mapa_encabezado_actividad(fila) -> dict[str, int] | None:
    alias = {
        "nro": ("nro", "no."),
        "codigo": ("cod", "codigo"),
        "equipo": ("equipo",),
        "actividad": ("actividad",),
        "tipo": ("tipo mantto", "tipo mant", "tipo"),
        "fecha": ("fecha programada", "fecha"),
    }
    mapa: dict[str, int] = {}
    for indice, valor in enumerate(list(fila)):
        normal = _norm_encabezado(valor)
        if not normal:
            continue
        for campo, opciones in alias.items():
            if campo in mapa:
                continue
            if any(normal == opcion or normal.startswith(opcion) for opcion in opciones):
                mapa[campo] = indice
    if "actividad" in mapa and "fecha" in mapa:
        return mapa
    return None


def directorio_programas_mtto() -> str:
    for carpeta in CARPETAS_PROGRAMA_MTTO:
        if os.path.isdir(carpeta):
            return carpeta
    return CARPETAS_PROGRAMA_MTTO[0]


def listar_archivos_programa_octubre() -> list[str]:
    carpeta = directorio_programas_mtto()
    if not os.path.isdir(carpeta):
        return []
    archivos = []
    for nombre in os.listdir(carpeta):
        if nombre.startswith("~$"):
            continue
        if nombre.lower().endswith(EXTENSIONES_PROGRAMA_EXCEL):
            archivos.append(nombre)
    return sorted(archivos, key=str.casefold)


def listar_pdfs_programa_octubre() -> list[str]:
    carpeta = directorio_programas_mtto()
    if not os.path.isdir(carpeta):
        return []
    return sorted(
        (
            nombre
            for nombre in os.listdir(carpeta)
            if nombre.lower().endswith(".pdf") and not nombre.startswith("~$")
        ),
        key=str.casefold,
    )


def _actividades_de_hoja(ruta: str, hoja: str, archivo: str, anio: int, mes: int) -> list[dict]:
    import warnings

    especialidad_archivo = especialidad_programa(archivo)
    especialidad_hoja = especialidad_programa(hoja)
    especialidad = especialidad_hoja or especialidad_archivo or "General"
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        muestra = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=20, dtype=object)
    encabezado = None
    fila_encabezado = None
    for indice, fila in muestra.iterrows():
        mapa = _mapa_encabezado_actividad(fila.tolist())
        if mapa:
            encabezado = mapa
            fila_encabezado = int(indice)
            break
    if not encabezado or fila_encabezado is None:
        return []
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(ruta, sheet_name=hoja, header=None, dtype=object)
    df = df.iloc[fila_encabezado + 1 :]
    actividades = []
    for _, fila in df.iterrows():
        valores = fila.tolist()
        actividad = _texto_actividad(valores[encabezado["actividad"]]) if encabezado["actividad"] < len(valores) else ""
        if not actividad:
            continue
        fecha = _fecha_actividad(valores[encabezado["fecha"]]) if encabezado["fecha"] < len(valores) else None
        if fecha is None or fecha.year != anio or fecha.month != mes:
            continue
        def celda(campo: str) -> str:
            indice = encabezado.get(campo)
            if indice is None or indice >= len(valores):
                return ""
            return _texto_actividad(valores[indice])

        actividades.append(
            {
                "especialidad": especialidad,
                "programa": f"{archivo} · {hoja.strip()}",
                "hoja": hoja.strip(),
                "nro": celda("nro"),
                "codigo": celda("codigo"),
                "equipo": celda("equipo"),
                "actividad": actividad,
                "tipo": celda("tipo"),
                "fecha": fecha.isoformat(),
                "prioridad": 0 if _norm_encabezado(hoja) == "historial" else 1,
            }
        )
    return actividades


@st.cache_data(show_spinner="Leyendo los programas de mantenimiento…")
def cargar_actividades_octubre(firmas: tuple, anio: int, mes: int) -> list[dict]:
    _ = firmas
    carpeta = directorio_programas_mtto()
    vistos: dict[tuple, dict] = {}
    for nombre in listar_archivos_programa_octubre():
        ruta = os.path.join(carpeta, nombre)
        try:
            hojas = hojas_programa_excel(ruta)
        except Exception:
            continue
        for hoja in hojas:
            if _norm_encabezado(hoja) in HOJAS_PROGRAMA_IGNORADAS:
                continue
            try:
                filas = _actividades_de_hoja(ruta, hoja, nombre, anio, mes)
            except Exception:
                continue
            for fila in filas:
                clave = (
                    fila["especialidad"],
                    fila["codigo"],
                    fila["fecha"],
                    fila["actividad"][:80],
                )
                previa = vistos.get(clave)
                if previa is None or fila["prioridad"] > previa["prioridad"]:
                    vistos[clave] = fila
    actividades = list(vistos.values())
    actividades.sort(key=lambda item: (item["especialidad"], item["fecha"], item["codigo"], item["nro"]))
    return actividades


def firmas_programa_octubre() -> tuple:
    carpeta = directorio_programas_mtto()
    firmas = []
    for nombre in listar_archivos_programa_octubre():
        ruta = os.path.join(carpeta, nombre)
        stat = os.stat(ruta)
        firmas.append((carpeta, nombre, stat.st_mtime, stat.st_size))
    return tuple(firmas)


def clave_control_actividad(fila: dict) -> str:
    base = "|".join(
        [
            str(fila.get("especialidad") or ""),
            str(fila.get("codigo") or ""),
            str(fila.get("fecha") or ""),
            str(fila.get("actividad") or "")[:80],
        ]
    )
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:16]


def leer_control_programa_octubre() -> dict:
    datos = leer_json(CONTROL_PROGRAMA_OCTUBRE_FILE, {})
    return datos if isinstance(datos, dict) else {}


def guardar_control_programa_octubre(datos: dict) -> None:
    escribir_json(CONTROL_PROGRAMA_OCTUBRE_FILE, datos)


def render_programa_octubre() -> None:
    st.markdown(
        '<div class="section-title">Programa de mantenimientos — octubre 2026</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Actividades leídas de programa_mantenimientos_octubre. "
        "Marca Realizado y anota observaciones; el control queda guardado en este equipo."
    )
    pdfs = listar_pdfs_programa_octubre()
    if pdfs:
        columnas = st.columns(min(len(pdfs), 3))
        for indice, nombre in enumerate(pdfs):
            ruta = os.path.join(directorio_programas_mtto(), nombre)
            with columnas[indice % len(columnas)]:
                with open(ruta, "rb") as archivo:
                    st.download_button(
                        f"Descargar PDF · {nombre}",
                        data=archivo.read(),
                        file_name=nombre,
                        mime="application/pdf",
                        key=f"pdf_octubre_{indice}",
                        use_container_width=True,
                    )
    carpeta_programa = directorio_programas_mtto()
    if not os.path.isdir(carpeta_programa):
        st.info("Coloca los archivos .xlsx en la carpeta PROGRAMAS MTTO OCTUBRE.")
        return
    archivos = listar_archivos_programa_octubre()
    if not archivos:
        st.info(f"No hay archivos Excel en {carpeta_programa}.")
        return

    anio_octubre, mes_octubre = MES_PROGRAMA_OCTUBRE
    actividades = cargar_actividades_octubre(firmas_programa_octubre(), anio_octubre, mes_octubre)
    if not actividades:
        st.warning("No se encontraron actividades de octubre 2026 en los Excel de la carpeta.")
        return

    especialidades = [nombre for nombre in ESPECIALIDADES_OCTUBRE if any(item["especialidad"] == nombre for item in actividades)]
    extra = sorted({item["especialidad"] for item in actividades} - set(especialidades))
    especialidades.extend(extra)
    control = leer_control_programa_octubre()
    pestanas = st.tabs(especialidades)
    for pestana, especialidad in zip(pestanas, especialidades):
        with pestana:
            filas = [item for item in actividades if item["especialidad"] == especialidad]
            origenes = sorted({item["programa"] for item in filas})
            if len(origenes) > 1:
                origen = st.selectbox(
                    "Programa o zona",
                    ["Todos"] + origenes,
                    key=f"origen_octubre_{especialidad}",
                )
                if origen != "Todos":
                    filas = [item for item in filas if item["programa"] == origen]
            registros = []
            for item in filas:
                clave = clave_control_actividad(item)
                guardado = control.get(clave) or {}
                registros.append(
                    {
                        "_id": clave,
                        "Fecha": datetime.fromisoformat(item["fecha"]).strftime("%d/%m/%Y"),
                        "Código": item["codigo"],
                        "Equipo": item["equipo"],
                        "Actividad": item["actividad"],
                        "Tipo": item["tipo"],
                        "Programa": item["hoja"],
                        "Realizado": bool(guardado.get("realizado")),
                        "Observaciones": str(guardado.get("observaciones") or ""),
                    }
                )
            tabla = pd.DataFrame(registros)
            hechos = int(tabla["Realizado"].sum()) if not tabla.empty else 0
            st.caption(f"{len(tabla)} actividades · {hechos} realizadas")
            editada = st.data_editor(
                tabla,
                column_order=[
                    "Realizado",
                    "Observaciones",
                    "Fecha",
                    "Código",
                    "Equipo",
                    "Actividad",
                    "Tipo",
                    "Programa",
                ],
                column_config={
                    "Realizado": st.column_config.CheckboxColumn("Realizado", default=False),
                    "Observaciones": st.column_config.TextColumn("Observaciones"),
                    "Actividad": st.column_config.TextColumn("Actividad", width="large"),
                },
                disabled=["Fecha", "Código", "Equipo", "Actividad", "Tipo", "Programa"],
                hide_index=True,
                use_container_width=True,
                key=f"editor_octubre_{especialidad}",
            )
            cambio = False
            for _, fila in editada.iterrows():
                clave = str(fila["_id"])
                observacion = fila["Observaciones"]
                if observacion is None:
                    texto_obs = ""
                else:
                    try:
                        texto_obs = "" if pd.isna(observacion) else str(observacion).strip()
                    except (TypeError, ValueError):
                        texto_obs = str(observacion).strip()
                previo = control.get(clave) if isinstance(control.get(clave), dict) else {}
                realizado = bool(fila["Realizado"])
                nuevo = dict(previo)
                nuevo["realizado"] = realizado
                nuevo["observaciones"] = texto_obs
                if realizado:
                    nuevo["estatus"] = "Cumplido"
                elif nuevo.get("estatus") == "Cumplido":
                    nuevo["estatus"] = "No Cumplido"
                if control.get(clave) == nuevo:
                    continue
                if _registro_control_vacio(nuevo):
                    if clave in control:
                        control.pop(clave, None)
                        cambio = True
                    continue
                control[clave] = nuevo
                cambio = True
            if cambio:
                guardar_control_programa_octubre(control)


def render_gestion_insumos_catalogo() -> None:
    st.markdown("**Alta de insumos**")
    st.caption(
        "Formulario libre para dar de alta materiales que no están en el catálogo inicial. "
        "Elige si es consumible (bandas C-115, C-112, C-240, Micro V, aceite, filtros) "
        "o componente rotativo / reparable (bomba de agua, alternador, marcha). "
        "Se guarda en almacen_insumos.json y aparece de inmediato en campo."
    )
    if st.session_state.pop("insumo_alta_ok", False):
        st.success("Insumo agregado al catálogo.")

    col_nom, col_tipo, col_uni, col_otra = st.columns([2, 1.4, 1, 1])
    with col_nom:
        nombre = st.text_input(
            "Nombre del insumo",
            placeholder="Ej. Filtro de aire, Banda C-115, Bomba de agua…",
            key="alta_insumo_nombre",
        )
    with col_tipo:
        tipo = st.selectbox("Categoría", TIPOS_INSUMO, key="alta_insumo_tipo")
    with col_uni:
        unidad = st.selectbox("Unidad", UNIDADES_INSUMO, key="alta_insumo_unidad")
    with col_otra:
        unidad_otra = ""
        if unidad == "Otra":
            unidad_otra = st.text_input(
                "Especificar unidad",
                placeholder="Ej. Gal, Par, Kit",
                key="alta_insumo_unidad_otra",
            )
        else:
            st.caption("Usa «Otra» si la unidad no está en la lista.")

    if st.button("➕ Dar de alta insumo", type="primary", key="btn_alta_insumo"):
        unidad_final = unidad_otra.strip() if unidad == "Otra" else unidad
        if unidad == "Otra" and not unidad_final:
            st.warning("Especifica la unidad del insumo.")
        else:
            error = alta_insumo_catalogo(nombre, unidad_final, tipo)
            if error:
                st.warning(error)
            else:
                st.session_state.insumo_alta_ok = True
                st.rerun()

    datos = cargar_almacen()
    filas = [
        {
            "Insumo": item["nombre"],
            "Categoría": tipo_insumo_item(item),
            "Unidad": item.get("unidad", ""),
            "Origen del catálogo": "Precargado" if item.get("precargado") else "Alta administrativa",
            "Activo": "Sí" if item.get("activo", True) else "No",
        }
        for item in datos["catalogo"]
    ]
    st.markdown("**Catálogo vigente**")
    st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)


def render_entradas_almacen() -> None:
    st.markdown("**Entrada de almacén**")
    st.caption(
        "Registra materiales que ingresan a inventario. El origen del insumo es obligatorio: "
        "PJP4 (Propio) u OPS (Proporcionado por el Cliente)."
    )
    if st.session_state.pop("entrada_almacen_ok", False):
        st.success("Entrada registrada en almacen_insumos.json.")

    catalogo = catalogo_insumos_activos()
    if not catalogo:
        st.info("No hay insumos activos. Da de alta al menos uno en Gestión de Insumos.")
        return

    nombres = [item["nombre"] for item in catalogo]
    col_f, col_i = st.columns(2)
    with col_f:
        fecha_ent = st.date_input(
            "Fecha de entrada",
            value=date.today(),
            format="DD/MM/YYYY",
            key="entrada_fecha",
        )
    with col_i:
        nombre_sel = st.selectbox("Insumo", nombres, key="entrada_insumo")

    item = next((i for i in catalogo if i["nombre"] == nombre_sel), catalogo[0])
    unidad = item.get("unidad") or "Pzas"
    col_c, col_o = st.columns(2)
    with col_c:
        if unidad_es_entera(unidad):
            cantidad = st.number_input(
                f"Cantidad ({unidad})",
                min_value=0,
                step=1,
                value=1,
                key=f"entrada_cantidad_{item['id']}",
            )
        else:
            paso = 0.1 if str(unidad).lower() == "kg" else 0.5
            cantidad = st.number_input(
                f"Cantidad ({unidad})",
                min_value=0.0,
                step=float(paso),
                value=float(paso),
                key=f"entrada_cantidad_{item['id']}",
            )
    with col_o:
        origen = st.radio(
            "Origen del insumo",
            ORIGENES_INSUMO,
            index=None,
            horizontal=True,
            key="entrada_origen",
        )

    nota = st.text_input(
        "Nota / referencia (opcional)",
        placeholder="Factura, remisión, pozo, cuadrilla…",
        key="entrada_nota",
    )

    if st.button("📥 Registrar entrada", type="primary", key="btn_entrada_almacen"):
        error = registrar_entrada_almacen(fecha_ent, item, float(cantidad), origen or "", nota)
        if error:
            st.warning(error)
        else:
            st.session_state.entrada_almacen_ok = True
            st.rerun()

    entradas = [
        {
            "Fecha": mov.get("fecha", ""),
            "Insumo": mov.get("insumo", ""),
            "Cantidad": mov.get("cantidad", 0),
            "Unidad": mov.get("unidad", ""),
            "Origen": mov.get("origen", ""),
            "Nota": mov.get("nota", ""),
        }
        for mov in reversed(cargar_almacen()["movimientos"])
        if str(mov.get("tipo", "")) == "entrada"
    ]
    st.markdown("**Entradas recientes**")
    if entradas:
        st.dataframe(pd.DataFrame(entradas), use_container_width=True, hide_index=True)
    else:
        st.info("Aún no hay entradas registradas.")


def render_reporte_almacen() -> None:
    st.markdown("**Reporte de insumos gastados**")
    st.caption(
        "Separa el consumo de campo según el origen: materiales PJP4 frente a los "
        "proporcionados por OPS. Fuente: almacen_insumos.json."
    )
    hoy = date.today()
    col_anio, col_mes = st.columns(2)
    with col_anio:
        anio = st.number_input(
            "Año",
            min_value=2024,
            max_value=2035,
            value=hoy.year,
            step=1,
            key="almacen_rep_anio",
        )
    with col_mes:
        mes = st.selectbox(
            "Mes",
            list(range(1, 13)),
            index=hoy.month - 1,
            format_func=lambda n: NOMBRES_MES[n - 1],
            key="almacen_rep_mes",
        )
    mes_clave = clave_mes(anio, mes)
    consumos = movimientos_consumo(mes_clave)
    tablas = resumen_consumo_por_origen(mes_clave)
    existencias = inventario_por_origen()

    tot_pjp4 = float(tablas["PJP4"]["Cantidad"].sum()) if not tablas["PJP4"].empty else 0
    tot_ops = float(tablas["OPS"]["Cantidad"].sum()) if not tablas["OPS"].empty else 0
    m1, m2, m3 = st.columns(3)
    m1.metric("Consumos del mes", len(consumos))
    m2.metric("Gastado PJP4", tot_pjp4)
    m3.metric("Gastado OPS", tot_ops)
    st.caption(f"{etiqueta_mes(mes_clave)} · Totales por origen del insumo.")

    col_p, col_o = st.columns(2)
    with col_p:
        st.markdown(
            '<div class="origen-pjp4"><h3>PJP4 (Propio)</h3>'
            "<p>Insumos gastados con cargo a PJP4.</p></div>",
            unsafe_allow_html=True,
        )
        if tablas["PJP4"].empty:
            st.info("Sin consumo PJP4 en el periodo.")
        else:
            st.dataframe(tablas["PJP4"], use_container_width=True, hide_index=True)
    with col_o:
        st.markdown(
            '<div class="origen-ops"><h3>OPS (Cliente)</h3>'
            "<p>Insumos proporcionados por el cliente y gastados en campo.</p></div>",
            unsafe_allow_html=True,
        )
        if tablas["OPS"].empty:
            st.info("Sin consumo OPS en el periodo.")
        else:
            st.dataframe(tablas["OPS"], use_container_width=True, hide_index=True)

    st.markdown("**Existencia neta por origen**")
    st.caption("Entradas menos consumos acumulados (todos los meses).")
    col_ep, col_eo = st.columns(2)
    with col_ep:
        st.caption("Inventario PJP4")
        if existencias["PJP4"].empty:
            st.info("Sin movimientos PJP4.")
        else:
            st.dataframe(existencias["PJP4"], use_container_width=True, hide_index=True)
    with col_eo:
        st.caption("Inventario OPS")
        if existencias["OPS"].empty:
            st.info("Sin movimientos OPS.")
        else:
            st.dataframe(existencias["OPS"], use_container_width=True, hide_index=True)

    detalle = [
        {
            "Fecha": mov.get("fecha", ""),
            "Insumo": mov.get("insumo", ""),
            "Cantidad": mov.get("cantidad", 0),
            "Unidad": mov.get("unidad", ""),
            "Origen": mov.get("origen", ""),
            "Cuadrilla": mov.get("cuadrilla", ""),
            "Sector": mov.get("sector", ""),
            "Secuencia": mov.get("secuencia", ""),
        }
        for mov in consumos
    ]
    st.markdown("**Detalle de consumos de campo**")
    if detalle:
        st.dataframe(pd.DataFrame(detalle), use_container_width=True, hide_index=True)
    else:
        st.info(f"No hay consumos registrados en {etiqueta_mes(mes_clave)}.")


def render_control_almacen() -> None:
    st.markdown(
        '<div class="section-title">Control de Almacén e Insumos</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Catálogo precargado, altas administrativas, entradas de almacén y reporte de consumo "
        f"por origen. Persistencia: {ALMACEN_FILE}."
    )
    sub = st.session_state.get("admin_sub_almacen")
    if not sub:
        render_cuadricula_explorador(
            SUBMENUS_ALMACEN,
            "admin_sub_almacen",
            "Elige una sección del almacén para continuar.",
        )
        return
    render_barra_volver(
        f'Almacén / <b>{html.escape(titulo_submenu(SUBMENUS_ALMACEN, sub))}</b>',
        "admin_sub_almacen",
        "← Volver a Almacén",
    )
    if sub == "gestion":
        render_gestion_insumos_catalogo()
    elif sub == "entradas":
        render_entradas_almacen()
    elif sub == "reporte":
        render_reporte_almacen()


def _mes_calendario_pozo() -> tuple[int, int]:
    hoy = date.today()
    return (
        int(st.session_state.get("ub_cal_anio") or hoy.year),
        int(st.session_state.get("ub_cal_mes") or hoy.month),
    )


def _desplazar_mes_calendario(delta: int) -> None:
    anio, mes = _mes_calendario_pozo()
    mes += delta
    while mes < 1:
        mes += 12
        anio -= 1
    while mes > 12:
        mes -= 12
        anio += 1
    st.session_state.ub_cal_anio = anio
    st.session_state.ub_cal_mes = mes
    st.session_state.ub_cal_dia = None


def _reset_explorador_pozos() -> None:
    st.session_state.ub_zona_explorador = None
    st.session_state.ub_ficha_id = None
    st.session_state.ub_cal_dia = None
    st.session_state.ub_cal_pozo = None


def render_volver_historial(ruta_html: str, nivel: str) -> None:
    st.markdown(f'<div class="explorer-crumb">{ruta_html}</div>', unsafe_allow_html=True)
    if st.button("← Volver", key=f"volver_hist_{nivel}"):
        if nivel == "zona":
            _reset_explorador_pozos()
        else:
            st.session_state.ub_ficha_id = None
            st.session_state.ub_cal_dia = None
        st.rerun()


def render_tarjeta_pozo(ficha: dict, zona: dict) -> None:
    prefijo = ficha.get("prefijo") or zona.get("prefijo") or ""
    mci = ficha.get("equipo_mci") or "Sin MCI"
    ub = ficha.get("equipo_ub") or "Sin UB"
    st.markdown(
        f"""
        <div class="pozo-card">
            <div class="prefijo">{html.escape(str(prefijo))}</div>
            <h3>{html.escape(ficha.get("nombre", ""))}</h3>
            <p>MCI: {html.escape(str(mci))}<br>UB: {html.escape(str(ub))}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Ver historial", key=f"abrir_pozo_{ficha['id']}", use_container_width=True):
        st.session_state.ub_zona_explorador = ficha.get("zona_id") or st.session_state.get("ub_zona_explorador")
        st.session_state.ub_ficha_id = ficha["id"]
        st.session_state.ub_cal_dia = None
        st.rerun()


def render_cuadricula_pozos(fichas: list[dict], datos: dict) -> None:
    zonas = mapa_zonas(datos["zonas"])
    if not fichas:
        st.info("No hay pozos en esta zona. Puedes darlos de alta más abajo.")
        return
    for inicio in range(0, len(fichas), 3):
        cols = st.columns(3, gap="medium")
        for col, ficha in zip(cols, fichas[inicio : inicio + 3]):
            with col:
                render_tarjeta_pozo(ficha, zonas.get(ficha.get("zona_id"), {}))


def render_pantalla_zonas_pozos(datos: dict) -> None:
    st.markdown(
        '<p class="explorer-meta">Explorador de zonas de pozos. Abre una tarjeta para ver el inventario de la zona.</p>',
        unsafe_allow_html=True,
    )
    busqueda = st.text_input(
        "Buscar pozo en todas las zonas",
        placeholder="Ej. PS515, PA-3, PGA-1",
        key="ub_buscar_explorador",
    ).strip()
    if busqueda:
        q = busqueda.casefold()
        coincidencias = sorted(
            [
                f
                for f in datos["fichas"].values()
                if q in f.get("nombre", "").casefold()
                or q in clave_ubicacion(f.get("nombre", "")).casefold()
            ],
            key=lambda f: orden_natural_pozo(f.get("nombre", "")),
        )
        st.caption(f"{len(coincidencias)} coincidencia(s) para «{busqueda}».")
        render_cuadricula_pozos(coincidencias, datos)
        return

    zonas_map = mapa_zonas(datos["zonas"])
    tarjetas = []
    for item in ZONAS_HISTORIAL_POZOS:
        zona = zonas_map.get(item["id"], {})
        n = len(fichas_de_zona(datos, item["id"]))
        sector = zona.get("sector") or item["detalle"]
        tarjetas.append(
            {
                "id": item["id"],
                "icono": item["icono"],
                "titulo": item["titulo"],
                "descripcion": f"{n} pozo(s) · {sector}",
            }
        )
    render_cuadricula_explorador(tarjetas, "ub_zona_explorador")

    extras = [z for z in datos["zonas"] if z["id"] not in {zdef["id"] for zdef in ZONAS_HISTORIAL_POZOS}]
    if extras:
        with st.expander("Instalaciones y otras ubicaciones"):
            extra_cards = []
            for zona in extras:
                n = len(fichas_de_zona(datos, zona["id"]))
                extra_cards.append(
                    {
                        "id": zona["id"],
                        "icono": "🏭" if zona["id"] == "instalaciones" else "📌",
                        "titulo": f"{zona.get('prefijo', '')} {zona.get('nombre', '')}".strip(),
                        "descripcion": f"{n} ubicación(es) fuera del inventario de pozos.",
                    }
                )
            render_cuadricula_explorador(extra_cards, "ub_zona_explorador")


def render_listado_pozos_zona(datos: dict, zona_id: str) -> None:
    zona = mapa_zonas(datos["zonas"]).get(zona_id, {})
    titulo = titulo_zona_historial(zona_id, zona)
    render_volver_historial(
        f'Historial y Estado de Pozos / <b>{html.escape(titulo)}</b>',
        "zona",
    )
    fichas = fichas_de_zona(datos, zona_id)
    filtro = st.text_input(
        "Filtrar pozos de esta zona",
        placeholder="Ej. 515, PS-12",
        key=f"ub_filtro_zona_{zona_id}",
    ).strip()
    if filtro:
        q = filtro.casefold()
        fichas = [
            f
            for f in fichas
            if q in f.get("nombre", "").casefold() or q in clave_ubicacion(f.get("nombre", "")).casefold()
        ]
        st.caption(f"{len(fichas)} pozo(s) en {titulo} que coinciden con «{filtro}».")
    else:
        st.caption(f"{len(fichas)} pozo(s) en {titulo}. Elige una tarjeta para abrir el detalle.")
    render_cuadricula_pozos(fichas, datos)
    render_alta_ubicacion(datos)


def render_detalle_estado_pozo(ficha: dict, datos: dict) -> None:
    zona = mapa_zonas(datos["zonas"]).get(ficha.get("zona_id"), {})
    titulo_zona = titulo_zona_historial(ficha.get("zona_id", ""), zona)
    render_volver_historial(
        f'Historial y Estado de Pozos / {html.escape(titulo_zona)} / '
        f'<b>{html.escape(ficha.get("nombre", ""))}</b>',
        "pozo",
    )

    hoy = date.today()
    actualizado = ficha.get("actualizado") or "sin actualizaciones de ficha"
    st.markdown(
        f"""
        <div class="ficha-hero">
            <div class="prefijo">{html.escape(titulo_zona)}</div>
            <h2>{html.escape(ficha.get("nombre", ""))}</h2>
            <p>{html.escape(ficha.get("sector") or "Sin sector asociado")} · Estatus: {html.escape(ficha.get("estatus") or "—")} · {html.escape(actualizado)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_mci, col_ub = st.columns(2, gap="medium")
    mci = ficha.get("equipo_mci") or "Sin motor MCI registrado"
    ub = ficha.get("equipo_ub") or "Sin unidad de bombeo registrada"
    cls_m = "equipo-card" if ficha.get("equipo_mci") else "equipo-card vacio"
    cls_u = "equipo-card" if ficha.get("equipo_ub") else "equipo-card vacio"
    with col_mci:
        st.markdown(
            f'<div class="{cls_m}"><div class="lbl">Motor MCI</div>'
            f"<h3>{html.escape(str(mci))}</h3>"
            f"<p>Equipo de combustión instalado en este pozo.</p></div>",
            unsafe_allow_html=True,
        )
    with col_ub:
        st.markdown(
            f'<div class="{cls_u}"><div class="lbl">Unidad de bombeo (UB)</div>'
            f"<h3>{html.escape(str(ub))}</h3>"
            f"<p>Unidad de bombeo (mecánica o hidráulica) asociada a la ubicación.</p></div>",
            unsafe_allow_html=True,
        )

    df_pozo = bitacora_de_pozo(cargar_bitacora(), ficha)
    cambios = cambios_de_pozo(datos, ficha)
    programados = programados_ordenados_pozo(ficha)
    eventos_hist = eventos_historial_campo_pozo(ficha)
    eventos = sorted(
        eventos_historial_pozo(df_pozo, cambios) + eventos_hist,
        key=lambda item: item.get("sort") or "0000-00-00",
        reverse=True,
    )
    dias_hist = {ev["fecha"] for ev in eventos_hist if ev.get("fecha")}
    dias_bit = fechas_bitacora_pozo(df_pozo) | dias_hist
    dias_prog = {fch for item in programados if (fch := fecha_de_programado(item))}
    pendientes = [p for p in programados if p.get("estatus") in {"No Cumplido", "Reprogramado"}]

    if st.session_state.get("ub_cal_pozo") != ficha["id"]:
        st.session_state.ub_cal_pozo = ficha["id"]
        fechas = dias_bit | dias_prog
        if fechas:
            ultimo = max(fechas)
            st.session_state.ub_cal_anio = ultimo.year
            st.session_state.ub_cal_mes = ultimo.month
        else:
            st.session_state.ub_cal_anio = hoy.year
            st.session_state.ub_cal_mes = hoy.month
        st.session_state.ub_cal_dia = None

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Actividades en bitácora", len(df_pozo))
    m2.metric("Histórico jul–sep", len(eventos_hist))
    m3.metric("Cambios de MCI / UB", len(cambios))
    m4.metric("Programados pendientes", len(pendientes))
    if eventos_hist:
        st.caption(
            "El mini reporte histórico sale de `assets/historial_pjp4.xlsx` "
            "(julio y agosto-septiembre), cruzado con esta ficha del catálogo."
        )

    anio, mes = _mes_calendario_pozo()
    clave_vista = clave_mes(anio, mes)
    prog_mes = [p for p in programados if str(p.get("mes") or "") == clave_vista]
    prog_mes_sin_dia = [p for p in prog_mes if fecha_de_programado(p) is None]
    dia_sel = parsear_fecha_valor(st.session_state.get("ub_cal_dia"))

    col_main, col_side = st.columns([1.7, 1], gap="large")
    with col_main:
        nav_a, nav_b, nav_c, nav_d = st.columns([1, 1, 1.4, 1])
        with nav_a:
            if st.button("← Mes", key=f"cal_prev_{ficha['id']}", use_container_width=True):
                _desplazar_mes_calendario(-1)
                st.rerun()
        with nav_b:
            if st.button("Hoy", key=f"cal_hoy_{ficha['id']}", use_container_width=True):
                st.session_state.ub_cal_anio = hoy.year
                st.session_state.ub_cal_mes = hoy.month
                st.session_state.ub_cal_dia = None
                st.rerun()
        with nav_c:
            st.markdown(
                f"<div style='text-align:center;padding-top:0.45rem;color:#e2e8f0;font-weight:700'>"
                f"{NOMBRES_MES[mes - 1]} {anio}</div>",
                unsafe_allow_html=True,
            )
        with nav_d:
            if st.button("Mes →", key=f"cal_next_{ficha['id']}", use_container_width=True):
                _desplazar_mes_calendario(1)
                st.rerun()

        if prog_mes_sin_dia:
            n = len(prog_mes_sin_dia)
            st.markdown(
                f'<div class="cal-banner">{n} mantenimiento(s) programado(s) en {NOMBRES_MES[mes - 1]} '
                f"sin día fijo. Quedan listados en el panel lateral y se marcan en el calendario "
                f"cuando exista fecha.</div>",
                unsafe_allow_html=True,
            )

        render_calendario_pozo(anio, mes, dias_bit, dias_prog, dia_sel)

        marcados = sorted(
            d for d in (dias_bit | dias_prog) if d.year == anio and d.month == mes
        )
        if marcados:
            st.caption("Días con registro en este mes. Pulsa uno para filtrar el historial.")
            botones_dia = [("Todos", None)] + [(f"{dia.day:02d}", dia) for dia in marcados]
            for inicio in range(0, len(botones_dia), 8):
                fila = botones_dia[inicio : inicio + 8]
                chips = st.columns(len(fila))
                for col, (etiqueta, dia) in zip(chips, fila):
                    with col:
                        if st.button(
                            etiqueta,
                            key=f"caldia_{ficha['id']}_{dia.isoformat() if dia else 'todos'}",
                            use_container_width=True,
                        ):
                            st.session_state.ub_cal_dia = dia.isoformat() if dia else None
                            st.rerun()

        st.markdown('<div class="section-title">Historial de bitácora</div>', unsafe_allow_html=True)
        if dia_sel:
            st.caption(f"Filtrado al {dia_sel.strftime('%d/%m/%Y')}.")
            visibles = [e for e in eventos if e.get("fecha") == dia_sel]
        else:
            st.caption(
                "Actividades ejecutadas, de la más reciente a la más antigua, "
                "incluyendo el histórico de campo julio–septiembre."
            )
            visibles = eventos

        if not visibles:
            st.info("No hay actividades registradas para este recorte.")
        else:
            for ev in visibles:
                cls = "hist-card"
                if ev["tipo"] == "cambio":
                    cls += " cambio"
                elif ev["tipo"] == "historial":
                    cls += " historico"
                elif ev.get("clase") == "No programado":
                    cls += " nprog"
                chips_html = "".join(
                    f'<span class="insumo-chip">{html.escape(ins)}</span>' for ins in ev.get("insumos") or []
                )
                meta = " · ".join(
                    p
                    for p in [
                        ev.get("clase") or "",
                        ev.get("cuadrilla") or "",
                        f"Sec. {ev['secuencia']}" if ev.get("secuencia") else "",
                        ev.get("estatus") or "",
                        (
                            f"Evidencia {ev.get('evidencia')} ({ev.get('fotos') or 0} foto(s))"
                            if ev.get("tipo") == "bitacora"
                            else ""
                        ),
                    ]
                    if p
                )
                extra = ""
                if ev["tipo"] == "cambio":
                    origen = ev.get("origen") or "Alta / reemplazo local"
                    anterior = ev.get("equipo") or "—"
                    extra = (
                        f"<p>Anterior: {html.escape(str(anterior))} · "
                        f"Origen: {html.escape(str(origen))}</p>"
                    )
                desc = ev.get("descripcion") or ""
                desc_html = f"<p>{html.escape(desc)}</p>" if desc else ""
                equipo = ev.get("equipo") if ev["tipo"] in {"bitacora", "historial"} else ""
                equipo_html = f"<p>{html.escape(equipo)}</p>" if equipo else ""
                tarjeta = (
                    f'<div class="{cls}">'
                    f'<div class="hist-date">{html.escape(ev["fecha_txt"])}</div>'
                    f"<h4>{html.escape(ev['titulo'])}</h4>"
                    f"<p>{html.escape(meta)}</p>"
                    f"{equipo_html}{desc_html}{extra}{chips_html}"
                    "</div>"
                )
                if hasattr(st, "html"):
                    st.html(tarjeta)
                else:
                    st.markdown(tarjeta, unsafe_allow_html=True)

            bit_visibles = [
                e for e in visibles if e.get("tipo") == "bitacora" and e.get("idx") is not None
            ]
            if bit_visibles:
                st.markdown("##### Completar o actualizar evidencia")
                st.caption("Elige el registro del calendario o del historial del pozo y anexa las fotos que le falten.")
                df_base = cargar_bitacora()
                opciones_idx = [e["idx"] for e in bit_visibles if e["idx"] in df_base.index]
                if opciones_idx:
                    elegido_idx = st.selectbox(
                        "Registro a completar",
                        opciones_idx,
                        format_func=lambda i: etiqueta_registro(df_base, i),
                        key=f"cal_ev_{ficha.get('id')}",
                    )
                    render_panel_evidencia_posterior(
                        df_base,
                        elegido_idx,
                        clave=f"cal_{ficha.get('id')}_{elegido_idx}",
                    )

    with col_side:
        st.markdown(
            '<div class="prog-panel"><h3>Mantenimientos programados</h3>'
            '<p class="hint">Pendientes y plan mensual de este pozo, sincronizados con el calendario.</p></div>',
            unsafe_allow_html=True,
        )
        if not programados:
            st.info("No hay mantenimientos programados ligados a este pozo.")
        else:
            for item in programados:
                estatus = str(item.get("estatus") or "—")
                cls = "prog-item"
                if estatus == "Cumplido":
                    cls += " cumplido"
                elif estatus == "Reprogramado":
                    cls += " reprog"
                if str(item.get("mes") or "") == clave_vista:
                    cls += " mes-activo"
                fch = fecha_de_programado(item)
                if fch:
                    cuando = fch.strftime("%d/%m/%Y")
                else:
                    cuando = etiqueta_mes(str(item.get("mes") or "")) or "Mes no indicado"
                obs = str(item.get("observacion") or "").strip()
                equipo = str(item.get("equipo") or "").strip()
                st.markdown(
                    f"""
                    <div class="{cls}">
                        <div class="mes">{html.escape(cuando)} · {html.escape(estatus)}</div>
                        <h4>{html.escape(str(item.get("actividad") or "Mantenimiento programado"))}</h4>
                        <p>{html.escape(str(item.get("ubicacion") or titulo_zona))}</p>
                        <p>{html.escape(equipo or "Sin equipo indicado")}{(" · " + html.escape(obs)) if obs else ""}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    personal = personal_de_pozo(df_pozo)
    with st.expander("Personal que intervino en este pozo", expanded=False):
        if not personal:
            st.info("Todavía no hay cuadrillas registradas en la bitácora para este pozo.")
        else:
            for item in personal:
                st.markdown(
                    f"""
                    <div class="personal-chip">
                        <div class="cuadrilla">{html.escape(item["cuadrilla"])}</div>
                        <p>{item["intervenciones"]} intervención(es) · Última: {html.escape(item["ultima"] or "s/f")}<br>
                        {html.escape(item.get("nombres") or "Sin personal de turno capturado")}<br>
                        {html.escape(item["actividades"])}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    render_editar_ubicacion(ficha, datos)
    render_form_equipo(ficha, datos)


def render_flujo_historial_pozos(datos: dict) -> None:
    zona_id = st.session_state.get("ub_zona_explorador")
    ficha_id = st.session_state.get("ub_ficha_id")
    if not zona_id:
        render_pantalla_zonas_pozos(datos)
        return
    if not ficha_id or ficha_id not in datos["fichas"]:
        render_listado_pozos_zona(datos, zona_id)
        return
    render_detalle_estado_pozo(datos["fichas"][ficha_id], datos)


def render_nav_ubicaciones(datos: dict) -> str | None:
    zonas = datos["zonas"]
    with st.sidebar:
        st.markdown("### Zonas y prefijos")
        st.caption("PS Soledad · PA Aragón · PGA Gallo · PSN · PPB · PC")
        busqueda = st.text_input(
            "Buscar ubicación",
            placeholder="Ej. PS-12, PA-3, PGA-1",
            key="ub_buscar",
        ).strip()
        etiquetas = [f"{z['prefijo']}: {z['nombre']}" for z in zonas]
        zona_idx = 0
        if "ub_zona_id" in st.session_state:
            ids = [z["id"] for z in zonas]
            if st.session_state.ub_zona_id in ids:
                zona_idx = ids.index(st.session_state.ub_zona_id)
        zona_sel = st.selectbox("Zona", etiquetas, index=zona_idx, key="ub_zona_sel")
        zona = zonas[etiquetas.index(zona_sel)]
        st.session_state.ub_zona_id = zona["id"]

        fichas = fichas_de_zona(datos, zona["id"])
        if busqueda:
            q = busqueda.casefold()
            fichas = [
                f
                for f in datos["fichas"].values()
                if q in f["nombre"].casefold() or q in clave_ubicacion(f["nombre"]).casefold()
            ]
            fichas = sorted(fichas, key=lambda f: f["nombre"])
        if not fichas:
            st.caption("No hay ubicaciones en esta zona.")
            return st.session_state.get("ub_ficha_id")

        opciones = {f["nombre"]: f["id"] for f in fichas}
        actual = st.session_state.get("ub_ficha_id")
        nombres = list(opciones.keys())
        idx = 0
        if actual in opciones.values():
            inv = {v: k for k, v in opciones.items()}
            if actual in inv and inv[actual] in nombres:
                idx = nombres.index(inv[actual])
        elegido = st.radio(
            "Ubicación (pozo)",
            nombres,
            index=idx,
            key=f"ub_radio_{zona['id']}_{clave_ubicacion(busqueda) or 'todas'}",
        )
        st.session_state.ub_ficha_id = opciones[elegido]
        st.session_state.ub_zona_id = datos["fichas"][opciones[elegido]].get("zona_id", zona["id"])
        return opciones[elegido]


def render_alta_ubicacion(datos: dict) -> None:
    with st.expander("➕ Dar de alta o corregir ubicación / pozo", expanded=False):
        st.caption(
            "Si la clave ya existe (por ejemplo PB129 en la zona equivocada), "
            "activa «Actualizar si ya existe» para corregir zona y sector sin perder el equipo."
        )
        zonas = datos["zonas"]
        etiquetas = [f"{z['prefijo']}: {z['nombre']}" for z in zonas]
        col_n, col_z = st.columns([2, 2])
        with col_n:
            nombre = st.text_input(
                "Clave o nombre",
                placeholder="Ej. PS-18, PA-9, PGA-6, PB129",
                key="ub_alta_nombre",
            )
        with col_z:
            zona_etq = st.selectbox("Zona / prefijo", etiquetas, key="ub_alta_zona")
        zona = zonas[etiquetas.index(zona_etq)]
        sectores = [s for s in sectores_actuales() if s != SECTOR_OTRO]
        sector_default = zona.get("sector") or ""
        idx_sec = sectores.index(sector_default) if sector_default in sectores else 0
        sector = st.selectbox(
            "Sector asociado",
            sectores or [""],
            index=min(idx_sec, max(len(sectores) - 1, 0)) if sectores else 0,
            key="ub_alta_sector",
        )
        actualizar = st.checkbox(
            "Actualizar si ya existe (corrige zona, sector o clave)",
            key="ub_alta_actualizar",
        )
        if st.button("Guardar ubicación", type="primary", key="btn_alta_ubicacion"):
            error, fid = guardar_o_actualizar_ubicacion(nombre, zona["id"], sector, actualizar)
            if error:
                st.warning(error)
            else:
                st.session_state.ub_ficha_id = fid
                st.session_state.ub_zona_id = zona["id"]
                accion = "actualizada" if actualizar and encontrar_ficha(datos, nombre) else "guardada"
                st.success(f"«{nombre.strip()}» quedó {accion} en {UBICACIONES_FILE}.")
                st.rerun()


def render_editar_ubicacion(ficha: dict, datos: dict) -> None:
    with st.expander("✏️ Editar o eliminar esta ubicación", expanded=False):
        st.caption(
            "Corrige la clave, la zona o el sector. El equipo instalado y el historial se conservan. "
            "Eliminar quita el registro de ubicaciones_equipos.json."
        )
        zonas = datos["zonas"]
        etiquetas = [f"{z['prefijo']}: {z['nombre']}" for z in zonas]
        ids = [z["id"] for z in zonas]
        idx_zona = ids.index(ficha["zona_id"]) if ficha.get("zona_id") in ids else 0
        sectores = [s for s in sectores_actuales() if s != SECTOR_OTRO]
        idx_sec = sectores.index(ficha["sector"]) if ficha.get("sector") in sectores else 0
        col_n, col_z = st.columns(2)
        with col_n:
            nombre = st.text_input(
                "Clave o nombre",
                value=ficha.get("nombre", ""),
                key=f"ub_edit_nombre_{ficha['id']}",
            )
        with col_z:
            zona_etq = st.selectbox(
                "Zona / prefijo",
                etiquetas,
                index=idx_zona,
                key=f"ub_edit_zona_{ficha['id']}",
            )
        zona = zonas[etiquetas.index(zona_etq)]
        sector = st.selectbox(
            "Sector asociado",
            sectores or [""],
            index=min(idx_sec, max(len(sectores) - 1, 0)) if sectores else 0,
            key=f"ub_edit_sector_{ficha['id']}",
        )
        col_ok, col_del = st.columns(2)
        with col_ok:
            if st.button("💾 Guardar cambios de ubicación", type="primary", key=f"btn_edit_ub_{ficha['id']}"):
                error = actualizar_datos_ubicacion(ficha["id"], nombre, zona["id"], sector)
                if error:
                    st.warning(error)
                else:
                    nuevo_id = clave_ubicacion(nombre.strip()) or ficha["id"]
                    st.session_state.ub_ficha_id = nuevo_id
                    st.session_state.ub_zona_id = zona["id"]
                    st.success(f"«{nombre.strip()}» actualizada en {UBICACIONES_FILE}.")
                    st.rerun()
        with col_del:
            confirma = st.checkbox(
                f"Confirmo eliminar «{ficha.get('nombre', '')}»",
                key=f"ub_del_ok_{ficha['id']}",
            )
            if st.button("🗑️ Eliminar ubicación", key=f"btn_del_ub_{ficha['id']}"):
                if not confirma:
                    st.warning("Marca la casilla de confirmación para eliminar este registro.")
                else:
                    error = eliminar_ubicacion(ficha["id"])
                    if error:
                        st.warning(error)
                    else:
                        st.session_state.ub_ficha_id = None
                        st.success(f"«{ficha.get('nombre', '')}» se eliminó de {UBICACIONES_FILE}.")
                        st.rerun()


def render_form_equipo(ficha: dict, datos: dict) -> None:
    with st.expander("Actualizar equipo instalado o registrar reubicación", expanded=False):
        cat = obtener_catalogos()
        mci_opts = [""] + list(cat.get("equipos_mci") or EQUIPOS_MCI_DEFAULT)
        ub_opts = [""] + opciones_equipos_ub()
        col_m, col_u, col_e = st.columns(3)
        with col_m:
            mci = st.selectbox(
                "Motor MCI actual",
                mci_opts,
                index=mci_opts.index(ficha["equipo_mci"]) if ficha.get("equipo_mci") in mci_opts else 0,
                key=f"ub_mci_{ficha['id']}",
            )
        with col_u:
            ub = st.selectbox(
                "Unidad de bombeo (UB)",
                ub_opts,
                index=ub_opts.index(ficha["equipo_ub"]) if ficha.get("equipo_ub") in ub_opts else 0,
                key=f"ub_ub_{ficha['id']}",
            )
        with col_e:
            estatus = st.selectbox(
                "Estatus de la ubicación",
                ESTATUS_OPCIONES + ["Sin equipo"],
                index=(ESTATUS_OPCIONES + ["Sin equipo"]).index(ficha["estatus"])
                if ficha.get("estatus") in ESTATUS_OPCIONES + ["Sin equipo"]
                else 0,
                key=f"ub_est_{ficha['id']}",
            )
        notas = st.text_area(
            "Notas técnicas",
            value=ficha.get("notas") or "",
            key=f"ub_notas_{ficha['id']}",
            height=80,
        )
        if st.button("💾 Guardar ficha de equipo", key=f"btn_ficha_{ficha['id']}"):
            actualizar_ficha(
                ficha["id"],
                {"equipo_mci": mci, "equipo_ub": ub, "estatus": estatus, "notas": notas},
            )
            st.success("Ficha actualizada en ubicaciones_equipos.json.")
            st.rerun()

        st.markdown("**Reubicación de equipo**")
        st.caption("Si un motor o UB se movió desde otro pozo, regístralo para conservar el historial.")
        col_t, col_f, col_o = st.columns(3)
        with col_t:
            tipo = st.radio("Tipo de equipo", ["MCI", "UB"], horizontal=True, key=f"ub_mov_tipo_{ficha['id']}")
        with col_f:
            fecha_mov = st.date_input(
                "Fecha del cambio",
                value=date.today(),
                format="DD/MM/YYYY",
                key=f"ub_mov_fecha_{ficha['id']}",
            )
        with col_o:
            otras = [
                f"{f['nombre']}"
                for f in sorted(datos["fichas"].values(), key=lambda x: x["nombre"])
                if f["id"] != ficha["id"]
            ]
            origen_nom = st.selectbox(
                "Reubicado desde (opcional)",
                [""] + otras,
                key=f"ub_mov_origen_{ficha['id']}",
            )
        opciones_eq = mci_opts if tipo == "MCI" else ub_opts
        equipo_nuevo = st.selectbox(
            "Equipo que queda instalado",
            opciones_eq,
            key=f"ub_mov_eq_{ficha['id']}",
        )
        motivo = st.text_input(
            "Motivo / observación",
            placeholder="Cambio por falla, reubicación de MCI, modernización…",
            key=f"ub_mov_motivo_{ficha['id']}",
        )
        if st.button("Registrar cambio de equipo", type="primary", key=f"btn_mov_{ficha['id']}"):
            origen_id = ""
            if origen_nom:
                origen_id = next(
                    (f["id"] for f in datos["fichas"].values() if f["nombre"] == origen_nom),
                    "",
                )
            error = registrar_cambio_equipo(
                ficha["id"], tipo, equipo_nuevo, origen_id, motivo, fecha_mov
            )
            if error:
                st.warning(error)
            else:
                st.success("Cambio de equipo guardado.")
                st.rerun()


def render_ficha_ubicacion(ficha: dict, datos: dict) -> None:
    zona = mapa_zonas(datos["zonas"]).get(ficha.get("zona_id"), {})
    zona_txt = f"{ficha.get('prefijo', '')}: {zona.get('nombre', 'Ubicación')}"
    actualizado = ficha.get("actualizado") or "sin actualizaciones de ficha"
    st.markdown(
        f"""
        <div class="ficha-hero">
            <div class="prefijo">{html.escape(zona_txt)}</div>
            <h2>{html.escape(ficha.get("nombre", ""))}</h2>
            <p>{html.escape(ficha.get("sector") or "Sin sector asociado")} · Estatus: {html.escape(ficha.get("estatus") or "—")} · {html.escape(actualizado)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_mci, col_ub = st.columns(2, gap="medium")
    mci = ficha.get("equipo_mci") or "Sin motor MCI registrado"
    ub = ficha.get("equipo_ub") or "Sin unidad de bombeo registrada"
    cls_m = "equipo-card" if ficha.get("equipo_mci") else "equipo-card vacio"
    cls_u = "equipo-card" if ficha.get("equipo_ub") else "equipo-card vacio"
    with col_mci:
        st.markdown(
            f'<div class="{cls_m}"><div class="lbl">Motor MCI</div>'
            f"<h3>{html.escape(mci)}</h3>"
            f"<p>Equipo de combustión instalado en este pozo.</p></div>",
            unsafe_allow_html=True,
        )
    with col_ub:
        st.markdown(
            f'<div class="{cls_u}"><div class="lbl">Unidad de bombeo (UB)</div>'
            f"<h3>{html.escape(ub)}</h3>"
            f"<p>Unidad de bombeo (mecánica o hidráulica) asociada a la ubicación.</p></div>",
            unsafe_allow_html=True,
        )
    if ficha.get("notas"):
        st.caption(f"Notas: {ficha['notas']}")

    render_editar_ubicacion(ficha, datos)
    render_form_equipo(ficha, datos)

    df = cargar_bitacora()
    df_pozo = bitacora_de_pozo(df, ficha)
    programados = programados_de_pozo(ficha)
    pendientes = [p for p in programados if p.get("estatus") in {"No Cumplido", "Reprogramado"}]
    historial_eq = cambios_de_pozo(datos, ficha)
    personal = personal_de_pozo(df_pozo)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Intervenciones en bitácora", len(df_pozo))
    m2.metric("Cambios de equipo", len(historial_eq))
    m3.metric("Programados pendientes", len(pendientes))
    m4.metric("Cuadrillas que intervinieron", len(personal))

    if st.session_state.get("ub_ficha_seccion_pozo") != ficha["id"]:
        st.session_state.ub_ficha_seccion = None
        st.session_state.ub_ficha_seccion_pozo = ficha["id"]

    seccion = st.session_state.get("ub_ficha_seccion")
    if not seccion:
        render_cuadricula_explorador(
            SUBMENUS_FICHA,
            "ub_ficha_seccion",
            "Historial y personal de esta ubicación. Abre una tarjeta para ver el detalle.",
        )
        return
    render_barra_volver(
        f'{html.escape(ficha.get("nombre", ""))} / <b>{html.escape(titulo_submenu(SUBMENUS_FICHA, seccion))}</b>',
        "ub_ficha_seccion",
        "← Volver a la ficha",
    )

    if seccion == "mantenimientos":
        st.caption(
            "Registros de bitacora_mantenimiento.xlsx ligados a este pozo "
            "(o al sector, si el renglón no trae clave de pozo)."
        )
        if df_pozo.empty:
            st.info("Aún no hay mantenimientos de campo para esta ubicación.")
        else:
            vista = df_pozo.copy()
            vista["Clase"] = vista.apply(clasificar_mantenimiento, axis=1)
            cols = [
                "FECHA",
                "SECUENCIA",
                "Clase",
                "TIPO DE ACTIVIDAD",
                "ACTIVIDAD CORTA",
                "CUADRILLA",
                "EQUIPO/UBICACIÓN",
                "EVALUACIÓN DEL ESTATUS",
                "DESCRIPCIÓN DE LA ACTIVIDAD",
            ]
            st.dataframe(vista[cols], use_container_width=True, hide_index=True)
            c1, c2 = st.columns(2)
            c1.metric("Programados en bitácora", int((vista["Clase"] == "Programado").sum()))
            c2.metric("No programados / emergencias", int((vista["Clase"] == "No programado").sum()))

    elif seccion == "equipos":
        if not historial_eq:
            st.info("Sin cambios de MCI/UB registrados. Usa el formulario de reubicación para dejar evidencia.")
        else:
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Fecha": c.get("fecha", ""),
                            "Tipo": c.get("tipo", ""),
                            "Anterior": c.get("equipo_anterior") or "—",
                            "Instalado": c.get("equipo_nuevo", ""),
                            "Vino de": c.get("origen") or "Alta / reemplazo local",
                            "Pozo destino": c.get("pozo", ""),
                            "Motivo": c.get("motivo", ""),
                        }
                        for c in historial_eq
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )

    elif seccion == "pendientes":
        st.caption(
            "Pendientes tomados de programados_ops.json. "
            "Si el programado solo tiene sector (sin pozo), se muestra en todos los pozos de esa zona."
        )
        if not pendientes:
            st.info("No hay mantenimientos pendientes o reprogramados para esta ubicación.")
        else:
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Mes": etiqueta_mes(str(p.get("mes", ""))),
                            "Actividad": p.get("actividad", ""),
                            "Ubicación plan": p.get("ubicacion", ""),
                            "Equipo": p.get("equipo", ""),
                            "Estatus": p.get("estatus", ""),
                            "Observación": p.get("observacion", ""),
                        }
                        for p in pendientes
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
        cumplidos = [p for p in programados if p.get("estatus") == "Cumplido"]
        if cumplidos:
            st.markdown("**Programados cumplidos (histórico OPS)**")
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Mes": etiqueta_mes(str(p.get("mes", ""))),
                            "Actividad": p.get("actividad", ""),
                            "Equipo": p.get("equipo", ""),
                            "Estatus": p.get("estatus", ""),
                        }
                        for p in cumplidos
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )

    elif seccion == "personal":
        if not personal:
            st.info("Todavía no hay cuadrillas registradas en la bitácora para este pozo.")
        else:
            for item in personal:
                st.markdown(
                    f"""
                    <div class="personal-chip">
                        <div class="cuadrilla">{html.escape(item["cuadrilla"])}</div>
                        <p>{item["intervenciones"]} intervención(es) · Última: {html.escape(item["ultima"] or "s/f")}<br>
                        {html.escape(item.get("nombres") or "Sin personal de turno capturado")}<br>
                        {html.escape(item["actividades"])}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )


def render_migracion_ubicaciones() -> None:
    st.markdown(
        '<div class="section-title">Migración de ubicaciones desde Excel</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Lee las tablas por zona (Aragón, Palo Blanco, Coyotes, Gallo, Soledad y Soledad Norte) "
        "con columnas UBICACIÓN, MCI y UB, y actualiza ubicaciones_equipos.json."
    )
    archivos = listar_excels_inventario_ubicaciones()
    if not archivos:
        st.info(
            "Coloca el inventario en la carpeta ubicacion_equipo/ "
            "(por ejemplo ubicaciones_equipos.json.xlsx)."
        )
        return
    etiquetas = [os.path.relpath(r) for r in archivos]
    elegido = st.selectbox("Archivo de inventario", etiquetas, key="mig_excel_sel")
    ruta = archivos[etiquetas.index(elegido)]
    mig = cargar_ubicaciones().get("migracion") or {}
    if mig:
        st.caption(
            f"Última importación: {mig.get('fecha', 's/f')} · "
            f"{mig.get('nuevos', 0)} altas · {mig.get('actualizados', 0)} actualizados · "
            f"{mig.get('total_excel', 0)} filas del Excel."
        )
    if st.button("📥 Importar / actualizar pozos desde Excel", type="primary", key="btn_migrar_excel"):
        resumen = importar_excel_ubicaciones(ruta, forzar=True)
        if not resumen.get("ok"):
            st.error(resumen.get("error") or "No se pudo importar.")
        else:
            st.success(
                f"Se cargaron {resumen['total']} pozos desde {resumen.get('archivo', 'Excel')}: "
                f"{resumen['nuevos']} nuevos y {resumen['actualizados']} actualizados "
                f"en {UBICACIONES_FILE}."
            )
            preview = resumen.get("preview") or []
            if preview:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "Ubicación": p["nombre"],
                                "Zona": p["zona_id"],
                                "MCI": p["equipo_mci"] or p["mci_raw"],
                                "UB": p["equipo_ub"] or p["ub_raw"],
                            }
                            for p in preview
                        ]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
            st.rerun()


def _tamano_archivo_legible(ruta: str) -> str:
    try:
        nbytes = os.path.getsize(ruta)
    except OSError:
        return ""
    if nbytes < 1024:
        return f"{nbytes} B"
    if nbytes < 1024 * 1024:
        return f"{nbytes / 1024:.0f} KB"
    return f"{nbytes / (1024 * 1024):.1f} MB"


def ruta_mapa_soledad() -> str | None:
    os.makedirs(MAPA_DIR, exist_ok=True)
    candidatos = [
        MAPA_SOLEDAD_FILE,
        os.path.join(MAPA_DIR, "mapa_soledad.pdf.pdf"),
        os.path.join(MAPA_DIR, "Mapa_Soledad.pdf"),
        os.path.join(MAPA_DIR, "mapa_soledad.PDF"),
    ]
    for ruta in candidatos:
        if os.path.isfile(ruta):
            return ruta
    try:
        nombres = os.listdir(MAPA_DIR)
    except OSError:
        return None
    for nombre in nombres:
        clave = nombre.casefold()
        if clave.startswith("mapa_soledad") and clave.endswith(".pdf"):
            return os.path.join(MAPA_DIR, nombre)
    for nombre in nombres:
        if nombre.casefold().endswith(".pdf"):
            return os.path.join(MAPA_DIR, nombre)
    return None


def asegurar_mapa_soledad() -> str | None:
    hallado = ruta_mapa_soledad()
    if not hallado:
        return None
    if os.path.normcase(os.path.abspath(hallado)) == os.path.normcase(os.path.abspath(MAPA_SOLEDAD_FILE)):
        return MAPA_SOLEDAD_FILE
    if os.path.isfile(MAPA_SOLEDAD_FILE):
        return MAPA_SOLEDAD_FILE
    try:
        shutil.copy2(hallado, MAPA_SOLEDAD_FILE)
        return MAPA_SOLEDAD_FILE
    except OSError:
        return hallado


def imagenes_desde_pdf(datos: bytes, escala: float = 2.0) -> list:
    import pypdfium2 as pdfium

    documento = pdfium.PdfDocument(datos)
    imagenes = []
    try:
        for indice in range(len(documento)):
            pagina = documento[indice]
            bitmap = pagina.render(scale=escala)
            imagenes.append(bitmap.to_pil())
    finally:
        documento.close()
    return imagenes


def html_visor_pdf(datos: bytes, titulo: str, altura: int = 820) -> str:
    b64 = base64.b64encode(datos).decode("ascii")
    titulo_attr = html.escape(titulo, quote=True)
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<title>{titulo_attr}</title>
<style>
  html, body {{
    margin: 0;
    height: 100%;
    background: #0b1220;
    color: #e2e8f0;
    font-family: "Segoe UI", system-ui, sans-serif;
  }}
  #barra {{
    display: flex;
    gap: 8px;
    align-items: center;
    padding: 8px 10px;
    background: #0f172a;
    border-bottom: 1px solid #334155;
  }}
  #barra button {{
    background: #1e3a5f;
    color: #e2e8f0;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 6px 10px;
    cursor: pointer;
    font-weight: 600;
  }}
  #barra button:hover {{ background: #16382c; }}
  #pagina {{ color: #94a3b8; font-size: 0.85rem; min-width: 72px; text-align: center; }}
  #visor {{
    height: calc(100% - 46px);
    overflow: auto;
    text-align: center;
    background: #0b1220;
  }}
  canvas {{
    margin: 10px auto 16px;
    max-width: none;
    background: #0f172a;
    box-shadow: 0 10px 28px rgba(2, 6, 23, 0.45);
    border: 1px solid #334155;
    border-radius: 8px;
  }}
  #error {{
    padding: 1.2rem;
    color: #fca5a5;
    display: none;
  }}
</style>
</head>
<body>
<div id="barra">
  <button id="prev" type="button">◀</button>
  <span id="pagina">1 / 1</span>
  <button id="next" type="button">▶</button>
  <button id="menos" type="button">−</button>
  <button id="mas" type="button">+</button>
  <span style="color:#64748b;font-size:0.8rem;margin-left:6px;">Ajustar zoom para leer el plano</span>
</div>
<div id="error">No se pudo renderizar el PDF en el visor. Usa «Descargar PDF».</div>
<div id="visor"><canvas id="lienzo"></canvas></div>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script>
  const errorBox = document.getElementById("error");
  function fallar(msg) {{
    errorBox.style.display = "block";
    errorBox.textContent = msg || "No se pudo renderizar el PDF.";
  }}
  try {{
    if (!window.pdfjsLib) throw new Error("No cargó el visor PDF.js");
    pdfjsLib.GlobalWorkerOptions.workerSrc = "";
    const raw = atob("{b64}");
    const bytes = new Uint8Array(raw.length);
    for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i);
    let pdfDoc = null, pageNum = 1, scale = 1.15;
    const canvas = document.getElementById("lienzo");
    const ctx = canvas.getContext("2d");
    async function pintar() {{
      const page = await pdfDoc.getPage(pageNum);
      const viewport = page.getViewport({{ scale: scale }});
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      await page.render({{ canvasContext: ctx, viewport: viewport }}).promise;
      document.getElementById("pagina").textContent = pageNum + " / " + pdfDoc.numPages;
    }}
    pdfjsLib.getDocument({{ data: bytes, disableWorker: true }}).promise.then((pdf) => {{
      pdfDoc = pdf;
      pintar();
    }}).catch((err) => fallar(String(err)));
    document.getElementById("prev").onclick = () => {{ if (pdfDoc && pageNum > 1) {{ pageNum -= 1; pintar(); }} }};
    document.getElementById("next").onclick = () => {{ if (pdfDoc && pageNum < pdfDoc.numPages) {{ pageNum += 1; pintar(); }} }};
    document.getElementById("menos").onclick = () => {{ scale = Math.max(0.5, scale - 0.2); pintar(); }};
    document.getElementById("mas").onclick = () => {{ scale = Math.min(3, scale + 0.2); pintar(); }};
  }} catch (err) {{
    fallar(String(err));
  }}
</script>
</body>
</html>"""


def parsear_nombre_ot_diario(nombre: str) -> date | None:
    base = os.path.basename(str(nombre or ""))
    while base.casefold().endswith(".pdf"):
        base = base[:-4]
    m = PATRON_OT_DIARIO.fullmatch(base.strip())
    if not m:
        return None
    dia, mes, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(anio, mes, dia)
    except ValueError:
        return None


def etiqueta_dia_ot(dia: date) -> str:
    return f"{dia.strftime('%d/%m/%Y')} · {NOMBRES_DIA[dia.weekday()]}"


def raices_reportes_ot() -> list[str]:
    """Carpetas donde pueden vivir los PDF diarios, de la preferida a las de respaldo."""
    candidatas = [BITACORA_PDF_DIR, HISTORIAL_OT_DIR, *HISTORIAL_OT_RUTAS_EXTRA]
    raices: list[str] = []
    vistas: set[str] = set()
    for ruta in candidatas:
        absoluta = os.path.abspath(ruta)
        clave = os.path.normcase(os.path.normpath(absoluta))
        if clave in vistas or not os.path.isdir(absoluta):
            continue
        vistas.add(clave)
        raices.append(absoluta)
    return raices


def iterar_pdfs_ot_diarios(carpeta: str):
    """PDF DDMMYYYY.pdf en la carpeta y en subcarpetas de mes (septiembre, octubre, …)."""
    try:
        recorridos = os.walk(carpeta)
    except OSError:
        return
    for dirpath, _dirnames, filenames in recorridos:
        for nombre in filenames:
            if not nombre.casefold().endswith(".pdf"):
                continue
            dia = parsear_nombre_ot_diario(nombre)
            if not dia:
                continue
            ruta = os.path.join(dirpath, nombre)
            if os.path.isfile(ruta):
                yield dia, ruta, nombre


def carpeta_historial_ot() -> str:
    for raiz in raices_reportes_ot():
        if any(True for _ in iterar_pdfs_ot_diarios(raiz)):
            return raiz
    preferida = os.path.abspath(BITACORA_PDF_DIR)
    os.makedirs(preferida, exist_ok=True)
    return preferida


def listar_reportes_ot_diarios() -> list[dict]:
    reportes: list[dict] = []
    vistos: set[date] = set()
    for raiz in raices_reportes_ot():
        for dia, ruta, nombre in iterar_pdfs_ot_diarios(raiz):
            if dia in vistos:
                continue
            vistos.add(dia)
            try:
                peso = os.path.getsize(ruta)
            except OSError:
                peso = 0
            reportes.append(
                {
                    "fecha": dia,
                    "ruta": ruta,
                    "nombre": nombre,
                    "bytes": peso,
                    "etiqueta": etiqueta_dia_ot(dia),
                    "liberado": False,
                }
            )
    for fecha_iso, meta in persistencia.leer_ot_liberadas().items():
        try:
            anio_iso, mes_iso, dia_iso = (int(parte) for parte in str(fecha_iso).split("-"))
            dia = date(anio_iso, mes_iso, dia_iso)
        except (TypeError, ValueError):
            continue
        if dia in vistos:
            continue
        vistos.add(dia)
        nombre = str((meta or {}).get("nombre") or "").strip() or f"{dia.day:02d}{dia.month:02d}{dia.year}.pdf"
        reportes.append(
            {
                "fecha": dia,
                "ruta": "",
                "nombre": nombre,
                "bytes": 0,
                "etiqueta": etiqueta_dia_ot(dia),
                "liberado": True,
            }
        )
    reportes.sort(key=lambda r: r["fecha"], reverse=True)
    return reportes


@st.cache_data(show_spinner=False)
def paginas_pdf_ot(ruta: str, mtime: float) -> int:
    import pypdfium2 as pdfium

    origen = _copia_excel_legible(ruta)
    documento = pdfium.PdfDocument(origen)
    try:
        return len(documento)
    finally:
        documento.close()


@st.cache_data(show_spinner=False)
def bytes_pdf_ot(ruta: str, mtime: float) -> bytes:
    with open(ruta, "rb") as archivo:
        return archivo.read()


@st.cache_data(show_spinner="Preparando hoja del reporte…")
def jpeg_pagina_reporte_ot(ruta: str, mtime: float, indice: int, escala: float = 1.35) -> tuple[str, int]:
    import pypdfium2 as pdfium

    origen = _copia_excel_legible(ruta)
    documento = pdfium.PdfDocument(origen)
    try:
        total = len(documento)
        if total <= 0:
            return "", 0
        indice = max(0, min(int(indice), total - 1))
        pagina = documento[indice]
        bitmap = pagina.render(scale=escala)
        imagen = bitmap.to_pil().convert("RGB")
    finally:
        documento.close()
    buf = io.BytesIO()
    imagen.save(buf, format="JPEG", quality=78, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii"), total


def ubicaciones_secuencia_bloque(bloque: pd.DataFrame) -> list[str]:
    if bloque is None or bloque.empty:
        return []
    filas = filas_secuencia_historial(bloque)
    sitios = []
    for indice, fila in enumerate(filas, start=1):
        sitio = etiqueta_ubicacion_fila(fila) or texto_visible(
            fila.get("UBICACIÓN") or fila.get("POZO") or fila.get("EQUIPO/UBICACIÓN") or fila.get("SECTOR"),
            "",
        )
        if not sitio or _sin_acentos(sitio).casefold() in {"sin pozo", "sin ubicacion"}:
            continue
        if sitios_operativos.es_vacio_sitio(sitio):
            continue
        numero = numero_secuencia_actividad(fila.get("SECUENCIA")) or indice
        sitios.append(f"{numero:02d} {sitio}")
    return sitios


def carpeta_mes_bitacora_pdf(dia: date) -> str:
    nombre_mes = NOMBRES_MES[dia.month - 1]
    base = os.path.abspath(BITACORA_PDF_DIR)
    os.makedirs(base, exist_ok=True)
    clave = _sin_acentos(nombre_mes).casefold()
    for entrada in os.listdir(base):
        ruta = os.path.join(base, entrada)
        if os.path.isdir(ruta) and _sin_acentos(entrada).casefold() == clave:
            return ruta
    destino = os.path.join(base, nombre_mes.casefold())
    os.makedirs(destino, exist_ok=True)
    return destino


def nombre_pdf_ot_diario(dia: date) -> str:
    return f"{dia.day:02d}{dia.month:02d}{dia.year}.pdf"


def guardar_pdf_ot_diario(archivo, dia: date) -> dict:
    if archivo is None:
        return {"ok": False, "mensaje": "Selecciona un archivo PDF.", "ruta": "", "fecha": dia}
    datos = archivo.getvalue() if hasattr(archivo, "getvalue") else b""
    if not datos:
        return {"ok": False, "mensaje": "El archivo está vacío.", "ruta": "", "fecha": dia}
    if not datos.startswith(b"%PDF"):
        return {"ok": False, "mensaje": "El archivo no tiene formato PDF.", "ruta": "", "fecha": dia}
    destino_dia = dia if isinstance(dia, date) else date.today()
    carpeta = carpeta_mes_bitacora_pdf(destino_dia)
    nombre = nombre_pdf_ot_diario(destino_dia)
    ruta = os.path.join(carpeta, nombre)
    with open(ruta, "wb") as salida:
        salida.write(datos)
    persistencia.olvidar_ot_liberada(destino_dia.isoformat())
    return {
        "ok": True,
        "mensaje": "",
        "ruta": ruta.replace("\\", "/"),
        "nombre": nombre,
        "fecha": destino_dia,
    }


def reporte_ot_de_fecha(dia: date) -> dict | None:
    if not isinstance(dia, date):
        return None
    for reporte in listar_reportes_ot_diarios():
        if reporte.get("fecha") == dia:
            return reporte
    return None


def render_carga_pdf_ot(dia_sugerido: date) -> None:
    aviso = st.session_state.pop("ot_pdf_aviso", None)
    if aviso:
        st.success(aviso)
    if "fecha_carga_pdf" not in st.session_state:
        sugerido = dia_sugerido if isinstance(dia_sugerido, date) else date.today()
        st.session_state.fecha_carga_pdf = sugerido
    st.markdown("##### Cargar PDF del día")
    st.caption(
        "Este calendario solo indica a qué día pertenece el archivo. "
        "No cambia el reporte que se está viendo. Al guardar, queda en `bitacora_pdf` "
        "como DDMMYYYY.pdf, dentro de la carpeta del mes."
    )
    col_fecha, col_estado = st.columns([1.05, 1.45], gap="medium")
    with col_fecha:
        dia = st.date_input(
            "Día del archivo a respaldar",
            format="DD/MM/YYYY",
            key="fecha_carga_pdf",
        )
    with col_estado:
        respaldo = reporte_ot_de_fecha(dia) if isinstance(dia, date) else None
        if respaldo:
            marca = (
                f'<div class="pdf-alerta ok">PDF respaldado · {html.escape(str(respaldo.get("nombre") or ""))}</div>'
            )
        else:
            marca = '<div class="pdf-alerta falta">⚠️ Falta evidencia de PDF para este día</div>'
        st.markdown(marca, unsafe_allow_html=True)
    carga = int(st.session_state.get("ot_pdf_carga") or 0)
    archivo = st.file_uploader(
        "PDF de la orden de trabajo",
        type=["pdf"],
        key=f"ot_pdf_archivo_{carga}",
    )
    if st.button("Guardar PDF en el historial", type="primary", key="ot_pdf_guardar"):
        if not isinstance(dia, date):
            st.warning("Elige la fecha a la que pertenece el PDF.")
            return
        resultado = guardar_pdf_ot_diario(archivo, dia)
        if not resultado.get("ok"):
            st.warning(resultado.get("mensaje") or "No se guardó el PDF.")
            return
        st.session_state.ot_pdf_carga = carga + 1
        st.session_state.ot_pdf_aviso = (
            f"El reporte {resultado['nombre']} quedó en {resultado['ruta']} "
            f"para el {resultado['fecha'].strftime('%d/%m/%Y')}."
        )
        st.rerun()


def _seleccionar_cuadrilla_ot(dia_iso: str, nombre: str, indice: int = 0) -> None:
    mismo = (
        st.session_state.get("ot_cuadrilla_dia") == dia_iso
        and st.session_state.get("ot_cuadrilla_nombre") == nombre
        and int(st.session_state.get("ot_cuadrilla_indice") or -1) == int(indice)
    )
    if mismo:
        st.session_state.ot_cuadrilla_nombre = ""
        st.session_state.ot_cuadrilla_indice = None
        return
    st.session_state.ot_cuadrilla_dia = dia_iso
    st.session_state.ot_cuadrilla_nombre = nombre
    st.session_state.ot_cuadrilla_indice = int(indice)


def _cerrar_cuadrilla_ot() -> None:
    st.session_state.ot_cuadrilla_nombre = ""


def fotos_actividad_ot(fila, dia: date, cuadrilla: str) -> list[str]:
    fotos = fotos_evidencia_octubre_secuencia(cuadrilla, dia, fila.get("SECUENCIA"))
    if fotos:
        return fotos
    return [ruta for ruta in rutas_fotos_fila(fila) if ruta and os.path.isfile(ruta)]


def render_grid_evidencia_actividad(fotos: list[str]) -> None:
    st.markdown("**Evidencia fotográfica**")
    if not fotos:
        st.caption("Sin fotografías para esta secuencia.")
        return
    st.caption(f"{len(fotos)} fotografía(s) de esta actividad.")
    for inicio in range(0, len(fotos), 3):
        columnas = st.columns(3, gap="small")
        for col, ruta in zip(columnas, fotos[inicio : inicio + 3]):
            with col:
                st.image(ruta, width="stretch")
                st.caption(os.path.basename(ruta))


def _texto_equipo_catalogo(valor: str) -> str:
    texto = str(valor or "").strip()
    if texto.casefold() in {"", "—", "-", "nan", "none", "ninguno", "n/a", "sin mci", "sin ub"}:
        return ""
    return texto


def ficha_por_ubicacion(datos: dict, ubicacion: str) -> dict | None:
    texto = str(ubicacion or "").strip()
    if not texto:
        return None
    ficha = encontrar_ficha(datos, texto)
    if ficha:
        return ficha
    indice = _indice_catalogo(datos)
    for clave in _variantes_clave_ubicacion(texto, canonizar_nombre_ubicacion(texto)):
        hallada = indice.get(clave)
        if hallada:
            return hallada
    return ficha_infraestructura(texto)


def equipo_catalogo_ubicacion(ubicacion: str, datos: dict | None = None) -> str:
    texto = str(ubicacion or "").strip()
    if not texto or texto in {"—", "-", "nan", "None"} or sitios_operativos.es_vacio_sitio(texto):
        return ""
    try:
        datos = datos if datos is not None else cargar_ubicaciones()
        ficha = ficha_por_ubicacion(datos, texto)
    except Exception:
        ficha = None
    if ficha:
        partes = [
            _texto_equipo_catalogo(ficha.get("equipo_mci")),
            _texto_equipo_catalogo(ficha.get("equipo_ub")),
        ]
        partes = [parte for parte in partes if parte]
        if partes:
            return " · ".join(partes)
        if sitios_operativos.es_instalacion(texto) or sitios_operativos.es_instalacion(ficha.get("nombre")):
            return ""
        return ""
    if sitios_operativos.es_instalacion(texto):
        return ""
    return "No especificado en catálogo"


def render_actividad_cuadrilla_ot(
    dia: date,
    fila,
    indice: int,
    cuadrilla: str,
    datos_ubicaciones: dict | None = None,
) -> None:
    numero = numero_secuencia_actividad(fila.get("SECUENCIA")) or indice
    ubicacion = etiqueta_ubicacion_fila(fila) or texto_visible(
        fila.get("UBICACIÓN") or fila.get("POZO") or fila.get("SECTOR") or fila.get("EQUIPO/UBICACIÓN")
    )
    equipo = equipo_catalogo_ubicacion(ubicacion, datos_ubicaciones)
    if not equipo or equipo == "No especificado en catálogo":
        respaldo = texto_visible(fila.get("EQUIPO/UBICACIÓN"), "")
        if (
            respaldo
            and sitios_operativos.es_equipo_explicito(respaldo)
            and _sin_acentos(respaldo).casefold() != _sin_acentos(ubicacion).casefold()
        ):
            equipo = respaldo
        else:
            equipo = ""
    personal = texto_visible(fila.get("PERSONAL"), "")
    if not personal:
        personal = texto_par_personal(
            str(fila.get("MECANICO EN TURNO") or fila.get("MECANICO TITULAR") or ""),
            str(fila.get("AYUDANTE EN TURNO") or fila.get("AYUDANTE TITULAR") or ""),
        )
    descripcion = texto_visible(
        fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD"),
        "Sin descripción de la OT",
    )
    comentarios = texto_visible(
        fila["COMENTARIOS"] if "COMENTARIOS" in getattr(fila, "index", []) else fila.get("OBSERVACIÓN REVISIÓN")
    )
    materiales = consumibles_de_fila(fila) or "Sin materiales registrados"
    fotos = fotos_actividad_ot(fila, dia, cuadrilla)
    with st.container(border=True):
        st.markdown(
            f"""
            <div class="ot-sec">Datos de la OT</div>
            <div class="ot-grid">
                <div class="ot-field ot-span"><div class="k">Secuencia</div>
                <div class="v">#{numero:02d}</div></div>
                <div class="ot-field ot-span"><div class="k">Descripción</div>
                <div class="v">{html.escape(descripcion)}</div></div>
                <div class="ot-field"><div class="k">Equipo</div>
                <div class="v">{html.escape(equipo or "—")}</div></div>
                <div class="ot-field"><div class="k">Ubicación</div>
                <div class="v">{html.escape(ubicacion)}</div></div>
                <div class="ot-field ot-span"><div class="k">Personal</div>
                <div class="v">{html.escape(personal)}</div></div>
            </div>
            <div class="ot-sec">Comentarios y materiales</div>
            <div class="ot-grid">
                <div class="ot-field ot-span"><div class="k">Comentarios de la orden de trabajo</div>
                <div class="v">{html.escape(comentarios)}</div></div>
                <div class="ot-field ot-span"><div class="k">Materiales empleados</div>
                <div class="v">{html.escape(materiales)}</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        render_grid_evidencia_actividad(fotos)


def render_desglose_cuadrilla_ot(
    dia: date,
    df_dia: pd.DataFrame,
    nombre: str,
    token: str = "",
) -> None:
    bloque = bloque_historial_cuadrilla(df_dia, nombre)
    etiqueta = etiqueta_cuadrilla_historial(nombre)
    datos_ubicaciones = cargar_ubicaciones()
    filas = filas_secuencia_historial(bloque) if not bloque.empty else []
    clave_btn = token or f"{dia.isoformat()}_{clave_hist_cuadrilla(nombre) or 'sin'}"
    cierra, titulo = st.columns([1, 1.6])
    with cierra:
        st.button(
            "Cerrar desglose",
            key=f"ot_cerrar_{clave_btn}",
            on_click=_cerrar_cuadrilla_ot,
            use_container_width=True,
        )
    with titulo:
        st.markdown(f"**{etiqueta}**")
    if not filas:
        st.info("No hay actividades de esta cuadrilla en la fecha seleccionada.")
        return
    st.caption(f"{len(filas)} actividad(es) en la jornada del {dia.strftime('%d/%m/%Y')}.")
    for indice, fila in enumerate(filas, start=1):
        render_actividad_cuadrilla_ot(
            dia, fila, indice, str(fila.get("CUADRILLA") or nombre), datos_ubicaciones
        )


def render_resumen_cuadrillas_ot_dia(dia: date, df_dia: pd.DataFrame | None = None) -> None:
    df = df_dia if df_dia is not None else bitacora_del_dia(bitacora_con_fechas(cargar_bitacora()), dia)
    grupos = resumen_cuadrillas_dia(df) if df is not None and not df.empty else []
    if not grupos:
        st.caption("No hay capturas de bitácora para este día. El PDF oficial se muestra a la izquierda.")
        return
    dia_iso = dia.isoformat()
    try:
        indice_sel = st.session_state.get("ot_cuadrilla_indice")
        indice_sel = int(indice_sel) if indice_sel is not None and indice_sel != "" else None
    except (TypeError, ValueError):
        indice_sel = None
    st.caption(
        f"{len(grupos)} cuadrilla(s) con reporte en bitácora el {dia.strftime('%d/%m/%Y')}. "
        "Elige una para ver sus actividades."
    )
    for indice, grupo in enumerate(grupos):
        bloque = df.loc[grupo["indices"]] if grupo.get("indices") else df.iloc[0:0]
        sitios = ubicaciones_secuencia_bloque(bloque)
        if sitios:
            detalle = " · ".join(html.escape(sitio) for sitio in sitios)
        else:
            detalle = "Sin ubicación registrada"
        nombre = str(grupo["cuadrilla"])
        slug = clave_hist_cuadrilla(nombre) or "sin"
        token = f"{dia_iso}_{indice}_{slug}"
        activa = (
            st.session_state.get("ot_cuadrilla_dia") == dia_iso
            and indice_sel == indice
            and str(st.session_state.get("ot_cuadrilla_nombre") or "") == nombre
        )
        clase = f"{grupo['marca']}{' activa' if activa else ''}"
        st.markdown(
            f"""
            <div class="cuadrilla-dia {html.escape(clase)}">
                <div class="lbl"><span class="semaforo {html.escape(grupo['marca'])}"></span>
                {html.escape(etiqueta_semaforo(grupo['marca']))}</div>
                <h3>{html.escape(etiqueta_cuadrilla_historial(nombre))}</h3>
                <p>{int(grupo['reportes'])} OT</p>
                <p>{detalle}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.button(
            "Ver actividades",
            key=f"ot_ver_act_{token}",
            on_click=_seleccionar_cuadrilla_ot,
            args=(dia_iso, nombre, indice),
            use_container_width=True,
        )
        if activa:
            render_desglose_cuadrilla_ot(dia, df, nombre, token)


def _seleccionar_dia_ot(etiqueta: str) -> None:
    st.session_state.ot_hist_sel = etiqueta
    st.session_state.ot_hist_pagina = 1


def _cambiar_hoja_ot(delta: int, total: int) -> None:
    actual = int(st.session_state.get("ot_hist_pagina") or 1)
    st.session_state.ot_hist_pagina = max(1, min(int(total), actual + int(delta)))


def render_historial_diario_ot() -> None:
    st.markdown('<div class="section-title">Historial diario de órdenes de trabajo</div>', unsafe_allow_html=True)
    carpeta = carpeta_historial_ot()
    st.markdown(
        f"""
        <div class="mapa-shell">
            <h3>Reportes diarios de OT</h3>
            <p>Se leen los PDF con nombre <b>DDMMYYYY.pdf</b> (ejemplo 01092026.pdf) en
            <code>{html.escape(carpeta)}</code> y en sus subcarpetas de mes
            (septiembre, octubre, …). Al descargar el PDF se retira del servidor.
            La bitácora del día permanece.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    render_generador_reporte_fotografico(key_prefix="ot_foto")
    reportes = listar_reportes_ot_diarios()
    if not reportes:
        st.warning(
            f"No hay PDF diarios todavía. Carga un archivo o colócalo como `01092026.pdf` en "
            f"`{BITACORA_PDF_DIR}/octubre/` u otra subcarpeta de mes."
        )
        render_carga_pdf_ot(date.today())
        return

    etiquetas = [r["etiqueta"] for r in reportes]
    if st.session_state.get("ot_hist_sel") not in etiquetas:
        st.session_state.ot_hist_sel = etiquetas[0]
    idx = etiquetas.index(st.session_state.ot_hist_sel)

    n1, n2, n3 = st.columns([1, 2.6, 1])
    with n1:
        st.button(
            "◀ Día anterior",
            use_container_width=True,
            disabled=idx >= len(etiquetas) - 1,
            key="ot_hist_prev",
            on_click=_seleccionar_dia_ot,
            args=(etiquetas[min(idx + 1, len(etiquetas) - 1)],),
        )
    with n2:
        elegido = st.selectbox(
            "Días disponibles",
            etiquetas,
            key="ot_hist_sel",
        )
    with n3:
        st.button(
            "Día siguiente ▶",
            use_container_width=True,
            disabled=idx <= 0,
            key="ot_hist_next",
            on_click=_seleccionar_dia_ot,
            args=(etiquetas[max(idx - 1, 0)],),
        )

    reporte = next((r for r in reportes if r["etiqueta"] == elegido), reportes[0])
    dia = reporte["fecha"]
    bitacora = bitacora_del_dia(bitacora_con_fechas(cargar_bitacora()), dia)
    grupos = resumen_cuadrillas_dia(bitacora) if bitacora is not None and not bitacora.empty else []
    ots = int(len(bitacora)) if bitacora is not None else 0

    k1, k2, k3 = st.columns(3, gap="medium")
    k1.markdown(kpi_html("Fecha", dia.strftime("%d/%m/%Y"), NOMBRES_DIA[dia.weekday()], "info"), unsafe_allow_html=True)
    k2.markdown(
        kpi_html("OT en bitácora", f"{ots}", f"{len(grupos)} cuadrilla(s)", "warn" if ots else "info"),
        unsafe_allow_html=True,
    )
    k3.markdown(
        kpi_html("Cuadrillas", f"{len(grupos)}", "Con actividad en la jornada", "ok" if grupos else "info"),
        unsafe_allow_html=True,
    )
    render_carga_pdf_ot(dia)

    if st.session_state.get("ot_hist_dia_ref") != dia.isoformat():
        st.session_state.ot_hist_dia_ref = dia.isoformat()
        st.session_state.ot_hist_pagina = 1

    ruta_pdf = str(reporte.get("ruta") or "")
    archivo_pdf = bool(ruta_pdf) and os.path.isfile(ruta_pdf)
    preview, total, mtime = "", 0, 0.0
    if archivo_pdf:
        try:
            mtime = os.path.getmtime(ruta_pdf)
            total = paginas_pdf_ot(ruta_pdf, mtime)
        except Exception as error:
            st.error(f"No se pudo abrir el PDF: {error}")

    cuadrilla_abierta = (
        st.session_state.get("ot_cuadrilla_dia") == dia.isoformat()
        and bool(st.session_state.get("ot_cuadrilla_nombre"))
    )
    visor, resumen = st.columns([1, 1.45] if cuadrilla_abierta else [1.35, 1], gap="large")
    with visor:
        st.markdown("**Reporte del día**")
        pagina_actual = 1
        if total > 1:
            if int(st.session_state.get("ot_hist_pagina") or 1) > total:
                st.session_state.ot_hist_pagina = 1
            h1, h2, h3 = st.columns([1, 1.3, 1])
            with h1:
                st.button(
                    "◀ Hoja",
                    use_container_width=True,
                    disabled=int(st.session_state.get("ot_hist_pagina") or 1) <= 1,
                    key="ot_hist_hoja_prev",
                    on_click=_cambiar_hoja_ot,
                    args=(-1, total),
                )
            with h2:
                pagina_actual = int(
                    st.number_input(
                        "Hoja",
                        min_value=1,
                        max_value=total,
                        step=1,
                        key="ot_hist_pagina",
                    )
                )
            with h3:
                st.button(
                    "Hoja ▶",
                    use_container_width=True,
                    disabled=int(st.session_state.get("ot_hist_pagina") or 1) >= total,
                    key="ot_hist_hoja_next",
                    on_click=_cambiar_hoja_ot,
                    args=(1, total),
                )
            st.caption(f"Hoja {pagina_actual} de {total}")
        elif total == 1:
            st.caption("1 hoja")
        if not archivo_pdf:
            st.info(
                "El PDF de este día ya se descargó y se retiró del servidor. "
                "Las actividades, pozos, insumos y estatus siguen en la bitácora."
            )
        elif total > 0:
            try:
                preview, total = jpeg_pagina_reporte_ot(ruta_pdf, mtime, pagina_actual - 1)
            except Exception as error:
                preview = ""
                st.error(f"No se pudo renderizar la hoja: {error}")
        if preview:
            st.markdown(
                f'<div class="ot-pagina mapa-preview"><img src="data:image/jpeg;base64,{preview}" '
                f'alt="OT {html.escape(dia.strftime("%d/%m/%Y"))}"/></div>',
                unsafe_allow_html=True,
            )
        if archivo_pdf:
            try:
                pdf_bytes = bytes_pdf_ot(ruta_pdf, mtime)
            except OSError:
                pdf_bytes = b""
            st.download_button(
                "Descargar PDF del día",
                data=pdf_bytes,
                file_name=reporte["nombre"] if reporte["nombre"].lower().endswith(".pdf") else f"{reporte['nombre']}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"ot_hist_dl_{dia.isoformat()}",
                on_click=_al_descargar_pdf_ot,
                args=(ruta_pdf, dia.isoformat(), reporte["nombre"]),
            )
    with resumen:
        st.markdown("**Cuadrillas de la jornada**")
        render_resumen_cuadrillas_ot_dia(dia, bitacora)


def jpeg_mapa_soledad(ruta: str, mtime: float, escala: float = 1.55) -> str:
    with open(ruta, "rb") as archivo:
        datos = archivo.read()
    paginas = imagenes_desde_pdf(datos, escala=escala)
    if not paginas:
        return ""
    buf = io.BytesIO()
    paginas[0].convert("RGB").save(buf, format="JPEG", quality=82, optimize=True)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _fmt_fecha_mapa(valor) -> str:
    if isinstance(valor, datetime):
        valor = valor.date()
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    return ""


def _programado_cubre_ficha(item: dict, ficha: dict) -> str:
    ubi = str(item.get("ubicacion") or "").strip()
    equipo = str(item.get("equipo") or "").strip()
    if ubi and _texto_coincide_pozo(ubi, ficha):
        return "pozo"
    if equipo and _texto_coincide_pozo(equipo, ficha):
        return "pozo"
    sector = str(ficha.get("sector") or "").strip()
    if sector and ubi:
        a, b = sector.casefold(), ubi.casefold()
        if a == b or a in b or b in a:
            return "zona"
    return ""


def inventario_inteligente_mapa(datos: dict | None = None) -> list[dict]:
    datos = datos or cargar_ubicaciones()
    bitacora = bitacora_con_fechas(cargar_bitacora())
    hoy = date.today()
    mes_clave = clave_mes(hoy.year, hoy.month)
    programados = programados_del_mes(mes_clave)
    zonas = mapa_zonas(datos.get("zonas"))
    inventario = []
    for ficha in datos.get("fichas", {}).values():
        if not isinstance(ficha, dict):
            continue
        nombre = str(ficha.get("nombre") or "").strip()
        if not nombre:
            continue
        zona_id = str(ficha.get("zona_id") or "otros")
        if zona_id not in {z["id"] for z in ZONAS_HISTORIAL_POZOS} and zona_id not in {"instalaciones"}:
            continue
        filas = bitacora_de_pozo(bitacora, ficha) if bitacora is not None and not bitacora.empty else pd.DataFrame()
        if not filas.empty and "_fecha" not in filas.columns:
            filas = bitacora_con_fechas(filas)
        filas_fecha = filas[filas["_fecha"].notna()] if not filas.empty and "_fecha" in filas.columns else filas
        ultima = None
        ultima_eval = ""
        ultima_rev = ""
        ultima_act = ""
        if filas_fecha is not None and not filas_fecha.empty:
            orden = filas_fecha.sort_values("_fecha")
            ultima_fila = orden.iloc[-1]
            ultima = ultima_fila["_fecha"]
            ultima_eval = str(ultima_fila.get("EVALUACIÓN DEL ESTATUS") or "").strip()
            ultima_rev = str(ultima_fila.get("REVISIÓN") or "").strip()
            ultima_act = str(ultima_fila.get("ACTIVIDAD CORTA") or ultima_fila.get("TIPO DE ACTIVIDAD") or "").strip()
        ots_hoy = int((filas_fecha["_fecha"] == hoy).sum()) if filas_fecha is not None and not filas_fecha.empty else 0
        ots_mes = 0
        if filas_fecha is not None and not filas_fecha.empty:
            ots_mes = int(
                filas_fecha["_fecha"].map(lambda d: isinstance(d, date) and d.year == hoy.year and d.month == hoy.month).sum()
            )
        cerradas_mes = 0
        if filas_fecha is not None and not filas_fecha.empty and "REVISIÓN" in filas_fecha.columns:
            en_mes = filas_fecha[
                filas_fecha["_fecha"].map(lambda d: isinstance(d, date) and d.year == hoy.year and d.month == hoy.month)
            ]
            cerradas_mes = int((en_mes["REVISIÓN"].astype(str) == "Validado").sum())

        preventivos = []
        cobertura = ""
        for item in programados:
            tipo = str(item.get("actividad") or "").strip()
            if tipo and tipo not in TIPOS_MANTENIMIENTO_PROGRAMADO:
                continue
            hit = _programado_cubre_ficha(item, ficha)
            if not hit:
                continue
            preventivos.append(item)
            if hit == "pozo" or not cobertura:
                cobertura = hit
        pendientes = [p for p in preventivos if str(p.get("estatus") or "") != "Cumplido"]

        estatus_ficha = str(ficha.get("estatus") or "Operativo").strip() or "Operativo"
        estatus_vivo = estatus_ficha
        if ultima and isinstance(ultima, date) and (hoy - ultima).days <= 10 and ultima_eval:
            estatus_vivo = ultima_eval
        marca_eq = clasificar_estatus_equipo(estatus_vivo)
        if marca_eq == "paro":
            semaforo = "paro"
            motivo = "Pozo fuera de operación"
        elif marca_eq == "mixto" or pendientes and cobertura == "pozo" or ultima_rev == "Con observaciones":
            semaforo = "mixto"
            if marca_eq == "mixto":
                motivo = "En mantenimiento"
            elif ultima_rev == "Con observaciones":
                motivo = "OT con observaciones / retraso"
            else:
                motivo = "Preventivo pendiente en el pozo"
        else:
            semaforo = "ok"
            motivo = "Operativo"

        inventario.append(
            {
                "id": str(ficha.get("id") or clave_ubicacion(nombre)),
                "nombre": nombre,
                "zona_id": zona_id,
                "zona": titulo_zona_historial(zona_id, zonas.get(zona_id, {})),
                "sector": str(ficha.get("sector") or ""),
                "equipo": str(ficha.get("equipo_mci") or ficha.get("equipo_ub") or "Sin equipo"),
                "estatus_ficha": estatus_ficha,
                "estatus_vivo": estatus_vivo,
                "semaforo": semaforo,
                "motivo": motivo,
                "ultima_ot": ultima if isinstance(ultima, date) else None,
                "ultima_act": ultima_act,
                "ultima_rev": ultima_rev,
                "ots_hoy": ots_hoy,
                "ots_mes": ots_mes,
                "cerradas_mes": cerradas_mes,
                "preventivos": preventivos,
                "pendientes": pendientes,
                "cobertura_prev": cobertura,
            }
        )
    orden_zona = {z["id"]: i for i, z in enumerate(ZONAS_HISTORIAL_POZOS)}
    inventario.sort(key=lambda p: (orden_zona.get(p["zona_id"], 99), orden_natural_pozo(p["nombre"])))
    return inventario


def filtrar_inventario_mapa(
    inventario: list[dict],
    modo: str,
    zona_id: str | None = None,
    recorte: str = "Hoy",
) -> list[dict]:
    vista = list(inventario)
    if zona_id and zona_id != "todas":
        vista = [p for p in vista if p["zona_id"] == zona_id]
    if modo.startswith("⚠️"):
        return [p for p in vista if p["pendientes"] or clasificar_estatus_equipo(p["estatus_vivo"]) == "mixto"]
    if modo.startswith("🔴"):
        return [p for p in vista if p["semaforo"] == "paro"]
    if modo.startswith("📅"):
        if recorte.startswith("Hoy"):
            return [p for p in vista if p["ots_hoy"] > 0]
        return [p for p in vista if p["ots_mes"] > 0]
    return vista


def html_esquema_mapa(items: list[dict]) -> str:
    grupos: dict[str, list[dict]] = {}
    for item in items:
        grupos.setdefault(item["zona_id"], []).append(item)
    bloques = []
    for zona in ZONAS_HISTORIAL_POZOS:
        filas = grupos.get(zona["id"]) or []
        if not filas:
            chips = '<span class="mapa-chip vacio">Sin pozos en el recorte</span>'
        else:
            chips = "".join(
                f'<span class="mapa-chip {html.escape(p["semaforo"])}" title="{html.escape(p["motivo"])}">'
                f'{html.escape(p["nombre"])}</span>'
                for p in filas
            )
        bloques.append(
            f'<div class="mapa-zona"><div class="mapa-zona-t">{zona["icono"]} {html.escape(zona["titulo"])}'
            f' · {len(filas)}</div><div class="mapa-chips">{chips}</div></div>'
        )
    extras = grupos.get("instalaciones") or []
    if extras:
        chips = "".join(
            f'<span class="mapa-chip {html.escape(p["semaforo"])}">{html.escape(p["nombre"])}</span>'
            for p in extras
        )
        bloques.append(
            f'<div class="mapa-zona"><div class="mapa-zona-t">🏭 Instalaciones · {len(extras)}</div>'
            f'<div class="mapa-chips">{chips}</div></div>'
        )
    return f'<div class="mapa-esquema">{"".join(bloques)}</div>'


def html_leyenda_mapa() -> str:
    return (
        '<div class="mapa-leyenda">'
        '<span><span class="semaforo ok"></span>Operativo</span>'
        '<span><span class="semaforo mixto"></span>Preventivo / en mantenimiento</span>'
        '<span><span class="semaforo paro"></span>Fuera de operación</span>'
        "</div>"
    )


def render_tarjeta_estado_mapa(item: dict) -> None:
    detalle = []
    if item.get("ultima_ot"):
        detalle.append(f"Última OT {_fmt_fecha_mapa(item['ultima_ot'])}")
    if item.get("ultima_act"):
        detalle.append(item["ultima_act"])
    if item.get("ots_mes"):
        detalle.append(f"{item['ots_mes']} OT en el mes")
    pendientes = item.get("pendientes") or []
    if pendientes:
        act = str(pendientes[0].get("actividad") or "Preventivo")
        est = str(pendientes[0].get("estatus") or "")
        detalle.append(f"{act} · {est}")
    st.markdown(
        f"""
        <div class="mapa-pozo {html.escape(item['semaforo'])}">
            <div class="prefijo"><span class="semaforo {html.escape(item['semaforo'])}"></span>
            {html.escape(item['zona'])}</div>
            <h3>{html.escape(item['nombre'])}</h3>
            <p>{html.escape(item['motivo'])}</p>
            <p class="s">{html.escape(" · ".join(detalle) if detalle else item.get("equipo") or "")}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Abrir ficha", key=f"mapa_ficha_{item['id']}", use_container_width=True):
        st.session_state.ub_sub_modulo = "fichas"
        st.session_state.ub_zona_explorador = item["zona_id"]
        st.session_state.ub_ficha_id = item["id"]
        st.rerun()


def render_kpis_mapa(modo: str, inventario: list[dict], filtrados: list[dict], recorte: str) -> None:
    total = len(inventario)
    operativos = sum(1 for p in inventario if p["semaforo"] == "ok")
    alertas = sum(1 for p in inventario if p["semaforo"] == "mixto")
    paros = sum(1 for p in inventario if p["semaforo"] == "paro")
    k1, k2, k3, k4 = st.columns(4, gap="medium")
    if modo.startswith("🗺️"):
        k1.markdown(kpi_html("Pozos en mapa", f"{total}", "Inventario de campo e instalaciones", "info"), unsafe_allow_html=True)
        k2.markdown(kpi_html("Operativos", f"{operativos}", "Semáforo verde", "ok"), unsafe_allow_html=True)
        k3.markdown(kpi_html("En alerta", f"{alertas}", "Preventivo o en mantenimiento", "warn"), unsafe_allow_html=True)
        k4.markdown(kpi_html("Fuera de operación", f"{paros}", "Requieren restablecimiento", "danger" if paros else "ok"), unsafe_allow_html=True)
        return
    if modo.startswith("⚠️"):
        planes = []
        vistos = set()
        for p in inventario:
            for prev in p.get("pendientes") or []:
                pid = prev.get("id") or id(prev)
                if pid in vistos:
                    continue
                vistos.add(pid)
                planes.append(prev)
        reprog = sum(1 for x in planes if str(x.get("estatus")) == "Reprogramado")
        cumplidos = 0
        for p in inventario:
            for prev in p.get("preventivos") or []:
                pid = prev.get("id") or id(prev)
                if pid in vistos:
                    continue
                if str(prev.get("estatus")) == "Cumplido":
                    cumplidos += 1
                    vistos.add(pid)
        k1.markdown(kpi_html("Pozos a intervenir", f"{len(filtrados)}", "Preventivo pendiente o en mantenimiento", "warn"), unsafe_allow_html=True)
        k2.markdown(kpi_html("Planes OPS abiertos", f"{len(planes)}", "No cumplido / reprogramado", "warn"), unsafe_allow_html=True)
        k3.markdown(kpi_html("Reprogramados", f"{reprog}", "Ajustes al plan del mes", "info"), unsafe_allow_html=True)
        k4.markdown(kpi_html("Cumplidos del mes", f"{cumplidos}", "Preventivos ya cerrados", "ok"), unsafe_allow_html=True)
        return
    if modo.startswith("🔴"):
        pct = round(100.0 * paros / max(total, 1), 0)
        ultimas = [p["ultima_ot"] for p in filtrados if p.get("ultima_ot")]
        k1.markdown(kpi_html("Pozos fuera", f"{paros}", f"{pct:.0f}% del inventario", "danger" if paros else "ok"), unsafe_allow_html=True)
        k2.markdown(kpi_html("En el recorte", f"{len(filtrados)}", "Filtrado actual", "danger" if filtrados else "ok"), unsafe_allow_html=True)
        k3.markdown(kpi_html("Último reporte", _fmt_fecha_mapa(max(ultimas)) if ultimas else "—", "Bitácora más reciente", "info"), unsafe_allow_html=True)
        k4.markdown(kpi_html("Operativos", f"{operativos}", "Continúan en servicio", "ok"), unsafe_allow_html=True)
        return
    ots = sum(p["ots_hoy"] if recorte.startswith("Hoy") else p["ots_mes"] for p in inventario)
    cerradas = sum(p["cerradas_mes"] for p in inventario) if not recorte.startswith("Hoy") else sum(
        1 for p in inventario if p["ots_hoy"] and p.get("ultima_rev") == "Validado"
    )
    k1.markdown(
        kpi_html("OT del recorte", f"{ots}", "Hoy" if recorte.startswith("Hoy") else "Mes en curso", "info"),
        unsafe_allow_html=True,
    )
    k2.markdown(kpi_html("Pozos con OT", f"{len(filtrados)}", "Ubicaciones con actividad", "warn" if filtrados else "ok"), unsafe_allow_html=True)
    k3.markdown(kpi_html("Cerradas", f"{cerradas}", "Validado en bitácora", "ok"), unsafe_allow_html=True)
    k4.markdown(
        kpi_html("Pendientes", f"{max(ots - cerradas, 0)}", "En proceso o con retraso", "warn" if ots > cerradas else "ok"),
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def jpeg_mapa_soledad_cached(ruta: str, mtime: float, escala: float = 1.55) -> str:
    return jpeg_mapa_soledad(ruta, mtime, escala)


def render_visor_mapa_soledad() -> None:
    st.markdown('<div class="section-title">Plano general inteligente</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="mapa-shell">
            <h3>Mapa del Área Contractual Soledad</h3>
            <p>Cambia el modo de visualización para cruzar el plano con el inventario, las OT y el plan preventivo OPS.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    modo = st.radio(
        "Modo de visualización",
        MODOS_MAPA,
        horizontal=True,
        key="mapa_modo_visual",
    )
    inventario = inventario_inteligente_mapa()
    zonas_opts = ["Todas las zonas"] + [z["titulo"] for z in ZONAS_HISTORIAL_POZOS]
    recorte = "Hoy"
    zona_id = "todas"
    if modo.startswith("📅"):
        hay_hoy = any(p["ots_hoy"] > 0 for p in inventario)
        recorte = st.radio(
            "Recorte de OT",
            ["Hoy", "Histórico (mes actual)"],
            index=0 if hay_hoy else 1,
            horizontal=True,
            key="mapa_recorte_ot",
        )
    zona_sel = st.selectbox("Zona", zonas_opts, key="mapa_zona_filtro")
    zona_id = next((z["id"] for z in ZONAS_HISTORIAL_POZOS if z["titulo"] == zona_sel), "todas")

    filtrados = filtrar_inventario_mapa(inventario, modo, zona_id, recorte)
    render_kpis_mapa(modo, inventario, filtrados, recorte)
    st.markdown(html_leyenda_mapa(), unsafe_allow_html=True)

    ruta = asegurar_mapa_soledad()
    col_mapa, col_esq = st.columns([1.15, 1], gap="large")
    with col_mapa:
        if ruta:
            try:
                mtime = os.path.getmtime(ruta)
                b64_img = jpeg_mapa_soledad_cached(ruta, mtime)
            except Exception:
                b64_img = ""
            if b64_img:
                st.markdown(
                    f'<div class="mapa-preview"><img src="data:image/jpeg;base64,{b64_img}" '
                    f'alt="Mapa del Área Contractual Soledad"/></div>',
                    unsafe_allow_html=True,
                )
            else:
                st.info("No se pudo renderizar el plano. Usa la descarga del PDF.")
        else:
            st.warning(f"Coloca `{MAPA_SOLEDAD_NOMBRE}` en la carpeta `{MAPA_DIR}/`.")
    with col_esq:
        st.markdown("**Esquema por zona**")
        st.markdown(html_esquema_mapa(filtrados), unsafe_allow_html=True)

    st.markdown("**Pozos del recorte**")
    if not filtrados:
        st.info("No hay pozos que coincidan con este modo. Cambia el filtro o carga bitácora / programados OPS.")
    else:
        for inicio in range(0, min(len(filtrados), 18), 3):
            cols = st.columns(3, gap="medium")
            for col, item in zip(cols, filtrados[inicio : inicio + 3]):
                with col:
                    render_tarjeta_estado_mapa(item)
        if len(filtrados) > 18:
            st.caption(f"Mostrando 18 de {len(filtrados)} pozos. Ajusta la zona para ver el resto.")

    if ruta:
        with st.expander("Descargar plano original (PDF)"):
            try:
                with open(ruta, "rb") as archivo:
                    pdf_bytes = archivo.read()
            except OSError:
                pdf_bytes = b""
            if pdf_bytes:
                st.download_button(
                    "Descargar PDF",
                    data=pdf_bytes,
                    file_name=MAPA_SOLEDAD_NOMBRE,
                    mime="application/pdf",
                    use_container_width=True,
                    key="mapa_soledad_descarga",
                )


def render_panel_ubicaciones(*, encabezado: bool = True) -> None:
    if encabezado:
        render_encabezado(
            "📍 Ubicación y Equipos",
            "Historial de pozos, plano general del área contractual y ficha de equipos",
        )
    if migracion_excel_pendiente() and "ub_excel_autohecho" not in st.session_state:
        auto = importar_excel_ubicaciones(forzar=True)
        st.session_state.ub_excel_autohecho = True
        if auto.get("ok") and not auto.get("omitido"):
            st.success(
                f"Inventario Excel importado: {auto['nuevos']} pozos nuevos y "
                f"{auto['actualizados']} actualizados en {UBICACIONES_FILE}."
            )
        elif not auto.get("ok"):
            st.warning(auto.get("error") or "No se pudo importar el Excel de ubicaciones.")

    datos = catalogo_con_historial()

    sub = st.session_state.get("ub_sub_modulo")
    if not sub:
        _reset_explorador_pozos()
        render_cuadricula_explorador(
            SUBMENUS_UBICACIONES,
            "ub_sub_modulo",
            "Explora el plano general, las zonas de pozos o importa el inventario desde Excel.",
        )
        return
    render_barra_volver(
        f'Ubicación y Equipos / <b>{html.escape(titulo_submenu(SUBMENUS_UBICACIONES, sub))}</b>',
        "ub_sub_modulo",
        "← Volver a Ubicación y Equipos",
    )
    if sub == "importar":
        render_migracion_ubicaciones()
        return
    if sub == "plano":
        render_visor_mapa_soledad()
        return

    st.caption(f"Inventario de ubicaciones en {UBICACIONES_FILE}, cruzado con bitácora y programados OPS.")
    render_flujo_historial_pozos(datos)


def _mes_historial_cuadrillas() -> tuple[int, int]:
    hoy = date.today()
    return (
        int(st.session_state.get("hc_cal_anio") or hoy.year),
        int(st.session_state.get("hc_cal_mes") or hoy.month),
    )


def _desplazar_mes_cuadrillas(delta: int) -> None:
    anio, mes = _mes_historial_cuadrillas()
    mes += delta
    while mes < 1:
        mes += 12
        anio -= 1
    while mes > 12:
        mes -= 12
        anio += 1
    st.session_state.hc_cal_anio = anio
    st.session_state.hc_cal_mes = mes
    dia = parsear_fecha_valor(st.session_state.get("hc_dia"))
    if dia and (dia.year != anio or dia.month != mes):
        st.session_state.hc_dia = None
        st.session_state.hc_cuadrilla = None


def _reset_historial_cuadrillas() -> None:
    st.session_state.hc_dia = None
    st.session_state.hc_cuadrilla = None


def render_galeria_reporte(evidencias: list[dict]) -> None:
    celdas = list(evidencias) + [{"ruta": "", "descripcion": ""}] * MAX_FOTOS
    for fila in range(3):
        cols = st.columns(3, gap="small")
        for col_i, col in enumerate(cols):
            item = celdas[fila * 3 + col_i]
            with col:
                ruta = str(item.get("ruta") or "").strip()
                desc = str(item.get("descripcion") or "").strip()
                if ruta and os.path.exists(ruta):
                    st.image(ruta, use_container_width=True)
                    st.markdown(
                        f'<p class="foto-cap">{html.escape(desc or os.path.basename(ruta))}</p>',
                        unsafe_allow_html=True,
                    )
                elif ruta:
                    st.markdown(
                        f'<div class="foto-vacia dark">No encontrada<br>{html.escape(os.path.basename(ruta))}</div>',
                        unsafe_allow_html=True,
                    )
                    if desc:
                        st.markdown(f'<p class="foto-cap">{html.escape(desc)}</p>', unsafe_allow_html=True)
                else:
                    st.markdown(
                        f'<div class="foto-vacia dark">Espacio fotográfico {fila * 3 + col_i + 1}</div>',
                        unsafe_allow_html=True,
                    )


def clave_hist_cuadrilla(nombre: str) -> str:
    return re.sub(r"\s+", " ", _sin_acentos(nombre).casefold()).strip()


def clave_cuadrilla_oficial(nombre: str) -> str:
    clave = clave_hist_cuadrilla(nombre)
    if clave in {"electrico", "electricos"}:
        return "electricos"
    if clave in {"retroexcavadora", "retro excavadora"}:
        return "retro excavadora"
    return clave


def orden_cuadrilla_oficial(nombre: str) -> tuple[int, str]:
    clave = clave_cuadrilla_oficial(nombre)
    rango = {clave_cuadrilla_oficial(item): indice for indice, item in enumerate(ORDEN_CUADRILLAS_OFICIAL)}
    return (rango.get(clave, 99), etiqueta_cuadrilla_historial(nombre).casefold())


def etiqueta_cuadrilla_historial(nombre: str) -> str:
    mapa = {
        "electricos": "Eléctricos",
        "electrico": "Eléctricos",
        "instrumentos": "Instrumentos",
        "mecanico 1": "Mecánico 1",
        "mecanico 2": "Mecánico 2",
        "mecanico 3": "Mecánico 3",
        "mecanico 4": "Mecánico 4",
        "montaje": "Montaje",
        "soldadura": "Soldadura",
        "retro excavadora": "Retro Excavadora",
        "retroexcavadora": "Retro Excavadora",
    }
    return mapa.get(clave_cuadrilla_oficial(nombre), str(nombre or "").strip() or "Sin cuadrilla")


def bloque_historial_cuadrilla(df: pd.DataFrame, cuadrilla: str) -> pd.DataFrame:
    if df is None or df.empty or "CUADRILLA" not in df.columns:
        return pd.DataFrame()
    objetivo = clave_cuadrilla_oficial(cuadrilla)
    return df.loc[df["CUADRILLA"].astype(str).map(lambda v: clave_cuadrilla_oficial(v) == objetivo)].copy()


def fechas_historial_cuadrilla(df: pd.DataFrame, cuadrilla: str) -> list[date]:
    base = df if df is not None and "_fecha" in getattr(df, "columns", []) else bitacora_con_fechas(df)
    bloque = bloque_historial_cuadrilla(base, cuadrilla)
    if bloque.empty or "_fecha" not in bloque.columns:
        return []
    return sorted({d for d in bloque["_fecha"] if isinstance(d, date)}, reverse=True)


def catalogo_cuadrillas_historial(df: pd.DataFrame) -> list[str]:
    vistos = {clave_cuadrilla_oficial(n): etiqueta_cuadrilla_historial(n) for n in CUADRILLAS}
    if df is not None and not df.empty and "CUADRILLA" in df.columns:
        for nombre in df["CUADRILLA"].astype(str):
            texto = str(nombre or "").strip()
            clave = clave_cuadrilla_oficial(texto)
            if texto and clave not in vistos:
                vistos[clave] = etiqueta_cuadrilla_historial(texto)
    return sorted(vistos.values(), key=orden_cuadrilla_oficial)


def nombre_carpeta_cuadrilla(cuadrilla: str) -> str:
    texto = etiqueta_cuadrilla_historial(cuadrilla)
    limpio = re.sub(r'[<>:"/\\|?*]+', " ", texto).strip().rstrip(".")
    return re.sub(r"\s+", " ", limpio) or "Sin cuadrilla"


def evidencia_fotografica_default() -> dict:
    return {"registros": {}}


def cargar_evidencia_fotografica() -> dict:
    datos = leer_json(EVIDENCIA_FOTOGRAFICA_FILE, evidencia_fotografica_default())
    if not isinstance(datos, dict) or not isinstance(datos.get("registros"), dict):
        return evidencia_fotografica_default()
    return datos


def guardar_evidencia_fotografica(datos: dict) -> None:
    escribir_json(EVIDENCIA_FOTOGRAFICA_FILE, datos)


def clave_evidencia_cuadrilla(cuadrilla: str, dia: date) -> str:
    return f"{clave_hist_cuadrilla(cuadrilla)}|{dia.isoformat()}"


def carpeta_evidencia_cuadrilla_dia(cuadrilla: str, dia: date) -> str:
    return os.path.join(EVIDENCIA_FOTOGRAFICA_DIR, nombre_carpeta_cuadrilla(cuadrilla), f"{dia.day:02d}")


def numero_secuencia_actividad(valor) -> int | None:
    texto = texto_secuencia_limpio(valor)
    digitos = re.search(r"\d+", texto)
    if not digitos:
        return None
    return int(digitos.group(0))


def carpeta_cuadrilla_octubre(cuadrilla: str) -> str:
    clave = clave_hist_cuadrilla(cuadrilla)
    if os.path.isdir(EVIDENCIA_OCTUBRE_DIR):
        for nombre in os.listdir(EVIDENCIA_OCTUBRE_DIR):
            ruta = os.path.join(EVIDENCIA_OCTUBRE_DIR, nombre)
            if os.path.isdir(ruta) and clave_hist_cuadrilla(nombre) == clave:
                return nombre
    return _sin_acentos(str(cuadrilla or "")).casefold().strip() or "sin cuadrilla"


def carpeta_evidencia_octubre_secuencia(cuadrilla: str, dia: date, secuencia) -> str:
    numero = numero_secuencia_actividad(secuencia)
    secuencia_txt = f"{numero:02d}" if numero is not None else "sin-secuencia"
    return os.path.join(
        EVIDENCIA_OCTUBRE_DIR,
        carpeta_cuadrilla_octubre(cuadrilla),
        dia.strftime("%d-%m-%Y"),
        secuencia_txt,
    )


def fotos_evidencia_octubre_secuencia(cuadrilla: str, dia: date, secuencia) -> list[str]:
    base = carpeta_evidencia_octubre_secuencia(cuadrilla, dia, secuencia)
    candidatos = [base]
    numero = numero_secuencia_actividad(secuencia)
    if numero is not None:
        alterna = os.path.join(os.path.dirname(base), str(numero))
        if alterna not in candidatos:
            candidatos.append(alterna)
    fotos = []
    for carpeta in candidatos:
        if not os.path.isdir(carpeta):
            continue
        for nombre in sorted(os.listdir(carpeta)):
            ruta = os.path.join(carpeta, nombre)
            if not os.path.isfile(ruta):
                continue
            if os.path.splitext(nombre)[1].lower() not in {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}:
                continue
            fotos.append(ruta.replace("\\", "/"))
        if fotos:
            break
    return fotos


def render_bloque_fotos_secuencia(cuadrilla: str, dia: date, secuencia) -> bool:
    carpeta = carpeta_evidencia_octubre_secuencia(cuadrilla, dia, secuencia).replace("\\", "/")
    fotos = fotos_evidencia_octubre_secuencia(cuadrilla, dia, secuencia)
    st.markdown("##### Evidencia fotográfica de esta secuencia")
    st.caption(f"`{carpeta}/`")
    if not fotos:
        st.info("Esta secuencia no tiene fotografías en su carpeta.")
        return False
    st.success(f"{len(fotos)} fotografía(s) de esta secuencia.")
    for inicio in range(0, len(fotos), 3):
        columnas = st.columns(3, gap="small")
        for col, ruta in zip(columnas, fotos[inicio : inicio + 3]):
            with col:
                st.image(ruta, use_container_width=True)
                st.caption(os.path.basename(ruta))
    return True


def fechas_fotos_cuadrilla(cuadrilla: str) -> set[date]:
    prefijo = clave_hist_cuadrilla(cuadrilla) + "|"
    dias: set[date] = set()
    for clave, item in cargar_evidencia_fotografica()["registros"].items():
        if not str(clave).startswith(prefijo) or not isinstance(item, dict):
            continue
        if not [ruta for ruta in (item.get("fotos") or []) if ruta]:
            continue
        fecha = parsear_fecha_valor(item.get("fecha"))
        if fecha:
            dias.add(fecha)
    carpeta = carpeta_cuadrilla_octubre(cuadrilla)
    if os.path.isdir(carpeta):
        for nombre in os.listdir(carpeta):
            dia = parsear_fecha_valor(nombre)
            if dia:
                dias.add(dia)
    return dias


def fotos_evidencia_cuadrilla(cuadrilla: str, dia: date) -> list[str]:
    item = cargar_evidencia_fotografica()["registros"].get(clave_evidencia_cuadrilla(cuadrilla, dia))
    if not isinstance(item, dict):
        return []
    return [ruta for ruta in (item.get("fotos") or []) if ruta and os.path.isfile(ruta)]


def guardar_lote_evidencia_cuadrilla(cuadrilla: str, dia: date, archivos: list) -> dict:
    fotos_nuevas = limitar_fotos(archivos)
    if not fotos_nuevas:
        return {"ok": False, "agregadas": 0, "mensaje": "Selecciona al menos una imagen."}
    carpeta = carpeta_evidencia_cuadrilla_dia(cuadrilla, dia)
    os.makedirs(carpeta, exist_ok=True)
    datos = cargar_evidencia_fotografica()
    clave = clave_evidencia_cuadrilla(cuadrilla, dia)
    actual = datos["registros"].get(clave)
    rutas = []
    if isinstance(actual, dict):
        rutas = [ruta for ruta in (actual.get("fotos") or []) if ruta and os.path.isfile(ruta)]
    disponibles = max(0, MAX_FOTOS - len(rutas))
    if disponibles <= 0:
        return {
            "ok": False,
            "agregadas": 0,
            "mensaje": "Este día ya tiene las 9 fotografías. No se pueden agregar más.",
        }
    tomadas = fotos_nuevas[:disponibles]
    sello = dia.strftime("%Y%m%d")
    for archivo in tomadas:
        extension = os.path.splitext(getattr(archivo, "name", "") or "")[1].lower() or ".jpg"
        if extension not in {".jpg", ".jpeg", ".png", ".webp"}:
            extension = ".jpg"
        destino = os.path.join(carpeta, f"{sello}_foto_{len(rutas) + 1:02d}{extension}")
        with open(destino, "wb") as salida:
            salida.write(archivo.getvalue())
        rutas.append(destino.replace("\\", "/"))
    datos["registros"][clave] = {
        "cuadrilla": etiqueta_cuadrilla_historial(cuadrilla),
        "fecha": dia.isoformat(),
        "dia": f"{dia.day:02d}",
        "carpeta": carpeta.replace("\\", "/"),
        "fotos": rutas,
        "actualizado": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    guardar_evidencia_fotografica(datos)
    return {
        "ok": True,
        "agregadas": len(tomadas),
        "total": len(rutas),
        "omitidas": max(0, len(fotos_nuevas) - len(tomadas)),
        "carpeta": carpeta.replace("\\", "/"),
        "mensaje": "",
    }


def slots_evidencia_cuadrilla(cuadrilla: str, dia: date) -> list[str]:
    item = cargar_evidencia_fotografica()["registros"].get(clave_evidencia_cuadrilla(cuadrilla, dia))
    if not isinstance(item, dict) or item.get("origen") != "baul":
        return []
    fotos = [str(ruta or "").strip() for ruta in (item.get("fotos") or [])][:MAX_FOTOS]
    while len(fotos) < MAX_FOTOS:
        fotos.append("")
    return [ruta if ruta and os.path.isfile(ruta) else "" for ruta in fotos]


def slots_fotos_fila(fila) -> list[str]:
    celdas = []
    for col in COLUMNAS_FOTOS:
        ruta = str(fila.get(col) or "").strip()
        if not ruta or ruta.lower() in {"nan", "none"}:
            celdas.append("")
        else:
            celdas.append(ruta)
    return celdas


def resumen_fechas_cuadrilla(df: pd.DataFrame, cuadrilla: str) -> list[dict]:
    fechas = set(fechas_historial_cuadrilla(df, cuadrilla))
    fechas.update(fechas_fotos_cuadrilla(cuadrilla))
    filas = []
    for dia in sorted(fechas, reverse=True):
        bloque = bloque_historial_cuadrilla(bitacora_del_dia(df, dia), cuadrilla)
        lote = bool(fotos_evidencia_cuadrilla(cuadrilla, dia))
        sin_fotos = 0
        if not bloque.empty and not lote:
            sin_fotos = int(sum(1 for i in bloque.index if estatus_evidencia_fila(bloque.loc[i]) != EVIDENCIA_COMPLETA))
        filas.append(
            {
                "dia": dia,
                "ots": int(len(bloque)),
                "marca": clasificar_bloque_cuadrilla(bloque) if not bloque.empty else "ok",
                "sin_fotos": sin_fotos,
                "lote": lote,
            }
        )
    return filas


def numero_secuencia_ot(valor, indice: int) -> tuple:
    texto = str(valor or "").strip()
    digitos = re.findall(r"\d+", texto)
    if digitos:
        return (0, int(digitos[0]), texto.lower())
    return (1, indice, texto.lower())


def etiqueta_secuencia_historial(fila, indice: int) -> str:
    sec = texto_secuencia_limpio(fila.get("SECUENCIA"))
    digitos = re.search(r"\d+", sec)
    num = f"{int(digitos.group(0)):02d}" if digitos else f"{indice:02d}"
    corto = str(fila.get("ACTIVIDAD CORTA") or fila.get("TIPO DE ACTIVIDAD") or "").strip()
    if not corto or corto.lower() in {"nan", "none", "actividad", "actividad de campo"}:
        descripcion = texto_visible(
            fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD"),
            "",
        )
        corto = (descripcion[:42] + "…") if len(descripcion) > 42 else (descripcion or "Actividad")
    sitio = etiqueta_ubicacion_fila(fila) or texto_visible(
        fila.get("UBICACIÓN") or fila.get("POZO") or fila.get("SECTOR") or fila.get("EQUIPO/UBICACIÓN"),
        "Sin ubicación",
    )
    evid = estatus_evidencia_fila(fila)
    fotos = len(rutas_fotos_fila(fila))
    return f"Secuencia #{num} ({corto} — {sitio} · {evid}, {fotos} foto(s))"


def filas_secuencia_historial(bloque: pd.DataFrame) -> list:
    filas = [bloque.loc[i] for i in bloque.index]
    return [
        fila
        for _, fila in sorted(
            enumerate(filas, start=1),
            key=lambda par: numero_secuencia_ot(par[1].get("SECUENCIA"), par[0]),
        )
    ]


def render_reporte_cuadrilla(dia: date, cuadrilla: str, df_dia: pd.DataFrame) -> None:
    bloque = bloque_historial_cuadrilla(df_dia, cuadrilla)
    if bloque.empty:
        st.info("No hay reportes de esta cuadrilla en la fecha seleccionada.")
        return
    filas = filas_secuencia_historial(bloque)
    marca = clasificar_bloque_cuadrilla(bloque)
    personal = personal_asignado_cuadrilla(dia, cuadrilla)
    opciones = []
    vistos: set[str] = set()
    for i, fila in enumerate(filas, start=1):
        etiqueta = etiqueta_secuencia_historial(fila, i)
        if etiqueta in vistos:
            etiqueta = f"{etiqueta} · {i}"
        vistos.add(etiqueta)
        opciones.append(etiqueta)

    st.success(
        f"Mostrando reporte de **{etiqueta_cuadrilla_historial(cuadrilla)}** para el **{dia.strftime('%d/%m/%Y')}** · "
        f"{'1 secuencia' if len(filas) == 1 else f'{len(filas)} secuencias'} · {etiqueta_semaforo(marca)}."
    )
    clave_sel = f"hc_sec_{clave_hist_cuadrilla(cuadrilla)}_{dia.isoformat()}"
    secuencia_elegida = st.selectbox(
        "🔍 Seleccione la Secuencia de Actividad del día:",
        opciones,
        key=clave_sel,
    )
    indice = opciones.index(secuencia_elegida)
    fila = filas[indice]
    materiales = materiales_de_filas([fila])
    evidencias = fotos_de_filas([fila])
    sec = texto_secuencia_limpio(fila.get("SECUENCIA")) or str(indice + 1)
    ubicacion = etiqueta_ubicacion_fila(fila) or texto_visible(
        fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN") or fila.get("SECTOR") or fila.get("POZO")
    )
    personal_txt = texto_visible(fila.get("PERSONAL"))
    desc = texto_visible(
        fila.get("DESCRIPCION DE LA ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD"),
        "Sin descripción capturada.",
    )
    comentarios = texto_visible(
        fila["COMENTARIOS"] if "COMENTARIOS" in getattr(fila, "index", []) else fila.get("OBSERVACIÓN REVISIÓN")
    )
    avance = avance_texto(fila["AVANCE %"] if "AVANCE %" in getattr(fila, "index", []) else None) or "—"
    estatus = texto_visible(fila.get("EVALUACIÓN DEL ESTATUS"))
    marca_ot = clasificar_estatus_equipo(estatus)
    if marca_ot == "paro":
        paro = "Equipo quedó fuera de operación. Hora de paro no capturada en bitácora."
    elif marca_ot == "mixto":
        paro = "Equipo en mantenimiento / estatus intermedio. Sin horario de paro registrado."
    else:
        paro = "Sin tiempo de paro. El equipo quedó operando."
    if PATRON_NO_PROGRAMADO.search(desc) and marca_ot != "ok":
        paro += " La descripción indica paro imprevisto o correctivo."
    pendiente_dia = personal or "Sin pendiente publicado para esta cuadrilla"
    gente = personal_de_registro(fila)
    titular_ot = texto_par_personal(gente["mecanico_titular"], gente["ayudante_titular"])
    real_ot = texto_par_personal(gente["mecanico_real"], gente["ayudante_real"])
    cobertura_ot = gente["cobertura"] or ("Sí" if real_ot != titular_ot and real_ot != "—" else "No")
    etiqueta_corta = secuencia_elegida.split(" (", 1)[0] if " (" in secuencia_elegida else secuencia_elegida[:13]
    st.divider()

    tab_act, tab_per, tab_ev = st.tabs(
        [
            f"⚙️ Detalle de {etiqueta_corta}",
            "👷 Personal y mano de obra",
            "📷 Evidencia fotográfica",
        ]
    )
    with tab_act:
        st.markdown(f"##### Información de la {secuencia_elegida}")
        st.markdown(
            f"""
            <div class="ot-sheet">
                <div class="ot-kicker">{html.escape(etiqueta_cuadrilla_historial(cuadrilla))} · {html.escape(dia.strftime("%d/%m/%Y"))} · Secuencia {html.escape(sec)}</div>
                <h3>{html.escape(ubicacion)}</h3>
                <div class="ot-grid">
                    <div class="ot-field"><div class="k">Cuadrilla</div>
                    <div class="v">{html.escape(etiqueta_cuadrilla_historial(gente["cuadrilla"] or cuadrilla))}</div></div>
                    <div class="ot-field"><div class="k">Personal</div>
                    <div class="v">{html.escape(personal_txt if personal_txt != "—" else real_ot)}</div></div>
                    <div class="ot-field"><div class="k">Secuencia</div>
                    <div class="v">{html.escape(sec)}</div></div>
                    <div class="ot-field"><div class="k">Ubicación</div>
                    <div class="v">{html.escape(ubicacion)}</div></div>
                    <div class="ot-field"><div class="k">Avance %</div>
                    <div class="v">{html.escape(avance)}</div></div>
                    <div class="ot-field"><div class="k">Fecha</div>
                    <div class="v">{html.escape(dia.strftime("%d/%m/%Y"))}</div></div>
                </div>
                <div class="ot-field"><div class="k">Descripción de la actividad</div>
                <div class="v">{html.escape(desc)}</div></div>
                <div class="ot-field"><div class="k">Comentarios de campo</div>
                <div class="v">{html.escape(comentarios)}</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if estatus != "—":
            st.caption(f"Estatus del equipo: {estatus}. {paro}")
        st.markdown('<div class="section-title">Insumos suministrados</div>', unsafe_allow_html=True)
        if not materiales:
            st.info("Esta secuencia no reportó materiales ni refacciones.")
        else:
            for item in materiales:
                st.markdown(
                    f'<div class="mat-row"><b>{html.escape(item["nombre"])}</b>'
                    f'<span>{html.escape(item["detalle"] or "Reportado")}</span></div>',
                    unsafe_allow_html=True,
                )
        render_bloque_fotos_secuencia(str(fila.get("CUADRILLA") or cuadrilla), dia, fila.get("SECUENCIA"))

    with tab_per:
        st.markdown("##### Personal encargado en esta secuencia")
        st.caption("Registro de turno interno. Sustituye la firma física tradicional para control de campo.")
        if personal:
            st.info(f"Pendiente publicado: {personal}")
        hay_gente = any(
            [
                gente["mecanico_titular"],
                gente["ayudante_titular"],
                gente["mecanico_real"],
                gente["ayudante_real"],
            ]
        )
        if hay_gente:
            st.markdown(
                f"""
                <div class="ot-sheet">
                    <div class="ot-kicker">{html.escape(gente["cuadrilla"] or cuadrilla)}</div>
                    <div class="ot-grid">
                        <div class="ot-field"><div class="k">Personal de la cuadrilla</div>
                        <div class="v">{html.escape(personal_txt)}</div></div>
                        <div class="ot-field"><div class="k">Personal titular</div>
                        <div class="v">{html.escape(titular_ot)}</div></div>
                        <div class="ot-field"><div class="k">Personal en turno (real)</div>
                        <div class="v">{html.escape(real_ot)}</div></div>
                        <div class="ot-field"><div class="k">Cobertura</div>
                        <div class="v">{html.escape(cobertura_ot)}</div></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("No hay nombres de turno capturados en esta secuencia.")
        render_nota_asistencia(
            gente["cuadrilla"] or cuadrilla,
            dia,
            str(fila.get("PERSONAL") or ""),
        )

    with tab_ev:
        st.markdown("##### Galería de evidencia fotográfica de la actividad")
        st.caption("Fotos de esta secuencia, leídas desde evidencia_fotografica_octubre.")
        encontro_octubre = render_bloque_fotos_secuencia(
            str(fila.get("CUADRILLA") or cuadrilla), dia, fila.get("SECUENCIA")
        )
        lote_slots = [] if encontro_octubre else slots_evidencia_cuadrilla(cuadrilla, dia)
        lote_dia = [] if encontro_octubre else fotos_evidencia_cuadrilla(cuadrilla, dia)
        if encontro_octubre:
            pass
        elif any(lote_slots):
            st.success(
                f"La cuadrilla ya tiene {len(lote_dia)} foto(s) del {dia.strftime('%d/%m/%Y')} "
                f"en `{carpeta_evidencia_cuadrilla_dia(cuadrilla, dia).replace(chr(92), '/')}`."
            )
            render_galeria_reporte(
                [{"ruta": ruta, "descripcion": f"Foto {i}"} for i, ruta in enumerate(lote_slots, start=1)]
            )
        elif lote_dia:
            st.success(
                f"La cuadrilla ya tiene {len(lote_dia)} foto(s) del {dia.strftime('%d/%m/%Y')} "
                f"en `{carpeta_evidencia_cuadrilla_dia(cuadrilla, dia).replace(chr(92), '/')}`."
            )
            render_galeria_reporte(
                [{"ruta": ruta, "descripcion": os.path.basename(ruta)} for ruta in lote_dia]
            )
        if not encontro_octubre:
            slots_seq = slots_fotos_fila(fila)
            if any(slots_seq):
                st.markdown("##### Fotos de esta secuencia")
                render_galeria_reporte(
                    [{"ruta": ruta, "descripcion": f"Foto {i}"} for i, ruta in enumerate(slots_seq, start=1)]
                )
            elif evidencias:
                st.markdown("##### Fotos de esta secuencia")
                render_galeria_reporte(evidencias)
            elif not lote_dia and not any(lote_slots):
                st.warning("Aún no hay fotografías cargadas para esta secuencia.")
        df_base = cargar_bitacora()
        idx_reg = fila.name
        if idx_reg in df_base.index:
            st.markdown("##### Completar o actualizar evidencia")
            render_panel_evidencia_posterior(df_base, idx_reg, clave=f"hc_{id_registro_fila(fila)}")
        st.markdown(
            """
            <div class="firma-row">
                <div class="firma-box">
                    <div class="k">Firma PJP4</div>
                    <div class="firma-linea">Responsable de cuadrilla / supervisor PJP4</div>
                </div>
                <div class="firma-box">
                    <div class="k">Firma cliente OPS</div>
                    <div class="firma-linea">Vo.Bo. representante OPS</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_listado_cuadrillas_dia(dia: date, df_dia: pd.DataFrame) -> None:
    grupos = resumen_cuadrillas_dia(df_dia)
    st.markdown('<div class="section-title">Cuadrillas del día</div>', unsafe_allow_html=True)
    if not grupos:
        st.info("No hay reportes de campo capturados en esta fecha.")
        return
    st.caption(
        f"{len(grupos)} cuadrilla(s) con bitácora el {dia.strftime('%d/%m/%Y')}. "
        "Verde: equipo operando · Rojo: fuera de operación."
    )
    for inicio in range(0, len(grupos), 3):
        cols = st.columns(3, gap="medium")
        for col, grupo in zip(cols, grupos[inicio : inicio + 3]):
            with col:
                sitios = ", ".join(grupo["pozos"][:3] or grupo["sectores"][:2]) or "Sin pozo"
                extra = f" +{len(grupo['pozos']) - 3}" if len(grupo["pozos"]) > 3 else ""
                st.markdown(
                    f"""
                    <div class="cuadrilla-dia {grupo['marca']}">
                        <div class="lbl"><span class="semaforo {grupo['marca']}"></span>{html.escape(etiqueta_semaforo(grupo['marca']))}</div>
                        <h3>{html.escape(grupo['cuadrilla'])}</h3>
                        <p>{grupo['reportes']} reporte(s) · {html.escape(sitios)}{html.escape(extra)}</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button(
                    "Ver reporte",
                    key=f"hc_abrir_{dia.isoformat()}_{clave_ubicacion(grupo['cuadrilla'])}",
                    use_container_width=True,
                ):
                    st.session_state.hc_cuadrilla = grupo["cuadrilla"]
                    st.rerun()


def render_subida_evidencia_cuadrilla(*, cuadrilla_fija: str = "", key_prefix: str = "evc") -> None:
    st.markdown('<div class="section-title">Evidencia fotográfica de cuadrilla</div>', unsafe_allow_html=True)
    st.caption(
        "Elige la cuadrilla y el día (01, 02…). Se guardan hasta 9 fotos en "
        f"`{EVIDENCIA_FOTOGRAFICA_DIR}/[Cuadrilla]/[Día]/` y el calendario marca ese día. "
        "Las fotos de una actividad se cargan directo en su registro."
    )
    catalogo = catalogo_cuadrillas_historial(bitacora_con_fechas(cargar_bitacora()))
    if cuadrilla_fija:
        cuadrilla = cuadrilla_fija
        st.markdown(f"**Cuadrilla:** {etiqueta_cuadrilla_historial(cuadrilla)}")
    else:
        etiquetas = [etiqueta_cuadrilla_historial(nombre) for nombre in catalogo]
        indice = 0
        previa = str(st.session_state.get(f"{key_prefix}_cuadrilla") or "")
        if previa in etiquetas:
            indice = etiquetas.index(previa)
        etiqueta = st.selectbox("Cuadrilla", etiquetas, index=indice, key=f"{key_prefix}_cuadrilla")
        cuadrilla = catalogo[etiquetas.index(etiqueta)]

    hoy = date.today()
    col_anio, col_mes, col_dia = st.columns(3)
    with col_anio:
        anio = int(
            st.selectbox(
                "Año",
                [hoy.year - 1, hoy.year, hoy.year + 1],
                index=1,
                key=f"{key_prefix}_anio",
            )
        )
    with col_mes:
        mes = int(
            st.selectbox(
                "Mes",
                list(range(1, 13)),
                index=hoy.month - 1,
                format_func=lambda numero: NOMBRES_MES[numero - 1],
                key=f"{key_prefix}_mes",
            )
        )
    ultimo = calendar.monthrange(anio, mes)[1]
    dias_txt = [f"{numero:02d}" for numero in range(1, ultimo + 1)]
    with col_dia:
        dia_txt = st.selectbox(
            "Día",
            dias_txt,
            index=min(hoy.day, ultimo) - 1,
            key=f"{key_prefix}_dia_{anio}_{mes}",
        )
    dia = date(anio, mes, int(dia_txt))
    base = bloque_historial_cuadrilla(bitacora_con_fechas(cargar_bitacora()), cuadrilla)
    marcas = mapa_dias_cuadrillas(base, anio, mes)
    dias_fotos = {fecha for fecha in fechas_fotos_cuadrilla(cuadrilla) if fecha.year == anio and fecha.month == mes}
    render_calendario_cuadrillas(anio, mes, marcas, dia, dias_fotos)

    existentes = fotos_evidencia_cuadrilla(cuadrilla, dia)
    carpeta = carpeta_evidencia_cuadrilla_dia(cuadrilla, dia).replace("\\", "/")
    if existentes:
        st.success(f"{dia_txt} · {len(existentes)} de {MAX_FOTOS} fotos en `{carpeta}`.")
        renderizar_galeria_rutas(existentes)
    else:
        st.info(f"El día {dia_txt} todavía no tiene fotos. Se guardarán en `{carpeta}`.")

    libres = max(0, MAX_FOTOS - len(existentes))
    archivos = None
    if libres <= 0:
        st.caption("Este día ya tiene las 9 fotografías.")
    else:
        archivos = st.file_uploader(
            "Lote de fotografías",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key=f"{key_prefix}_archivos_{clave_hist_cuadrilla(cuadrilla)}_{dia.isoformat()}",
            help=f"Hasta {MAX_FOTOS} imágenes por día. Quedan {libres} espacio(s).",
        )
        if archivos and len(archivos) > libres:
            st.caption(f"Solo se guardarán {libres} archivo(s); el día admite máximo {MAX_FOTOS}.")
    if st.button(
        "Guardar evidencia del día",
        type="primary",
        use_container_width=True,
        key=f"{key_prefix}_guardar_{clave_hist_cuadrilla(cuadrilla)}_{dia.isoformat()}",
        disabled=libres <= 0,
    ):
        resultado = guardar_lote_evidencia_cuadrilla(cuadrilla, dia, archivos or [])
        if not resultado.get("ok"):
            st.warning(resultado.get("mensaje") or "No se guardaron fotografías.")
        else:
            extra = f" Se omitieron {resultado['omitidas']}." if resultado.get("omitidas") else ""
            st.success(
                f"Se guardaron {resultado['agregadas']} foto(s) en `{resultado['carpeta']}`. "
                f"El día {dia_txt} queda marcado en el calendario.{extra}"
            )
            st.rerun()


def actividades_reporte_fotografico(inicio: date, fin: date) -> list[dict]:
    if fin < inicio:
        inicio, fin = fin, inicio
    df = bitacora_con_fechas(cargar_bitacora())
    if df is None or df.empty or "_fecha" not in df.columns:
        return []
    actividades = []
    for _, fila in df.iterrows():
        dia = fila.get("_fecha")
        if not isinstance(dia, date) or dia < inicio or dia > fin:
            continue
        cuadrilla_bruta = str(fila.get("CUADRILLA") or "").strip()
        cuadrilla = etiqueta_cuadrilla_historial(cuadrilla_bruta)
        secuencia = numero_secuencia_actividad(fila.get("SECUENCIA")) or 0
        descripcion = texto_visible(
            fila.get("DESCRIPCION DE LA ACTIVIDAD")
            or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD")
            or fila.get("ACTIVIDAD CORTA")
            or fila.get("TIPO DE ACTIVIDAD"),
            "",
        )
        descripcion = sitios_operativos.quitar_ninguno(descripcion)
        campos = textos_ubicacion_ot(fila)
        ubicacion = etiqueta_ubicacion_fila(fila) or formato_ubicacion_reporte(fila)
        if not ubicacion:
            ubicacion = sitios_operativos.texto_ubicacion_final(
                campos.get("pozo", ""),
                campos.get("ubicacion", ""),
                campos.get("equipo ubicacion", ""),
                campos.get("sector", ""),
            )
        if "AVANCE %" in fila.index:
            avance = avance_texto(fila.get("AVANCE %"))
        else:
            avance = avance_texto(fila.get("AVANCE"))
        fotos = [ruta for ruta in rutas_fotos_fila(fila) if ruta and os.path.isfile(ruta)]
        actividades.append(
            {
                "fecha": dia,
                "orden": orden_cuadrilla_oficial(cuadrilla)[0],
                "secuencia": secuencia,
                "ot": etiqueta_secuencia_consecutiva(secuencia) if secuencia else "",
                "ubicacion": ubicacion,
                "pozo": campos.get("pozo", ""),
                "sector": campos.get("sector", ""),
                "equipo": campos.get("equipo ubicacion", ""),
                "descripcion": descripcion,
                "avance": avance,
                "cuadrilla": cuadrilla,
                "fotos": fotos,
            }
        )
    return actividades


def modulo_reporte_fotografico():
    """Streamlit no recarga módulos ya importados; sin esto el PDF sale con el diseño anterior."""
    mtime_foto = _mtime_archivo(reporte_fotografico.__file__)
    mtime_sitios = _mtime_archivo(sitios_operativos.__file__)
    firma = (mtime_foto, mtime_sitios)
    if getattr(reporte_fotografico, "_mtime_cargado", None) != firma:
        importlib.reload(sitios_operativos)
        importlib.reload(reporte_fotografico)
        reporte_fotografico._mtime_cargado = firma
    return reporte_fotografico


def _al_descargar_reporte_fotografico(rutas: tuple, aviso_key: str) -> None:
    st.session_state[aviso_key] = persistencia.liberar_archivos_pesados(list(rutas or ()))


def _al_descargar_pdf_ot(ruta: str, fecha_iso: str, nombre: str) -> None:
    if persistencia.liberar_archivos_pesados([ruta]):
        persistencia.registrar_ot_liberada(fecha_iso, nombre)


def _al_descargar_copia_respaldo(rutas: tuple) -> None:
    persistencia.liberar_archivos_pesados(list(rutas or ()))


def render_generador_reporte_fotografico(*, key_prefix: str = "foto_rep") -> None:
    generador = modulo_reporte_fotografico()
    st.markdown(
        '<div class="section-title">Reporte fotográfico</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Genera el machote oficial con las actividades del periodo, "
        "ordenadas por cuadrilla: Eléctricos, Instrumentos, Mecánico 1 a 4, "
        "Montaje, Soldadura y Retro Excavadora. Las fotos quedan en filas de 3. "
        "Al descargar el Excel o el PDF, esas fotografías se retiran del servidor. "
        "La bitácora (fecha, pozo, actividad, insumos y estatus) se conserva."
    )
    disponibles = fechas_disponibles_operacion()
    ultima = disponibles[0] if disponibles else date.today()
    primera = disponibles[-1] if disponibles else date.today()
    if disponibles:
        st.caption(
            "Jornadas en bitácora/evidencia: "
            + ", ".join(d.strftime("%d/%m/%Y") for d in disponibles[:12])
            + ("…" if len(disponibles) > 12 else "")
        )
    modo = st.radio(
        "Periodo",
        ["Un día", "Rango de fechas"],
        horizontal=True,
        key=f"{key_prefix}_modo",
    )
    if modo == "Un día":
        etiquetas = [d.strftime("%d/%m/%Y") for d in disponibles]
        etiquetas.append("Otra fecha…")
        elegido = st.selectbox(
            "Día con registros",
            etiquetas,
            key=f"{key_prefix}_dia_lista",
        )
        if elegido == "Otra fecha…":
            dia = st.date_input(
                "Fecha",
                value=ultima,
                format="DD/MM/YYYY",
                key=f"{key_prefix}_dia",
            )
        else:
            dia = parsear_fecha_valor(elegido) or ultima
        inicio = fin = dia
    else:
        col_desde, col_hasta = st.columns(2)
        with col_desde:
            inicio = st.date_input(
                "Desde",
                value=primera,
                format="DD/MM/YYYY",
                key=f"{key_prefix}_desde",
            )
        with col_hasta:
            fin = st.date_input(
                "Hasta",
                value=ultima,
                format="DD/MM/YYYY",
                key=f"{key_prefix}_hasta",
            )
    opcion_todas = "Todas las cuadrillas (Reporte General)"
    cuadrilla_filtro = st.selectbox(
        "Cuadrilla",
        [opcion_todas, *ORDEN_CUADRILLAS_OFICIAL],
        key=f"{key_prefix}_cuadrilla",
        help="El reporte general incluye todas las cuadrillas de la jornada, en el orden oficial.",
    )
    if st.button("Generar reporte fotográfico", type="primary", key=f"{key_prefix}_generar"):
        st.session_state.pop(f"{key_prefix}_liberadas", None)
        actividades = actividades_reporte_fotografico(inicio, fin)
        if cuadrilla_filtro != opcion_todas:
            objetivo = clave_cuadrilla_oficial(cuadrilla_filtro)
            actividades = [
                item
                for item in actividades
                if clave_cuadrilla_oficial(item.get("cuadrilla") or "") == objetivo
            ]
        if not actividades:
            st.session_state.pop(f"{key_prefix}_paquete", None)
            if cuadrilla_filtro == opcion_todas:
                st.warning("No hay actividades de bitácora en el periodo seleccionado.")
            else:
                st.warning(
                    f"No hay actividades de {cuadrilla_filtro} en el periodo seleccionado."
                )
        else:
            with st.spinner("Armando el Excel y el PDF…"):
                try:
                    xlsx, pdf, resumen = generador.generar_reporte_fotografico(actividades)
                except Exception as exc:
                    st.session_state.pop(f"{key_prefix}_paquete", None)
                    st.error(f"No se pudo generar el reporte: {exc}")
                else:
                    fotos_rutas = []
                    vistas = set()
                    for item in actividades:
                        for ruta_foto in item.get("fotos") or []:
                            if ruta_foto and ruta_foto not in vistas:
                                vistas.add(ruta_foto)
                                fotos_rutas.append(ruta_foto)
                    st.session_state[f"{key_prefix}_paquete"] = {
                        "xlsx": xlsx,
                        "pdf": pdf,
                        "resumen": resumen,
                        "inicio": min(inicio, fin),
                        "fin": max(inicio, fin),
                        "cuadrilla": "" if cuadrilla_filtro == opcion_todas else cuadrilla_filtro,
                        "version": generador.VERSION_DISENO_PDF,
                        "fotos_rutas": tuple(fotos_rutas),
                    }
    paquete = st.session_state.get(f"{key_prefix}_paquete")
    if paquete and paquete.get("version") != generador.VERSION_DISENO_PDF:
        st.session_state.pop(f"{key_prefix}_paquete", None)
        st.info("El diseño del reporte se actualizó; vuelve a presionar «Generar reporte fotográfico».")
        return
    if not paquete:
        return
    resumen = paquete["resumen"]
    st.success(
        f"{resumen['actividades']} actividad(es) · {resumen['hojas']} hoja(s) · "
        f"{resumen['fotos']} fotografía(s)."
    )
    if resumen.get("omitidas"):
        st.caption(f"{resumen['omitidas']} fotografía(s) no se pudieron insertar.")
    marca = paquete["inicio"].strftime("%Y%m%d")
    if paquete["fin"] != paquete["inicio"]:
        marca += "_" + paquete["fin"].strftime("%Y%m%d")
    if paquete.get("cuadrilla"):
        marca += "_" + clave_cuadrilla_oficial(paquete["cuadrilla"]).replace(" ", "_")
        st.caption(f"Cuadrilla: {paquete['cuadrilla']}.")
    fotos_rutas = tuple(paquete.get("fotos_rutas") or ())
    aviso_key = f"{key_prefix}_liberadas"
    col_xlsx, col_pdf = st.columns(2)
    with col_xlsx:
        st.download_button(
            "Descargar Excel (.xlsx)",
            data=paquete["xlsx"],
            file_name=f"reporte_fotografico_{marca}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"{key_prefix}_xlsx",
            on_click=_al_descargar_reporte_fotografico,
            args=(fotos_rutas, aviso_key),
        )
    with col_pdf:
        st.download_button(
            "Descargar PDF",
            data=paquete["pdf"],
            file_name=f"reporte_fotografico_{marca}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key=f"{key_prefix}_pdf",
            on_click=_al_descargar_reporte_fotografico,
            args=(fotos_rutas, aviso_key),
        )
    liberadas = st.session_state.get(aviso_key)
    if liberadas:
        st.caption(
            f"Se retiraron {liberadas} fotografía(s) del servidor. "
            "El archivo descargado las conserva y la bitácora no se modifica."
        )


def render_historial_cuadrillas(*, encabezado: bool = True) -> None:
    if encabezado:
        render_encabezado(
            "📋 Historial y Reportes de Cuadrillas",
            "Consulta centralizada de órdenes de trabajo, insumos y evidencias por cuadrilla y fecha",
        )
    else:
        st.markdown("### 📋 Historial y Reportes de Cuadrillas - Base Álamo")
        st.markdown(
            "Consulta centralizada de órdenes de trabajo, insumos y evidencias fotográficas por cuadrilla y fecha."
        )
        st.divider()

    render_generador_reporte_fotografico(key_prefix="hc_foto")
    df = bitacora_con_fechas(cargar_bitacora())
    cuadrilla = str(st.session_state.get("hc_cuadrilla") or "").strip()
    dia_sel = parsear_fecha_valor(st.session_state.get("hc_dia"))

    if not cuadrilla:
        st.markdown("#### 🛠️ Seleccione una Cuadrilla Operativa")
        st.caption("Carpeta por cuadrilla. Elige una para ver sus jornadas con bitácora.")
        catalogo = catalogo_cuadrillas_historial(df)
        col1, col2, col3 = st.columns(3)
        for i, nombre in enumerate(catalogo):
            col = [col1, col2, col3][i % 3]
            jornadas = fechas_historial_cuadrilla(df, nombre)
            clave = clave_hist_cuadrilla(nombre)
            with col:
                if st.button(
                    f"📁 {etiqueta_cuadrilla_historial(nombre)}",
                    key=f"hc_folder_{clave}",
                    use_container_width=True,
                ):
                    st.session_state.hc_cuadrilla = nombre
                    st.session_state.hc_dia = None
                    st.rerun()
                if jornadas:
                    ultimo = jornadas[0].strftime("%d/%m/%Y")
                    st.caption(f"{len(jornadas)} jornada(s) · última {ultimo}")
                else:
                    st.caption("Sin reportes en bitácora")
        return

    etiqueta = etiqueta_cuadrilla_historial(cuadrilla)

    if dia_sel is None:
        if st.button("⬅️ Volver a Lista de Cuadrillas", key="hc_volver_lista"):
            st.session_state.hc_cuadrilla = None
            st.session_state.hc_dia = None
            st.rerun()
        st.markdown(f"#### 📅 Historial de Reportes: {etiqueta}")
        st.markdown(
            "Seleccione una fecha para consultar el detalle de actividades, insumos y evidencias."
        )
        render_nota_asistencia(cuadrilla)
        resumen_cal = resumen_jornadas_cuadrilla(cuadrilla)
        if resumen_cal and resumen_cal.get("dias"):
            with st.expander(f"Calendario de asistencia · {resumen_cal.get('periodo') or ''}"):
                filas_dias = []
                for dia_item in resumen_cal["dias"]:
                    codigo_dia = str(dia_item.get("codigo") or "").strip()
                    if dia_item.get("laboral"):
                        estatus_dia = "Jornada"
                    elif _es_franco_asistencia(codigo_dia):
                        estatus_dia = "Franco (F/O)"
                    else:
                        estatus_dia = "Sin registro"
                    fecha_txt = str(dia_item.get("fecha") or "")
                    try:
                        fecha_txt = date.fromisoformat(fecha_txt).strftime("%d/%m/%Y")
                    except ValueError:
                        pass
                    nombres_dia = " · ".join(nombres_de_codigo_personal(codigo_dia))
                    filas_dias.append(
                        {
                            "Fecha": fecha_txt,
                            "Día": dia_item.get("dia_semana") or "",
                            "Código": codigo_dia or "—",
                            "Personal": nombres_dia or "—",
                            "Estatus": estatus_dia,
                        }
                    )
                st.dataframe(pd.DataFrame(filas_dias), hide_index=True, use_container_width=True)
        render_subida_evidencia_cuadrilla(
            cuadrilla_fija=cuadrilla,
            key_prefix=f"hc_ev_{clave_hist_cuadrilla(cuadrilla)}",
        )
        st.divider()
        historial = resumen_fechas_cuadrilla(df, cuadrilla)
        if not historial:
            st.info("Esta cuadrilla aún no tiene reportes capturados en la bitácora.")
        for item in historial:
            dia = item["dia"]
            if item.get("lote"):
                extra_ev = " · evidencia fotográfica"
            elif item.get("sin_fotos"):
                extra_ev = f" · {item['sin_fotos']} sin fotos"
            else:
                extra_ev = " · evidencia completa"
            estado = f" · {etiqueta_semaforo(item['marca'])}" if item.get("ots") else ""
            if st.button(
                f"📄 {dia.strftime('%d/%m/%Y')} · {item['ots']} OT{estado}{extra_ev}",
                key=f"hc_fecha_{clave_hist_cuadrilla(cuadrilla)}_{dia.isoformat()}",
                use_container_width=True,
            ):
                st.session_state.hc_dia = dia.isoformat()
                st.rerun()
        col_f, col_g = st.columns([3, 1])
        with col_f:
            elegido = st.date_input(
                "O ir a otra fecha",
                value=date.today(),
                format="DD/MM/YYYY",
                key="hc_date_input",
            )
        with col_g:
            st.markdown("<div style='height:1.7rem'></div>", unsafe_allow_html=True)
            if st.button("Ir", key="hc_ir_fecha", use_container_width=True) and isinstance(elegido, date):
                st.session_state.hc_dia = elegido.isoformat()
                st.rerun()
        return

    if st.button("⬅️ Volver al Historial de Fechas", key="hc_volver_fechas"):
        st.session_state.hc_dia = None
        st.rerun()
    st.markdown(
        f'<div class="explorer-crumb">Historial de Cuadrillas / '
        f"<b>{html.escape(etiqueta)}</b> / {dia_sel.strftime('%d/%m/%Y')}</div>",
        unsafe_allow_html=True,
    )
    render_reporte_cuadrilla(dia_sel, cuadrilla, bitacora_del_dia(df, dia_sel))


def soporte_default() -> dict:
    return {
        "anio": date.today().year,
        "proyecto": [],
        "aceite": {},
        "po_baker": [],
        "fuera_po": [],
        "fuera_adendum": [],
        "actualizado": "",
    }


def cargar_soporte() -> dict:
    datos = leer_json(SOPORTE_FILE, soporte_default())
    if not isinstance(datos, dict):
        return soporte_default()
    base = soporte_default()
    base.update({k: datos.get(k, base[k]) for k in base})
    for clave in ("proyecto", "po_baker", "fuera_po", "fuera_adendum"):
        if not isinstance(base[clave], list):
            base[clave] = []
    if not isinstance(base.get("aceite"), dict):
        base["aceite"] = {}
    return base


def guardar_soporte(datos: dict) -> None:
    datos = dict(datos)
    datos["actualizado"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    escribir_json(SOPORTE_FILE, datos)


def etiqueta_pozo_soporte(nombre: str, zona: dict | None = None, zona_id: str = "") -> str:
    texto = str(nombre or "").strip()
    if not texto:
        return "Sin pozo"
    zona = zona or {}
    corto = ""
    zid = zona_id or str(zona.get("id") or "")
    for item in ZONAS_HISTORIAL_POZOS:
        if item["id"] == zid:
            corto = str(item.get("detalle") or "").replace("Pozos ", "").strip()
            break
    if not corto:
        corto = str(zona.get("detalle") or zona.get("nombre") or "").replace("Pozos ", "").strip()
    m = re.search(r"(\d+[A-Za-z]?)$", clave_ubicacion(texto).replace("-", ""))
    if corto and m:
        return f"{corto} {m.group(1)}"
    return texto


def ficha_para_soporte(datos: dict, pozo: str) -> dict | None:
    ficha = encontrar_ficha(datos, pozo)
    if ficha:
        return ficha
    for cand in (datos.get("fichas") or {}).values():
        if _texto_coincide_pozo(pozo, cand):
            return cand
    return None


def etiquetas_pozos_soporte(zona_id: str | None = None) -> list[str]:
    datos_ub = cargar_ubicaciones()
    zonas = mapa_zonas(datos_ub.get("zonas"))
    etiquetas = []
    for ficha in (datos_ub.get("fichas") or {}).values():
        zid = str(ficha.get("zona_id") or "")
        if zona_id and zid != zona_id:
            continue
        if not zona_id and zid == "instalaciones":
            continue
        etiquetas.append(etiqueta_pozo_soporte(ficha.get("nombre", ""), zonas.get(zid, {}), zid))
    return etiquetas


def _num_insumo(valor) -> float:
    try:
        return float(valor or 0)
    except (TypeError, ValueError):
        return 0.0


def periodo_soporte(anio: int) -> tuple[date, date]:
    return date(anio, 8, 1), date(anio, 12, 31)


def filas_proyecto_desde_bitacora(anio: int) -> list[dict]:
    inicio, fin = periodo_soporte(anio)
    df = bitacora_con_fechas(cargar_bitacora())
    datos_ub = cargar_ubicaciones()
    zonas = mapa_zonas(datos_ub.get("zonas"))
    filas = []
    if df.empty:
        return filas
    for idx, fila in df.iterrows():
        fch = fila.get("_fecha")
        if not isinstance(fch, date) or fch < inicio or fch > fin:
            continue
        pozo = str(fila.get("POZO") or "").strip()
        ficha = ficha_para_soporte(datos_ub, pozo) if pozo else None
        zona_id = (ficha or {}).get("zona_id", "")
        zona = zonas.get(zona_id, {})
        item = {
            "id": f"bit-{idx}-{fch.isoformat()}-{fila.get('SECUENCIA', '')}-{pozo}",
            "origen": "bitacora",
            "Pozo": etiqueta_pozo_soporte(pozo or fila.get("SECTOR"), zona, zona_id),
            "Zona": titulo_zona_historial(zona_id, zona) if ficha else str(fila.get("SECTOR") or ""),
            "OT": str(fila.get("SECUENCIA") or "").strip(),
            "Trabajo": str(fila.get("TIPO DE ACTIVIDAD") or fila.get("DESCRIPCIÓN DE LA ACTIVIDAD") or "").strip(),
            "Fecha": fch.strftime("%Y-%m-%d"),
            "Cuadrilla": str(fila.get("CUADRILLA") or "").strip(),
        }
        for col in COLUMNAS_INSUMOS:
            item[ETIQUETAS_INSUMO_SOPORTE[col]] = _num_insumo(fila.get(col))
        item["Bandas (PZA)"] = _num_insumo(fila.get("BANDAS CANTIDAD"))
        filas.append(item)
    return filas


def mezclar_proyecto(guardadas: list[dict], bitacora: list[dict]) -> list[dict]:
    manuales = [f for f in guardadas if str(f.get("origen") or "") != "bitacora"]
    auto = [f for f in guardadas if str(f.get("origen") or "") == "bitacora"]
    por_id = {str(f.get("id")): f for f in auto}
    mezcla = []
    ids_nuevos = set()
    for fila in bitacora:
        actual = por_id.get(str(fila["id"]), fila)
        if actual is not fila:
            combinada = dict(fila)
            combinada.update({k: actual.get(k, v) for k, v in actual.items() if k not in {"id", "origen"}})
            mezcla.append(combinada)
        else:
            mezcla.append(fila)
        ids_nuevos.add(str(fila["id"]))
    for fila in auto:
        if str(fila.get("id")) not in ids_nuevos:
            mezcla.append(fila)
    mezcla.extend(manuales)
    return mezcla


def filas_aceite_desde_bitacora(anio: int, mes: int) -> dict[str, dict[str, float]]:
    df = bitacora_con_fechas(cargar_bitacora())
    datos_ub = cargar_ubicaciones()
    zonas = mapa_zonas(datos_ub.get("zonas"))
    matriz: dict[str, dict[str, float]] = {}
    if df.empty:
        return matriz
    for _, fila in df.iterrows():
        fch = fila.get("_fecha")
        if not isinstance(fch, date) or fch.year != anio or fch.month != mes:
            continue
        litros = _num_insumo(fila.get("ACEITE (LTS)"))
        if litros <= 0:
            continue
        pozo = str(fila.get("POZO") or "").strip() or str(fila.get("SECTOR") or "Sin pozo")
        ficha = ficha_para_soporte(datos_ub, pozo) if pozo else None
        zona_id = (ficha or {}).get("zona_id", "")
        zona = zonas.get(zona_id, {})
        etiqueta = etiqueta_pozo_soporte(pozo, zona, zona_id)
        dia = str(fch.day)
        matriz.setdefault(etiqueta, {})
        matriz[etiqueta][dia] = round(matriz[etiqueta].get(dia, 0) + litros, 2)
    return matriz


def filas_po_default() -> list[dict]:
    filas = []
    for i, item in enumerate(catalogo_insumos_activos(), start=1):
        unidad = str(item.get("unidad") or "PZA").upper()
        if unidad in {"PZAS", "PZA", "PIEZAS"}:
            unidad = "PZA"
        filas.append(
            {
                "Ítem": f"{i:02d}",
                "Descripción": item.get("nombre", ""),
                "Unidad": unidad if unidad in UNIDADES_SOPORTE else "PZA",
                "Precio unitario": 0.0,
                "Cantidad requerida": 0.0,
                "Cantidad suministrada": 0.0,
            }
        )
    extras = [
        ("Filtro de combustible", "PZA"),
        ("Empaque de culata", "JGO"),
        ("Aceite SAE 50 a granel", "TAMBO"),
    ]
    n = len(filas)
    for j, (nombre, unidad) in enumerate(extras, start=1):
        if any(f["Descripción"].casefold() == nombre.casefold() for f in filas):
            continue
        n += 1
        filas.append(
            {
                "Ítem": f"{n:02d}",
                "Descripción": nombre,
                "Unidad": unidad,
                "Precio unitario": 0.0,
                "Cantidad requerida": 0.0,
                "Cantidad suministrada": 0.0,
            }
        )
    return filas


def completar_finanzas(filas: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(filas) if filas else pd.DataFrame(
        columns=["Ítem", "Descripción", "Unidad", "Precio unitario", "Cantidad requerida", "Cantidad suministrada"]
    )
    for col in ("Precio unitario", "Cantidad requerida", "Cantidad suministrada"):
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    for col in ("Ítem", "Descripción", "Unidad"):
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str)
    df["Subtotal"] = (df["Precio unitario"] * df["Cantidad requerida"]).round(2)
    df["Cantidad excedida"] = (df["Cantidad suministrada"] - df["Cantidad requerida"]).clip(lower=0).round(2)
    return df[["Ítem", "Descripción", "Unidad", "Precio unitario", "Cantidad requerida", "Subtotal", "Cantidad suministrada", "Cantidad excedida"]]


def filas_finanzas_editor(df: pd.DataFrame) -> list[dict]:
    salida = []
    for _, fila in df.iterrows():
        desc = str(fila.get("Descripción") or "").strip()
        if not desc or desc.lower() in {"nan", "none"}:
            continue
        salida.append(
            {
                "Ítem": str(fila.get("Ítem") or "").strip(),
                "Descripción": desc,
                "Unidad": str(fila.get("Unidad") or "PZA").strip() or "PZA",
                "Precio unitario": _num_insumo(fila.get("Precio unitario")),
                "Cantidad requerida": _num_insumo(fila.get("Cantidad requerida")),
                "Cantidad suministrada": _num_insumo(fila.get("Cantidad suministrada")),
            }
        )
    return salida


def money_html(titulo: str, monto: float) -> str:
    return (
        f'<div class="money-card"><div class="k">{html.escape(titulo)}</div>'
        f"<h3>${monto:,.2f}</h3></div>"
    )


def exportar_excel_soporte(datos: dict, anio: int, mes_aceite: int) -> bytes:
    buf = io.BytesIO()
    proyecto_filas = [
        f for f in (datos.get("proyecto") or [])
        if not str(f.get("Fecha") or "").strip() or str(f.get("Fecha")).startswith(str(anio))
    ]
    proyecto = pd.DataFrame(proyecto_filas)
    po = completar_finanzas(datos.get("po_baker") or [])
    fuera = completar_finanzas(datos.get("fuera_po") or [])
    adendum = completar_finanzas(datos.get("fuera_adendum") or [])
    aceite_json = (datos.get("aceite") or {}).get(clave_mes(anio, mes_aceite), {})
    filas_aceite = []
    dias = calendar.monthrange(anio, mes_aceite)[1]
    for pozo, dias_map in sorted(aceite_json.items()):
        fila = {"Pozo": pozo}
        total = 0.0
        for d in range(1, dias + 1):
            val = _num_insumo((dias_map or {}).get(str(d)))
            fila[str(d)] = val
            total += val
        fila["Total L"] = round(total, 2)
        filas_aceite.append(fila)
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (proyecto if not proyecto.empty else pd.DataFrame({"Aviso": ["Sin filas"]})).to_excel(
            writer, sheet_name="Proyecto Refacciones", index=False
        )
        (pd.DataFrame(filas_aceite) if filas_aceite else pd.DataFrame({"Aviso": ["Sin litros"]})).to_excel(
            writer, sheet_name="Aceite de Relleno", index=False
        )
        po.to_excel(writer, sheet_name="PO Baker", index=False)
        fuera.to_excel(writer, sheet_name="Fuera PO Baker", index=False)
        adendum.to_excel(writer, sheet_name="Fuera Adendum", index=False)
    return buf.getvalue()


def render_tab_proyecto_refacciones(datos: dict, anio: int) -> None:
    st.caption(
        f"Órdenes de {NOMBRES_MES[7]} a {NOMBRES_MES[11]} {anio}, cruzadas con bitácora, "
        "cuadrilla y el desglose de refacciones / consumibles."
    )
    zonas = ["Todas"] + [z["titulo"] for z in ZONAS_HISTORIAL_POZOS]
    filtro = st.selectbox("Filtrar por zona", zonas, key="sop_zona_proy")
    filas = []
    for fila in datos.get("proyecto") or []:
        fecha = str(fila.get("Fecha") or "").strip()
        if fecha and not fecha.startswith(str(anio)):
            continue
        if filtro != "Todas" and str(fila.get("Zona") or "") != filtro:
            continue
        filas.append(fila)
    cols_base = ["Pozo", "Zona", "OT", "Trabajo", "Fecha", "Cuadrilla"]
    cols_ins = list(ETIQUETAS_INSUMO_SOPORTE.values()) + ["Bandas (PZA)"]
    if not filas:
        vacia = {c: "" for c in cols_base}
        vacia.update({c: 0.0 for c in cols_ins})
        df = pd.DataFrame([vacia])
    else:
        df = pd.DataFrame(filas)
        for col in cols_base + ["id", "origen"]:
            if col not in df.columns:
                df[col] = ""
        for col in cols_ins:
            if col not in df.columns:
                df[col] = 0.0
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    visibles = cols_base + cols_ins
    editor = st.data_editor(
        df[visibles],
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"sop_editor_proy_{anio}_{filtro}",
        column_config={
            "Fecha": st.column_config.TextColumn(help="AAAA-MM-DD"),
            **{c: st.column_config.NumberColumn(format="%.2f", min_value=0) for c in cols_ins},
        },
    )
    m1, m2, m3 = st.columns(3)
    m1.metric("OT / renglones", int(len(editor)))
    m2.metric("Aceite SAE-50 (L)", float(pd.to_numeric(editor.get("Aceite SAE-50 (L)", 0), errors="coerce").fillna(0).sum()))
    m3.metric("Filtros (PZA)", float(pd.to_numeric(editor.get("Filtro aceite (PZA)", 0), errors="coerce").fillna(0).sum()))
    if st.button("💾 Guardar proyecto de refacciones", type="primary", key="btn_sop_proy"):
        guardadas = []
        origenes = list(filas)
        for i, fila in editor.reset_index(drop=True).iterrows():
            pozo = str(fila.get("Pozo") or "").strip()
            trabajo = str(fila.get("Trabajo") or "").strip()
            if not pozo and not trabajo:
                continue
            prev = origenes[i] if i < len(origenes) else {}
            item = {
                "id": prev.get("id") or uuid.uuid4().hex[:10],
                "origen": prev.get("origen") or "manual",
                "Pozo": pozo,
                "Zona": str(fila.get("Zona") or "").strip(),
                "OT": str(fila.get("OT") or "").strip(),
                "Trabajo": trabajo,
                "Fecha": str(fila.get("Fecha") or "").strip(),
                "Cuadrilla": str(fila.get("Cuadrilla") or "").strip(),
            }
            for col in cols_ins:
                item[col] = _num_insumo(fila.get(col))
            guardadas.append(item)
        resto = []
        for previa in datos.get("proyecto") or []:
            fecha = str(previa.get("Fecha") or "").strip()
            misma_zona = filtro == "Todas" or str(previa.get("Zona") or "") == filtro
            mismo_anio = (not fecha) or fecha.startswith(str(anio))
            if misma_zona and mismo_anio:
                continue
            resto.append(previa)
        datos["proyecto"] = resto + guardadas
        guardar_soporte(datos)
        st.success(f"Proyecto guardado en {SOPORTE_FILE}.")
        st.rerun()


def render_tab_aceite_relleno(datos: dict, anio: int) -> None:
    meses = [m for m in MESES_SOPORTE]
    etiquetas = [f"{NOMBRES_MES[m - 1]} {anio}" for m in meses]
    f1, f2 = st.columns(2)
    with f1:
        mes = st.selectbox("Mes de auditoría", etiquetas, key="sop_mes_aceite")
    zonas_filtro = ["Todas"] + [z["titulo"] for z in ZONAS_HISTORIAL_POZOS]
    with f2:
        zona_opt = st.selectbox("Zona", zonas_filtro, index=1 if len(zonas_filtro) > 1 else 0, key="sop_zona_aceite")
    mes_n = meses[etiquetas.index(mes)]
    zona_id = next((z["id"] for z in ZONAS_HISTORIAL_POZOS if z["titulo"] == zona_opt), None)
    dias = calendar.monthrange(anio, mes_n)[1]
    clave = clave_mes(anio, mes_n)
    guardado = (datos.get("aceite") or {}).get(clave, {})
    bit = filas_aceite_desde_bitacora(anio, mes_n)
    catalogo = set(etiquetas_pozos_soporte(zona_id if zona_opt != "Todas" else None))
    extras = set(guardado) | set(bit)
    if zona_opt != "Todas":
        extras = {p for p in extras if p in catalogo}
    pozos = sorted(catalogo | extras)
    filas = []
    for pozo in pozos:
        if not pozo:
            continue
        fila = {"Pozo": pozo}
        total = 0.0
        for d in range(1, dias + 1):
            ds = str(d)
            if pozo in guardado and ds in (guardado.get(pozo) or {}):
                val = _num_insumo(guardado[pozo][ds])
            else:
                val = _num_insumo((bit.get(pozo) or {}).get(ds))
            fila[ds] = val
            total += val
        fila["Total L"] = round(total, 2)
        filas.append(fila)
    if not filas:
        fila = {"Pozo": "Soledad 107"}
        for d in range(1, dias + 1):
            fila[str(d)] = 0.0
        fila["Total L"] = 0.0
        filas.append(fila)
    df = pd.DataFrame(filas)
    st.caption("Matriz diaria de litros de aceite SAE-50 por pozo. Las celdas se cruzan con la bitácora de cuadrillas mecánicas.")
    editor = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=f"sop_editor_aceite_{clave}",
        column_config={
            "Pozo": st.column_config.TextColumn(required=True, width="medium"),
            "Total L": st.column_config.NumberColumn(disabled=True, format="%.1f"),
            **{str(d): st.column_config.NumberColumn(format="%.1f", min_value=0) for d in range(1, dias + 1)},
        },
    )
    tot_mes = 0.0
    pozos_con = 0
    for _, fila in editor.iterrows():
        litros = 0.0
        for d in range(1, dias + 1):
            if str(d) in editor.columns:
                litros += _num_insumo(fila.get(str(d)))
        tot_mes += litros
        if litros > 0:
            pozos_con += 1
    c1, c2 = st.columns(2)
    c1.metric(f"Litros en {NOMBRES_MES[mes_n - 1]}", round(tot_mes, 2))
    c2.metric("Pozos con relleno", pozos_con)
    if st.button("💾 Guardar matriz de aceite", type="primary", key="btn_sop_aceite"):
        mes_map = {} if zona_opt == "Todas" else dict((datos.get("aceite") or {}).get(clave) or {})
        if zona_opt != "Todas":
            for pozo_cat in catalogo:
                mes_map.pop(pozo_cat, None)
        for _, fila in editor.iterrows():
            pozo = str(fila.get("Pozo") or "").strip()
            if not pozo:
                continue
            mes_map[pozo] = {str(d): _num_insumo(fila.get(str(d))) for d in range(1, dias + 1)}
        datos.setdefault("aceite", {})[clave] = mes_map
        guardar_soporte(datos)
        st.success(f"Aceite de relleno de {NOMBRES_MES[mes_n - 1]} guardado en {SOPORTE_FILE}.")
        st.rerun()


def render_editor_finanzas(titulo: str, ayuda: str, filas: list[dict], key: str) -> pd.DataFrame:
    st.markdown(f'<div class="section-title">{html.escape(titulo)}</div>', unsafe_allow_html=True)
    st.caption(ayuda)
    df = completar_finanzas(filas)
    editor = st.data_editor(
        df.drop(columns=["Subtotal", "Cantidad excedida"]),
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key=key,
        column_config={
            "Precio unitario": st.column_config.NumberColumn(format="$%.2f", min_value=0),
            "Cantidad requerida": st.column_config.NumberColumn(format="%.2f", min_value=0),
            "Cantidad suministrada": st.column_config.NumberColumn(format="%.2f", min_value=0),
            "Unidad": st.column_config.SelectboxColumn(options=UNIDADES_SOPORTE),
        },
    )
    calc = completar_finanzas(filas_finanzas_editor(editor))
    st.dataframe(calc[["Ítem", "Descripción", "Subtotal", "Cantidad excedida"]], use_container_width=True, hide_index=True)
    return calc


def render_tab_po_baker(datos: dict) -> None:
    st.caption(
        "Control financiero de inventario y compras para proformas OPS. "
        "El subtotal es Precio unitario × Cantidad requerida. "
        "La cantidad excedida es lo suministrado por encima de lo requerido."
    )
    po = render_editor_finanzas(
        "P.O. Baker e insumos del pedido",
        "Materiales contemplados en la orden Baker.",
        datos.get("po_baker") or filas_po_default(),
        "sop_editor_po",
    )
    fuera = render_editor_finanzas(
        "Materiales fuera de P.O. Baker",
        "Consumos o compras que no viajan en la P.O. original.",
        datos.get("fuera_po") or [],
        "sop_editor_fuera",
    )
    adendum = render_editor_finanzas(
        "Materiales fuera de Adéndum",
        "Ítems adicionales al adéndum contractual.",
        datos.get("fuera_adendum") or [],
        "sop_editor_adendum",
    )
    t_po = float(po["Subtotal"].sum()) if not po.empty else 0.0
    t_fuera = float(fuera["Subtotal"].sum()) if not fuera.empty else 0.0
    t_ad = float(adendum["Subtotal"].sum()) if not adendum.empty else 0.0
    g1, g2, g3, g4 = st.columns(4)
    g1.markdown(money_html("Subtotal P.O. Baker", t_po), unsafe_allow_html=True)
    g2.markdown(money_html("Fuera de P.O.", t_fuera), unsafe_allow_html=True)
    g3.markdown(money_html("Fuera de Adéndum", t_ad), unsafe_allow_html=True)
    g4.markdown(money_html("Subtotal general acumulado", t_po + t_fuera + t_ad), unsafe_allow_html=True)
    if st.button("💾 Guardar control P.O. Baker", type="primary", key="btn_sop_po"):
        datos["po_baker"] = filas_finanzas_editor(po)
        datos["fuera_po"] = filas_finanzas_editor(fuera)
        datos["fuera_adendum"] = filas_finanzas_editor(adendum)
        guardar_soporte(datos)
        st.success(f"Control financiero guardado en {SOPORTE_FILE}.")
        st.rerun()


def render_soporte_refacciones() -> None:
    st.markdown('<div class="section-title">Soporte de Refacciones Agosto — Diciembre</div>', unsafe_allow_html=True)
    st.caption(
        "Réplica gerencial del control Excel: proyecto por pozo/OT, aceite de relleno diario "
        f"y P.O. Baker. Los cambios viven en {SOPORTE_FILE} y se exportan a proforma."
    )
    hoy = date.today()
    anio = st.selectbox(
        "Año del soporte",
        [hoy.year - 1, hoy.year, hoy.year + 1],
        index=1,
        key="sop_anio",
    )
    datos = cargar_soporte()
    datos["anio"] = anio
    hay_anio = any(str(f.get("Fecha") or "").startswith(str(anio)) for f in (datos.get("proyecto") or []))
    if not hay_anio:
        semilla = filas_proyecto_desde_bitacora(anio)
        if semilla:
            datos["proyecto"] = mezclar_proyecto(datos.get("proyecto") or [], semilla)
            guardar_soporte(datos)
    if not datos.get("po_baker"):
        datos["po_baker"] = filas_po_default()
        guardar_soporte(datos)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        if st.button("🔄 Actualizar desde bitácora", key="btn_sop_sync"):
            datos["proyecto"] = mezclar_proyecto(datos.get("proyecto") or [], filas_proyecto_desde_bitacora(anio))
            guardar_soporte(datos)
            st.success("Proyecto sincronizado con bitacora_mantenimiento.xlsx. Las filas manuales se conservaron.")
            st.rerun()
    with col_b:
        mes_exp = st.selectbox(
            "Mes a exportar (aceite)",
            [NOMBRES_MES[m - 1] for m in MESES_SOPORTE],
            key="sop_mes_exp",
        )
        mes_n = MESES_SOPORTE[[NOMBRES_MES[m - 1] for m in MESES_SOPORTE].index(mes_exp)]
    with col_c:
        st.download_button(
            "⬇️ Exportar Excel proforma",
            data=exportar_excel_soporte(datos, anio, mes_n),
            file_name=f"soporte_refacciones_{anio}_{mes_n:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_sop_xlsx",
        )
    if datos.get("actualizado"):
        st.caption(f"Última actualización: {datos['actualizado']}")

    tab_p, tab_a, tab_po = st.tabs(
        ["Proyecto Refacciones", "Aceite de Relleno", "P.O. Baker y Insumos"]
    )
    with tab_p:
        render_tab_proyecto_refacciones(datos, anio)
    with tab_a:
        render_tab_aceite_relleno(datos, anio)
    with tab_po:
        render_tab_po_baker(datos)


def _num_combustible(valor) -> float:
    try:
        if valor is None or (isinstance(valor, str) and not valor.strip()):
            return 0.0
        numero = float(valor)
    except (TypeError, ValueError):
        return 0.0
    if numero != numero or numero in (float("inf"), float("-inf")):
        return 0.0
    return numero


def redondear_mitad_arriba(valor, decimales: int = 2) -> float:
    """Redondeo aritmético: desde .5 en el siguiente decimal, aleja de cero."""
    cuantia = Decimal("1").scaleb(-int(decimales))
    if not isinstance(valor, Decimal):
        valor = Decimal(str(_num_combustible(valor)))
    return float(valor.quantize(cuantia, rounding=ROUND_HALF_UP))


def costo_total_combustible(litros: float, precio: float) -> float:
    """Litros por precio, en centavos. 44.135 × 27 = 1191.645 → 1191.65."""
    litros_d = Decimal(str(_num_combustible(litros))).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
    precio_d = Decimal(str(_num_combustible(precio))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return float((litros_d * precio_d).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def normalizar_tipo_combustible(valor) -> str:
    texto = str(valor or "").strip()
    if not texto or texto.lower() in {"nan", "none"}:
        return ""
    clave = _sin_acentos(texto).casefold()
    if "diesel" in clave:
        return "Diesel"
    if "gasolin" in clave:
        return "Gasolina"
    return texto


def cargar_parque_vehicular(ruta: str) -> pd.DataFrame:
    columnas = ["No. económico", "Combustible asignado", "Modelo"]
    try:
        crudo = pd.read_excel(_copia_excel_legible(ruta), sheet_name="Listas")
    except Exception:
        return pd.DataFrame(columns=columnas)
    if crudo.empty:
        return pd.DataFrame(columns=columnas)
    unidad = crudo.iloc[:, 0].astype(str).str.strip()
    tipo = crudo.iloc[:, 1].map(normalizar_tipo_combustible) if crudo.shape[1] > 1 else ""
    if crudo.shape[1] > 3:
        modelo = crudo.iloc[:, 3].astype(str).str.strip()
    else:
        modelo = pd.Series([""] * len(crudo), index=crudo.index)
    df = pd.DataFrame({"No. económico": unidad, "Combustible asignado": tipo, "Modelo": modelo})
    df["Modelo"] = df["Modelo"].fillna("").astype(str).str.strip()
    df.loc[df["Modelo"].str.casefold().isin({"", "nan", "none", "modelo"}), "Modelo"] = ""
    df = df[~df["No. económico"].str.casefold().isin({"", "nan", "none", "numero_economico", "número económico"})]
    return df.drop_duplicates(subset=["No. económico"]).reset_index(drop=True)


def cargar_registro_combustible(ruta: str) -> pd.DataFrame:
    crudo = pd.read_excel(_copia_excel_legible(ruta), sheet_name="Registro Diario", header=3)
    if crudo.empty:
        return pd.DataFrame()
    alias = {
        "item": "Item",
        "fecha": "Fecha",
        "numero economico": "No. económico",
        "numero_economico": "No. económico",
        "kilometraje inicial": "Km inicial",
        "kilometraje actual (final)": "Km actual",
        "kilometraje recorrido": "Km recorrido",
        "tipo de combustible": "Combustible",
        "litros cargados": "Litros",
        "precio por litro": "Precio $/L",
        "total pesos": "Total $",
        "rendimiento (km/l)": "Rendimiento Excel (km/l)",
        "comentarios": "Comentarios",
        "notas": "Comentarios",
        "notas o comentarios": "Comentarios",
    }
    renombre = {}
    for col in crudo.columns:
        clave = _sin_acentos(str(col)).casefold().strip()
        if clave in alias:
            renombre[col] = alias[clave]
    df = crudo.rename(columns=renombre)
    if "Fecha" not in df.columns or "No. económico" not in df.columns:
        return pd.DataFrame()
    df["_fila_excel"] = df.index + 5
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["No. económico"] = df["No. económico"].astype(str).str.strip()
    df = df[df["Fecha"].notna()]
    df = df[~df["No. económico"].str.casefold().isin({"", "nan", "none"})]
    for col in ("Km inicial", "Km actual", "Km recorrido", "Litros", "Precio $/L", "Total $", "Rendimiento Excel (km/l)"):
        if col not in df.columns:
            df[col] = 0.0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    if "Combustible" in df.columns:
        df["Combustible"] = df["Combustible"].map(normalizar_tipo_combustible)
    else:
        df["Combustible"] = ""
    if "Comentarios" not in df.columns:
        df["Comentarios"] = ""
    else:
        df["Comentarios"] = df["Comentarios"].fillna("").astype(str).str.strip()
        df.loc[df["Comentarios"].str.casefold().isin({"nan", "none"}), "Comentarios"] = ""
    con_precio = (df["Litros"] > 0) & (df["Precio $/L"] > 0)
    if con_precio.any():
        df.loc[con_precio, "Total $"] = [
            costo_total_combustible(litros, precio)
            for litros, precio in zip(df.loc[con_precio, "Litros"], df.loc[con_precio, "Precio $/L"])
        ]
    km0 = df["Km inicial"]
    litros = df["Litros"]
    recorrido = df["Km recorrido"]
    real = pd.Series(index=df.index, dtype="float64")
    mask = (km0 > 0) & (litros > 0) & (recorrido > 0)
    real.loc[mask] = (recorrido.loc[mask] / litros.loc[mask]).round(2)
    df["Rendimiento (km/l)"] = real
    if "Item" in df.columns:
        df["Item"] = pd.to_numeric(df["Item"], errors="coerce")
    return df.sort_values(["Fecha", "No. económico"], ignore_index=True)


@st.cache_data(ttl=60, show_spinner="Leyendo control de combustible…")
def cargar_control_combustible(ruta: str, mtime: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    _ = mtime
    return cargar_registro_combustible(ruta), cargar_parque_vehicular(ruta)


def filtrar_registro_combustible(
    df: pd.DataFrame,
    unidades: list[str] | None,
    inicio: date | None,
    fin: date | None,
    combustible: str = "Todos",
) -> pd.DataFrame:
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    vista = df.copy()
    if unidades:
        vista = vista[vista["No. económico"].isin(unidades)]
    if inicio:
        vista = vista[vista["Fecha"] >= pd.Timestamp(inicio)]
    if fin:
        vista = vista[vista["Fecha"] <= pd.Timestamp(fin) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)]
    if combustible and combustible != "Todos":
        vista = vista[vista["Combustible"] == combustible]
    return vista.reset_index(drop=True)


def resumen_semanal_combustible(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    base = df.copy()
    base["Semana"] = base["Fecha"].dt.to_period("W-SUN").apply(lambda p: p.start_time.date())
    filas = []
    for (semana, unidad), bloque in base.groupby(["Semana", "No. económico"], dropna=False):
        fin_sem = semana + pd.Timedelta(days=6)
        km_real = float(bloque.loc[bloque["Km inicial"] > 0, "Km recorrido"].sum())
        litros = float(bloque["Litros"].sum())
        real = bloque["Rendimiento (km/l)"].dropna()
        filas.append(
            {
                "Semana": f"{semana.strftime('%d/%m/%Y')} — {fin_sem.strftime('%d/%m/%Y')}",
                "Inicio": semana,
                "No. económico": unidad,
                "Cargas": int(len(bloque)),
                "Km recorridos": round(km_real, 1),
                "Litros": round(litros, 2),
                "Gasto $": round(float(bloque["Total $"].sum()), 2),
                "Rendimiento (km/l)": round(float(real.mean()), 2) if not real.empty else None,
                "Diesel L": round(float(bloque.loc[bloque["Combustible"] == "Diesel", "Litros"].sum()), 2),
                "Gasolina L": round(float(bloque.loc[bloque["Combustible"] == "Gasolina", "Litros"].sum()), 2),
            }
        )
    return pd.DataFrame(filas).sort_values(["Inicio", "No. económico"], ignore_index=True)


def resumen_mensual_combustible(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    base = df.copy()
    base["Mes"] = base["Fecha"].dt.to_period("M")
    filas = []
    for (mes, unidad), bloque in base.groupby(["Mes", "No. económico"], dropna=False):
        km_real = float(bloque.loc[bloque["Km inicial"] > 0, "Km recorrido"].sum())
        litros = float(bloque["Litros"].sum())
        real = bloque["Rendimiento (km/l)"].dropna()
        filas.append(
            {
                "Mes": etiqueta_mes(f"{mes.year:04d}-{mes.month:02d}"),
                "Periodo": str(mes),
                "No. económico": unidad,
                "Cargas": int(len(bloque)),
                "Km recorridos": round(km_real, 1),
                "Litros": round(litros, 2),
                "Gasto $": round(float(bloque["Total $"].sum()), 2),
                "Rendimiento (km/l)": round(float(real.mean()), 2) if not real.empty else None,
            }
        )
    return pd.DataFrame(filas).sort_values(["Periodo", "No. económico"], ignore_index=True)


def gasto_registros_combustible(df: pd.DataFrame, referencia: pd.DataFrame | None = None) -> float:
    """Suma el dinero del periodo: Total $ del registro, o litros por precio unitario."""
    if df is None or df.empty or "Litros" not in df.columns:
        return 0.0
    litros = pd.to_numeric(df["Litros"], errors="coerce").fillna(0.0)
    if "Total $" in df.columns:
        total = pd.to_numeric(df["Total $"], errors="coerce").fillna(0.0)
    else:
        total = pd.Series(0.0, index=df.index)
    if "Precio $/L" in df.columns:
        precio = pd.to_numeric(df["Precio $/L"], errors="coerce").fillna(0.0)
    else:
        precio = pd.Series(0.0, index=df.index)
    producto = pd.Series(
        [
            costo_total_combustible(litros.at[idx], precio.at[idx]) if litros.at[idx] > 0 and precio.at[idx] > 0 else 0.0
            for idx in df.index
        ],
        index=df.index,
    )
    monto = total.where(total > 0, producto)
    faltante = (monto <= 0) & (litros > 0)
    if faltante.any():
        base = referencia if referencia is not None and not referencia.empty else df
        for idx in df.index[faltante]:
            tipo = normalizar_tipo_combustible(df.at[idx, "Combustible"]) if "Combustible" in df.columns else ""
            unitario = ultimo_precio_combustible(base, tipo) if tipo in {"Diesel", "Gasolina"} else 0.0
            if unitario > 0:
                monto.at[idx] = costo_total_combustible(litros.at[idx], unitario)
    return redondear_mitad_arriba(float(monto.sum()), 2)


def ultima_fecha_combustible(df: pd.DataFrame) -> date | None:
    if df is None or df.empty or "Fecha" not in getattr(df, "columns", []):
        return None
    fechas = pd.to_datetime(df["Fecha"], errors="coerce").dropna()
    if fechas.empty:
        return None
    valor = fechas.max()
    if isinstance(valor, pd.Timestamp):
        return valor.date()
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return None


def metricas_combustible(df: pd.DataFrame) -> dict[str, float]:
    if df is None or df.empty:
        return {"cargas": 0, "unidades": 0, "km": 0.0, "litros": 0.0, "gasto": 0.0, "rend": 0.0}
    real = df["Rendimiento (km/l)"].dropna() if "Rendimiento (km/l)" in df.columns else pd.Series(dtype=float)
    return {
        "cargas": int(len(df)),
        "unidades": int(df["No. económico"].nunique()),
        "km": float(df.loc[df["Km inicial"] > 0, "Km recorrido"].sum()) if "Km inicial" in df.columns else 0.0,
        "litros": float(df["Litros"].sum()),
        "gasto": gasto_registros_combustible(df),
        "rend": float(real.mean()) if not real.empty else 0.0,
    }


def unidades_combustible(diario: pd.DataFrame, parque: pd.DataFrame) -> list[str]:
    vistos: set[str] = set()
    for origen in (parque, diario):
        if origen is None or origen.empty or "No. económico" not in origen.columns:
            continue
        for valor in origen["No. económico"].dropna().astype(str).str.strip():
            if valor and valor.casefold() not in {"nan", "none"}:
                vistos.add(valor)
    return sorted(vistos, key=lambda u: (u.upper(), u))


def ultimo_odometro_unidad(diario: pd.DataFrame, unidad: str) -> float:
    if diario is None or diario.empty or not unidad:
        return 0.0
    bloque = diario[diario["No. económico"].astype(str).str.strip() == str(unidad).strip()]
    if bloque.empty:
        return 0.0
    orden = ["Fecha"] + (["Item"] if "Item" in bloque.columns else [])
    ultimo = bloque.sort_values(orden, na_position="last").iloc[-1]
    return _num_combustible(ultimo.get("Km actual"))


def ultimo_precio_combustible(diario: pd.DataFrame, tipo: str) -> float:
    base = 27.0 if tipo == "Diesel" else 23.98
    if diario is None or diario.empty:
        return base
    bloque = diario[diario["Combustible"] == tipo]
    if bloque.empty:
        return base
    precio = _num_combustible(bloque.sort_values("Fecha").iloc[-1].get("Precio $/L"))
    return precio if precio > 0 else base


def combustible_asignado_unidad(parque: pd.DataFrame, unidad: str) -> str:
    if parque is None or parque.empty or not unidad:
        return "Gasolina"
    fila = parque[parque["No. económico"].astype(str).str.strip() == str(unidad).strip()]
    if fila.empty:
        return "Gasolina"
    tipo = str(fila.iloc[0].get("Combustible asignado") or "").strip()
    return tipo if tipo in {"Gasolina", "Diesel"} else "Gasolina"


def es_unidad_otros(unidad: str) -> bool:
    return str(unidad or "").strip().casefold() == "otros"


def calcular_carga_combustible(km_inicial: float, km_actual: float, litros: float, precio: float) -> tuple[float, float, float]:
    recorrido = max(0.0, _num_combustible(km_actual) - _num_combustible(km_inicial))
    total = costo_total_combustible(litros, precio)
    rendimiento = round(recorrido / float(litros), 4) if _num_combustible(litros) > 0 else 0.0
    return round(recorrido, 2), total, rendimiento


def _primera_fila_libre_registro(ws) -> int:
    ultima = max(int(ws.max_row or 4), 4)
    for fila in range(5, ultima + 1):
        if ws.cell(fila, 3).value in (None, ""):
            return fila
    return ultima + 1


def _copiar_estilo_fila(ws, origen: int, destino: int, columnas: int = 11) -> None:
    from copy import copy

    if origen < 5 or destino == origen:
        return
    for col in range(1, columnas + 1):
        src = ws.cell(origen, col)
        dst = ws.cell(destino, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)


def _guardar_libro_excel(wb, ruta: str) -> None:
    tmp = os.path.join(tempfile.gettempdir(), "pjp4_combustible_write.xlsx")
    wb.save(tmp)
    try:
        persistencia.copiar_seguro(tmp, ruta)
    except PermissionError:
        alt = os.path.join(os.path.dirname(ruta), "~pjp4_combustible_write.xlsx")
        persistencia.copiar_seguro(tmp, alt)
        persistencia.reemplazar_seguro(alt, ruta)


def _escribir_comentarios_carga(ws, fila: int, comentarios: str) -> None:
    if ws.cell(4, 12).value in (None, ""):
        ws.cell(4, 12, "Comentarios")
    ws.cell(fila, 12, comentarios or None)


def _escribir_carga_en_fila(
    ws,
    fila: int,
    fecha_carga: date,
    unidad: str,
    km_inicial: float,
    km_actual: float,
    tipo: str,
    litros: float,
    precio: float,
    recorrido: float,
    total: float,
    rendimiento: float,
    comentarios: str,
    *,
    escribir_item: bool,
) -> None:
    if escribir_item:
        ws.cell(fila, 1, fila - 4)
    ws.cell(fila, 2, datetime.combine(fecha_carga, datetime.min.time()))
    ws.cell(fila, 2).number_format = "DD/MM/YYYY"
    ws.cell(fila, 3, unidad)
    ws.cell(fila, 4, km_inicial)
    ws.cell(fila, 5, km_actual)
    ws.cell(fila, 6, recorrido)
    ws.cell(fila, 7, tipo)
    ws.cell(fila, 8, litros)
    ws.cell(fila, 9, precio)
    ws.cell(fila, 10, total)
    ws.cell(fila, 11, rendimiento)
    if es_unidad_otros(unidad) or comentarios:
        _escribir_comentarios_carga(ws, fila, comentarios)
    elif ws.cell(fila, 12).value not in (None, ""):
        ws.cell(fila, 12, None)


def guardar_carga_combustible(
    ruta: str,
    fecha_carga: date,
    unidad: str,
    km_inicial: float,
    km_actual: float,
    tipo: str,
    litros: float,
    precio: float,
    recorrido: float,
    total: float,
    rendimiento: float,
    comentarios: str = "",
) -> None:
    from openpyxl import load_workbook

    origen = _copia_excel_legible(ruta)
    wb = load_workbook(origen)
    try:
        if "Registro Diario" not in wb.sheetnames:
            raise ValueError("El archivo no tiene la hoja Registro Diario.")
        ws = wb["Registro Diario"]
        fila = _primera_fila_libre_registro(ws)
        _copiar_estilo_fila(ws, max(5, fila - 1), fila)
        _escribir_carga_en_fila(
            ws,
            fila,
            fecha_carga,
            unidad,
            km_inicial,
            km_actual,
            tipo,
            litros,
            precio,
            recorrido,
            total,
            rendimiento,
            comentarios,
            escribir_item=True,
        )
        _guardar_libro_excel(wb, ruta)
    finally:
        wb.close()


def actualizar_carga_combustible(
    ruta: str,
    fila_excel: int,
    unidad_original: str,
    fecha_carga: date,
    unidad: str,
    km_inicial: float,
    km_actual: float,
    tipo: str,
    litros: float,
    precio: float,
    recorrido: float,
    total: float,
    rendimiento: float,
    comentarios: str = "",
) -> None:
    from openpyxl import load_workbook

    if int(fila_excel) < 5:
        raise ValueError("El renglón a corregir no es válido.")
    origen = _copia_excel_legible(ruta)
    wb = load_workbook(origen)
    try:
        if "Registro Diario" not in wb.sheetnames:
            raise ValueError("El archivo no tiene la hoja Registro Diario.")
        ws = wb["Registro Diario"]
        actual = str(ws.cell(int(fila_excel), 3).value or "").strip()
        esperado = str(unidad_original or "").strip()
        if actual and esperado and actual != esperado:
            raise ValueError(
                "Ese renglón ya no coincide con el registro seleccionado. Recarga la página e inténtalo de nuevo."
            )
        _escribir_carga_en_fila(
            ws,
            int(fila_excel),
            fecha_carga,
            unidad,
            km_inicial,
            km_actual,
            tipo,
            litros,
            precio,
            recorrido,
            total,
            rendimiento,
            comentarios,
            escribir_item=False,
        )
        _guardar_libro_excel(wb, ruta)
    finally:
        wb.close()


def _ultima_fila_catalogo(ws) -> int:
    ultima = 1
    for fila in range(2, (ws.max_row or 1) + 1):
        if str(ws.cell(fila, 1).value or "").strip():
            ultima = fila
    return ultima


def _extender_listas_combustible(wb, ultima_fila: int) -> None:
    if ultima_fila < 2:
        return
    definiciones = getattr(wb, "defined_names", None)
    if definiciones is not None and "Lista_Vehiculos" in definiciones:
        definiciones["Lista_Vehiculos"].attr_text = f"Listas!$A$2:$A${ultima_fila}"
    if "Registro Diario" not in wb.sheetnames:
        return
    patron = re.compile(r"Listas!\$A\$2:\$C\$\d+")
    nuevo = f"Listas!$A$2:$C${ultima_fila}"
    ws = wb["Registro Diario"]
    for fila in ws.iter_rows(min_row=5, max_row=ws.max_row or 5, max_col=12):
        for celda in fila:
            valor = celda.value
            if isinstance(valor, str) and patron.search(valor):
                celda.value = patron.sub(nuevo, valor)


def agregar_unidad_parque(ruta: str, numero: str, modelo: str, combustible: str) -> None:
    from openpyxl import load_workbook

    numero = str(numero or "").strip()
    modelo = str(modelo or "").strip()
    combustible = normalizar_tipo_combustible(combustible)
    if not numero or not modelo:
        raise ValueError("Escribe el número económico y el modelo.")
    if combustible not in {"Gasolina", "Diesel"}:
        raise ValueError("El tipo de combustible debe ser Gasolina o Diesel.")
    if numero.casefold() in {"numero_economico", "número económico", "nan", "none"}:
        raise ValueError("Ese número económico no es válido.")

    origen = _copia_excel_legible(ruta)
    wb = load_workbook(origen)
    try:
        if "Listas" not in wb.sheetnames:
            raise ValueError("El archivo no tiene la hoja Listas.")
        ws = wb["Listas"]
        for fila in range(2, (ws.max_row or 1) + 1):
            existente = str(ws.cell(fila, 1).value or "").strip()
            if existente.casefold() == numero.casefold():
                raise ValueError(f"{numero} ya está en el catálogo.")
        destino = _ultima_fila_catalogo(ws) + 1
        if ws.cell(1, 4).value in (None, ""):
            ws.cell(1, 4, "Modelo")
        ws.cell(destino, 1, numero)
        ws.cell(destino, 2, combustible)
        ws.cell(destino, 3, combustible)
        ws.cell(destino, 4, modelo)
        _extender_listas_combustible(wb, destino)
        _guardar_libro_excel(wb, ruta)
    finally:
        wb.close()


def _opciones_unidad_combustible(unidades: list[str]) -> list[str]:
    opciones = [str(u).strip() for u in unidades if str(u).strip()]
    if not any(es_unidad_otros(u) for u in opciones):
        opciones.append("otros")
    return opciones


def _etiqueta_carga_combustible(fila) -> str:
    fecha = fila.get("Fecha")
    if isinstance(fecha, pd.Timestamp):
        fecha_txt = fecha.strftime("%d/%m/%Y")
    elif isinstance(fecha, datetime):
        fecha_txt = fecha.strftime("%d/%m/%Y")
    elif isinstance(fecha, date):
        fecha_txt = fecha.strftime("%d/%m/%Y")
    else:
        fecha_txt = str(fecha or "Sin fecha")
    item = fila.get("Item")
    try:
        item_txt = str(int(item))
    except (TypeError, ValueError):
        item_txt = "—"
    unidad = str(fila.get("No. económico") or "").strip()
    litros = _num_combustible(fila.get("Litros"))
    total = _num_combustible(fila.get("Total $"))
    nota = str(fila.get("Comentarios") or "").strip()
    extra = f" · {nota}" if nota else ""
    return f"Item {item_txt} · {fecha_txt} · {unidad} · {litros:.3f} L · ${total:,.2f}{extra}"


def render_alta_unidad_combustible(parque: pd.DataFrame) -> None:
    with st.expander("➕ Dar de alta unidad", expanded=False):
        st.caption(
            "Número económico, modelo y combustible quedan en la hoja Listas "
            "y se agregan solos a los selectores."
        )
        fid = int(st.session_state.get("alta_unidad_id", 0))
        c1, c2, c3 = st.columns(3)
        with c1:
            numero = st.text_input("Número económico", key=f"alta_unidad_num_{fid}").strip()
        with c2:
            modelo = st.text_input("Modelo", key=f"alta_unidad_modelo_{fid}").strip()
        with c3:
            combustible = st.selectbox(
                "Tipo de combustible",
                ["Gasolina", "Diesel"],
                key=f"alta_unidad_tipo_{fid}",
            )
        if parque is not None and not parque.empty and numero:
            ya = parque["No. económico"].astype(str).str.strip().str.casefold()
            if numero.casefold() in set(ya):
                st.caption(f"{numero} ya está en el padrón.")
        if st.button("Guardar unidad", type="primary", key=f"alta_unidad_ok_{fid}"):
            try:
                agregar_unidad_parque(COMBUSTIBLE_FILE, numero, modelo, combustible)
            except PermissionError:
                st.error("No se pudo escribir el Excel. Ciérralo en Excel u OneDrive e inténtalo de nuevo.")
            except Exception as exc:
                st.error(str(exc))
            else:
                cargar_control_combustible.clear()
                st.session_state.alta_unidad_id = fid + 1
                st.session_state.comb_just_saved = (
                    f"Unidad {numero} dada de alta ({modelo}, {combustible}). Ya aparece en los selectores."
                )
                st.rerun()


def render_correccion_carga_combustible(diario: pd.DataFrame, parque: pd.DataFrame) -> None:
    with st.expander("✏️ Corregir registro guardado", expanded=False):
        st.caption("Corrige una carga ya guardada si el personal capturó mal litros, kilometraje o la unidad.")
        if diario is None or diario.empty or "_fila_excel" not in diario.columns:
            st.info("No hay cargas para corregir.")
            return
        busqueda = st.text_input(
            "Buscar registro",
            placeholder="Unidad, fecha, item o comentario",
            key="comb_corr_buscar",
        ).strip().casefold()
        orden = diario.sort_values(
            ["Fecha", "Item"] if "Item" in diario.columns else ["Fecha"],
            ascending=False,
        )
        opciones: list[tuple[int, str]] = []
        filas: dict[int, pd.Series] = {}
        for _, fila in orden.iterrows():
            etiqueta = _etiqueta_carga_combustible(fila)
            if busqueda and busqueda not in etiqueta.casefold():
                continue
            fila_excel = int(_num_combustible(fila.get("_fila_excel")))
            if fila_excel < 5 or fila_excel in filas:
                continue
            opciones.append((fila_excel, etiqueta))
            filas[fila_excel] = fila
        if not opciones:
            st.info("Ningún registro coincide con la búsqueda.")
            return
        etiquetas = [etiqueta for _, etiqueta in opciones]
        indice = st.selectbox(
            "Registro",
            range(len(opciones)),
            format_func=lambda i: etiquetas[i],
            key="comb_corr_sel",
        )
        fila_excel, _etiqueta = opciones[int(indice)]
        fila = filas[fila_excel]
        fecha_base = fila.get("Fecha")
        if isinstance(fecha_base, pd.Timestamp):
            fecha_base = fecha_base.date()
        elif isinstance(fecha_base, datetime):
            fecha_base = fecha_base.date()
        if not isinstance(fecha_base, date):
            fecha_base = date.today()
        unidades = _opciones_unidad_combustible(unidades_combustible(diario, parque))
        unidad_actual = str(fila.get("No. económico") or "").strip()
        if unidad_actual and unidad_actual not in unidades:
            unidades = [unidad_actual] + unidades
        rid = fila_excel
        c1, c2, c3 = st.columns(3)
        with c1:
            fecha_carga = st.date_input(
                "Fecha",
                value=fecha_base,
                format="DD/MM/YYYY",
                key=f"comb_edit_fecha_{rid}",
            )
            unidad = st.selectbox(
                "Número económico",
                unidades,
                index=unidades.index(unidad_actual) if unidad_actual in unidades else 0,
                key=f"comb_edit_unidad_{rid}",
            )
        with c2:
            km_inicial = st.number_input(
                "Kilometraje inicial",
                min_value=0.0,
                value=redondear_mitad_arriba(fila.get("Km inicial"), 1),
                step=0.1,
                format="%.1f",
                key=f"comb_edit_km0_{rid}",
            )
            km_actual = st.number_input(
                "Kilometraje actual",
                min_value=0.0,
                value=redondear_mitad_arriba(fila.get("Km actual"), 1),
                step=0.1,
                format="%.1f",
                key=f"comb_edit_km_{rid}",
            )
        tipo_actual = normalizar_tipo_combustible(fila.get("Combustible")) or "Gasolina"
        with c3:
            tipo = st.selectbox(
                "Tipo de combustible",
                ["Gasolina", "Diesel"],
                index=0 if tipo_actual != "Diesel" else 1,
                key=f"comb_edit_tipo_{rid}",
            )
            litros = st.number_input(
                "Litros cargados",
                min_value=0.0,
                value=redondear_mitad_arriba(fila.get("Litros"), 3),
                step=0.001,
                format="%.3f",
                key=f"comb_edit_litros_{rid}",
            )
            precio = st.number_input(
                "Precio por litro",
                min_value=0.0,
                value=redondear_mitad_arriba(fila.get("Precio $/L"), 2),
                step=0.01,
                format="%.2f",
                key=f"comb_edit_precio_{rid}",
            )
        comentarios = ""
        if es_unidad_otros(unidad):
            comentarios = st.text_area(
                "Notas o comentarios",
                value=str(fila.get("Comentarios") or ""),
                key=f"comb_edit_notas_{rid}",
            ).strip()
        recorrido, total, rendimiento = calcular_carga_combustible(km_inicial, km_actual, litros, precio)
        m1, m2, m3 = st.columns(3)
        m1.metric("Km recorrido", f"{recorrido:,.1f}")
        m2.metric("Total pesos", f"${total:,.2f}")
        m3.metric("Rendimiento", f"{rendimiento:.2f} km/l")
        if st.button("Guardar corrección", type="primary", key=f"comb_edit_ok_{rid}"):
            if litros <= 0:
                st.warning("Captura los litros cargados.")
            elif precio <= 0:
                st.warning("Captura el precio por litro.")
            elif km_inicial > 0 and km_actual < km_inicial:
                st.warning(
                    f"El kilometraje actual ({km_actual:,.1f}) no puede ser menor que el inicial "
                    f"({km_inicial:,.1f})."
                )
            else:
                try:
                    actualizar_carga_combustible(
                        COMBUSTIBLE_FILE,
                        fila_excel,
                        unidad_actual,
                        fecha_carga,
                        unidad,
                        km_inicial,
                        km_actual,
                        tipo,
                        litros,
                        precio,
                        recorrido,
                        total,
                        rendimiento,
                        comentarios,
                    )
                except PermissionError:
                    st.error("No se pudo escribir el Excel. Ciérralo en Excel u OneDrive e inténtalo de nuevo.")
                except Exception as exc:
                    st.error(f"No se pudo corregir el registro: {exc}")
                else:
                    cargar_control_combustible.clear()
                    st.session_state.comb_just_saved = (
                        f"Registro de {unidad} corregido: {recorrido:,.1f} km, {litros:g} L, ${total:,.2f}."
                    )
                    st.rerun()


def render_formulario_carga_combustible(diario: pd.DataFrame, parque: pd.DataFrame) -> None:
    unidades = _opciones_unidad_combustible(unidades_combustible(diario, parque))

    with st.expander("➕ Capturar nueva carga", expanded=st.session_state.get("comb_form_abierto", False)):
        st.caption(
            "Se calcula solo el km recorrido, el total en pesos y el rendimiento. "
            f"El renglón se agrega a {os.path.basename(COMBUSTIBLE_FILE)}."
        )
        fid = st.session_state.get("comb_form_id", 0)
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            fecha_carga = st.date_input(
                "Fecha",
                value=date.today(),
                format="DD/MM/YYYY",
                key=f"comb_cap_fecha_{fid}",
            )
            unidad = st.selectbox(
                "Número económico",
                unidades,
                key=f"comb_cap_unidad_{fid}",
            )
        tipo_default = combustible_asignado_unidad(parque, unidad)
        km_inicial = ultimo_odometro_unidad(diario, unidad)
        with col_b:
            km_actual = st.number_input(
                "Kilometraje actual",
                min_value=0.0,
                value=redondear_mitad_arriba(km_inicial, 1) if km_inicial > 0 else 0.0,
                step=0.1,
                format="%.1f",
                key=f"comb_cap_km_{fid}_{unidad}",
            )
            tipo = st.selectbox(
                "Tipo de combustible",
                ["Gasolina", "Diesel"],
                index=0 if tipo_default == "Gasolina" else 1,
                key=f"comb_cap_tipo_{fid}_{unidad}",
            )
        with col_c:
            litros = st.number_input(
                "Litros cargados",
                min_value=0.0,
                value=0.0,
                step=0.001,
                format="%.3f",
                key=f"comb_cap_litros_{fid}",
            )
            precio = st.number_input(
                "Precio por litro",
                min_value=0.0,
                value=float(ultimo_precio_combustible(diario, tipo)),
                step=0.01,
                format="%.2f",
                key=f"comb_cap_precio_{fid}_{tipo}",
            )

        comentarios = ""
        if es_unidad_otros(unidad):
            comentarios = st.text_area(
                "Notas o comentarios",
                placeholder="Describe la unidad, el equipo o el motivo de esta carga.",
                key=f"comb_cap_notas_{fid}",
            ).strip()

        recorrido, total, rendimiento = calcular_carga_combustible(km_inicial, km_actual, litros, precio)
        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Km inicial (última carga)", f"{km_inicial:,.1f}")
        p2.metric("Km recorrido", f"{recorrido:,.1f}")
        p3.metric("Total pesos", f"${total:,.2f}")
        p4.metric("Rendimiento", f"{rendimiento:.2f} km/l")

        if st.button("💾 Guardar Carga", type="primary", key=f"comb_cap_guardar_{fid}"):
            if not unidad:
                st.warning("Selecciona el número económico.")
            elif litros <= 0:
                st.warning("Captura los litros cargados.")
            elif precio <= 0:
                st.warning("Captura el precio por litro.")
            elif km_inicial > 0 and km_actual < km_inicial:
                st.warning(
                    f"El kilometraje actual ({km_actual:,.1f}) no puede ser menor que el último odómetro "
                    f"de {unidad} ({km_inicial:,.1f})."
                )
            else:
                try:
                    guardar_carga_combustible(
                        COMBUSTIBLE_FILE,
                        fecha_carga,
                        unidad,
                        km_inicial,
                        km_actual,
                        tipo,
                        litros,
                        precio,
                        recorrido,
                        total,
                        rendimiento,
                        comentarios,
                    )
                    cargar_control_combustible.clear()
                    st.session_state.comb_just_saved = (
                        f"Carga de {unidad} guardada: {recorrido:,.1f} km, {litros:g} L, "
                        f"${total:,.2f}, {rendimiento:.2f} km/l."
                    )
                    st.session_state.comb_form_id = fid + 1
                    st.session_state.comb_form_abierto = False
                    st.rerun()
                except PermissionError:
                    st.error(
                        "No se pudo escribir el Excel. Ciérralo en Excel u OneDrive e inténtalo de nuevo."
                    )
                except Exception as exc:
                    st.error(f"No se pudo guardar la carga: {exc}")


def render_control_combustible() -> None:
    st.markdown('<div class="section-title">Control de Combustible</div>', unsafe_allow_html=True)
    st.caption(
        "Auditoría de cargas, kilometraje y rendimiento de la flota a partir de "
        f"{COMBUSTIBLE_FILE}."
    )
    if not os.path.exists(COMBUSTIBLE_FILE):
        st.warning(
            f"No se encontró {COMBUSTIBLE_FILE}. "
            "Coloca el Excel en la carpeta control_combustible/ y recarga."
        )
        return

    mtime = os.path.getmtime(COMBUSTIBLE_FILE)
    diario, parque = cargar_control_combustible(COMBUSTIBLE_FILE, mtime)
    actualizado = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")
    st.caption(f"Archivo leído: {os.path.basename(COMBUSTIBLE_FILE)} · actualizado {actualizado}")
    aviso = st.session_state.pop("comb_just_saved", None)
    if aviso:
        st.success(aviso)

    render_alta_unidad_combustible(parque)
    render_formulario_carga_combustible(diario, parque)

    if diario.empty:
        st.info("El registro diario no tiene cargas capturadas. Usa el formulario de arriba para registrar la primera.")
        if not parque.empty:
            st.dataframe(parque, use_container_width=True, hide_index=True)
        return

    render_correccion_carga_combustible(diario, parque)

    unidades = unidades_combustible(diario, parque)
    combustibles = ["Todos"] + sorted({c for c in diario["Combustible"] if c})
    fmin, fmax = diario["Fecha"].min().date(), diario["Fecha"].max().date()
    fmax = max(fmax, date.today())

    c1, c2, c3 = st.columns([1.3, 1.8, 1])
    with c1:
        sel_unidades = st.multiselect(
            "Número económico",
            unidades,
            default=[],
            placeholder="Todas las unidades",
            key="comb_unidades",
        )
    with c2:
        modo_fecha = st.radio(
            "Filtro de fecha",
            ["Día específico", "Rango de fechas"],
            index=1,
            horizontal=True,
            key="comb_modo_fecha",
        )
        if modo_fecha == "Día específico":
            dia = min(max(date.today(), fmin), fmax)
            dia_sel = st.date_input(
                "Día",
                value=dia,
                min_value=fmin,
                max_value=fmax,
                format="DD/MM/YYYY",
                key="comb_dia",
            )
            inicio, fin = dia_sel, dia_sel
        else:
            rango = st.date_input(
                "Rango de fechas",
                value=(fmin, fmax),
                min_value=fmin,
                max_value=fmax,
                format="DD/MM/YYYY",
                key="comb_rango",
            )
            if isinstance(rango, (list, tuple)) and len(rango) == 2:
                inicio, fin = rango[0], rango[1]
            else:
                inicio, fin = fmin, fmax
    with c3:
        tipo_comb = st.selectbox("Combustible", combustibles, key="comb_tipo")

    vista = filtrar_registro_combustible(diario, sel_unidades, inicio, fin, tipo_comb)
    met = metricas_combustible(vista)
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Cargas", int(met["cargas"]))
    k2.metric("Unidades", int(met["unidades"]))
    k3.metric("Km recorridos", f"{met['km']:,.0f}")
    k4.metric("Litros", f"{met['litros']:,.1f}")
    k5.metric("Gasto", f"${met['gasto']:,.2f}")
    if met["rend"]:
        st.caption(
            f"Rendimiento promedio real del recorte: **{met['rend']:.2f} km/l**. "
            "Las primeras cargas (odómetro inicial en 0) no entran al promedio."
        )

    semanal = resumen_semanal_combustible(vista)
    mensual = resumen_mensual_combustible(vista)
    tab_d, tab_s, tab_m, tab_p = st.tabs(
        ["Registro Diario", "Resumen Semanal", "Resumen Mensual", "Parque vehicular"]
    )
    cols_diario = [
        "Fecha",
        "No. económico",
        "Combustible",
        "Km inicial",
        "Km actual",
        "Km recorrido",
        "Litros",
        "Precio $/L",
        "Total $",
        "Rendimiento (km/l)",
        "Comentarios",
    ]
    with tab_d:
        st.caption("Cargas del Excel (Registro Diario). El rendimiento real omite altas de odómetro.")
        if vista.empty:
            st.info("No hay cargas con los filtros seleccionados.")
        else:
            tabla = vista.copy()
            tabla["Fecha"] = tabla["Fecha"].dt.strftime("%d/%m/%Y")
            visibles = [c for c in cols_diario if c in tabla.columns]
            if "Comentarios" in visibles:
                notas = tabla["Comentarios"].fillna("").astype(str).str.strip().str.casefold()
                if not notas.isin({"", "nan", "none"}).eq(False).any():
                    visibles = [c for c in visibles if c != "Comentarios"]
            st.dataframe(
                tabla[visibles],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Litros": st.column_config.NumberColumn(format="%.2f"),
                    "Precio $/L": st.column_config.NumberColumn(format="$%.2f"),
                    "Total $": st.column_config.NumberColumn(format="$%.2f"),
                    "Rendimiento (km/l)": st.column_config.NumberColumn(format="%.2f"),
                    "Km recorrido": st.column_config.NumberColumn(format="%.0f"),
                },
            )
    with tab_s:
        st.caption("Consumo semanal por unidad (lunes a domingo), cruzado con litros de diésel y gasolina.")
        if semanal.empty:
            st.info("No hay semanas con cargas en el recorte.")
        else:
            st.dataframe(
                semanal.drop(columns=["Inicio"], errors="ignore"),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Litros": st.column_config.NumberColumn(format="%.2f"),
                    "Gasto $": st.column_config.NumberColumn(format="$%.2f"),
                    "Rendimiento (km/l)": st.column_config.NumberColumn(format="%.2f"),
                },
            )
    with tab_m:
        st.caption("Totales mensuales por número económico: km, litros, gasto y rendimiento promedio.")
        if mensual.empty:
            st.info("No hay meses con cargas en el recorte.")
        else:
            st.dataframe(
                mensual.drop(columns=["Periodo"], errors="ignore"),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Litros": st.column_config.NumberColumn(format="%.2f"),
                    "Gasto $": st.column_config.NumberColumn(format="$%.2f"),
                    "Rendimiento (km/l)": st.column_config.NumberColumn(format="%.2f"),
                },
            )
            g1, g2 = st.columns(2)
            g1.markdown(money_html("Gasto del recorte", met["gasto"]), unsafe_allow_html=True)
            g2.markdown(
                f'<div class="money-card"><div class="k">Litros del recorte</div>'
                f'<h3>{met["litros"]:,.1f} L</h3></div>',
                unsafe_allow_html=True,
            )
    with tab_p:
        st.caption("Padrón de unidades de la hoja Listas (vehículos y equipos auxiliares).")
        if parque.empty:
            st.info("No se pudo leer la hoja Listas.")
        else:
            if vista.empty:
                consumo = pd.DataFrame(columns=["No. económico", "Cargas", "Litros", "Gasto $"])
            else:
                consumo = vista.groupby("No. económico", as_index=False).agg(
                    Cargas=("Litros", "count"),
                    Litros=("Litros", "sum"),
                    **{"Gasto $": ("Total $", "sum")},
                )
            padrones = parque.merge(consumo, on="No. económico", how="left")
            padrones["Cargas"] = padrones["Cargas"].fillna(0).astype(int)
            padrones["Litros"] = padrones["Litros"].fillna(0).round(2)
            padrones["Gasto $"] = padrones["Gasto $"].fillna(0).round(2)
            preferidas = [
                c for c in ("No. económico", "Modelo", "Combustible asignado") if c in padrones.columns
            ]
            padrones = padrones[preferidas + [c for c in padrones.columns if c not in preferidas]]
            st.dataframe(
                padrones,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Litros": st.column_config.NumberColumn(format="%.2f"),
                    "Gasto $": st.column_config.NumberColumn(format="$%.2f"),
                },
            )

    recorte_xlsx = vista.drop(columns=["_fila_excel"], errors="ignore").copy()
    if not recorte_xlsx.empty:
        recorte_xlsx["Fecha"] = recorte_xlsx["Fecha"].dt.strftime("%Y-%m-%d")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        (recorte_xlsx if not recorte_xlsx.empty else pd.DataFrame({"Aviso": ["Sin cargas"]})).to_excel(
            writer, sheet_name="Registro Diario", index=False
        )
        (semanal.drop(columns=["Inicio"], errors="ignore") if not semanal.empty else pd.DataFrame({"Aviso": ["Sin semanas"]})).to_excel(
            writer, sheet_name="Resumen Semanal", index=False
        )
        (mensual.drop(columns=["Periodo"], errors="ignore") if not mensual.empty else pd.DataFrame({"Aviso": ["Sin meses"]})).to_excel(
            writer, sheet_name="Resumen Mensual", index=False
        )
    st.download_button(
        "⬇️ Exportar recorte a Excel",
        data=buf.getvalue(),
        file_name="control_combustible_recorte.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_comb_xlsx",
    )


def vehiculos_default() -> dict:
    return {
        "intervalo_km": INTERVALO_MTO_KM_DEFAULT,
        "umbral_alerta_km": UMBRAL_ALERTA_KM_DEFAULT,
        "unidades": {},
        "ots": [],
        "ingresos": [],
        "asignaciones": [],
    }


def cargar_vehiculos() -> dict:
    datos = leer_json(VEHICULOS_FILE, vehiculos_default())
    if not isinstance(datos, dict):
        return vehiculos_default()
    base = vehiculos_default()
    base["intervalo_km"] = max(1, int(_num_combustible(datos.get("intervalo_km")) or INTERVALO_MTO_KM_DEFAULT))
    base["umbral_alerta_km"] = max(1, int(_num_combustible(datos.get("umbral_alerta_km")) or UMBRAL_ALERTA_KM_DEFAULT))
    unidades = datos.get("unidades") or {}
    base["unidades"] = unidades if isinstance(unidades, dict) else {}
    ots = datos.get("ots") or []
    base["ots"] = [ot for ot in ots if isinstance(ot, dict)] if isinstance(ots, list) else []
    ingresos = datos.get("ingresos") or []
    base["ingresos"] = [ing for ing in ingresos if isinstance(ing, dict)] if isinstance(ingresos, list) else []
    asignaciones = datos.get("asignaciones") or []
    base["asignaciones"] = (
        [asg for asg in asignaciones if isinstance(asg, dict)] if isinstance(asignaciones, list) else []
    )
    return base


def guardar_vehiculos(datos: dict) -> None:
    escribir_json(VEHICULOS_FILE, datos)


def config_unidad_veh(datos: dict, unidad: str) -> dict:
    raw = (datos.get("unidades") or {}).get(unidad) or {}
    if not isinstance(raw, dict):
        raw = {}
    intervalo = int(_num_combustible(raw.get("intervalo_km")) or datos.get("intervalo_km") or INTERVALO_MTO_KM_DEFAULT)
    umbral = int(_num_combustible(raw.get("umbral_alerta_km")) or datos.get("umbral_alerta_km") or UMBRAL_ALERTA_KM_DEFAULT)
    return {
        "intervalo_km": max(1, intervalo),
        "umbral_alerta_km": max(1, umbral),
        "ultimo_mto_km": max(0.0, _num_combustible(raw.get("ultimo_mto_km"))),
    }


def folio_ot_vehicular(datos: dict, anio: int | None = None) -> str:
    anio = anio or date.today().year
    prefijo = f"OT-VEH-{anio}-"
    maximo = 0
    for ot in datos.get("ots") or []:
        folio = str(ot.get("id") or "")
        if folio.startswith(prefijo):
            try:
                maximo = max(maximo, int(folio[len(prefijo):]))
            except ValueError:
                continue
    return f"{prefijo}{maximo + 1:03d}"


def km_diario_unidad(diario: pd.DataFrame, unidad: str, ventana_dias: int = 30) -> float:
    if diario is None or diario.empty or not unidad:
        return 0.0
    bloque = diario[diario["No. económico"].astype(str).str.strip() == str(unidad).strip()]
    real = bloque[bloque["Km inicial"] > 0] if "Km inicial" in bloque.columns else bloque
    if real.empty:
        return 0.0
    fmax = real["Fecha"].max()
    ventana = real[real["Fecha"] >= fmax - pd.Timedelta(days=ventana_dias)]
    if ventana.empty:
        ventana = real
    km = float(ventana["Km recorrido"].sum()) if "Km recorrido" in ventana.columns else 0.0
    if km <= 0:
        return 0.0
    dias = int((ventana["Fecha"].max() - ventana["Fecha"].min()).days)
    return round(km / max(dias, 1), 2)


def proximo_mto_km(odometro: float, ultimo_mto_km: float, intervalo: int) -> float:
    intervalo = max(1, int(intervalo))
    if ultimo_mto_km > 0:
        objetivo = float(ultimo_mto_km) + intervalo
        while odometro > objetivo + 0.5:
            objetivo += intervalo
        return objetivo
    if odometro <= 0:
        return float(intervalo)
    resto = odometro % intervalo
    if resto <= 0.5:
        return float(odometro)
    return float(odometro + (intervalo - resto))


def estado_mto_km(km_restantes: float, umbral: float) -> str:
    if km_restantes <= 0:
        return "vencido"
    if km_restantes <= umbral:
        return "alerta"
    return "ok"


def fecha_estimada_mto(km_restantes: float, km_diario: float, hoy: date | None = None) -> date | None:
    hoy = hoy or date.today()
    if km_restantes <= 0:
        return hoy
    if km_diario <= 0:
        return None
    dias = max(0, int(round(km_restantes / km_diario)))
    return hoy + timedelta(days=dias)


def monitorear_flota(diario: pd.DataFrame, parque: pd.DataFrame, datos: dict) -> list[dict]:
    unidades = unidades_combustible(diario, parque)
    filas = []
    for unidad in unidades:
        cfg = config_unidad_veh(datos, unidad)
        odometro = ultimo_odometro_unidad(diario, unidad)
        objetivo = proximo_mto_km(odometro, cfg["ultimo_mto_km"], cfg["intervalo_km"])
        restantes = round(objetivo - odometro, 1)
        km_dia = km_diario_unidad(diario, unidad)
        estado = estado_mto_km(restantes, cfg["umbral_alerta_km"])
        estimada = fecha_estimada_mto(restantes, km_dia)
        filas.append(
            {
                "unidad": unidad,
                "odometro": odometro,
                "ultimo_mto_km": cfg["ultimo_mto_km"],
                "intervalo_km": cfg["intervalo_km"],
                "umbral_alerta_km": cfg["umbral_alerta_km"],
                "km_objetivo": objetivo,
                "km_restantes": restantes,
                "km_diario": km_dia,
                "estado": estado,
                "fecha_estimada": estimada,
            }
        )
    orden = {"vencido": 0, "alerta": 1, "ok": 2}
    return sorted(filas, key=lambda f: (orden.get(f["estado"], 9), f["km_restantes"], f["unidad"]))


def sincronizar_ots_preventivas(diario: pd.DataFrame, parque: pd.DataFrame, datos: dict) -> dict:
    monitores = monitorear_flota(diario, parque, datos)
    existentes = datos.get("ots") or []
    creadas = 0
    for mon in monitores:
        if mon["estado"] not in {"alerta", "vencido"}:
            continue
        ya = False
        for ot in existentes:
            if (
                ot.get("tipo") == "preventivo"
                and str(ot.get("unidad")) == mon["unidad"]
                and abs(_num_combustible(ot.get("km_objetivo")) - mon["km_objetivo"]) < 1
                and str(ot.get("estatus")) != ESTATUS_OT_CERRADA
            ):
                ot["km_actual"] = mon["odometro"]
                ot["km_restantes"] = mon["km_restantes"]
                ot["km_diario"] = mon["km_diario"]
                ot["fecha_estimada"] = mon["fecha_estimada"].isoformat() if mon["fecha_estimada"] else ""
                ya = True
                break
        if ya:
            continue
        estimada = mon["fecha_estimada"]
        aviso = "vencido" if mon["estado"] == "vencido" else "próximo"
        existentes.append(
            {
                "id": folio_ot_vehicular(datos),
                "tipo": "preventivo",
                "unidad": mon["unidad"],
                "estatus": ESTATUS_OT_ABIERTA,
                "fecha_creacion": date.today().isoformat(),
                "fecha_estimada": estimada.isoformat() if estimada else "",
                "km_actual": mon["odometro"],
                "km_objetivo": mon["km_objetivo"],
                "km_restantes": mon["km_restantes"],
                "km_diario": mon["km_diario"],
                "descripcion": (
                    f"Mantenimiento preventivo {aviso} a los {mon['km_objetivo']:,.0f} km "
                    f"(intervalo {mon['intervalo_km']:,} km)."
                ),
                "falla": "",
                "taller": "",
                "mecanico": "",
                "refacciones": [],
                "costo_mano_obra": 0.0,
                "costo_refacciones": 0.0,
                "costo_total": 0.0,
                "reporte_final": "",
                "fecha_cierre": "",
                "origen": "automatica",
            }
        )
        creadas += 1
        datos["ots"] = existentes
    datos["ots"] = existentes
    if creadas:
        guardar_vehiculos(datos)
    return datos


def costo_ot_vehicular(ot: dict) -> float:
    refs = ot.get("refacciones") or []
    costo_ref = sum(_num_combustible(r.get("costo")) for r in refs if isinstance(r, dict))
    if costo_ref <= 0:
        costo_ref = _num_combustible(ot.get("costo_refacciones"))
    mano = _num_combustible(ot.get("costo_mano_obra"))
    return round(mano + costo_ref, 2)


def ots_por_estatus(datos: dict, estatus: str | None = None, tipo: str | None = None) -> list[dict]:
    salida = []
    for ot in datos.get("ots") or []:
        if estatus and str(ot.get("estatus")) != estatus:
            continue
        if tipo and str(ot.get("tipo")) != tipo:
            continue
        salida.append(ot)
    return sorted(salida, key=lambda o: str(o.get("fecha_creacion") or ""), reverse=True)


def cargar_datos_flota() -> tuple[pd.DataFrame, pd.DataFrame]:
    if not os.path.exists(COMBUSTIBLE_FILE):
        return pd.DataFrame(), pd.DataFrame()
    try:
        return cargar_control_combustible(COMBUSTIBLE_FILE, os.path.getmtime(COMBUSTIBLE_FILE))
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def html_fecha(valor) -> str:
    dia = parsear_fecha_valor(valor)
    return dia.strftime("%d/%m/%Y") if dia else "—"


def etiqueta_estado_mto(estado: str) -> tuple[str, str]:
    if estado == "vencido":
        return "due", "Vencido / en ventana crítica"
    if estado == "alerta":
        return "warn", "Alerta · 500 km o menos"
    return "ok", "Dentro de programa"


def render_ficha_ot_vehicular(ot: dict) -> None:
    tipo = "Preventivo" if ot.get("tipo") == "preventivo" else (
        "Servicio externo" if ot.get("tipo") == "externo" else "Correctivo"
    )
    estatus = str(ot.get("estatus") or "—").replace("_", " ").title()
    refs = ot.get("refacciones") or []
    filas_ref = ""
    if refs:
        for ref in refs:
            if not isinstance(ref, dict):
                continue
            nombre = html.escape(str(ref.get("nombre") or "—"))
            cant = html.escape(str(ref.get("cantidad") or "1"))
            costo = _num_combustible(ref.get("costo"))
            filas_ref += (
                f'<div class="mat-row"><b>{nombre}</b><span>{cant} · ${costo:,.2f}</span></div>'
            )
    else:
        filas_ref = '<div class="mat-row"><b>Sin refacciones capturadas</b><span>—</span></div>'
    st.markdown(
        f"""
        <div class="ot-sheet">
            <div class="ot-kicker">{html.escape(str(ot.get("id") or "OT"))} · {tipo} · {html.escape(estatus)}</div>
            <h3>{html.escape(str(ot.get("unidad") or "Unidad"))}</h3>
            <div class="ot-grid">
                <div class="ot-field"><div class="k">Fecha de alta</div>
                <div class="v">{html.escape(html_fecha(ot.get("fecha_creacion")))}</div></div>
                <div class="ot-field"><div class="k">Fecha estimada / programada</div>
                <div class="v">{html.escape(html_fecha(ot.get("fecha_estimada")))}</div></div>
                <div class="ot-field"><div class="k">Odómetro al generar</div>
                <div class="v">{_num_combustible(ot.get("km_actual")):,.0f} km</div></div>
                <div class="ot-field"><div class="k">Km objetivo / servicio</div>
                <div class="v">{_num_combustible(ot.get("km_objetivo")):,.0f} km</div></div>
                <div class="ot-field"><div class="k">Taller</div>
                <div class="v">{html.escape(str(ot.get("taller") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Mecánico responsable</div>
                <div class="v">{html.escape(str(ot.get("mecanico") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Costo total</div>
                <div class="v">${costo_ot_vehicular(ot):,.2f}</div></div>
                <div class="ot-field"><div class="k">Cierre</div>
                <div class="v">{html.escape(html_fecha(ot.get("fecha_cierre")))}</div></div>
            </div>
            <div class="ot-field"><div class="k">Descripción / falla</div>
            <div class="v">{html.escape(str(ot.get("falla") or ot.get("descripcion") or "—"))}</div></div>
            <div class="ot-field" style="margin-top:0.65rem"><div class="k">Refacciones</div></div>
            {filas_ref}
            <div class="ot-field" style="margin-top:0.55rem"><div class="k">Reporte final</div>
            <div class="v">{html.escape(str(ot.get("reporte_final") or "Pendiente de cierre."))}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _cerrar_ot_vehicular(datos: dict, folio: str, mecanico: str, taller: str, reporte: str, costo_mo: float, refs: list[dict]) -> None:
    for ot in datos.get("ots") or []:
        if str(ot.get("id")) != folio:
            continue
        ot["estatus"] = ESTATUS_OT_CERRADA
        ot["fecha_cierre"] = date.today().isoformat()
        ot["mecanico"] = mecanico.strip()
        ot["taller"] = taller.strip()
        ot["reporte_final"] = reporte.strip()
        ot["refacciones"] = refs
        ot["costo_mano_obra"] = round(costo_mo, 2)
        ot["costo_refacciones"] = round(sum(_num_combustible(r.get("costo")) for r in refs), 2)
        ot["costo_total"] = costo_ot_vehicular(ot)
        if ot.get("tipo") == "preventivo":
            unidad = str(ot.get("unidad") or "")
            unidades = datos.setdefault("unidades", {})
            cfg = dict(unidades.get(unidad) or {})
            km_cierre = _num_combustible(ot.get("km_actual")) or _num_combustible(ot.get("km_objetivo"))
            cfg["ultimo_mto_km"] = km_cierre
            unidades[unidad] = cfg
        break
    guardar_vehiculos(datos)


def render_cierre_ot_form(ot: dict, key_pref: str) -> None:
    with st.expander(f"Cerrar {ot.get('id')} · {ot.get('unidad')}", expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            mecanico = st.text_input("Mecánico responsable", value=str(ot.get("mecanico") or ""), key=f"{key_pref}_mec")
            taller = st.text_input("Taller", value=str(ot.get("taller") or ""), key=f"{key_pref}_taller")
            costo_mo = st.number_input(
                "Costo mano de obra",
                min_value=0.0,
                value=float(_num_combustible(ot.get("costo_mano_obra"))),
                step=50.0,
                format="%.2f",
                key=f"{key_pref}_mo",
            )
        with c2:
            reporte = st.text_area(
                "Reporte final",
                value=str(ot.get("reporte_final") or ""),
                height=120,
                key=f"{key_pref}_rep",
            )
        refs_prev = ot.get("refacciones") or [{"nombre": "", "cantidad": 1, "costo": 0.0}]
        if not refs_prev:
            refs_prev = [{"nombre": "", "cantidad": 1, "costo": 0.0}]
        editor = st.data_editor(
            pd.DataFrame(refs_prev),
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            key=f"{key_pref}_refs",
            column_config={
                "nombre": st.column_config.TextColumn("Refacción"),
                "cantidad": st.column_config.NumberColumn("Cantidad", min_value=0, step=1),
                "costo": st.column_config.NumberColumn("Costo $", min_value=0.0, format="$%.2f"),
            },
        )
        if st.button("Marcar OT como realizada", type="primary", key=f"{key_pref}_ok"):
            refs = []
            for _, fila in editor.iterrows():
                nombre = str(fila.get("nombre") or "").strip()
                if not nombre:
                    continue
                refs.append(
                    {
                        "nombre": nombre,
                        "cantidad": _num_combustible(fila.get("cantidad")) or 1,
                        "costo": _num_combustible(fila.get("costo")),
                    }
                )
            if not mecanico.strip():
                st.warning("Captura el mecánico responsable para cerrar la OT.")
            elif not reporte.strip():
                st.warning("Captura el reporte final del servicio.")
            else:
                datos = cargar_vehiculos()
                _cerrar_ot_vehicular(datos, str(ot.get("id")), mecanico, taller, reporte, costo_mo, refs)
                st.session_state.veh_aviso = f"{ot.get('id')} cerrada. Quedó en el historial de OT realizadas."
                st.rerun()


def render_alertas_preventivos() -> None:
    st.markdown('<div class="section-title">Alertas y Preventivos por Kilometraje</div>', unsafe_allow_html=True)
    st.caption(
        "El sistema vigila el odómetro de cada unidad. Si faltan 500 km o menos para el servicio "
        f"(intervalo base {INTERVALO_MTO_KM_DEFAULT:,} km), genera una OT preventiva y estima la fecha "
        "con el kilometraje diario de uso."
    )
    diario, parque = cargar_datos_flota()
    datos = sincronizar_ots_preventivas(diario, parque, cargar_vehiculos())
    aviso = st.session_state.pop("veh_aviso", None)
    if aviso:
        st.success(aviso)

    monitores = monitorear_flota(diario, parque, datos)
    alertas = [m for m in monitores if m["estado"] == "alerta"]
    vencidas = [m for m in monitores if m["estado"] == "vencido"]
    abiertas = ots_por_estatus(datos, tipo="preventivo")
    abiertas = [o for o in abiertas if o.get("estatus") != ESTATUS_OT_CERRADA]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Unidades monitoreadas", len(monitores))
    k2.metric("En alerta (≤ 500 km)", len(alertas))
    k3.metric("Vencidas", len(vencidas))
    k4.metric("OT preventivas abiertas", len(abiertas))

    with st.expander("⚙️ Intervalo y último servicio por unidad", expanded=False):
        g1, g2 = st.columns(2)
        with g1:
            intervalo_g = st.number_input(
                "Intervalo general (km)",
                min_value=500,
                value=int(datos.get("intervalo_km") or INTERVALO_MTO_KM_DEFAULT),
                step=500,
                key="veh_int_g",
            )
        with g2:
            umbral_g = st.number_input(
                "Umbral de alerta (km)",
                min_value=50,
                value=int(datos.get("umbral_alerta_km") or UMBRAL_ALERTA_KM_DEFAULT),
                step=50,
                key="veh_umb_g",
            )
        filas_cfg = []
        for mon in monitores:
            filas_cfg.append(
                {
                    "Unidad": mon["unidad"],
                    "Odómetro": mon["odometro"],
                    "Último MTO (km)": mon["ultimo_mto_km"],
                    "Intervalo km": mon["intervalo_km"],
                    "Umbral alerta": mon["umbral_alerta_km"],
                }
            )
        editor = st.data_editor(
            pd.DataFrame(filas_cfg) if filas_cfg else pd.DataFrame(
                columns=["Unidad", "Odómetro", "Último MTO (km)", "Intervalo km", "Umbral alerta"]
            ),
            hide_index=True,
            use_container_width=True,
            disabled=["Unidad", "Odómetro"],
            key="veh_cfg_unidades",
        )
        if st.button("Guardar configuración de preventivos", key="veh_cfg_save"):
            datos["intervalo_km"] = int(intervalo_g)
            datos["umbral_alerta_km"] = int(umbral_g)
            unidades = datos.setdefault("unidades", {})
            for _, fila in editor.iterrows():
                unidad = str(fila.get("Unidad") or "").strip()
                if not unidad:
                    continue
                unidades[unidad] = {
                    "intervalo_km": int(_num_combustible(fila.get("Intervalo km")) or intervalo_g),
                    "umbral_alerta_km": int(_num_combustible(fila.get("Umbral alerta")) or umbral_g),
                    "ultimo_mto_km": _num_combustible(fila.get("Último MTO (km)")),
                }
            guardar_vehiculos(datos)
            st.session_state.veh_aviso = "Configuración de preventivos actualizada. Se recalcularon las alertas."
            st.rerun()

    def _tarjeta_unidad(mon: dict) -> None:
        tono, etiqueta = etiqueta_estado_mto(mon["estado"])
        fecha_txt = mon["fecha_estimada"].strftime("%d/%m/%Y") if mon["fecha_estimada"] else "sin km diario suficiente"
        uso = f"{mon['km_diario']:.1f} km/día" if mon["km_diario"] else "sin historial reciente"
        st.markdown(
            f"""
            <div class="veh-alert {tono}">
                <div class="kicker">{html.escape(etiqueta)}</div>
                <h3>{html.escape(mon["unidad"])} · {mon["odometro"]:,.0f} km</h3>
                <p>Próximo servicio a los <b>{mon["km_objetivo"]:,.0f} km</b> · restan
                <b>{mon["km_restantes"]:,.0f} km</b> · uso {html.escape(uso)} ·
                fecha estimada <b>{html.escape(fecha_txt)}</b></p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("**Unidades en ventana de preventivo**")
    criticas = [m for m in monitores if m["estado"] in {"alerta", "vencido"}]
    ok_unidades = [m for m in monitores if m["estado"] == "ok"]
    if not monitores:
        st.info("No hay unidades en el parque o en el registro de combustible.")
    elif not criticas:
        st.success("Ninguna unidad está a 500 km o menos de su próximo servicio.")
    else:
        for mon in criticas:
            _tarjeta_unidad(mon)
    if ok_unidades:
        with st.expander(f"Unidades dentro de programa ({len(ok_unidades)})", expanded=False):
            for mon in ok_unidades:
                _tarjeta_unidad(mon)

    st.markdown('<div class="section-title">Órdenes de trabajo preventivas</div>', unsafe_allow_html=True)
    if not abiertas:
        st.info("No hay OT preventivas abiertas. Se generan solas cuando una unidad entra en la ventana de 500 km.")
        return
    for ot in abiertas:
        render_ficha_ot_vehicular(ot)
        render_cierre_ot_form(ot, f"prev_{ot.get('id')}")


def render_mantenimientos_correctivos() -> None:
    st.markdown('<div class="section-title">Mantenimientos Correctivos</div>', unsafe_allow_html=True)
    st.caption("Registra fallas mecánicas imprevistas, taller, refacciones y costos. Cada alta genera una OT correctiva.")
    diario, parque = cargar_datos_flota()
    datos = cargar_vehiculos()
    aviso = st.session_state.pop("veh_aviso", None)
    if aviso:
        st.success(aviso)

    unidades = unidades_combustible(diario, parque)
    abiertas = [o for o in ots_por_estatus(datos, tipo="correctivo") if o.get("estatus") != ESTATUS_OT_CERRADA]
    k1, k2 = st.columns(2)
    k1.metric("OT correctivas abiertas", len(abiertas))
    k2.metric(
        "Costo estimado abierto",
        f"${sum(costo_ot_vehicular(o) for o in abiertas):,.2f}",
    )

    with st.expander("➕ Registrar falla / correctivo", expanded=not abiertas):
        if not unidades:
            st.warning("No hay números económicos cargados. Revisa la hoja Listas del control de combustible.")
        else:
            fid = st.session_state.get("veh_corr_id", 0)
            c1, c2, c3 = st.columns(3)
            with c1:
                unidad = st.selectbox("Número económico", unidades, key=f"corr_un_{fid}")
                fecha_falla = st.date_input("Fecha de la falla", value=date.today(), format="DD/MM/YYYY", key=f"corr_f_{fid}")
            with c2:
                taller = st.text_input("Taller", key=f"corr_t_{fid}")
                mecanico = st.text_input("Mecánico / responsable", key=f"corr_m_{fid}")
            with c3:
                costo_mo = st.number_input("Costo mano de obra", min_value=0.0, value=0.0, step=50.0, format="%.2f", key=f"corr_mo_{fid}")
                km_act = ultimo_odometro_unidad(diario, unidad) if unidad else 0.0
                st.metric("Odómetro actual", f"{km_act:,.0f} km")
            falla = st.text_area("Descripción de la falla", height=90, key=f"corr_d_{fid}")
            refs_ed = st.data_editor(
                pd.DataFrame([{"nombre": "", "cantidad": 1, "costo": 0.0}]),
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                key=f"corr_refs_{fid}",
                column_config={
                    "nombre": st.column_config.TextColumn("Refacción"),
                    "cantidad": st.column_config.NumberColumn("Cantidad", min_value=0, step=1),
                    "costo": st.column_config.NumberColumn("Costo $", min_value=0.0, format="$%.2f"),
                },
            )
            if st.button("Generar OT correctiva", type="primary", key=f"corr_ok_{fid}"):
                if not unidad:
                    st.warning("Selecciona la unidad.")
                elif not str(falla).strip():
                    st.warning("Describe la falla mecánica.")
                else:
                    refs = []
                    for _, fila in refs_ed.iterrows():
                        nombre = str(fila.get("nombre") or "").strip()
                        if not nombre:
                            continue
                        refs.append(
                            {
                                "nombre": nombre,
                                "cantidad": _num_combustible(fila.get("cantidad")) or 1,
                                "costo": _num_combustible(fila.get("costo")),
                            }
                        )
                    ot = {
                        "id": folio_ot_vehicular(datos),
                        "tipo": "correctivo",
                        "unidad": unidad,
                        "estatus": ESTATUS_OT_ABIERTA,
                        "fecha_creacion": fecha_falla.isoformat(),
                        "fecha_estimada": fecha_falla.isoformat(),
                        "km_actual": km_act,
                        "km_objetivo": km_act,
                        "km_restantes": 0,
                        "km_diario": km_diario_unidad(diario, unidad),
                        "descripcion": f"Correctivo imprevisto — {unidad}.",
                        "falla": str(falla).strip(),
                        "taller": taller.strip(),
                        "mecanico": mecanico.strip(),
                        "refacciones": refs,
                        "costo_mano_obra": round(float(costo_mo), 2),
                        "costo_refacciones": round(sum(_num_combustible(r.get("costo")) for r in refs), 2),
                        "costo_total": 0.0,
                        "reporte_final": "",
                        "fecha_cierre": "",
                        "origen": "manual",
                    }
                    ot["costo_total"] = costo_ot_vehicular(ot)
                    datos.setdefault("ots", []).append(ot)
                    guardar_vehiculos(datos)
                    st.session_state.veh_corr_id = fid + 1
                    st.session_state.veh_aviso = f"{ot['id']} generada para {unidad}."
                    st.rerun()

    st.markdown("**OT correctivas abiertas**")
    if not abiertas:
        st.info("No hay correctivos abiertos.")
        return
    for ot in abiertas:
        render_ficha_ot_vehicular(ot)
        render_cierre_ot_form(ot, f"corr_{ot.get('id')}")


def render_historial_ot_vehiculos() -> None:
    st.markdown('<div class="section-title">Historial de OT Cerradas / Realizadas</div>', unsafe_allow_html=True)
    st.caption("Consulta el reporte final, el mecánico responsable y el costo total de cada orden concluida.")
    datos = cargar_vehiculos()
    cerradas = ots_por_estatus(datos, estatus=ESTATUS_OT_CERRADA)
    if not cerradas:
        st.info("Aún no hay OT cerradas. Al concluir un preventivo o un correctivo aparecen aquí.")
        return

    unidades = sorted({str(o.get("unidad") or "") for o in cerradas if o.get("unidad")})
    tipos = ["Todas", "Preventivo", "Correctivo", "Servicio externo"]
    f1, f2, f3 = st.columns(3)
    with f1:
        tipo_f = st.selectbox("Tipo", tipos, key="hist_ot_tipo")
    with f2:
        unidad_f = st.selectbox("Unidad", ["Todas"] + unidades, key="hist_ot_un")
    with f3:
        meses = sorted(
            {
                str(parsear_fecha_valor(o.get("fecha_cierre") or o.get("fecha_creacion")))[:7]
                for o in cerradas
                if parsear_fecha_valor(o.get("fecha_cierre") or o.get("fecha_creacion"))
            },
            reverse=True,
        )
        mes_f = st.selectbox(
            "Mes de cierre",
            ["Todos"] + [etiqueta_mes(m) for m in meses],
            key="hist_ot_mes",
        )
    vista = []
    for ot in cerradas:
        if tipo_f == "Preventivo" and ot.get("tipo") != "preventivo":
            continue
        if tipo_f == "Correctivo" and ot.get("tipo") != "correctivo":
            continue
        if tipo_f == "Servicio externo" and ot.get("tipo") != "externo":
            continue
        if unidad_f != "Todas" and str(ot.get("unidad")) != unidad_f:
            continue
        cierre = parsear_fecha_valor(ot.get("fecha_cierre") or ot.get("fecha_creacion"))
        if mes_f != "Todos" and (not cierre or etiqueta_mes(cierre.strftime("%Y-%m")) != mes_f):
            continue
        vista.append(ot)

    gasto = sum(costo_ot_vehicular(o) for o in vista)
    c1, c2 = st.columns(2)
    c1.metric("OT realizadas", len(vista))
    c2.metric("Costo total del recorte", f"${gasto:,.2f}")

    if not vista:
        st.info("No hay OT cerradas con esos filtros.")
        return
    tabla = pd.DataFrame(
        [
            {
                "Folio": o.get("id"),
                "Tipo": (
                    "Preventivo" if o.get("tipo") == "preventivo"
                    else "Servicio externo" if o.get("tipo") == "externo"
                    else "Correctivo"
                ),
                "Unidad": o.get("unidad"),
                "Cierre": html_fecha(o.get("fecha_cierre")),
                "Mecánico": o.get("mecanico") or "—",
                "Taller": o.get("taller") or "—",
                "Costo $": costo_ot_vehicular(o),
            }
            for o in vista
        ]
    )
    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
        column_config={"Costo $": st.column_config.NumberColumn(format="$%.2f")},
    )
    st.markdown("**Reportes finales**")
    for ot in vista:
        render_ficha_ot_vehicular(ot)


def render_mantenimientos_mensuales() -> None:
    st.markdown('<div class="section-title">Mantenimientos Mensuales</div>', unsafe_allow_html=True)
    st.caption("Concentrado de gasto de combustible, OT de flota y estado de las unidades en el mes.")
    diario, parque = cargar_datos_flota()
    datos = sincronizar_ots_preventivas(diario, parque, cargar_vehiculos())
    hoy = date.today()
    c1, c2 = st.columns(2)
    with c1:
        anio = st.number_input("Año", min_value=2024, max_value=hoy.year + 1, value=hoy.year, step=1, key="veh_mes_anio")
    with c2:
        mes_n = st.selectbox(
            "Mes",
            list(range(1, 13)),
            index=hoy.month - 1,
            format_func=lambda n: NOMBRES_MES[n - 1],
            key="veh_mes_n",
        )
    inicio = date(int(anio), int(mes_n), 1)
    fin = date(int(anio), int(mes_n), calendar.monthrange(int(anio), int(mes_n))[1])
    comb = filtrar_registro_combustible(diario, None, inicio, fin, "Todos")
    met = metricas_combustible(comb)
    ots_mes = []
    for ot in datos.get("ots") or []:
        ref = parsear_fecha_valor(ot.get("fecha_cierre") or ot.get("fecha_creacion"))
        if ref and inicio <= ref <= fin:
            ots_mes.append(ot)
    gasto_ot = sum(costo_ot_vehicular(o) for o in ots_mes)
    prev = [o for o in ots_mes if o.get("tipo") == "preventivo"]
    corr = [o for o in ots_mes if o.get("tipo") == "correctivo"]
    cerradas = [o for o in ots_mes if o.get("estatus") == ESTATUS_OT_CERRADA]
    monitores = monitorear_flota(diario, parque, datos)
    n_alerta = sum(1 for m in monitores if m["estado"] == "alerta")
    n_venc = sum(1 for m in monitores if m["estado"] == "vencido")

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Gasto combustible", f"${met['gasto']:,.2f}")
    k2.metric("Gasto OT flota", f"${gasto_ot:,.2f}")
    k3.metric("OT del mes", len(ots_mes))
    k4.metric("OT cerradas", len(cerradas))
    k5.metric("Unidades en alerta / vencidas", f"{n_alerta} / {n_venc}")

    g1, g2, g3 = st.columns(3)
    g1.markdown(money_html("Combustible del mes", met["gasto"]), unsafe_allow_html=True)
    g2.markdown(money_html("Mantenimiento del mes", gasto_ot), unsafe_allow_html=True)
    g3.markdown(money_html("Gasto total flota", met["gasto"] + gasto_ot), unsafe_allow_html=True)

    st.markdown("**Estado actual de la flota**")
    if monitores:
        estado_df = pd.DataFrame(
            [
                {
                    "No. económico": m["unidad"],
                    "Odómetro": m["odometro"],
                    "Próximo MTO (km)": m["km_objetivo"],
                    "Km restantes": m["km_restantes"],
                    "Km / día": m["km_diario"],
                    "Fecha estimada": m["fecha_estimada"].strftime("%d/%m/%Y") if m["fecha_estimada"] else "—",
                    "Estatus": etiqueta_estado_mto(m["estado"])[1],
                }
                for m in monitores
            ]
        )
        st.dataframe(estado_df, use_container_width=True, hide_index=True)
    else:
        st.info("Sin unidades para el concentrado.")

    st.markdown("**OT del mes**")
    if ots_mes:
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Folio": o.get("id"),
                        "Tipo": (
                    "Preventivo" if o.get("tipo") == "preventivo"
                    else "Servicio externo" if o.get("tipo") == "externo"
                    else "Correctivo"
                ),
                        "Unidad": o.get("unidad"),
                        "Estatus": str(o.get("estatus") or "").replace("_", " ").title(),
                        "Fecha": html_fecha(o.get("fecha_cierre") or o.get("fecha_creacion")),
                        "Mecánico": o.get("mecanico") or "—",
                        "Costo $": costo_ot_vehicular(o),
                    }
                    for o in ots_mes
                ]
            ),
            use_container_width=True,
            hide_index=True,
            column_config={"Costo $": st.column_config.NumberColumn(format="$%.2f")},
        )
    else:
        st.info(f"No hay OT de flota en {etiqueta_mes(f'{int(anio):04d}-{int(mes_n):02d}')}.")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                {
                    "Mes": etiqueta_mes(f"{int(anio):04d}-{int(mes_n):02d}"),
                    "Gasto combustible $": round(met["gasto"], 2),
                    "Litros": round(met["litros"], 2),
                    "Km": round(met["km"], 1),
                    "Gasto OT $": round(gasto_ot, 2),
                    "OT preventivas": len(prev),
                    "OT correctivas": len(corr),
                    "OT cerradas": len(cerradas),
                    "Unidades alerta": n_alerta,
                    "Unidades vencidas": n_venc,
                }
            ]
        ).to_excel(writer, sheet_name="Resumen", index=False)
        if monitores:
            estado_df.to_excel(writer, sheet_name="Flota", index=False)
        if ots_mes:
            pd.DataFrame(
                [
                    {
                        "Folio": o.get("id"),
                        "Tipo": o.get("tipo"),
                        "Unidad": o.get("unidad"),
                        "Estatus": o.get("estatus"),
                        "Mecanico": o.get("mecanico"),
                        "Costo": costo_ot_vehicular(o),
                    }
                    for o in ots_mes
                ]
            ).to_excel(writer, sheet_name="OT", index=False)
    st.download_button(
        "⬇️ Exportar concentrado mensual",
        data=buf.getvalue(),
        file_name=f"mantenimientos_mensuales_{int(anio):04d}-{int(mes_n):02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="veh_mes_xlsx",
    )


def folio_ingreso_vehicular(datos: dict, anio: int | None = None) -> str:
    anio = anio or date.today().year
    prefijo = f"ING-VEH-{anio}-"
    maximo = 0
    for ing in datos.get("ingresos") or []:
        folio = str(ing.get("id") or "")
        if folio.startswith(prefijo):
            try:
                maximo = max(maximo, int(folio[len(prefijo):]))
            except ValueError:
                continue
    return f"{prefijo}{maximo + 1:03d}"


def folio_solicitud_ref(datos: dict, anio: int | None = None) -> str:
    anio = anio or date.today().year
    prefijo = f"SOL-REF-{anio}-"
    maximo = 0
    for ing in datos.get("ingresos") or []:
        folio = str((ing.get("solicitud") or {}).get("id") or "")
        if folio.startswith(prefijo):
            try:
                maximo = max(maximo, int(folio[len(prefijo):]))
            except ValueError:
                continue
    return f"{prefijo}{maximo + 1:03d}"


def existencia_insumo(nombre: str) -> float:
    nombre_n = str(nombre or "").strip().casefold()
    if not nombre_n:
        return 0.0
    total = 0.0
    for tabla in inventario_por_origen().values():
        if tabla is None or tabla.empty:
            continue
        match = tabla[tabla["Insumo"].astype(str).str.strip().str.casefold() == nombre_n]
        if not match.empty:
            total += float(match["Existencia"].sum())
    return total


def guardar_evidencias_taller(folio: str, fase: str, archivos: list) -> list[str]:
    if not archivos:
        return []
    carpeta = os.path.join(TALLER_EVIDENCIAS_DIR, str(folio), fase)
    os.makedirs(carpeta, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    rutas: list[str] = []
    for i, archivo in enumerate(list(archivos)[:8], start=1):
        extension = os.path.splitext(getattr(archivo, "name", "") or "")[1].lower() or ".jpg"
        destino = os.path.join(carpeta, f"{fase}_{stamp}_{i}{extension}")
        with open(destino, "wb") as salida:
            salida.write(archivo.getvalue())
        rutas.append(destino.replace("\\", "/"))
    return rutas


def opciones_mecanicos_taller() -> list[str]:
    nombres = list(bolsa_nombres_personal())
    vistos = {n.casefold() for n in nombres}
    for cuadrilla in CUADRILLAS:
        if cuadrilla.casefold() not in vistos:
            nombres.append(cuadrilla)
            vistos.add(cuadrilla.casefold())
    return nombres


def opciones_responsables_taller() -> list[str]:
    return opciones_mecanicos_taller() + ["Parque vehicular", "Administración"]


def unidades_taller(diario: pd.DataFrame, parque: pd.DataFrame, datos: dict) -> list[str]:
    unidades = list(unidades_combustible(diario, parque))
    vistos = {u.casefold() for u in unidades}
    for origen in (
        (datos.get("unidades") or {}).keys(),
        [i.get("unidad") for i in datos.get("ingresos") or []],
        [a.get("unidad") for a in datos.get("asignaciones") or []],
    ):
        for valor in origen:
            texto = str(valor or "").strip()
            if texto and texto.casefold() not in vistos:
                unidades.append(texto)
                vistos.add(texto.casefold())
    return sorted(unidades, key=lambda u: (u.upper(), u))


def actualizar_unidad_taller(datos: dict, unidad: str, **campos) -> None:
    if not unidad:
        return
    unidades = datos.setdefault("unidades", {})
    cfg = dict(unidades.get(unidad) or {})
    for clave, valor in campos.items():
        cfg[clave] = valor
    unidades[unidad] = cfg


def aplicar_semaforo_unidad(datos: dict, unidad: str, semaforo: str, desde: date | None = None) -> None:
    semaforo = semaforo if semaforo in ETIQUETA_SEMAFORO_UNIDAD else SEMAFORO_OPERATIVA
    raw = dict((datos.get("unidades") or {}).get(unidad) or {})
    campos: dict = {"semaforo": semaforo}
    if semaforo in {SEMAFORO_DETENIDA, SEMAFORO_EXTERNO}:
        previa = parsear_fecha_valor(raw.get("detenida_desde"))
        campos["detenida_desde"] = (previa or desde or date.today()).isoformat()
    else:
        campos["detenida_desde"] = ""
    actualizar_unidad_taller(datos, unidad, **campos)


def ingresos_abiertos(datos: dict, unidad: str = "") -> list[dict]:
    salida = []
    for ing in datos.get("ingresos") or []:
        if str(ing.get("estatus") or "") == ESTATUS_INGRESO_CERRADO:
            continue
        if unidad and str(ing.get("unidad") or "") != unidad:
            continue
        salida.append(ing)
    return sorted(salida, key=lambda i: str(i.get("fecha_entrada") or ""), reverse=True)


def semaforo_efectivo_unidad(datos: dict, unidad: str) -> dict:
    abiertos = ingresos_abiertos(datos, unidad)
    if abiertos:
        ing = abiertos[0]
        semaforo = str(ing.get("semaforo") or SEMAFORO_DETENIDA)
        desde = parsear_fecha_valor(ing.get("fecha_entrada"))
        return {"semaforo": semaforo, "ingreso": ing, "desde": desde}
    raw = dict((datos.get("unidades") or {}).get(unidad) or {})
    semaforo = str(raw.get("semaforo") or SEMAFORO_OPERATIVA)
    if semaforo not in ETIQUETA_SEMAFORO_UNIDAD:
        semaforo = SEMAFORO_OPERATIVA
    return {
        "semaforo": semaforo,
        "ingreso": None,
        "desde": parsear_fecha_valor(raw.get("detenida_desde")),
    }


def dias_downtime(semaforo: str, desde: date | None, hoy: date | None = None) -> int:
    hoy = hoy or date.today()
    if semaforo not in {SEMAFORO_DETENIDA, SEMAFORO_EXTERNO} or not desde:
        return 0
    return max(0, (hoy - desde).days)


def etiqueta_tipo_taller(tipo: str) -> str:
    if tipo == TIPO_TALLER_PREVENTIVO:
        return "Preventivo"
    if tipo == TIPO_TALLER_EXTERNO:
        return "Servicio externo"
    return "Correctivo"


def ot_desde_ingreso_taller(datos: dict, ingreso: dict) -> dict:
    tipo = str(ingreso.get("tipo") or TIPO_TALLER_CORRECTIVO)
    return {
        "id": folio_ot_vehicular(datos),
        "tipo": tipo,
        "unidad": ingreso.get("unidad"),
        "estatus": ESTATUS_OT_PROCESO,
        "fecha_creacion": ingreso.get("fecha_entrada") or date.today().isoformat(),
        "fecha_estimada": ingreso.get("fecha_programacion") or ingreso.get("fecha_entrada") or "",
        "km_actual": _num_combustible(ingreso.get("km_entrada")),
        "km_objetivo": _num_combustible(ingreso.get("km_entrada")),
        "km_restantes": 0,
        "descripcion": str(ingreso.get("falla") or etiqueta_tipo_taller(tipo)),
        "falla": ingreso.get("falla") or "",
        "taller": str((ingreso.get("externo") or {}).get("taller") or "Taller PJP4"),
        "mecanico": ingreso.get("mecanico") or "",
        "refacciones": ingreso.get("refacciones") or [],
        "costo_mano_obra": 0.0,
        "costo_refacciones": 0.0,
        "costo_total": 0.0,
        "reporte_final": "",
        "fecha_cierre": "",
        "origen": "taller",
        "ingreso_id": ingreso.get("id"),
    }


def persistir_ingreso_taller(datos: dict, ingreso: dict, crear_ot: bool = False) -> dict:
    lista = datos.setdefault("ingresos", [])
    idx = next((i for i, item in enumerate(lista) if str(item.get("id")) == str(ingreso.get("id"))), None)
    if crear_ot and not ingreso.get("ot_id"):
        ot = ot_desde_ingreso_taller(datos, ingreso)
        datos.setdefault("ots", []).append(ot)
        ingreso["ot_id"] = ot["id"]
    if idx is None:
        lista.append(ingreso)
    else:
        lista[idx] = ingreso
    aplicar_semaforo_unidad(
        datos,
        str(ingreso.get("unidad") or ""),
        str(ingreso.get("semaforo") or SEMAFORO_DETENIDA),
        parsear_fecha_valor(ingreso.get("fecha_entrada")),
    )
    placas = str(ingreso.get("placas") or "").strip()
    if placas:
        actualizar_unidad_taller(datos, str(ingreso.get("unidad") or ""), placas=placas)
    guardar_vehiculos(datos)
    return ingreso


def cerrar_ingreso_taller(datos: dict, folio: str, km_salida: float, horometro_salida: float, notas: str) -> None:
    for ing in datos.get("ingresos") or []:
        if str(ing.get("id")) != folio:
            continue
        ing["estatus"] = ESTATUS_INGRESO_CERRADO
        ing["semaforo"] = SEMAFORO_OPERATIVA
        ing["fecha_salida"] = date.today().isoformat()
        ing["km_salida"] = km_salida
        ing["horometro_salida"] = horometro_salida
        if notas.strip():
            ing["reporte_cierre"] = notas.strip()
        aplicar_semaforo_unidad(datos, str(ing.get("unidad") or ""), SEMAFORO_OPERATIVA)
        ot_id = str(ing.get("ot_id") or "")
        if ot_id:
            for ot in datos.get("ots") or []:
                if str(ot.get("id")) == ot_id and str(ot.get("estatus")) != ESTATUS_OT_CERRADA:
                    ot["estatus"] = ESTATUS_OT_CERRADA
                    ot["fecha_cierre"] = date.today().isoformat()
                    ot["reporte_final"] = notas.strip() or "Unidad liberada desde taller."
                    ot["mecanico"] = ing.get("mecanico") or ot.get("mecanico") or ""
        break
    guardar_vehiculos(datos)


def render_evidencias_taller(ingreso: dict) -> None:
    evidencias = ingreso.get("evidencias") or {}
    cols = st.columns(3)
    for col, fase in zip(cols, ("antes", "durante", "despues")):
        with col:
            st.markdown(f'<div class="taller-fase">{fase}</div>', unsafe_allow_html=True)
            fotos = [f for f in (evidencias.get(fase) or []) if f]
            if not fotos:
                st.caption("Sin fotos")
                continue
            for ruta in fotos:
                if os.path.isfile(ruta):
                    st.image(ruta, use_container_width=True)
                else:
                    st.caption(os.path.basename(str(ruta)))


def render_ficha_ingreso_taller(ingreso: dict) -> None:
    tipo = etiqueta_tipo_taller(str(ingreso.get("tipo") or ""))
    semaforo = str(ingreso.get("semaforo") or SEMAFORO_DETENIDA)
    clase = CLASE_SEMAFORO_UNIDAD.get(semaforo, "mixto")
    etiqueta = ETIQUETA_SEMAFORO_UNIDAD.get(semaforo, semaforo)
    downtime = dias_downtime(semaforo, parsear_fecha_valor(ingreso.get("fecha_entrada")))
    solicitud = ingreso.get("solicitud") or {}
    externo = ingreso.get("externo") or {}
    st.markdown(
        f"""
        <div class="ot-sheet">
            <div class="ot-kicker">{html.escape(str(ingreso.get("id") or "ING"))} · {html.escape(tipo)} ·
            {html.escape(ETIQUETA_ESTATUS_INGRESO.get(str(ingreso.get("estatus") or ""), "—"))}</div>
            <h3><span class="semaforo {clase}"></span> {html.escape(str(ingreso.get("unidad") or "Unidad"))}
            · {html.escape(etiqueta)}</h3>
            <div class="ot-grid">
                <div class="ot-field"><div class="k">Placas</div>
                <div class="v">{html.escape(str(ingreso.get("placas") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Entrada</div>
                <div class="v">{html.escape(html_fecha(ingreso.get("fecha_entrada")))}</div></div>
                <div class="ot-field"><div class="k">Km / horómetro entrada</div>
                <div class="v">{_num_combustible(ingreso.get("km_entrada")):,.0f} km ·
                {_num_combustible(ingreso.get("horometro_entrada")):,.0f} h</div></div>
                <div class="ot-field"><div class="k">Responsable / cuadrilla</div>
                <div class="v">{html.escape(str(ingreso.get("responsable") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Mecánico asignado</div>
                <div class="v">{html.escape(str(ingreso.get("mecanico") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Programación / lead time</div>
                <div class="v">{html.escape(html_fecha(ingreso.get("fecha_programacion")))} ·
                {_num_combustible(ingreso.get("horas_estimadas")):,.0f} h</div></div>
                <div class="ot-field"><div class="k">Días de inactividad</div>
                <div class="v">{downtime} día(s) fuera de operación</div></div>
                <div class="ot-field"><div class="k">OT vinculada</div>
                <div class="v">{html.escape(str(ingreso.get("ot_id") or "—"))}</div></div>
            </div>
            <div class="ot-field"><div class="k">Daños / fallas</div>
            <div class="v">{html.escape(str(ingreso.get("falla") or "—"))}</div></div>
            <div class="ot-field" style="margin-top:0.45rem"><div class="k">Refacciones / solicitud</div>
            <div class="v">{html.escape(str(solicitud.get("id") or "Sin solicitud"))} ·
            {html.escape(str(solicitud.get("estatus") or "—"))}</div></div>
            <div class="ot-field"><div class="k">Servicio externo</div>
            <div class="v">{html.escape(str(externo.get("taller") or "—"))} ·
            {html.escape(str(externo.get("cotizacion") or "sin cotización"))} ·
            {html.escape(str(externo.get("autorizacion") or "—"))}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_semaforo_parque(datos: dict, diario: pd.DataFrame, parque: pd.DataFrame) -> None:
    unidades = unidades_taller(diario, parque, datos)
    filas = []
    for unidad in unidades:
        estado = semaforo_efectivo_unidad(datos, unidad)
        semaforo = estado["semaforo"]
        downtime = dias_downtime(semaforo, estado["desde"])
        ing = estado["ingreso"]
        filas.append(
            {
                "unidad": unidad,
                "semaforo": semaforo,
                "downtime": downtime,
                "odometro": ultimo_odometro_unidad(diario, unidad),
                "ingreso": ing,
            }
        )
    n_prec = sum(1 for f in filas if f["semaforo"] == SEMAFORO_PRECAUCION)
    n_det = sum(1 for f in filas if f["semaforo"] == SEMAFORO_DETENIDA)
    n_ext = sum(1 for f in filas if f["semaforo"] == SEMAFORO_EXTERNO)
    dias_tot = sum(f["downtime"] for f in filas)
    k1, k2, k3, k4 = st.columns(4, gap="medium")
    k1.markdown(kpi_html("Precaución", str(n_prec), "Pueden rodar temporalmente", "warn"), unsafe_allow_html=True)
    k2.markdown(kpi_html("Detenidas en base", str(n_det), "Inhabilitadas hasta reparación", "danger"), unsafe_allow_html=True)
    k3.markdown(kpi_html("Taller externo", str(n_ext), "Fuera de la base", "info"), unsafe_allow_html=True)
    k4.markdown(kpi_html("Días de inactividad", str(dias_tot), "Acumulado de unidades detenidas/externas", "warn" if dias_tot else "ok"), unsafe_allow_html=True)

    if not filas:
        st.info("No hay números económicos en combustible ni ingresos de taller.")
        return
    orden = {SEMAFORO_DETENIDA: 0, SEMAFORO_EXTERNO: 1, SEMAFORO_PRECAUCION: 2, SEMAFORO_OPERATIVA: 3}

    def _tarjeta(fila: dict) -> None:
        semaforo = fila["semaforo"]
        tono = TONO_ALERTA_SEMAFORO.get(semaforo, "")
        etiqueta = ETIQUETA_SEMAFORO_UNIDAD[semaforo]
        clase = CLASE_SEMAFORO_UNIDAD[semaforo]
        ing = fila["ingreso"]
        if ing:
            extra = (
                f"{etiqueta_tipo_taller(str(ing.get('tipo')))} {ing.get('id')} · "
                f"{ETIQUETA_ESTATUS_INGRESO.get(str(ing.get('estatus') or ''), '')} · "
                f"programada {html_fecha(ing.get('fecha_programacion'))} · "
                f"{_num_combustible(ing.get('horas_estimadas')):,.0f} h estimadas"
            )
        elif fila["downtime"]:
            extra = f"Sin ingreso abierto · {fila['downtime']} día(s) fuera"
        else:
            extra = "Sin intervención abierta"
        st.markdown(
            f"""
            <div class="veh-alert {tono}">
                <div class="kicker"><span class="semaforo {clase}"></span> {html.escape(etiqueta)}
                · {fila['downtime']} día(s) de inactividad</div>
                <h3>{html.escape(fila['unidad'])} · {fila['odometro']:,.0f} km</h3>
                <p>{html.escape(extra)}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    ordenadas = sorted(filas, key=lambda f: (orden.get(f["semaforo"], 9), -f["downtime"], f["unidad"]))
    criticas = [f for f in ordenadas if f["semaforo"] != SEMAFORO_OPERATIVA]
    operativas = [f for f in ordenadas if f["semaforo"] == SEMAFORO_OPERATIVA]
    if not criticas:
        st.success("Ninguna unidad está detenida, en externo o con precaución.")
    else:
        for fila in criticas:
            _tarjeta(fila)
    if operativas:
        with st.expander(f"Unidades operativas ({len(operativas)})", expanded=False):
            for fila in operativas:
                _tarjeta(fila)


def render_form_ingreso_taller(datos: dict, diario: pd.DataFrame, parque: pd.DataFrame) -> None:
    unidades = unidades_taller(diario, parque, datos)
    if not unidades:
        st.warning("No hay unidades en el parque. Carga el control de combustible o captura un número económico.")
        return
    fid = int(st.session_state.get("veh_ing_id") or 0)
    c1, c2, c3 = st.columns(3)
    with c1:
        unidad = st.selectbox("Número económico / unidad", unidades, key=f"ing_un_{fid}")
        placas = st.text_input("Placas", value=str(((datos.get("unidades") or {}).get(unidad) or {}).get("placas") or ""), key=f"ing_pl_{fid}")
        fecha_ent = st.date_input("Fecha de entrada", value=date.today(), format="DD/MM/YYYY", key=f"ing_fe_{fid}")
    with c2:
        km_ent = st.number_input(
            "Kilometraje de entrada",
            min_value=0.0,
            value=float(ultimo_odometro_unidad(diario, unidad) or 0),
            step=1.0,
            key=f"ing_km_{fid}",
        )
        horo_ent = st.number_input("Horómetro de entrada (h)", min_value=0.0, value=0.0, step=1.0, key=f"ing_ho_{fid}")
        responsables = opciones_responsables_taller()
        responsable = st.selectbox("Responsable / cuadrilla que tenía la unidad", responsables, key=f"ing_res_{fid}")
    with c3:
        tipo_lbl = st.radio(
            "Tipo de mantenimiento",
            ["Preventivo (refacciones en almacén)", "Correctivo (puede faltar pieza)", "Servicio externo"],
            key=f"ing_tipo_{fid}",
        )
        if "Preventivo" in tipo_lbl:
            tipo = TIPO_TALLER_PREVENTIVO
            semaforo_def = SEMAFORO_PRECAUCION
        elif "externo" in tipo_lbl:
            tipo = TIPO_TALLER_EXTERNO
            semaforo_def = SEMAFORO_EXTERNO
        else:
            tipo = TIPO_TALLER_CORRECTIVO
            semaforo_def = SEMAFORO_DETENIDA
        semaforo_ops = [SEMAFORO_PRECAUCION, SEMAFORO_DETENIDA, SEMAFORO_EXTERNO]
        semaforo = st.selectbox(
            "Semáforo de la unidad",
            semaforo_ops,
            index=semaforo_ops.index(semaforo_def),
            format_func=lambda s: ETIQUETA_SEMAFORO_UNIDAD[s],
            key=f"ing_sem_{fid}",
        )
        mecanicos = opciones_mecanicos_taller()
        mecanico = st.selectbox("Mecánico asignado", mecanicos, key=f"ing_mec_{fid}")

    falla = st.text_area("Descripción de daños o fallas", height=90, key=f"ing_falla_{fid}")
    p1, p2 = st.columns(2)
    with p1:
        fecha_prog = st.date_input("Fecha de programación (intervención)", value=date.today(), format="DD/MM/YYYY", key=f"ing_fp_{fid}")
    with p2:
        horas = st.number_input("Tiempo asignado / lead time (horas estimadas)", min_value=1.0, value=8.0, step=1.0, key=f"ing_hrs_{fid}")

    catalogo = [item["nombre"] for item in catalogo_insumos_activos()]
    refs: list[dict] = []
    solicitud: dict = {}
    externo: dict = {}
    en_almacen = True
    if tipo == TIPO_TALLER_PREVENTIVO:
        st.caption("El preventivo asume que las refacciones del programa ya están en almacén. Solo agenda fecha y mecánico.")
        en_almacen = st.checkbox("Confirmo que las refacciones del programa están en almacén", value=True, key=f"ing_alm_{fid}")
        if catalogo:
            elegidas = st.multiselect("Refacciones del programa (catálogo de almacén)", catalogo, key=f"ing_prev_ref_{fid}")
            refs = [{"nombre": n, "cantidad": 1, "en_almacen": True} for n in elegidas]
        if not en_almacen:
            st.warning("Si falta material, el ingreso pasará a la cola de solicitud de refacciones.")
    elif tipo == TIPO_TALLER_CORRECTIVO:
        st.caption("Si alguna pieza no tiene existencia, la unidad entra a la cola de solicitud de refacciones.")
        editor = st.data_editor(
            pd.DataFrame([{"nombre": catalogo[0] if catalogo else "", "cantidad": 1}]),
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            key=f"ing_corr_ref_{fid}",
            column_config={
                "nombre": st.column_config.TextColumn("Refacción / pieza"),
                "cantidad": st.column_config.NumberColumn("Cantidad", min_value=1, step=1),
            },
        )
        faltantes = []
        for _, fila in editor.iterrows():
            nombre = str(fila.get("nombre") or "").strip()
            if not nombre:
                continue
            cant = _num_combustible(fila.get("cantidad")) or 1
            hay = existencia_insumo(nombre)
            ok = hay >= cant
            refs.append({"nombre": nombre, "cantidad": cant, "en_almacen": ok, "existencia": hay})
            if not ok:
                faltantes.append(nombre)
        if faltantes:
            st.warning("Sin existencia suficiente: " + ", ".join(faltantes) + ". Se generará solicitud de refacciones.")
    else:
        e1, e2 = st.columns(2)
        with e1:
            motivo = st.text_area("Motivo de envío a taller externo", height=80, key=f"ing_ext_mot_{fid}")
            taller_ext = st.text_input("Taller / proveedor externo", key=f"ing_ext_t_{fid}")
        with e2:
            cotizacion = st.text_input("No. de cotización / presupuesto", key=f"ing_ext_cot_{fid}")
            autorizacion = st.selectbox("Estatus de autorización", AUTORIZACION_EXTERNO, key=f"ing_ext_aut_{fid}")
            fecha_envio = st.date_input("Fecha de envío", value=date.today(), format="DD/MM/YYYY", key=f"ing_ext_fe_{fid}")
        externo = {
            "motivo": str(motivo or "").strip(),
            "taller": str(taller_ext or "").strip(),
            "cotizacion": str(cotizacion or "").strip(),
            "autorizacion": autorizacion,
            "fecha_envio": fecha_envio.isoformat(),
            "fecha_retorno": "",
        }

    st.markdown("**Evidencia fotográfica**")
    f_antes, f_dur, f_des = st.columns(3)
    with f_antes:
        fotos_antes = st.file_uploader("Antes", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key=f"ing_fa_{fid}")
    with f_dur:
        fotos_durante = st.file_uploader("Durante", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key=f"ing_fd_{fid}")
    with f_des:
        fotos_despues = st.file_uploader("Después", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key=f"ing_fev_{fid}")

    if st.button("Registrar ingreso a taller", type="primary", key=f"ing_ok_{fid}"):
        if not str(falla).strip():
            st.warning("Describe los daños o la falla de entrada.")
            return
        if tipo == TIPO_TALLER_EXTERNO and not (externo.get("motivo") and externo.get("taller")):
            st.warning("Captura motivo y taller/proveedor externo.")
            return
        folio = folio_ingreso_vehicular(datos)
        faltan = [r for r in refs if not r.get("en_almacen")]
        if tipo == TIPO_TALLER_PREVENTIVO and en_almacen and not faltan:
            estatus = ESTATUS_INGRESO_TALLER
        elif tipo == TIPO_TALLER_EXTERNO:
            estatus = ESTATUS_INGRESO_EXTERNO
        elif faltan or (tipo == TIPO_TALLER_PREVENTIVO and not en_almacen):
            estatus = ESTATUS_INGRESO_ESPERA
            solicitud = {
                "id": folio_solicitud_ref(datos),
                "estatus": SOLICITUD_PENDIENTE,
                "fecha": fecha_ent.isoformat(),
                "piezas": [{"nombre": r["nombre"], "cantidad": r.get("cantidad") or 1} for r in (faltan or refs)],
            }
        else:
            estatus = ESTATUS_INGRESO_TALLER
        evidencias = {
            "antes": guardar_evidencias_taller(folio, "antes", fotos_antes or []),
            "durante": guardar_evidencias_taller(folio, "durante", fotos_durante or []),
            "despues": guardar_evidencias_taller(folio, "despues", fotos_despues or []),
        }
        ingreso = {
            "id": folio,
            "unidad": unidad,
            "placas": placas.strip(),
            "fecha_entrada": fecha_ent.isoformat(),
            "fecha_salida": "",
            "km_entrada": km_ent,
            "km_salida": 0,
            "horometro_entrada": horo_ent,
            "horometro_salida": 0,
            "responsable": responsable,
            "falla": str(falla).strip(),
            "tipo": tipo,
            "semaforo": semaforo,
            "fecha_programacion": fecha_prog.isoformat(),
            "horas_estimadas": horas,
            "mecanico": mecanico,
            "estatus": estatus,
            "ot_id": "",
            "refacciones": refs,
            "solicitud": solicitud,
            "externo": externo,
            "evidencias": evidencias,
            "origen": "taller",
        }
        persistir_ingreso_taller(datos, ingreso, crear_ot=True)
        st.session_state.veh_ing_id = fid + 1
        st.session_state.veh_aviso = (
            f"{folio} registrada para {unidad}. "
            + ETIQUETA_ESTATUS_INGRESO.get(estatus, "")
        )
        st.rerun()


def render_tablero_taller(datos: dict) -> None:
    abiertos = ingresos_abiertos(datos)
    if not abiertos:
        st.info("No hay unidades en taller. El diagnóstico de entrada crea el folio y la OT vinculada.")
        return
    filtro = st.selectbox(
        "Filtrar por tipo",
        ["Todos", "Preventivo", "Correctivo", "Servicio externo"],
        key="taller_tablero_tipo",
    )
    for ing in abiertos:
        if filtro == "Preventivo" and ing.get("tipo") != TIPO_TALLER_PREVENTIVO:
            continue
        if filtro == "Correctivo" and ing.get("tipo") != TIPO_TALLER_CORRECTIVO:
            continue
        if filtro == "Servicio externo" and ing.get("tipo") != TIPO_TALLER_EXTERNO:
            continue
        render_ficha_ingreso_taller(ing)
        render_evidencias_taller(ing)
        folio = str(ing.get("id"))
        with st.expander(f"Actualizar / liberar {folio}", expanded=False):
            n1, n2, n3 = st.columns(3)
            with n1:
                nueva_prog = st.date_input(
                    "Fecha de programación",
                    value=parsear_fecha_valor(ing.get("fecha_programacion")) or date.today(),
                    format="DD/MM/YYYY",
                    key=f"tb_fp_{folio}",
                )
                nuevas_hrs = st.number_input(
                    "Horas estimadas",
                    min_value=1.0,
                    value=float(_num_combustible(ing.get("horas_estimadas")) or 8),
                    step=1.0,
                    key=f"tb_h_{folio}",
                )
            with n2:
                mecs = opciones_mecanicos_taller()
                mec_act = str(ing.get("mecanico") or "")
                idx_m = mecs.index(mec_act) if mec_act in mecs else 0
                nuevo_mec = st.selectbox("Mecánico", mecs, index=idx_m, key=f"tb_m_{folio}")
                fotos_d = st.file_uploader("Agregar fotos Durante", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key=f"tb_fd_{folio}")
            with n3:
                km_sal = st.number_input("Km de salida", min_value=0.0, value=float(_num_combustible(ing.get("km_entrada"))), step=1.0, key=f"tb_kms_{folio}")
                horo_sal = st.number_input("Horómetro de salida", min_value=0.0, value=float(_num_combustible(ing.get("horometro_entrada"))), step=1.0, key=f"tb_hs_{folio}")
                fotos_des = st.file_uploader("Fotos Después", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key=f"tb_fe_{folio}")
            notas = st.text_input("Nota de liberación / reporte", key=f"tb_n_{folio}")
            b1, b2 = st.columns(2)
            with b1:
                if st.button("Guardar planeación", key=f"tb_save_{folio}"):
                    ing["fecha_programacion"] = nueva_prog.isoformat()
                    ing["horas_estimadas"] = nuevas_hrs
                    ing["mecanico"] = nuevo_mec
                    evidencias = dict(ing.get("evidencias") or {"antes": [], "durante": [], "despues": []})
                    evidencias["durante"] = list(evidencias.get("durante") or []) + guardar_evidencias_taller(folio, "durante", fotos_d or [])
                    evidencias["despues"] = list(evidencias.get("despues") or []) + guardar_evidencias_taller(folio, "despues", fotos_des or [])
                    ing["evidencias"] = evidencias
                    persistir_ingreso_taller(datos, ing)
                    st.session_state.veh_aviso = f"Planeación de {folio} actualizada."
                    st.rerun()
            with b2:
                if st.button("Liberar unidad (cierre)", type="primary", key=f"tb_close_{folio}"):
                    evidencias = dict(ing.get("evidencias") or {"antes": [], "durante": [], "despues": []})
                    evidencias["despues"] = list(evidencias.get("despues") or []) + guardar_evidencias_taller(folio, "despues", fotos_des or [])
                    ing["evidencias"] = evidencias
                    persistir_ingreso_taller(datos, ing)
                    cerrar_ingreso_taller(datos, folio, km_sal, horo_sal, notas)
                    st.session_state.veh_aviso = f"{folio} cerrada. {ing.get('unidad')} quedó operativa."
                    st.rerun()


def render_cola_refacciones(datos: dict) -> None:
    cola = [
        ing
        for ing in ingresos_abiertos(datos)
        if str((ing.get("solicitud") or {}).get("estatus") or "") == SOLICITUD_PENDIENTE
        or str(ing.get("estatus")) == ESTATUS_INGRESO_ESPERA
    ]
    st.caption("Unidades detenidas o en espera porque la pieza no está en almacén. Al surtir, pasan a programación de taller.")
    if not cola:
        st.success("No hay solicitudes de refacción pendientes.")
        return
    for ing in cola:
        solicitud = dict(ing.get("solicitud") or {})
        piezas = solicitud.get("piezas") or ing.get("refacciones") or []
        render_ficha_ingreso_taller(ing)
        if piezas:
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Pieza": p.get("nombre"),
                            "Cantidad": p.get("cantidad") or 1,
                            "Existencia almacén": existencia_insumo(str(p.get("nombre") or "")),
                        }
                        for p in piezas
                        if isinstance(p, dict)
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
        folio = str(ing.get("id"))
        if st.button(f"Marcar refacción surtida · {folio}", key=f"sol_ok_{folio}"):
            solicitud["estatus"] = SOLICITUD_SURTIDA
            solicitud["fecha_surtido"] = date.today().isoformat()
            ing["solicitud"] = solicitud
            ing["estatus"] = ESTATUS_INGRESO_TALLER
            persistir_ingreso_taller(datos, ing)
            st.session_state.veh_aviso = f"{folio}: material surtido. La unidad entra a taller local."
            st.rerun()


def render_seguimiento_externo(datos: dict) -> None:
    externos = [
        ing for ing in ingresos_abiertos(datos)
        if ing.get("tipo") == TIPO_TALLER_EXTERNO or str(ing.get("estatus")) == ESTATUS_INGRESO_EXTERNO
    ]
    st.caption("Reparaciones que rebasan la capacidad local: cotización, autorización y retorno a base.")
    if not externos:
        st.info("No hay unidades en servicio externo.")
        return
    for ing in externos:
        render_ficha_ingreso_taller(ing)
        folio = str(ing.get("id"))
        ext = dict(ing.get("externo") or {})
        with st.expander(f"Seguimiento externo {folio}", expanded=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                taller = st.text_input("Taller / proveedor", value=str(ext.get("taller") or ""), key=f"ex_t_{folio}")
                cotizacion = st.text_input("Cotización", value=str(ext.get("cotizacion") or ""), key=f"ex_c_{folio}")
            with c2:
                auts = AUTORIZACION_EXTERNO
                actual = str(ext.get("autorizacion") or auts[0])
                autorizacion = st.selectbox("Autorización", auts, index=auts.index(actual) if actual in auts else 0, key=f"ex_a_{folio}")
                envio = st.date_input(
                    "Fecha de envío",
                    value=parsear_fecha_valor(ext.get("fecha_envio")) or date.today(),
                    format="DD/MM/YYYY",
                    key=f"ex_e_{folio}",
                )
            with c3:
                motivo = st.text_area("Motivo", value=str(ext.get("motivo") or ""), height=90, key=f"ex_m_{folio}")
            if st.button("Guardar seguimiento externo", key=f"ex_save_{folio}"):
                ing["externo"] = {
                    "motivo": motivo.strip(),
                    "taller": taller.strip(),
                    "cotizacion": cotizacion.strip(),
                    "autorizacion": autorizacion,
                    "fecha_envio": envio.isoformat(),
                    "fecha_retorno": ext.get("fecha_retorno") or "",
                }
                if autorizacion == "Retornada a base":
                    ing["externo"]["fecha_retorno"] = date.today().isoformat()
                    ing["semaforo"] = SEMAFORO_PRECAUCION
                    ing["estatus"] = ESTATUS_INGRESO_TALLER
                elif autorizacion in {"Autorizada", "En taller externo"}:
                    ing["semaforo"] = SEMAFORO_EXTERNO
                    ing["estatus"] = ESTATUS_INGRESO_EXTERNO
                persistir_ingreso_taller(datos, ing)
                st.session_state.veh_aviso = f"Seguimiento externo de {folio} actualizado."
                st.rerun()


def folio_asignacion_unidad(datos: dict, anio: int | None = None) -> str:
    anio = anio or date.today().year
    prefijo = f"ASG-VEH-{anio}-"
    maximo = 0
    for asg in datos.get("asignaciones") or []:
        folio = str(asg.get("id") or "")
        if folio.startswith(prefijo):
            try:
                maximo = max(maximo, int(folio[len(prefijo):]))
            except ValueError:
                continue
    return f"{prefijo}{maximo + 1:03d}"


def asignaciones_abiertas(datos: dict, unidad: str = "") -> list[dict]:
    salida = []
    for asg in datos.get("asignaciones") or []:
        if str(asg.get("estatus") or "") == ESTATUS_ASIG_CERRADA:
            continue
        if unidad and str(asg.get("unidad") or "") != unidad:
            continue
        salida.append(asg)
    return sorted(salida, key=lambda a: str(a.get("fecha_salida") or ""), reverse=True)


def asignaciones_de_unidad(datos: dict, unidad: str) -> list[dict]:
    return sorted(
        [asg for asg in (datos.get("asignaciones") or []) if str(asg.get("unidad") or "") == unidad],
        key=lambda a: str(a.get("fecha_salida") or ""),
        reverse=True,
    )


def persistir_asignacion_unidad(datos: dict, asignacion: dict) -> dict:
    lista = datos.setdefault("asignaciones", [])
    idx = next((i for i, item in enumerate(lista) if str(item.get("id")) == str(asignacion.get("id"))), None)
    if idx is None:
        lista.append(asignacion)
    else:
        lista[idx] = asignacion
    guardar_vehiculos(datos)
    return asignacion


def render_evidencias_asignacion(asg: dict) -> None:
    cols = st.columns(2)
    for col, fase, clave in (
        (cols[0], "Salida (check-in)", "fotos_salida"),
        (cols[1], "Cierre (check-out)", "fotos_cierre"),
    ):
        with col:
            st.markdown(f'<div class="taller-fase">{fase}</div>', unsafe_allow_html=True)
            fotos = [f for f in (asg.get(clave) or []) if f]
            if not fotos:
                st.caption("Sin fotos")
                continue
            for ruta in fotos:
                if os.path.isfile(ruta):
                    st.image(ruta, use_container_width=True)
                else:
                    st.caption(os.path.basename(str(ruta)))


def render_ficha_asignacion_unidad(asg: dict) -> None:
    abierta = str(asg.get("estatus") or "") != ESTATUS_ASIG_CERRADA
    tono = "warn" if abierta else "ok"
    km_out = _num_combustible(asg.get("km_salida"))
    km_in = _num_combustible(asg.get("km_cierre"))
    recorridos = _num_combustible(asg.get("km_recorridos"))
    if recorridos <= 0 and km_in > km_out:
        recorridos = km_in - km_out
    st.markdown(
        f"""
        <div class="ot-sheet">
            <div class="ot-kicker">{html.escape(str(asg.get("id") or "ASG"))} ·
            {"En ruta" if abierta else "Cerrada"} · {html.escape(str(asg.get("proposito") or "Asignación"))}</div>
            <h3>{html.escape(str(asg.get("unidad") or "Unidad"))}</h3>
            <div class="ot-grid">
                <div class="ot-field"><div class="k">Cuadrilla / responsable</div>
                <div class="v">{html.escape(str(asg.get("cuadrilla") or "—"))} ·
                {html.escape(str(asg.get("responsable") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Check-in de salida</div>
                <div class="v">{html.escape(html_fecha(asg.get("fecha_salida")))} ·
                {km_out:,.0f} km · comb. {html.escape(str(asg.get("combustible_salida") or "—"))} ·
                fluidos {html.escape(str(asg.get("fluidos_salida") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Check-out de cierre</div>
                <div class="v">{html.escape(html_fecha(asg.get("fecha_cierre")) if asg.get("fecha_cierre") else "Pendiente")} ·
                {km_in:,.0f} km · comb. {html.escape(str(asg.get("combustible_cierre") or "—"))} ·
                fluidos {html.escape(str(asg.get("fluidos_cierre") or "—"))}</div></div>
                <div class="ot-field"><div class="k">Km recorridos</div>
                <div class="v">{recorridos:,.0f} km</div></div>
            </div>
            <div class="ot-field"><div class="k">Observación de salida</div>
            <div class="v">{html.escape(str(asg.get("obs_salida") or "—"))}</div></div>
            <div class="ot-field"><div class="k">Observación de cierre</div>
            <div class="v">{html.escape(str(asg.get("obs_cierre") or "—"))}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _ = tono


def render_control_asignacion_unidades() -> None:
    st.markdown('<div class="section-title">Check-in / check-out de unidades</div>', unsafe_allow_html=True)
    st.caption(
        "Para cuadrillas de campo: recibe la unidad (relevo, jornada o carga de combustible), captura km, "
        "combustible, fluidos y fotos. Al devolverla, el sistema calcula los kilómetros recorridos y lo guarda en el historial."
    )
    diario, parque = cargar_datos_flota()
    datos = cargar_vehiculos()
    aviso = st.session_state.pop("asig_aviso", None) or st.session_state.pop("veh_aviso", None)
    if aviso:
        st.success(aviso)

    abiertas = asignaciones_abiertas(datos)
    hoy = date.today()
    cerradas_hoy = [
        a
        for a in (datos.get("asignaciones") or [])
        if str(a.get("estatus")) == ESTATUS_ASIG_CERRADA and parsear_fecha_valor(a.get("fecha_cierre")) == hoy
    ]
    km_hoy = sum(_num_combustible(a.get("km_recorridos")) for a in cerradas_hoy)
    k1, k2, k3 = st.columns(3, gap="medium")
    k1.markdown(kpi_html("En ruta", str(len(abiertas)), "Asignaciones abiertas", "warn" if abiertas else "ok"), unsafe_allow_html=True)
    k2.markdown(kpi_html("Cerradas hoy", str(len(cerradas_hoy)), "Check-out del día", "info"), unsafe_allow_html=True)
    k3.markdown(kpi_html("Km de hoy", f"{km_hoy:,.0f}", "Recorrido al devolver", "ok"), unsafe_allow_html=True)

    unidades = unidades_taller(diario, parque, datos)
    if not unidades:
        st.warning("No hay números económicos. Revisa el control de combustible o el parque.")
        return

    st.markdown("**1. Check-in de salida · inicio de ruta**")
    fid = int(st.session_state.get("asig_form_id") or 0)
    pref_un = str(st.session_state.get("hangar_unidad") or "")
    pref_pr = st.session_state.pop("asig_proposito_pref", None)
    key_un = f"asg_un_{fid}"
    key_pr = f"asg_pr_{fid}"
    if pref_un and pref_un in unidades and key_un not in st.session_state:
        st.session_state[key_un] = pref_un
    if pref_pr and pref_pr in PROPOSITOS_ASIGNACION and key_pr not in st.session_state:
        st.session_state[key_pr] = pref_pr
    c1, c2 = st.columns(2)
    with c1:
        unidad = st.selectbox("Unidad que se asigna", unidades, key=key_un)
        proposito = st.selectbox("Propósito de la asignación", PROPOSITOS_ASIGNACION, key=key_pr)
        if proposito == PROPOSITO_CARGA_COMBUSTIBLE:
            st.caption("Salida a estación de servicio. Se guardan km, combustible inicial/final, fotos y comentarios en el historial de la unidad.")
        cuadrilla = st.selectbox("Cuadrilla que recibe", CUADRILLAS, key=f"asg_cu_{fid}")
        responsables = opciones_responsables_taller()
        responsable = st.selectbox("Personal responsable", responsables, key=f"asg_re_{fid}")
    with c2:
        fecha_sal = st.date_input("Fecha de salida", value=date.today(), format="DD/MM/YYYY", key=f"asg_fs_{fid}")
        km_sal = st.number_input(
            "Kilometraje de salida",
            min_value=0.0,
            value=float(ultimo_odometro_unidad(diario, unidad) or 0),
            step=1.0,
            key=f"asg_km_{fid}",
        )
        comb_sal = st.select_slider(
            "Nivel de combustible al salir" if proposito == PROPOSITO_CARGA_COMBUSTIBLE else "Nivel de combustible",
            options=NIVELES_COMBUSTIBLE_UNIDAD,
            value="3/4",
            key=f"asg_cb_{fid}",
        )
        flu_sal = st.select_slider("Nivel de fluidos", options=NIVELES_FLUIDOS_UNIDAD, value="Correctos", key=f"asg_fl_{fid}")
    obs_sal = st.text_area(
        "Comentarios de salida (estación, ticket o novedad)" if proposito == PROPOSITO_CARGA_COMBUSTIBLE
        else "Observaciones al recibir (opcional)",
        height=70,
        key=f"asg_os_{fid}",
    )
    fotos_sal = st.file_uploader(
        "Fotos de las condiciones al recibir la unidad",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        key=f"asg_fsf_{fid}",
    )
    if st.button("Registrar check-in de salida", type="primary", key=f"asg_ok_{fid}"):
        if asignaciones_abiertas(datos, unidad):
            st.warning(f"{unidad} ya tiene una asignación abierta. Ciérrala antes de entregar otra vez.")
        else:
            folio = folio_asignacion_unidad(datos)
            asg = {
                "id": folio,
                "unidad": unidad,
                "proposito": proposito,
                "cuadrilla": cuadrilla,
                "responsable": responsable,
                "fecha_salida": fecha_sal.isoformat(),
                "km_salida": km_sal,
                "combustible_salida": comb_sal,
                "fluidos_salida": flu_sal,
                "obs_salida": str(obs_sal or "").strip(),
                "fotos_salida": guardar_evidencias_taller(folio, "salida", fotos_sal or []),
                "fecha_cierre": "",
                "km_cierre": 0,
                "combustible_cierre": "",
                "fluidos_cierre": "",
                "obs_cierre": "",
                "fotos_cierre": [],
                "km_recorridos": 0,
                "estatus": ESTATUS_ASIG_RUTA,
                "origen": "campo",
            }
            persistir_asignacion_unidad(datos, asg)
            st.session_state.asig_form_id = fid + 1
            st.session_state.asig_aviso = f"{folio}: {unidad} quedó asignada a {cuadrilla}. El check-in ya está en el historial."
            st.rerun()

    st.markdown("**2. Check-out de cierre · devolución**")
    if not abiertas:
        st.info("No hay unidades en ruta. El check-out aparece aquí cuando hay un check-in abierto.")
    else:
        for asg in abiertas:
            render_ficha_asignacion_unidad(asg)
            folio = str(asg.get("id"))
            es_carga = str(asg.get("proposito") or "") == PROPOSITO_CARGA_COMBUSTIBLE
            with st.expander(f"Cerrar / devolver {asg.get('unidad')} · {folio}", expanded=len(abiertas) == 1):
                if es_carga:
                    st.caption("Carga de combustible: registra el nivel final después de surtir y, si aplica, ticket o comentario.")
                x1, x2, x3 = st.columns(3)
                with x1:
                    fecha_ci = st.date_input(
                        "Fecha de cierre",
                        value=date.today(),
                        format="DD/MM/YYYY",
                        key=f"asg_fc_{folio}",
                    )
                    km_ci = st.number_input(
                        "Kilometraje de cierre",
                        min_value=float(_num_combustible(asg.get("km_salida"))),
                        value=float(_num_combustible(asg.get("km_salida"))),
                        step=1.0,
                        key=f"asg_kmc_{folio}",
                    )
                with x2:
                    comb_ci = st.select_slider(
                        "Combustible al salir de la estación" if es_carga else "Combustible al entregar",
                        options=NIVELES_COMBUSTIBLE_UNIDAD,
                        value="Lleno" if es_carga else "1/2",
                        key=f"asg_cbc_{folio}",
                    )
                    flu_ci = st.select_slider(
                        "Fluidos al entregar",
                        options=NIVELES_FLUIDOS_UNIDAD,
                        value="Correctos",
                        key=f"asg_flc_{folio}",
                    )
                with x3:
                    recorridos = max(0.0, float(km_ci) - _num_combustible(asg.get("km_salida")))
                    st.metric("Km recorridos", f"{recorridos:,.0f}")
                    fotos_ci = st.file_uploader(
                        "Fotos de cierre (ticket o tanque)" if es_carga else "Fotos de cierre",
                        type=["jpg", "jpeg", "png", "webp"],
                        accept_multiple_files=True,
                        key=f"asg_ffc_{folio}",
                    )
                obs_ci = st.text_input(
                    "Comentarios de cierre (litros, ticket o novedad)" if es_carga else "Observación de devolución",
                    key=f"asg_oc_{folio}",
                )
                if st.button("Registrar check-out y devolver", type="primary", key=f"asg_close_{folio}"):
                    if float(km_ci) < _num_combustible(asg.get("km_salida")):
                        st.warning("El kilometraje de cierre no puede ser menor al de salida.")
                    else:
                        asg["fecha_cierre"] = fecha_ci.isoformat()
                        asg["km_cierre"] = float(km_ci)
                        asg["combustible_cierre"] = comb_ci
                        asg["fluidos_cierre"] = flu_ci
                        asg["obs_cierre"] = str(obs_ci or "").strip()
                        asg["fotos_cierre"] = list(asg.get("fotos_cierre") or []) + guardar_evidencias_taller(
                            folio, "cierre", fotos_ci or []
                        )
                        asg["km_recorridos"] = round(float(km_ci) - _num_combustible(asg.get("km_salida")), 1)
                        asg["estatus"] = ESTATUS_ASIG_CERRADA
                        persistir_asignacion_unidad(datos, asg)
                        st.session_state.asig_aviso = (
                            f"{folio} quitada de ruta. {asg.get('unidad')} recorrió {asg['km_recorridos']:,.0f} km."
                        )
                        st.rerun()

    historial = sorted(
        [a for a in (datos.get("asignaciones") or []) if isinstance(a, dict)],
        key=lambda a: str(a.get("fecha_salida") or ""),
        reverse=True,
    )
    if historial:
        st.markdown("**Historial de asignaciones**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Folio": a.get("id"),
                        "Unidad": a.get("unidad"),
                        "Propósito": a.get("proposito"),
                        "Cuadrilla": a.get("cuadrilla"),
                        "Salida": html_fecha(a.get("fecha_salida")),
                        "Cierre": html_fecha(a.get("fecha_cierre")),
                        "Km salida": _num_combustible(a.get("km_salida")),
                        "Km cierre": _num_combustible(a.get("km_cierre")),
                        "Km recorridos": _num_combustible(a.get("km_recorridos")),
                        "Comb. salida": a.get("combustible_salida") or "—",
                        "Comb. cierre": a.get("combustible_cierre") or "—",
                        "Estatus": "En ruta" if a.get("estatus") == ESTATUS_ASIG_RUTA else "Cerrada",
                    }
                    for a in historial[:40]
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )


TARJETAS_GESTION_TALLER = [
    {
        "id": "semaforo",
        "icono": "🚦",
        "titulo": "Semáforo y días de inactividad",
        "descripcion": "Estatus de cada unidad: precaución, detenida en base o taller externo.",
    },
    {
        "id": "diagnostico",
        "icono": "📋",
        "titulo": "Diagnóstico de entrada",
        "descripcion": "Registra km, responsable, fallas y evidencia fotográfica del ingreso.",
    },
    {
        "id": "tablero",
        "icono": "🛠️",
        "titulo": "Tablero de taller",
        "descripcion": "Programa fecha, lead time y mecánico; libera la unidad al cierre.",
    },
    {
        "id": "refacciones",
        "icono": "📦",
        "titulo": "Solicitud de refacciones",
        "descripcion": "Cola de compras cuando la pieza no está en almacén.",
    },
    {
        "id": "externo",
        "icono": "🏭",
        "titulo": "Servicio externo",
        "descripcion": "Cotización, autorización y seguimiento hasta el retorno a base.",
    },
    {
        "id": "asignacion",
        "icono": "🔑",
        "titulo": "Check-in / Check-out de Unidades",
        "descripcion": "Asigna una unidad a la cuadrilla, registra salida y cierra la devolución.",
    },
]


def ingresos_de_unidad(datos: dict, unidad: str) -> list[dict]:
    return sorted(
        [ing for ing in (datos.get("ingresos") or []) if str(ing.get("unidad") or "") == unidad],
        key=lambda i: str(i.get("fecha_entrada") or ""),
        reverse=True,
    )


def filas_filtro_taller(datos: dict, diario: pd.DataFrame, parque: pd.DataFrame, filtro: str) -> list[dict]:
    abiertos = ingresos_abiertos(datos)
    if filtro == "taller":
        origen = abiertos
    elif filtro == "espera":
        origen = [ing for ing in abiertos if str(ing.get("estatus")) == ESTATUS_INGRESO_ESPERA]
    elif filtro == "externo":
        origen = [
            ing
            for ing in abiertos
            if ing.get("tipo") == TIPO_TALLER_EXTERNO or str(ing.get("estatus")) == ESTATUS_INGRESO_EXTERNO
        ]
    elif filtro == "inactividad":
        filas_inact: list[dict] = []
        for unidad in unidades_taller(diario, parque, datos):
            estado = semaforo_efectivo_unidad(datos, unidad)
            dias = dias_downtime(estado["semaforo"], estado["desde"])
            if dias <= 0:
                continue
            filas_inact.append({"unidad": unidad, "ingreso": estado.get("ingreso"), "dias": dias})
        return filas_inact
    else:
        origen = []
    filas = []
    for ing in origen:
        filas.append(
            {
                "unidad": str(ing.get("unidad") or ""),
                "ingreso": ing,
                "dias": dias_downtime(str(ing.get("semaforo") or SEMAFORO_DETENIDA), parsear_fecha_valor(ing.get("fecha_entrada"))),
            }
        )
    return filas


def render_historial_unidad_taller(datos: dict, unidad: str) -> None:
    ingresos = ingresos_de_unidad(datos, unidad)
    ots = [ot for ot in (datos.get("ots") or []) if str(ot.get("unidad") or "") == unidad]
    estado = semaforo_efectivo_unidad(datos, unidad)
    dias = dias_downtime(estado["semaforo"], estado["desde"])
    asignaciones = asignaciones_de_unidad(datos, unidad)
    st.markdown(f'<div class="section-title">Historial · {html.escape(unidad)}</div>', unsafe_allow_html=True)
    st.caption(
        f"{ETIQUETA_SEMAFORO_UNIDAD.get(estado['semaforo'], estado['semaforo'])} · "
        f"{dias} día(s) de inactividad · {len(ingresos)} ingreso(s) · "
        f"{len(asignaciones)} asignación(es) · {len(ots)} OT de flota."
    )
    if not ingresos and not ots and not asignaciones:
        st.info("Esta unidad no tiene ingresos de taller, asignaciones ni OT de flota.")
        return
    if asignaciones:
        st.markdown("**Asignaciones y check-in / check-out**")
        for asg in asignaciones:
            render_ficha_asignacion_unidad(asg)
            render_evidencias_asignacion(asg)
    if ingresos:
        st.markdown("**Ingresos a taller**")
        for ing in ingresos:
            render_ficha_ingreso_taller(ing)
            render_evidencias_taller(ing)
    if ots:
        st.markdown("**Órdenes de trabajo de la unidad**")
        st.dataframe(
            pd.DataFrame(
                [
                    {
                        "Folio": ot.get("id"),
                        "Tipo": etiqueta_tipo_taller(str(ot.get("tipo") or "")),
                        "Estatus": str(ot.get("estatus") or "").replace("_", " ").title(),
                        "Alta": html_fecha(ot.get("fecha_creacion")),
                        "Cierre": html_fecha(ot.get("fecha_cierre")),
                        "Mecánico": ot.get("mecanico") or "—",
                    }
                    for ot in ots
                ]
            ),
            use_container_width=True,
            hide_index=True,
        )


def render_listado_filtro_taller(datos: dict, diario: pd.DataFrame, parque: pd.DataFrame, filtro: str) -> None:
    titulos = {
        "taller": "Unidades en taller",
        "espera": "Unidades en espera de refacción",
        "externo": "Unidades en servicio externo",
        "inactividad": "Unidades con días de inactividad",
    }
    st.markdown(f'<div class="section-title">{titulos.get(filtro, "Unidades")}</div>', unsafe_allow_html=True)
    filas = filas_filtro_taller(datos, diario, parque, filtro)
    if not filas:
        st.info("No hay unidades en esta situación.")
        return
    for fila in filas:
        unidad = fila["unidad"]
        ing = fila.get("ingreso")
        extra = (
            f"{etiqueta_tipo_taller(str(ing.get('tipo')))} · {ing.get('id')} · "
            f"{ETIQUETA_ESTATUS_INGRESO.get(str(ing.get('estatus') or ''), '')}"
            if ing
            else "Sin folio de ingreso abierto"
        )
        st.markdown(
            f"""
            <div class="veh-alert {TONO_ALERTA_SEMAFORO.get(str((ing or {}).get('semaforo') or SEMAFORO_DETENIDA), 'due')}">
                <div class="kicker">{html.escape(unidad)} · {fila['dias']} día(s) de inactividad</div>
                <h3>{html.escape(unidad)}</h3>
                <p>{html.escape(extra)}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button(f"Ver historial de {unidad}", key=f"taller_hist_{filtro}_{unidad}", use_container_width=True):
            st.session_state.taller_unidad = unidad
            st.rerun()


def _toggle_filtro_taller(filtro: str) -> None:
    if st.session_state.get("taller_filtro") == filtro and not st.session_state.get("taller_unidad"):
        st.session_state.taller_filtro = None
    else:
        st.session_state.taller_filtro = filtro
        st.session_state.taller_seccion = None
    st.session_state.taller_unidad = None


def render_kpi_filtro_taller(titulo: str, valor: str, detalle: str, tono: str, filtro: str, key: str) -> None:
    activo = st.session_state.get("taller_filtro") == filtro
    st.markdown(
        kpi_html(titulo, valor, detalle, tono, extra_clase="activa" if activo else ""),
        unsafe_allow_html=True,
    )
    etiqueta = "Ocultar listado" if activo and not st.session_state.get("taller_unidad") else "Ver unidades"
    if st.button(etiqueta, key=key, use_container_width=True):
        _toggle_filtro_taller(filtro)
        st.rerun()


def html_vistazo_flota() -> str:
    return """
        <div class="flota-vistazo">
            <div class="lbl">Vistazo general</div>
            <div class="silueta">
                <svg viewBox="0 0 240 110" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                    <path d="M18 72h14c3 12 14 20 26 20s23-8 26-20h56c3 12 14 20 26 20s23-8 26-20h16"
                        stroke="#38bdf8" stroke-width="3.2" stroke-linecap="round"/>
                    <path d="M22 72V54c0-6 4-12 10-14l28-8h58l36 22h30c6 0 10 4 10 10v8"
                        stroke="#86efac" stroke-width="3.2" stroke-linejoin="round"/>
                    <circle cx="58" cy="78" r="11" stroke="#7dd3fc" stroke-width="3"/>
                    <circle cx="166" cy="78" r="11" stroke="#7dd3fc" stroke-width="3"/>
                    <path d="M62 42h48" stroke="#93c5fd" stroke-width="2.4" stroke-linecap="round"/>
                </svg>
            </div>
            <h3>Resumen de flota</h3>
            <p>Consulta el taller, las compras y las unidades fuera de operación. Elige un recuadro para ver las unidades de esa situación.</p>
        </div>
        """


def render_gestion_parque_vehicular() -> None:
    st.markdown('<div class="section-title">Gestión del parque vehicular</div>', unsafe_allow_html=True)
    st.caption(
        "Taller, mecánicos, compras y auditoría: diagnóstico de entrada, preventivo/correctivo/externo, "
        "programación, semáforo de operación y días de inactividad."
    )
    diario, parque = cargar_datos_flota()
    datos = cargar_vehiculos()
    aviso = st.session_state.pop("veh_aviso", None)
    if aviso:
        st.success(aviso)

    abiertos = ingresos_abiertos(datos)
    espera = [i for i in abiertos if i.get("estatus") == ESTATUS_INGRESO_ESPERA]
    externos = [
        i
        for i in abiertos
        if i.get("tipo") == TIPO_TALLER_EXTERNO or str(i.get("estatus")) == ESTATUS_INGRESO_EXTERNO
    ]
    dias = 0
    for unidad in unidades_taller(diario, parque, datos):
        est = semaforo_efectivo_unidad(datos, unidad)
        dias += dias_downtime(est["semaforo"], est["desde"])

    vistazo, kpis = st.columns([1.05, 1.15], gap="large")
    with vistazo:
        st.markdown(html_vistazo_flota(), unsafe_allow_html=True)
    with kpis:
        r1a, r1b = st.columns(2, gap="medium")
        with r1a:
            render_kpi_filtro_taller("En taller", str(len(abiertos)), "Ingresos abiertos", "info", "taller", "kpi_taller_abiertos")
        with r1b:
            render_kpi_filtro_taller(
                "Espera refacción",
                str(len(espera)),
                "Cola de compras",
                "warn" if espera else "ok",
                "espera",
                "kpi_taller_espera",
            )
        r2a, r2b = st.columns(2, gap="medium")
        with r2a:
            render_kpi_filtro_taller("Servicio externo", str(len(externos)), "Fuera de base", "info", "externo", "kpi_taller_ext")
        with r2b:
            render_kpi_filtro_taller(
                "Días de inactividad",
                f"{dias} d",
                "Detenidas + externas",
                "danger" if dias else "ok",
                "inactividad",
                "kpi_taller_dias",
            )

    filtro = st.session_state.get("taller_filtro")
    unidad_sel = st.session_state.get("taller_unidad")
    if unidad_sel:
        if st.button("← Volver al listado", key="taller_back_unidad"):
            st.session_state.taller_unidad = None
            st.rerun()
        render_historial_unidad_taller(datos, str(unidad_sel))
    elif filtro:
        render_listado_filtro_taller(datos, diario, parque, str(filtro))

    st.markdown('<div class="flota-nav-sep"></div>', unsafe_allow_html=True)
    seccion = st.session_state.get("taller_seccion")
    if not seccion:
        st.markdown('<p class="explorer-meta">Elige un apartado del taller para abrir el detalle.</p>', unsafe_allow_html=True)
        for inicio in range(0, len(TARJETAS_GESTION_TALLER), 3):
            cols = st.columns(3, gap="medium")
            for col, modulo in zip(cols, TARJETAS_GESTION_TALLER[inicio : inicio + 3]):
                with col:
                    st.markdown(
                        f"""
                        <div class="explorer-card">
                            <div class="icon-tile">{modulo["icono"]}</div>
                            <h3>{html.escape(modulo["titulo"])}</h3>
                            <p>{html.escape(modulo["descripcion"])}</p>
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )
                    if st.button("Abrir", key=f"explorar_taller_seccion_{modulo['id']}", use_container_width=True):
                        st.session_state.taller_seccion = modulo["id"]
                        st.session_state.taller_filtro = None
                        st.session_state.taller_unidad = None
                        st.rerun()
        return
    render_barra_volver(
        f'Taller / <b>{html.escape(titulo_submenu(TARJETAS_GESTION_TALLER, seccion))}</b>',
        "taller_seccion",
        "← Volver a apartados",
    )
    if seccion == "semaforo":
        render_semaforo_parque(datos, diario, parque)
    elif seccion == "diagnostico":
        render_form_ingreso_taller(datos, diario, parque)
    elif seccion == "tablero":
        render_tablero_taller(datos)
    elif seccion == "refacciones":
        render_cola_refacciones(datos)
    elif seccion == "externo":
        render_seguimiento_externo(datos)
    elif seccion == "asignacion":
        render_control_asignacion_unidades()
    else:
        st.session_state.taller_seccion = None
        st.rerun()


def tarjetas_menu_vehiculos(datos: dict, diario: pd.DataFrame, parque: pd.DataFrame) -> list[dict]:
    monitores = monitorear_flota(diario, parque, datos)
    n_alerta = sum(1 for m in monitores if m["estado"] in {"alerta", "vencido"})
    n_corr = len([o for o in ots_por_estatus(datos, tipo="correctivo") if o.get("estatus") != ESTATUS_OT_CERRADA])
    n_cerr = len(ots_por_estatus(datos, estatus=ESTATUS_OT_CERRADA))
    n_taller = len(ingresos_abiertos(datos))
    n_espera = len(
        [i for i in ingresos_abiertos(datos) if str(i.get("estatus")) == ESTATUS_INGRESO_ESPERA]
    )
    n_asig = len(asignaciones_abiertas(datos))
    hoy = date.today()
    inicio = date(hoy.year, hoy.month, 1)
    met = metricas_combustible(filtrar_registro_combustible(diario, None, inicio, hoy, "Todos"))
    gasto_ot_mes = 0.0
    for ot in datos.get("ots") or []:
        ref = parsear_fecha_valor(ot.get("fecha_cierre") or ot.get("fecha_creacion"))
        if ref and ref.year == hoy.year and ref.month == hoy.month:
            gasto_ot_mes += costo_ot_vehicular(ot)
    return [
        {
            "id": "taller",
            "icono": "🛠️",
            "titulo": "Gestión del Parque Vehicular",
            "descripcion": "Diagnóstico de entrada, preventivo/correctivo/externo, semáforo, refacciones y días de inactividad.",
            "badge": (
                f"{n_taller} en taller · {n_espera} espera refacción"
                if n_taller else "Sin ingresos abiertos"
            ),
            "badge_tono": "due" if n_espera else ("warn" if n_taller else ""),
        },
        {
            "id": "asignacion",
            "icono": "🔑",
            "titulo": "Check-in / Check-out de Unidades",
            "descripcion": "Asigna unidad a la cuadrilla, registra salida y cierra la devolución con km y fotos.",
            "badge": f"{n_asig} en ruta" if n_asig else "Sin unidades en ruta",
            "badge_tono": "warn" if n_asig else "",
        },
        {
            "id": "combustible",
            "icono": "⛽",
            "titulo": "Control de Combustible",
            "descripcion": "Cargas, litros, kilometraje y rendimiento real (km/l) de cada número económico.",
            "badge": f"{int(met['cargas'])} cargas este mes" if met["cargas"] else "Sin cargas del mes",
        },
        {
            "id": "preventivos",
            "icono": "🚨",
            "titulo": "Alertas y Preventivos por Kilometraje",
            "descripcion": "Monitorea el odómetro y genera OT preventiva si la unidad está a 500 km o menos del servicio.",
            "badge": f"{n_alerta} unidad(es) en ventana" if n_alerta else "Flota dentro de programa",
            "badge_tono": "due" if n_alerta else "",
        },
        {
            "id": "correctivos",
            "icono": "🔧",
            "titulo": "Mantenimientos Correctivos",
            "descripcion": "Fallas mecánicas imprevistas: taller, refacciones, costos y OT de reparación.",
            "badge": f"{n_corr} OT abiertas" if n_corr else "Sin correctivos abiertos",
            "badge_tono": "warn" if n_corr else "",
        },
        {
            "id": "historial",
            "icono": "📁",
            "titulo": "Historial de OT Cerradas / Realizadas",
            "descripcion": "Consulta órdenes concluidas con reporte final, mecánico responsable y costo total.",
            "badge": f"{n_cerr} OT realizadas" if n_cerr else "Sin OT cerradas",
        },
        {
            "id": "mensuales",
            "icono": "📅",
            "titulo": "Mantenimientos Mensuales",
            "descripcion": "Concentrado mensual de gasto de flota, estado de unidades y OT del periodo.",
            "badge": f"${met['gasto'] + gasto_ot_mes:,.0f} en el mes",
        },
    ]


def render_modulo_vehiculos() -> None:
    st.markdown('<div class="section-title">Vehículos</div>', unsafe_allow_html=True)
    st.caption("Parque vehicular: taller integral, combustible, preventivos, correctivos e historial de OT.")
    diario, parque = cargar_datos_flota()
    datos = sincronizar_ots_preventivas(diario, parque, cargar_vehiculos())
    sub = st.session_state.get("admin_sub_vehiculos")
    tarjetas = tarjetas_menu_vehiculos(datos, diario, parque)
    if not sub:
        render_cuadricula_explorador(
            tarjetas,
            "admin_sub_vehiculos",
            "Elige un apartado del parque vehicular. Las tarjetas muestran el recorte del mes en curso.",
        )
        return
    render_barra_volver(
        f'Vehículos / <b>{html.escape(titulo_submenu(tarjetas, sub))}</b>',
        "admin_sub_vehiculos",
        "← Volver a Vehículos",
    )
    if sub == "taller":
        render_gestion_parque_vehicular()
    elif sub == "asignacion":
        render_control_asignacion_unidades()
    elif sub == "combustible":
        render_control_combustible()
    elif sub == "preventivos":
        render_alertas_preventivos()
    elif sub == "correctivos":
        render_mantenimientos_correctivos()
    elif sub == "historial":
        render_historial_ot_vehiculos()
    elif sub == "mensuales":
        render_mantenimientos_mensuales()
    else:
        st.session_state.admin_sub_vehiculos = None
        st.rerun()


ESTATUS_DASHBOARD = ["Completadas", "En proceso", "Con retraso"]
COLORES_ESTATUS_DASH = {
    "Completadas": "#34d399",
    "En proceso": "#60a5fa",
    "Con retraso": "#fbbf24",
}


def kpi_html(titulo: str, valor: str, detalle: str = "", tono: str = "", extra_clase: str = "") -> str:
    clase = " ".join(p for p in ("kpi-card", tono, extra_clase) if p).strip()
    extra = f'<p class="s">{html.escape(detalle)}</p>' if detalle else ""
    return (
        f'<div class="{clase}"><div class="k">{html.escape(titulo)}</div>'
        f"<h3>{html.escape(valor)}</h3>{extra}</div>"
    )


def estilo_figura_dash(fig, titulo: str = "") -> None:
    fig.update_layout(
        title=dict(text=titulo, font=dict(size=15, color="#f8fafc"), x=0.01, xanchor="left"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.35)",
        font=dict(color="#cbd5e1", family="Segoe UI, system-ui, sans-serif", size=12),
        margin=dict(l=12, r=12, t=52, b=8),
        legend=dict(
            bgcolor="rgba(15,23,42,0.55)",
            bordercolor="#334155",
            orientation="h",
            yanchor="bottom",
            y=-0.22,
            x=0,
        ),
        xaxis=dict(gridcolor="#1e293b", zerolinecolor="#334155", color="#94a3b8"),
        yaxis=dict(gridcolor="#1e293b", zerolinecolor="#334155", color="#94a3b8"),
        hoverlabel=dict(bgcolor="#0f172a", font_color="#e2e8f0", bordercolor="#334155"),
    )


def estatus_ot_campo(fila) -> str:
    revision = str(fila.get("REVISIÓN") or "Pendiente").strip()
    if revision == "Validado":
        return "Completadas"
    if revision == "Con observaciones":
        return "Con retraso"
    return "En proceso"


def estatus_ot_flota(ot: dict) -> str:
    estado = str(ot.get("estatus") or "").strip()
    if estado == ESTATUS_OT_CERRADA:
        return "Completadas"
    estimada = parsear_fecha_valor(ot.get("fecha_estimada"))
    if estimada and estimada < date.today() and estado != ESTATUS_OT_CERRADA:
        return "Con retraso"
    return "En proceso"


def ots_campo_dashboard(inicio: date, fin: date) -> pd.DataFrame:
    df = bitacora_con_fechas(cargar_bitacora())
    if df is None or df.empty or "_fecha" not in df.columns:
        return pd.DataFrame()
    vista = df[df["_fecha"].notna()].copy()
    vista = vista[(vista["_fecha"] >= inicio) & (vista["_fecha"] <= fin)]
    if vista.empty:
        return vista
    vista["Origen"] = "Campo"
    vista["Estatus dash"] = vista.apply(estatus_ot_campo, axis=1)
    vista["Etiqueta"] = vista["ACTIVIDAD CORTA"].astype(str).str.strip()
    vacio = vista["Etiqueta"].isin(["", "nan", "None"])
    vista.loc[vacio, "Etiqueta"] = vista.loc[vacio, "TIPO DE ACTIVIDAD"].astype(str).str.strip()
    vista.loc[vista["Etiqueta"].isin(["", "nan", "None"]), "Etiqueta"] = "Sin actividad"
    vista["Serie"] = vista["_fecha"]
    return vista


def ots_flota_dashboard(inicio: date, fin: date) -> pd.DataFrame:
    datos = cargar_vehiculos()
    filas = []
    for ot in datos.get("ots") or []:
        ref = parsear_fecha_valor(ot.get("fecha_cierre") or ot.get("fecha_creacion"))
        if not ref or ref < inicio or ref > fin:
            continue
        filas.append(
            {
                "Origen": "Flota",
                "CUADRILLA": "Parque vehicular",
                "Etiqueta": "Preventivo" if ot.get("tipo") == "preventivo" else "Correctivo",
                "Estatus dash": estatus_ot_flota(ot),
                "_fecha": ref,
                "Serie": ref,
                "Folio": ot.get("id") or "",
                "Unidad": ot.get("unidad") or "",
            }
        )
    return pd.DataFrame(filas)


def filtrar_ots_dashboard(
    campo: pd.DataFrame,
    flota: pd.DataFrame,
    cuadrilla: str,
    estatus: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if cuadrilla and cuadrilla != "Todas":
        if cuadrilla == "Parque vehicular":
            campo = campo.iloc[0:0] if campo is not None else pd.DataFrame()
        else:
            flota = flota.iloc[0:0] if flota is not None else pd.DataFrame()
            if campo is not None and not campo.empty:
                campo = campo[campo["CUADRILLA"].astype(str) == cuadrilla]
    if estatus and estatus != "Todos":
        if campo is not None and not campo.empty:
            campo = campo[campo["Estatus dash"] == estatus]
        if flota is not None and not flota.empty:
            flota = flota[flota["Estatus dash"] == estatus]
    return (
        campo if campo is not None else pd.DataFrame(),
        flota if flota is not None else pd.DataFrame(),
    )


def cumplimiento_dashboard(inicio: date, fin: date) -> tuple[float, int, int]:
    meses = []
    cursor = date(inicio.year, inicio.month, 1)
    tope = date(fin.year, fin.month, 1)
    while cursor <= tope:
        meses.append(clave_mes(cursor.year, cursor.month))
        cursor = date(cursor.year + (1 if cursor.month == 12 else 0), 1 if cursor.month == 12 else cursor.month + 1, 1)
    filas = [item for item in cargar_programados_ops() if str(item.get("mes", "")) in meses]
    if not filas:
        return 0.0, 0, 0
    cumplidos = sum(1 for item in filas if str(item.get("estatus", "")) == "Cumplido")
    return porcentaje_cumplimiento(filas), cumplidos, len(filas)


def figura_barras_cuadrilla(campo: pd.DataFrame, flota: pd.DataFrame):
    bloques = []
    if campo is not None and not campo.empty:
        copia = campo.copy()
        copia["CUADRILLA"] = copia["CUADRILLA"].map(lambda valor: etiqueta_cuadrilla_historial(str(valor or "")))
        bloques.append(copia.groupby("CUADRILLA", dropna=False).size().reset_index(name="OT"))
    if flota is not None and not flota.empty:
        bloques.append(pd.DataFrame([{"CUADRILLA": "Parque vehicular", "OT": int(len(flota))}]))
    if not bloques:
        return None
    tabla = pd.concat(bloques, ignore_index=True).groupby("CUADRILLA", as_index=False)["OT"].sum()
    nombres = sorted(tabla["CUADRILLA"].astype(str), key=orden_cuadrilla_oficial)
    fig = px.bar(
        tabla,
        x="OT",
        y="CUADRILLA",
        orientation="h",
        color="OT",
        color_continuous_scale=["#16382c", "#34d399"],
        text="OT",
    )
    fig.update_traces(textposition="outside", hovertemplate="%{y}<br>%{x} OT<extra></extra>", marker_line_width=0)
    fig.update_layout(coloraxis_showscale=False, xaxis_title="Órdenes de trabajo", yaxis_title="")
    fig.update_yaxes(categoryorder="array", categoryarray=list(reversed(nombres)))
    estilo_figura_dash(fig, "Distribución de servicios por cuadrilla")
    return fig


def figura_actividades(campo: pd.DataFrame, flota: pd.DataFrame):
    partes = []
    if campo is not None and not campo.empty:
        partes.append(campo[["Etiqueta"]])
    if flota is not None and not flota.empty:
        partes.append(flota[["Etiqueta"]])
    if not partes:
        return None
    tabla = pd.concat(partes, ignore_index=True)
    tabla = tabla[tabla["Etiqueta"].astype(str).str.strip().isin(["", "nan", "None"]) == False]
    conteo = tabla["Etiqueta"].value_counts().head(8).sort_values(ascending=True)
    if conteo.empty:
        return None
    fig = px.bar(
        x=conteo.values,
        y=conteo.index.astype(str),
        orientation="h",
        color=conteo.values,
        color_continuous_scale=["#1e3a5f", "#7dd3fc"],
        text=conteo.values,
    )
    fig.update_traces(textposition="outside", hovertemplate="%{y}<br>%{x} servicios<extra></extra>", marker_line_width=0)
    fig.update_layout(coloraxis_showscale=False, xaxis_title="Frecuencia", yaxis_title="")
    estilo_figura_dash(fig, "Actividades más frecuentes")
    return fig


def figura_tendencia(campo: pd.DataFrame, flota: pd.DataFrame, combustible: pd.DataFrame, inicio: date, fin: date):
    dias = pd.date_range(inicio, fin, freq="D")
    usar_semana = (fin - inicio).days > 45
    def serie_ots() -> pd.Series:
        fechas = []
        if campo is not None and not campo.empty:
            fechas.extend(campo["_fecha"].tolist())
        if flota is not None and not flota.empty:
            fechas.extend(flota["_fecha"].tolist())
        if not fechas:
            return pd.Series(dtype=float)
        s = pd.Series(1, index=pd.to_datetime(fechas))
        return s.resample("W-SUN" if usar_semana else "D").sum()

    ot = serie_ots().reindex(dias if not usar_semana else None, fill_value=0)
    if usar_semana:
        idx = pd.date_range(inicio, fin, freq="W-SUN")
        ot = serie_ots().reindex(idx, fill_value=0)
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=ot.index,
            y=ot.values,
            name="OT generadas",
            mode="lines+markers",
            line=dict(color="#7dd3fc", width=2.4),
            fill="tozeroy",
            fillcolor="rgba(125, 211, 252, 0.18)",
            hovertemplate="%{x|%d/%m/%Y}<br>%{y:.0f} OT<extra></extra>",
        )
    )
    if combustible is not None and not combustible.empty:
        comb = combustible.copy()
        comb["_dia"] = pd.to_datetime(comb["Fecha"]).dt.normalize()
        litros = comb.groupby("_dia")["Litros"].sum()
        if usar_semana:
            litros = litros.resample("W-SUN").sum()
        litros = litros.reindex(ot.index, fill_value=0)
        fig.add_trace(
            go.Scatter(
                x=litros.index,
                y=litros.values,
                name="Litros cargados",
                mode="lines",
                line=dict(color="#86efac", width=2.2),
                yaxis="y2",
                hovertemplate="%{x|%d/%m/%Y}<br>%{y:.1f} L<extra></extra>",
            )
        )
        fig.update_layout(
            yaxis2=dict(
                overlaying="y",
                side="right",
                title="Litros",
                gridcolor="rgba(0,0,0,0)",
                color="#86efac",
            )
        )
    fig.update_layout(yaxis_title="OT", xaxis_title="")
    estilo_figura_dash(fig, "Tendencia de OT y combustible")
    return fig


def figura_donut(campo: pd.DataFrame, flota: pd.DataFrame):
    estados = []
    if campo is not None and not campo.empty:
        estados.extend(campo["Estatus dash"].tolist())
    if flota is not None and not flota.empty:
        estados.extend(flota["Estatus dash"].tolist())
    if not estados:
        return None
    conteo = pd.Series(estados).value_counts().reindex(ESTATUS_DASHBOARD, fill_value=0)
    conteo = conteo[conteo > 0]
    if conteo.empty:
        return None
    fig = go.Figure(
        go.Pie(
            labels=list(conteo.index),
            values=list(conteo.values),
            hole=0.62,
            marker=dict(colors=[COLORES_ESTATUS_DASH[k] for k in conteo.index], line=dict(color="#0f172a", width=2)),
            textinfo="percent",
            hovertemplate="%{label}<br>%{value} OT (%{percent})<extra></extra>",
        )
    )
    fig.update_layout(
        annotations=[
            dict(
                text=f"<b>{int(conteo.sum())}</b><br>OT",
                x=0.5,
                y=0.5,
                font=dict(size=16, color="#f8fafc"),
                showarrow=False,
            )
        ]
    )
    estilo_figura_dash(fig, "Estatus de las órdenes de trabajo")
    return fig


_ESPEJO_INSUMO_BITACORA = {
    "ACEITE (LTS)": "ACEITE SAE 50 PJP4",
    "BANDAS CANTIDAD": "BANDAS C144 PJP4",
    "GRASA (KG)": "GRASA OPS",
    "REFRIGERANTE (LTS)": "REFRIGERANTE OPS",
}
_ETIQUETAS_REFACCION = {
    "aceite sae 50 pjp4": "Aceite SAE 50",
    "aceite lts": "Aceite SAE 50",
    "bandas c144 pjp4": "Bandas C144",
    "bandas c112 pjp4": "Bandas C112",
    "bandas cantidad": "Bandas",
    "bomba de agua reparada pjp4": "Bomba de agua",
    "filtro de aceite pjp4": "Filtro de aceite",
    "filtro de aceite ops": "Filtro de aceite OPS",
    "filtros pzas": "Filtro de aceite",
    "bujias pjp4": "Bujías",
    "bujias pzas": "Bujías",
    "juego de cables para bujias": "Cables de bujías",
    "grasa ops": "Grasa (Kg)",
    "grasa kg": "Grasa (Kg)",
    "refrigerante ops": "Refrigerante",
    "refrigerante lts": "Refrigerante",
    "alternador cantidad": "Alternador",
    "filtro de aire": "Filtro de aire",
    "tapa de distribuidor": "Tapa de distribuidor",
    "rotor": "Rotor",
    "polea loca": "Polea loca",
    "boton de arranque": "Botón de arranque",
}


RUBRO_INSUMO = "Insumo"
RUBRO_REPARACION = "Reparación"
_CLAVES_REPARACION_MCI = (
    "alternador",
    "radiador",
    "bomba de agua",
    "bomba de aceite",
    "bomba de combustible",
    "marcha",
    "motor de arranque",
    "boton de arranque",
    "arranque",
    "tapa de distribuidor",
    "distribuidor",
    "rotor",
    "polea",
    "bobina",
    "carburador",
    "culata",
    "piston",
    "biela",
    "ciguenal",
    "volante",
    "embrague",
    "termostato",
    "inyector",
    "solenoide",
    "regulador",
    "reparad",
    "sustitui",
)


def _es_grasa(nombre: str) -> bool:
    return "grasa" in _sin_acentos(str(nombre or "")).casefold()


def clasificar_rubro_refaccion(nombre: str) -> str:
    """Separa consumibles de componentes del MCI reparados o sustituidos.

    Una columna nueva de la bitácora entra sola: si su nombre corresponde a una
    pieza del motor (alternador, radiador, bomba, marcha, etc.) se cuenta como
    reparación; el resto se cuenta como insumo.
    """
    clave = _clave_columna_bitacora(nombre)
    if any(token in clave for token in _CLAVES_REPARACION_MCI):
        return RUBRO_REPARACION
    return RUBRO_INSUMO


def cantidad_grasa_kg(valor) -> float:
    """Pasa a kilogramos la grasa capturada en gramos.

    Un valor mayor de 100, o con unidad g/gramos, se divide entre 1,000.
    Un valor de 100 o menos, o marcado en kg, se conserva en kilogramos.
    """
    if isinstance(valor, bool):
        return 0.0
    if isinstance(valor, (int, float)):
        try:
            if pd.isna(valor):
                return 0.0
        except TypeError:
            pass
        numero = float(valor)
        if numero <= 0:
            return 0.0
        return numero / 1000.0 if numero > 100 else numero
    texto = str(valor or "").strip().replace(",", ".")
    if not texto or texto.lower() in {"nan", "none", "<na>", "nat"}:
        return 0.0
    coincidencia = re.search(
        r"(\d+(?:\.\d+)?)\s*(kilogramos?|kilos?|kg|gramos?|grs|gr|g)?\b",
        texto,
        flags=re.IGNORECASE,
    )
    if not coincidencia:
        return 0.0
    numero = float(coincidencia.group(1))
    if numero <= 0:
        return 0.0
    unidad = (coincidencia.group(2) or "").casefold()
    if unidad in {"kg", "kilo", "kilos", "kilogramo", "kilogramos"}:
        return numero
    if unidad in {"g", "gr", "grs", "gramo", "gramos"} or numero > 100:
        return numero / 1000.0
    return numero


def etiqueta_refaccion(columna: str) -> str:
    if _es_grasa(columna):
        return "Grasa (Kg)"
    clave = _clave_columna_bitacora(columna)
    if clave in _ETIQUETAS_REFACCION:
        return _ETIQUETAS_REFACCION[clave]
    if columna in ETIQUETAS_INSUMO_SOPORTE:
        return ETIQUETAS_INSUMO_SOPORTE[columna]
    texto = re.sub(r"\s*(\(|\b)(PJP4|OPS)\)?\s*$", "", str(columna or ""), flags=re.IGNORECASE).strip()
    palabras = []
    for indice, palabra in enumerate(texto.split()):
        plano = _sin_acentos(palabra).casefold()
        if indice and plano in {"de", "para", "del", "la", "el", "y", "a"}:
            palabras.append(palabra.lower())
        elif re.search(r"\d", palabra):
            palabras.append(palabra.upper())
        else:
            palabras.append(palabra.capitalize())
    titulo = " ".join(palabras).replace("Bujias", "Bujías").replace("Bujía", "Bujía")
    return titulo or str(columna or "").strip()


def _es_columna_cantidad_refaccion(nombre: str) -> bool:
    clave = _clave_columna_bitacora(nombre)
    if not clave or "descripcion" in clave or "estado origen" in clave or clave == "otros insumos":
        return False
    if nombre in COLUMNAS_INSUMOS or nombre in {"BANDAS CANTIDAD", "ALTERNADOR CANTIDAD"}:
        return True
    return es_columna_insumo_excel(nombre)


def _numero_insumo(valor) -> float:
    if valor is None:
        return 0.0
    try:
        if pd.isna(valor):
            return 0.0
    except TypeError:
        pass
    texto = str(valor).strip().replace(",", ".")
    if not texto or texto.lower() in {"nan", "none", "<na>", "nat"}:
        return 0.0
    try:
        numero = float(texto)
    except ValueError:
        return 0.0
    return numero if numero > 0 else 0.0


def _partes_otros_insumos(valor) -> list[tuple[str, float]]:
    texto = texto_visible(valor, "")
    if not texto:
        return []
    salida = []
    for parte in re.split(r"[,;\n|/]+", texto):
        pieza = parte.strip()
        if not pieza or re.fullmatch(r"\d+(?:[.,]\d+)?", pieza):
            continue
        numeros = re.findall(r"\d+(?:[.,]\d+)?", pieza)
        cantidad = float(numeros[-1].replace(",", ".")) if numeros else 1.0
        nombre = re.sub(r"\d+(?:[.,]\d+)?", " ", pieza)
        nombre = re.sub(r"\b(kg|kilos?|kilogramos?|gramos?|grs|gr|g)\b", " ", nombre, flags=re.IGNORECASE)
        nombre = re.sub(r"\s+", " ", nombre).strip(" -:xX")
        if not nombre or cantidad <= 0:
            continue
        if _es_grasa(nombre) or _es_grasa(pieza):
            cantidad = cantidad_grasa_kg(pieza)
            if cantidad <= 0:
                continue
        salida.append((etiqueta_refaccion(nombre), cantidad))
    return salida


def _ubicacion_fila_refaccion(fila) -> str:
    return etiqueta_ubicacion_fila(fila) or "Sin ubicación"


def detalle_insumos_suministrados(df: pd.DataFrame | None) -> pd.DataFrame:
    vacio = pd.DataFrame(columns=["fecha", "ubicacion", "rubro", "refaccion", "cantidad"])
    if df is None or df.empty:
        return vacio
    base = df if "_fecha" in df.columns else bitacora_con_fechas(df)
    columnas = []
    for columna in base.columns:
        if not _es_columna_cantidad_refaccion(columna):
            continue
        origen = _ESPEJO_INSUMO_BITACORA.get(columna)
        if origen and origen in base.columns:
            continue
        columnas.append(columna)
    filas = []
    for _, fila in base.iterrows():
        dia = fila.get("_fecha")
        if not isinstance(dia, date):
            continue
        ubicacion = _ubicacion_fila_refaccion(fila)
        for columna in columnas:
            bruto = fila.get(columna)
            cantidad = cantidad_grasa_kg(bruto) if _es_grasa(columna) else _numero_insumo(bruto)
            if cantidad <= 0:
                continue
            filas.append(
                {
                    "fecha": dia,
                    "ubicacion": ubicacion,
                    "rubro": clasificar_rubro_refaccion(columna),
                    "refaccion": etiqueta_refaccion(columna),
                    "cantidad": cantidad,
                }
            )
        for nombre, cantidad in _partes_otros_insumos(fila.get("OTROS INSUMOS")):
            filas.append(
                {
                    "fecha": dia,
                    "ubicacion": ubicacion,
                    "rubro": clasificar_rubro_refaccion(nombre),
                    "refaccion": nombre,
                    "cantidad": cantidad,
                }
            )
    if not filas:
        return vacio
    return pd.DataFrame(filas)


def _cantidad_visible(valor: float) -> int | float:
    if abs(valor - round(valor)) < 1e-6:
        return int(round(valor))
    return round(float(valor), 2)


def _texto_cantidad_refaccion(refaccion: str, valor: float) -> str:
    visible = _cantidad_visible(valor)
    if _es_grasa(refaccion):
        return f"{visible} Kg"
    return str(visible)


def _vista_insumos_periodo(
    detalle: pd.DataFrame,
    inicio: date,
    fin: date,
    ubicacion: str = "",
    rubro: str = "",
) -> pd.DataFrame:
    if detalle is None or detalle.empty:
        return pd.DataFrame(columns=["fecha", "ubicacion", "rubro", "refaccion", "cantidad"])
    vista = detalle.loc[(detalle["fecha"] >= inicio) & (detalle["fecha"] <= fin)].copy()
    if ubicacion and ubicacion != "Todas las ubicaciones":
        vista = vista.loc[vista["ubicacion"].astype(str) == ubicacion]
    if rubro:
        vista = vista.loc[vista["rubro"] == rubro]
    return vista


def _barras_desde_vista(vista: pd.DataFrame) -> pd.DataFrame:
    if vista is None or vista.empty:
        return pd.DataFrame(columns=["Refacción", "Cantidad"])
    barras = (
        vista.groupby("refaccion", as_index=False)["cantidad"]
        .sum()
        .rename(columns={"refaccion": "Refacción", "cantidad": "Cantidad"})
        .sort_values("Cantidad", ascending=False)
    )
    barras["Cantidad"] = barras["Cantidad"].map(_cantidad_visible)
    return barras


def _desglose_desde_vista(vista: pd.DataFrame, con_rubro: bool = False) -> pd.DataFrame:
    columnas = ["Equipo", "Rubro", "Refacción", "Cantidad"] if con_rubro else ["Equipo", "Refacción", "Cantidad"]
    if vista is None or vista.empty:
        return pd.DataFrame(columns=columnas)
    claves = ["ubicacion", "rubro", "refaccion"] if con_rubro else ["ubicacion", "refaccion"]
    desglose = (
        vista.groupby(claves, as_index=False)["cantidad"]
        .sum()
        .rename(
            columns={
                "ubicacion": "Equipo",
                "rubro": "Rubro",
                "refaccion": "Refacción",
                "cantidad": "Cantidad",
            }
        )
    )
    desglose["Cantidad"] = desglose["Cantidad"].map(_cantidad_visible)
    desglose["_orden"] = desglose["Equipo"].map(orden_natural_pozo)
    orden = ["_orden", "Rubro", "Cantidad"] if con_rubro else ["_orden", "Cantidad"]
    ascendente = [True, True, False] if con_rubro else [True, False]
    desglose = desglose.sort_values(orden, ascending=ascendente).drop(columns="_orden")
    return desglose[columnas]


def resumen_insumos_periodo(
    detalle: pd.DataFrame,
    inicio: date,
    fin: date,
    ubicacion: str = "",
    rubro: str = "",
) -> tuple[pd.DataFrame, pd.DataFrame]:
    vista = _vista_insumos_periodo(detalle, inicio, fin, ubicacion, rubro)
    return _barras_desde_vista(vista), _desglose_desde_vista(vista, con_rubro=False)


def semanas_de_mes(anio: int, mes: int) -> list[tuple[str, date, date]]:
    inicio = date(anio, mes, 1)
    fin = date(anio + (mes == 12), 1 if mes == 12 else mes + 1, 1) - timedelta(days=1)
    cursor = inicio - timedelta(days=inicio.weekday())
    semanas = []
    numero = 1
    while cursor <= fin:
        recorte_ini = max(cursor, inicio)
        recorte_fin = min(cursor + timedelta(days=6), fin)
        etiqueta = f"Semana {numero} · {recorte_ini.strftime('%d/%m')} al {recorte_fin.strftime('%d/%m')}"
        semanas.append((etiqueta, recorte_ini, recorte_fin))
        numero += 1
        cursor += timedelta(days=7)
    return semanas


def figura_barras_insumos(barras: pd.DataFrame, titulo: str = "Insumos y refacciones suministradas"):
    if barras is None or barras.empty:
        return None
    fig = px.bar(
        barras,
        x="Refacción",
        y="Cantidad",
        text=[
            _texto_cantidad_refaccion(nombre, valor)
            for nombre, valor in zip(barras["Refacción"], barras["Cantidad"])
        ],
    )
    fig.update_traces(
        marker_color="#1e3a8a",
        marker_line_color="#172554",
        marker_line_width=0,
        textposition="outside",
        textfont=dict(color="#f8fafc", size=13),
        cliponaxis=False,
        hovertemplate="%{x}<br>%{text}<extra></extra>",
    )
    maximo = float(pd.to_numeric(barras["Cantidad"], errors="coerce").max() or 0)
    estilo_figura_dash(fig, titulo)
    fig.update_layout(
        height=500,
        margin=dict(l=12, r=16, t=68, b=110),
        xaxis_title="",
        yaxis_title="Cantidad suministrada",
        xaxis_tickangle=-32,
        showlegend=False,
    )
    fig.update_yaxes(range=[0, maximo * 1.22 if maximo else 1])
    return fig


def _tabla_cantidades_visibles(desglose: pd.DataFrame) -> pd.DataFrame:
    tabla = desglose.copy()
    if tabla.empty or "Cantidad" not in tabla.columns:
        return tabla
    tabla["Cantidad"] = [
        _texto_cantidad_refaccion(nombre, valor)
        for nombre, valor in zip(tabla["Refacción"], tabla["Cantidad"])
    ]
    return tabla


def _pintar_barras_y_desglose(barras: pd.DataFrame, desglose: pd.DataFrame, titulo: str, vacio: str) -> None:
    if barras.empty:
        st.info(vacio)
        return
    with st.container(border=True):
        figura = figura_barras_insumos(barras, titulo)
        st.plotly_chart(figura, use_container_width=True, config={"displayModeBar": False})
    st.markdown("**Desglose por equipo**")
    st.dataframe(_tabla_cantidades_visibles(desglose), hide_index=True, use_container_width=True)


def render_control_insumos_refacciones() -> None:
    st.markdown(
        '<div class="section-title">Control de Insumos y Refacciones Suministradas</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Separa los consumibles (aceite, filtros, refrigerante, grasa y similares) "
        "de las reparaciones del MCI (alternador, radiador, bomba de agua, marcha y otras piezas). "
        "Una columna nueva de la bitácora entra sola en el rubro que le corresponde. "
        "La grasa capturada en gramos se muestra en kilogramos."
    )
    detalle = detalle_insumos_suministrados(bitacora_con_fechas(cargar_bitacora()))
    if detalle.empty:
        st.info("La bitácora todavía no tiene insumos o reparaciones con cantidad.")
        return
    meses = sorted({(dia.year, dia.month) for dia in detalle["fecha"]})
    etiquetas_mes = [f"{NOMBRES_MES[mes - 1]} {anio}" for anio, mes in meses]
    hoy = date.today()
    preferido = f"{NOMBRES_MES[hoy.month - 1]} {hoy.year}"
    indice_mes = etiquetas_mes.index(preferido) if preferido in etiquetas_mes else len(etiquetas_mes) - 1
    c_mes, c_sem, c_ubi = st.columns([1.1, 1.4, 1.5])
    with c_mes:
        etiqueta_mes = st.selectbox("Mes", etiquetas_mes, index=indice_mes, key="insref_mes")
    anio, mes = meses[etiquetas_mes.index(etiqueta_mes)]
    semanas = semanas_de_mes(anio, mes)
    with c_sem:
        etiqueta_semana = st.selectbox(
            "Semana",
            ["Todas las semanas", *[item[0] for item in semanas]],
            key=f"insref_semana_{anio}_{mes}",
        )
    if etiqueta_semana == "Todas las semanas":
        inicio, fin = semanas[0][1], semanas[-1][2]
    else:
        inicio, fin = next((ini, fin_) for nombre, ini, fin_ in semanas if nombre == etiqueta_semana)
    ubicaciones = sorted(set(detalle["ubicacion"].astype(str)), key=orden_natural_pozo)
    with c_ubi:
        ubicacion = st.selectbox(
            "Ubicación / pozo",
            ["Todas las ubicaciones", *ubicaciones],
            key="insref_ubicacion",
            filter_mode="contains",
            help="Escribe un pozo o una batería, por ejemplo PS626 o BSNI.",
        )
    vista = _vista_insumos_periodo(detalle, inicio, fin, ubicacion)
    st.caption(f"Periodo {inicio.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}.")
    if vista.empty:
        st.info("No hay insumos ni reparaciones en ese periodo y ubicación.")
        return
    insumos = vista.loc[vista["rubro"] == RUBRO_INSUMO]
    reparaciones = vista.loc[vista["rubro"] == RUBRO_REPARACION]
    k1, k2 = st.columns(2)
    k1.metric(
        "Insumos Suministrados",
        int(insumos["refaccion"].nunique()) if not insumos.empty else 0,
        help="Consumibles distintos reportados en el periodo: aceite, filtros, refrigerante, grasa y similares.",
    )
    k2.metric(
        "Reparaciones Realizadas",
        int(round(pd.to_numeric(reparaciones["cantidad"], errors="coerce").sum())) if not reparaciones.empty else 0,
        help="Piezas del MCI reparadas o sustituidas en el periodo. Una columna nueva (radiador, marcha, alternador) se suma en cuanto la bitácora traiga cantidad.",
    )
    tab_insumos, tab_reparaciones, tab_general = st.tabs(
        ["Insumos Suministrados", "Reparaciones Realizadas", "Tabla General"]
    )
    with tab_insumos:
        _pintar_barras_y_desglose(
            _barras_desde_vista(insumos),
            _desglose_desde_vista(insumos),
            "Insumos suministrados",
            "No hay insumos suministrados en ese periodo y ubicación.",
        )
    with tab_reparaciones:
        _pintar_barras_y_desglose(
            _barras_desde_vista(reparaciones),
            _desglose_desde_vista(reparaciones),
            "Reparaciones realizadas",
            "No hay reparaciones en ese periodo. Aparecerán al cargar alternador, radiador, bomba de agua, marcha u otra pieza del MCI.",
        )
    with tab_general:
        st.markdown("**Tabla general**")
        st.dataframe(
            _tabla_cantidades_visibles(_desglose_desde_vista(vista, con_rubro=True)),
            hide_index=True,
            use_container_width=True,
        )


def render_dashboard_ordenes() -> None:
    if asegurar_datos_demo_si_faltan():
        st.rerun()
    st.markdown('<div class="section-title">Dashboard Ejecutivo</div>', unsafe_allow_html=True)
    st.caption(
        "Vista corporativa de órdenes de campo, OT de flota, combustible y cumplimiento OPS. "
        "Los filtros recortan métricas y gráficos al mismo periodo."
    )
    if st.session_state.get("demo_auto"):
        st.info("Se cargaron datos de demostración de septiembre 2026 porque la bitácora no tenía registros suficientes.")

    bitacora = bitacora_con_fechas(cargar_bitacora())
    fechas_bit = [f for f in bitacora["_fecha"].tolist() if f] if not bitacora.empty and "_fecha" in bitacora.columns else []
    diario, _parque = cargar_datos_flota()
    fmin, fmax = date(date.today().year, 1, 1), date.today()
    if fechas_bit:
        fmin, fmax = min(fechas_bit), max(max(fechas_bit), date.today())
    if diario is not None and not diario.empty:
        dmin, dmax = diario["Fecha"].min().date(), diario["Fecha"].max().date()
        fmin, fmax = min(fmin, dmin), max(fmax, dmax, date.today())

    hoy = date.today()
    inicio_def = max(fmin, date(hoy.year, hoy.month, 1))
    if inicio_def > fmax:
        inicio_def = fmin

    f1, f2, f3 = st.columns([1.8, 1.1, 1.1])
    with f1:
        modo_fecha = st.radio(
            "Filtro de fecha",
            ["Día específico", "Rango de fechas"],
            index=1,
            horizontal=True,
            key="dash_modo_fecha",
        )
        if modo_fecha == "Día específico":
            dia = min(max(hoy, fmin), fmax)
            dia_sel = st.date_input(
                "Día",
                value=dia,
                min_value=fmin,
                max_value=fmax,
                format="DD/MM/YYYY",
                key="dash_dia",
            )
            inicio, fin = dia_sel, dia_sel
        else:
            rango = st.date_input(
                "Rango de fechas",
                value=(inicio_def, min(hoy, fmax)),
                min_value=fmin,
                max_value=fmax,
                format="DD/MM/YYYY",
                key="dash_rango",
            )
            if isinstance(rango, (list, tuple)) and len(rango) == 2:
                inicio, fin = rango[0], rango[1]
            else:
                inicio, fin = inicio_def, min(hoy, fmax)
            if inicio > fin:
                inicio, fin = fin, inicio
    cuadrillas = ["Todas"] + list(CUADRILLAS) + ["Parque vehicular"]
    with f2:
        cuadrilla = st.selectbox("Cuadrilla", cuadrillas, key="dash_cuadrilla")
    with f3:
        estatus = st.selectbox("Estatus de OT", ["Todos"] + ESTATUS_DASHBOARD, key="dash_estatus")

    campo_raw = ots_campo_dashboard(inicio, fin)
    flota_raw = ots_flota_dashboard(inicio, fin)
    campo, flota = filtrar_ots_dashboard(campo_raw, flota_raw, cuadrilla, estatus)
    comb = filtrar_registro_combustible(diario, None, inicio, fin, "Todos")
    met_comb = metricas_combustible(comb)
    total_ot = int(len(campo) + len(flota))
    cerradas = int(
        ((campo["Estatus dash"] == "Completadas").sum() if not campo.empty else 0)
        + ((flota["Estatus dash"] == "Completadas").sum() if not flota.empty else 0)
    )
    pendientes = total_ot - cerradas
    cumplimiento, cumplidos, programados = cumplimiento_dashboard(inicio, fin)
    if programados == 0 and not campo.empty:
        validados = int((campo["REVISIÓN"].astype(str) == "Validado").sum()) if "REVISIÓN" in campo.columns else cerradas
        cumplimiento = round(100.0 * validados / max(len(campo), 1), 1)
        cumplidos, programados = validados, len(campo)

    k1, k2, k3, k4 = st.columns(4, gap="medium")
    k1.markdown(
        kpi_html(
            "OT del día" if inicio == fin else "OT del periodo",
            f"{total_ot:,}",
            f"Campo {len(campo)} · Flota {len(flota)}",
            "info",
        ),
        unsafe_allow_html=True,
    )
    k2.markdown(
        kpi_html(
            "Cerradas vs pendientes",
            f"{cerradas} / {pendientes}",
            f"{(100.0 * cerradas / total_ot):.0f}% cerradas" if total_ot else "Sin OT en el recorte",
            "ok" if pendientes == 0 and total_ot else "warn" if pendientes else "",
        ),
        unsafe_allow_html=True,
    )
    k3.markdown(
        kpi_html(
            "Gasto de combustible",
            f"${met_comb['gasto']:,.0f}",
            f"{met_comb['litros']:,.1f} L · {int(met_comb['cargas'])} cargas",
            "ok",
        ),
        unsafe_allow_html=True,
    )
    k4.markdown(
        kpi_html(
            "Cumplimiento OPS",
            f"{cumplimiento:.0f}%",
            f"{cumplidos} de {programados} programados" if programados else "Sin plan OPS en el recorte",
            "ok" if cumplimiento >= 80 else "warn",
        ),
        unsafe_allow_html=True,
    )

    st.markdown("")
    c_donut, c_bar = st.columns([1.05, 1.35], gap="large")
    with c_donut:
        with st.container(border=True):
            fig_d = figura_donut(campo, flota)
            if fig_d is None:
                st.markdown("**Estatus de las órdenes de trabajo**")
                st.info("No hay OT en el recorte seleccionado.")
            else:
                st.plotly_chart(fig_d, use_container_width=True, config={"displayModeBar": False})
    with c_bar:
        with st.container(border=True):
            fig_c = figura_barras_cuadrilla(campo, flota)
            if fig_c is None:
                st.markdown("**Distribución de servicios por cuadrilla**")
                st.info("No hay suficientes registros para graficar cuadrillas.")
            else:
                st.plotly_chart(fig_c, use_container_width=True, config={"displayModeBar": False})

    with st.container(border=True):
        fig_a = figura_actividades(campo, flota)
        if fig_a is None:
            st.markdown("**Actividades más frecuentes**")
            st.info("No hay suficientes registros para graficar actividades.")
        else:
            st.plotly_chart(fig_a, use_container_width=True, config={"displayModeBar": False})

    with st.container(border=True):
        fig_t = figura_tendencia(campo, flota, comb, inicio, fin)
        st.plotly_chart(fig_t, use_container_width=True, config={"displayModeBar": False})

    with st.container(border=True):
        st.markdown("**Detalle del recorte**")
        if campo.empty and flota.empty:
            st.caption("Ajusta el rango o la cuadrilla para ver el listado de OT.")
        else:
            filas = []
            if not campo.empty:
                for _, fila in campo.sort_values("_fecha", ascending=False).head(20).iterrows():
                    filas.append(
                        {
                            "Fecha": fila["_fecha"].strftime("%d/%m/%Y") if fila.get("_fecha") else "",
                            "Origen": "Campo",
                            "Cuadrilla / unidad": str(fila.get("CUADRILLA") or ""),
                            "Actividad": str(fila.get("Etiqueta") or ""),
                            "Estatus": str(fila.get("Estatus dash") or ""),
                        }
                    )
            if not flota.empty:
                for _, fila in flota.sort_values("_fecha", ascending=False).head(8).iterrows():
                    filas.append(
                        {
                            "Fecha": fila["_fecha"].strftime("%d/%m/%Y") if fila.get("_fecha") else "",
                            "Origen": "Flota",
                            "Cuadrilla / unidad": str(fila.get("Unidad") or "Parque vehicular"),
                            "Actividad": str(fila.get("Etiqueta") or ""),
                            "Estatus": str(fila.get("Estatus dash") or ""),
                        }
                    )
            st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)


def render_dashboard_ejecutivo() -> None:
    tab_ordenes, tab_insumos = st.tabs(
        ["Órdenes y combustible", "Control de Insumos y Refacciones Suministradas"]
    )
    with tab_ordenes:
        render_dashboard_ordenes()
    with tab_insumos:
        render_control_insumos_refacciones()


MODULOS_ADMIN = [
    {
        "id": "registro_bitacora",
        "icono": "➕",
        "titulo": "Registrar Actividad en Bitácora",
        "descripcion": "Captura la actividad del día, la cuadrilla, los consumibles y las fotos de evidencia.",
    },
    {
        "id": "dashboard",
        "icono": "📈",
        "titulo": "Dashboard Ejecutivo",
        "descripcion": "KPIs de OT, combustible y el control de insumos y refacciones suministradas.",
    },
    {
        "id": "insumos_refacciones",
        "icono": "📦",
        "titulo": "Insumos y Refacciones",
        "descripcion": "Consumo de refacciones e insumos suministrados en cada mantenimiento, con filtro de mes, semana y pozo.",
    },
    {
        "id": "ot_diario",
        "icono": "📑",
        "titulo": "Historial Diario de Órdenes de Trabajo",
        "descripcion": "Consulta los PDF diarios de OT (DDMMYYYY) y el resumen de cuadrillas de cada jornada.",
    },
    {
        "id": "asignaciones",
        "icono": "📅",
        "titulo": "Editor de Asignaciones Diarias",
        "descripcion": "Publica pendientes por cuadrilla y ausencias del día para el personal de campo.",
    },
    {
        "id": "almacen",
        "icono": "📦",
        "titulo": "Control de Almacén e Insumos",
        "descripcion": "Catálogo, entradas de almacén y reporte de consumo PJP4 frente a OPS.",
    },
    {
        "id": "ubicaciones",
        "icono": "📍",
        "titulo": "Ubicación y Equipos",
        "descripcion": "Ficha técnica por pozo: MCI, UB, historial, programados e importación Excel.",
    },
    {
        "id": "historial_base",
        "icono": "📚",
        "titulo": "Historial",
        "descripcion": "Historial de OTs, bitácora general y calendario de asistencia de cuadrillas.",
    },
    {
        "id": "programados",
        "icono": "📊",
        "titulo": "Programados vs No Programados",
        "descripcion": "Compara el plan mensual OPS con paros imprevistos y correctivos de bitácora.",
    },
    {
        "id": "programas",
        "icono": "🗂️",
        "titulo": "Consulta de Programas Mensuales",
        "descripcion": "Abre los Excel de mantenimiento y consúltalos con su disposición original.",
    },
    {
        "id": "programa_octubre",
        "icono": "📅",
        "titulo": "Programa de octubre",
        "descripcion": "Actividades de octubre por especialidad, con control de realizado y observaciones.",
    },
    {
        "id": "mantenimientos_programados",
        "icono": "🗓️",
        "titulo": "Mantenimientos Programados",
        "descripcion": "Programas de Zona Norte y Zona Sur, confirmación en campo y avance acumulado.",
    },
    {
        "id": "reportes_ejecutivos",
        "icono": "📄",
        "titulo": "Reportes ejecutivos",
        "descripcion": "Reporte formal para OPS, con la bitácora cruzada por fechas y por criterio.",
    },
    {
        "id": "calendario_asistencia",
        "icono": "📅",
        "titulo": "Calendario de asistencia",
        "descripcion": "Salidas a campo y descansos F/O por cuadrilla, semana y día del mes.",
    },
    {
        "id": "auditoria",
        "icono": "🔒",
        "titulo": "Auditoría y Accesos",
        "descripcion": "Revisión de bitácoras, gestión de catálogos y control de permisos.",
    },
    {
        "id": "gestionar_actividades",
        "icono": "📋",
        "titulo": "Gestionar Actividades",
        "descripcion": "Alta, edición y baja del catálogo de actividades de mantenimiento por especialidad.",
    },
    {
        "id": "cuadrillas",
        "icono": "👷",
        "titulo": "Historial de Cuadrillas",
        "descripcion": "Carpeta por cuadrilla: fechas, OT, personal e evidencia fotográfica.",
    },
    {
        "id": "soporte",
        "icono": "🧰",
        "titulo": "Soporte de Refacciones Agosto - Diciembre",
        "descripcion": "Proyecto por OT, aceite de relleno diario y control financiero P.O. Baker para proformas OPS.",
    },
    {
        "id": "vehiculos",
        "icono": "🚗",
        "titulo": "Vehículos",
        "descripcion": "Taller integral, combustible, preventivos por kilometraje, correctivos e historial de OT de la flota.",
    },
]
SUBMENUS_ALMACEN = [
    {
        "id": "gestion",
        "icono": "📋",
        "titulo": "Gestión de Insumos",
        "descripcion": "Alta libre de materiales y catálogo vigente guardado en almacen_insumos.json.",
    },
    {
        "id": "entradas",
        "icono": "📥",
        "titulo": "Entradas de almacén",
        "descripcion": "Registra ingresos de inventario con origen obligatorio PJP4 u OPS.",
    },
    {
        "id": "reporte",
        "icono": "📑",
        "titulo": "Reporte PJP4 vs OPS",
        "descripcion": "Separa el consumo de campo e inventario neto por origen del insumo.",
    },
]
SUBMENUS_PROGRAMADOS = [
    {
        "id": "ops",
        "icono": "✅",
        "titulo": "Mantenimientos Programados (OPS)",
        "descripcion": "Plan mensual del cliente: actividad, ubicación, equipo y estatus de facturación.",
    },
    {
        "id": "imprevistos",
        "icono": "🚨",
        "titulo": "Mantenimientos No Programados",
        "descripcion": "Paros, correctivos y restablecimientos capturados en la bitácora de campo.",
    },
]
SUBMENUS_AUDITORIA = [
    {
        "id": "revision",
        "icono": "📝",
        "titulo": "Revisión de Bitácoras",
        "descripcion": "Audita bitacora_operativa_soledad.xlsx: fecha, cuadrilla, personal, ubicación, consumibles y avance.",
    },
    {
        "id": "catalogos",
        "icono": "📚",
        "titulo": "Gestión de Catálogos",
        "descripcion": "Ubicaciones, motores MCI y unidades de bombeo que usa el formulario.",
    },
    {
        "id": "permisos",
        "icono": "🔑",
        "titulo": "Control de Permisos",
        "descripcion": "Registro local de usuarios, roles y accesos activos.",
    },
]
SUBMENUS_CATALOGOS = [
    {
        "id": "ubicaciones",
        "icono": "🗺️",
        "titulo": "Ubicaciones",
        "descripcion": "Sectores y zonas, y si requieren registro de pozo.",
    },
    {
        "id": "mci",
        "icono": "⚙️",
        "titulo": "Equipos MCI",
        "descripcion": "Motores de combustión interna disponibles en el selector de equipo.",
    },
    {
        "id": "ub",
        "icono": "🛢️",
        "titulo": "Equipos UB",
        "descripcion": "Unidades de bombeo (mecánicas e hidráulicas) del catálogo.",
    },
    {
        "id": "actividades",
        "icono": "📋",
        "titulo": "Gestionar Actividades",
        "descripcion": "Agrega, edita, desactiva o elimina actividades. El cambio entra de inmediato en la bitácora.",
    },
    {
        "id": "personal",
        "icono": "👷",
        "titulo": "Personal y coberturas",
        "descripcion": "Dotación titular por posición oficial (Mecánico 1–4) y bolsa de personal que puede cubrir.",
    },
]
SUBMENUS_UBICACIONES = [
    {
        "id": "plano",
        "icono": "🗺️",
        "titulo": "Plano general",
        "descripcion": "Mapa interactivo del Área Contractual Soledad: zonas, preventivos, paros y estatus de OT.",
    },
    {
        "id": "fichas",
        "icono": "🛢️",
        "titulo": "Historial y Estado de Pozos",
        "descripcion": "Explora zonas, abre un pozo y consulta calendario, bitácora y programados.",
    },
    {
        "id": "importar",
        "icono": "📥",
        "titulo": "Importar inventario Excel",
        "descripcion": "Carga masiva de pozos y equipos desde las tablas por zona.",
    },
]
SUBMENUS_FICHA = [
    {
        "id": "mantenimientos",
        "icono": "🔧",
        "titulo": "Mantenimientos realizados",
        "descripcion": "Historial de bitácora: programados y no programados de este pozo.",
    },
    {
        "id": "equipos",
        "icono": "🔁",
        "titulo": "Historial de equipos",
        "descripcion": "Cambios y reubicaciones de MCI o UB registradas en la ficha.",
    },
    {
        "id": "pendientes",
        "icono": "📌",
        "titulo": "Pendientes / programados",
        "descripcion": "Actividades OPS no cumplidas, reprogramadas y cumplidas.",
    },
    {
        "id": "personal",
        "icono": "👷",
        "titulo": "Personal que intervino",
        "descripcion": "Cuadrillas de campo que han trabajado en esta ubicación.",
    },
]
CLAVES_SUBMENU_ADMIN = (
    "admin_sub_almacen",
    "admin_sub_programados",
    "admin_sub_auditoria",
    "admin_sub_catalogos",
    "admin_sub_vehiculos",
    "ub_sub_modulo",
    "ub_ficha_seccion",
    "ub_zona_explorador",
    "ub_ficha_id",
    "ub_cal_dia",
    "ub_cal_pozo",
    "hc_dia",
    "hc_cuadrilla",
)


def render_tarjeta_explorador(modulo: dict, state_key: str) -> None:
    badge = str(modulo.get("badge") or "").strip()
    tono = str(modulo.get("badge_tono") or "").strip()
    clase_badge = f"badge-count {tono}".strip() if badge else ""
    badge_html = f'<div class="{clase_badge}">{html.escape(badge)}</div>' if badge else ""
    st.markdown(
        f"""
        <div class="explorer-card">
            <div class="icon-tile">{modulo["icono"]}</div>
            <h3>{html.escape(modulo["titulo"])}</h3>
            <p>{html.escape(modulo["descripcion"])}</p>
            {badge_html}
        </div>
        """,
        unsafe_allow_html=True,
    )
    if st.button("Abrir", key=f"explorar_{state_key}_{modulo['id']}", use_container_width=True):
        st.session_state[state_key] = modulo["id"]
        st.rerun()


def render_cuadricula_explorador(modulos: list[dict], state_key: str, meta: str = "") -> None:
    if meta:
        st.markdown(f'<p class="explorer-meta">{meta}</p>', unsafe_allow_html=True)
    for inicio in range(0, len(modulos), 3):
        cols = st.columns(3, gap="medium")
        for col, modulo in zip(cols, modulos[inicio : inicio + 3]):
            with col:
                render_tarjeta_explorador(modulo, state_key)


def render_barra_volver(ruta_html: str, state_key: str, etiqueta: str = "← Volver") -> None:
    st.markdown(f'<div class="explorer-crumb">{ruta_html}</div>', unsafe_allow_html=True)
    if st.button(etiqueta, key=f"volver_{state_key}"):
        st.session_state[state_key] = None
        st.rerun()


def titulo_submenu(modulos: list[dict], submenu_id: str) -> str:
    for item in modulos:
        if item["id"] == submenu_id:
            return item["titulo"]
    return submenu_id


def reset_submodulos_admin() -> None:
    for clave in CLAVES_SUBMENU_ADMIN:
        st.session_state[clave] = None


def abrir_modulo_admin(modulo_id: str, **subs) -> None:
    reset_submodulos_admin()
    st.session_state.admin_modulo = modulo_id
    for clave, valor in subs.items():
        st.session_state[clave] = valor
    st.rerun()


def render_seccion_admin(kicker: str, titulo: str, ayuda: str) -> None:
    st.markdown(
        f"""
        <div class="admin-compact">
            <div class="kicker">{html.escape(kicker)}</div>
            <h4>{html.escape(titulo)}</h4>
            <p>{html.escape(ayuda)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


_PATRON_POZO_CAMPO = re.compile(r"^(PSN|PGA|PPB|PS|PA|PB|PC)-?\d", re.IGNORECASE)


def delta_vs_previo(actual: float, previo: float) -> int | None:
    if actual > 0 and previo > 0:
        return round(((actual - previo) / previo) * 100)
    return None


def texto_rango_registros(inicio: date, ultima: date | None) -> str:
    """Inicio del periodo y fecha del último registro que entra en el total."""
    ini = inicio.strftime("%d/%m/%Y")
    if isinstance(ultima, date):
        return f"Desde {ini} hasta {ultima.strftime('%d/%m/%Y')}"
    return f"Desde {ini} · sin registros"


def ubicacion_de_registro(fila) -> str:
    return etiqueta_ubicacion_fila(fila) or sitios_operativos.texto_ubicacion_final(
        texto_visible(fila.get("POZO"), ""),
        texto_visible(fila.get("UBICACIÓN") or fila.get("UBICACION"), ""),
        texto_visible(fila.get("SECTOR"), ""),
    )


def es_pozo_de_campo(nombre: str) -> bool:
    clave = clave_ubicacion(nombre).replace(" ", "")
    if not clave:
        return False
    if inferir_zona_id(nombre) not in {"instalaciones", "otros", ""}:
        return True
    return bool(_PATRON_POZO_CAMPO.match(clave))


def _texto_trabajo_registro(fila) -> str:
    partes = []
    for col in (
        "TIPO DE ACTIVIDAD",
        "ACTIVIDAD CORTA",
        "DESCRIPCIÓN DE LA ACTIVIDAD",
        "DESCRIPCION DE LA ACTIVIDAD",
        "EQUIPO/UBICACIÓN",
    ):
        partes.append(texto_visible(fila.get(col), ""))
    return " ".join(p for p in partes if p).casefold()


def fila_suministra_aceite_mci(fila) -> bool:
    """Aceite SAE-50 entregado a un MCI en pozo. El hidráulico va en otra columna."""
    if _num_insumo(fila.get("ACEITE (LTS)")) <= 0:
        return False
    if not es_pozo_de_campo(ubicacion_de_registro(fila)):
        return False
    texto = _texto_trabajo_registro(fila)
    if re.search(r"\bmci\b", texto):
        return True
    if re.search(r"\bubm\b|unidad de bombeo", texto):
        return False
    return True


def resumen_aceite_mci(inicio: date, fin: date) -> dict:
    df = bitacora_con_fechas(cargar_bitacora())
    total = 0.0
    ultima: date | None = None
    if df is not None and not df.empty and "_fecha" in df.columns:
        for _, fila in df.iterrows():
            fch = fila.get("_fecha")
            if not isinstance(fch, date) or fch < inicio or fch > fin:
                continue
            if not fila_suministra_aceite_mci(fila):
                continue
            total += _num_insumo(fila.get("ACEITE (LTS)"))
            if ultima is None or fch > ultima:
                ultima = fch
    return {"litros": round(total, 2), "ultima": ultima}


def metricas_nucleo_admin(jornada: date) -> dict:
    inicio_mes = date(jornada.year, jornada.month, 1)
    if jornada.month == 1:
        prev_ini = date(jornada.year - 1, 12, 1)
        prev_ultimo = date(jornada.year - 1, 12, 31)
    else:
        prev_ini = date(jornada.year, jornada.month - 1, 1)
        prev_ultimo = inicio_mes - timedelta(days=1)
    prev_fin = date(prev_ini.year, prev_ini.month, min(jornada.day, prev_ultimo.day))
    diario, _parque = cargar_datos_flota()
    mes_comb = filtrar_registro_combustible(diario, None, inicio_mes, jornada)
    prev_comb = filtrar_registro_combustible(diario, None, prev_ini, prev_fin)
    litros_mes = metricas_combustible(mes_comb)["litros"]
    litros_prev = metricas_combustible(prev_comb)["litros"]
    gasto_mes = gasto_registros_combustible(mes_comb, diario)
    gasto_prev = gasto_registros_combustible(prev_comb, diario)
    ultima_carga = ultima_fecha_combustible(mes_comb)
    delta = delta_vs_previo(litros_mes, litros_prev)
    delta_gasto = delta_vs_previo(gasto_mes, gasto_prev)
    aceite = resumen_aceite_mci(inicio_mes, jornada)
    aceite_prev = resumen_aceite_mci(prev_ini, prev_fin)
    delta_aceite = delta_vs_previo(aceite["litros"], aceite_prev["litros"])
    alertas = 0
    for tabla in inventario_por_origen().values():
        if tabla is not None and not tabla.empty:
            alertas += int((pd.to_numeric(tabla["Existencia"], errors="coerce").fillna(0) <= 0).sum())
    bit = bitacora_del_dia(bitacora_con_fechas(cargar_bitacora()), jornada)
    reportes = listar_reportes_ot_diarios()
    ot = next((r for r in reportes if r["fecha"] == jornada), None)
    fechas_bit = set()
    df_bit = bitacora_con_fechas(cargar_bitacora())
    if df_bit is not None and not df_bit.empty and "_fecha" in df_bit.columns:
        fechas_bit = {
            d
            for d in df_bit["_fecha"]
            if isinstance(d, date) and d.year == jornada.year and d.month == jornada.month
        }
    fechas_ot = {r["fecha"] for r in reportes if r["fecha"].year == jornada.year and r["fecha"].month == jornada.month}
    asig = asignacion_del_dia(jornada)
    return {
        "litros_mes": litros_mes,
        "delta_litros": delta,
        "gasto_mes": gasto_mes,
        "delta_gasto": delta_gasto,
        "ultima_carga": ultima_carga,
        "aceite_mes": aceite["litros"],
        "delta_aceite": delta_aceite,
        "ultima_aceite": aceite["ultima"],
        "alertas": alertas,
        "bitacoras": 0 if bit is None else int(len(bit)),
        "ot": ot,
        "fechas_bit": fechas_bit,
        "fechas_ot": fechas_ot,
        "asignacion": hay_asignacion_publicada(asig),
    }


def _html_delta_metrica(delta: str | None) -> str:
    if not delta:
        return ""
    texto = str(delta).strip()
    if texto.startswith("-"):
        clase = "down"
    elif texto.startswith("+"):
        clase = "up"
    else:
        clase = "flat"
    return f'<div class="delta {clase}">{html.escape(delta)}</div>'


def html_tarjeta_combustible(
    titulo: str,
    valor: str,
    delta: str | None,
    corte: str,
    extra: str | None = None,
    delta_extra: str | None = None,
) -> str:
    extra_html = f'<div class="sub">{html.escape(extra)}</div>' if extra else ""
    return (
        f'<div class="metric-fuel">'
        f'<div class="lbl">{html.escape(titulo)}</div>'
        f'<div class="val">{html.escape(valor)}</div>'
        f"{extra_html}"
        f"{_html_delta_metrica(delta)}"
        f"{_html_delta_metrica(delta_extra)}"
        f'<div class="corte">{html.escape(corte)}</div>'
        f"</div>"
    )


def render_mini_calendario_admin(jornada: date, fechas_bit: set, fechas_ot: set) -> None:
    hoy = date.today()
    cal = calendar.Calendar(firstweekday=calendar.MONDAY)
    reglas = []
    for dia in cal.itermonthdates(jornada.year, jornada.month):
        if dia.month != jornada.month:
            continue
        clave = f"caldia_{dia.strftime('%Y%m%d')}"
        extra = []
        if dia in fechas_bit and dia in fechas_ot:
            extra.append(
                "background: linear-gradient(135deg, rgba(20, 83, 45, 0.45) 0%, rgba(30, 58, 95, 0.75) 100%);"
                "border-color: #86efac;"
            )
        elif dia in fechas_bit:
            extra.append("background: rgba(20, 83, 45, 0.38); border-color: #2d6a4f;")
        elif dia in fechas_ot:
            extra.append("background: rgba(30, 58, 95, 0.72); border-color: #3b82f6;")
        if dia == hoy:
            extra.append("box-shadow: 0 0 0 1px rgba(134, 239, 172, 0.35); border-color: #86efac;")
        if dia == jornada:
            extra.append("outline: 2px solid #fbbf24; outline-offset: 1px;")
        if extra:
            reglas.append(f'[class*="st-key-{clave}"] button {{ {" ".join(extra)} }}')
    if reglas:
        st.markdown(f"<style>{''.join(reglas)}</style>", unsafe_allow_html=True)
    cabecera = st.columns(7)
    for col, nombre in zip(cabecera, ("L", "M", "X", "J", "V", "S", "D")):
        col.markdown(f'<div class="cal-dow">{nombre}</div>', unsafe_allow_html=True)
    semanas = cal.monthdatescalendar(jornada.year, jornada.month)
    for semana in semanas:
        cols = st.columns(7)
        for col, dia in zip(cols, semana):
            with col:
                if dia.month != jornada.month:
                    st.markdown('<div class="cal-day empty"></div>', unsafe_allow_html=True)
                    continue
                if st.button(
                    str(dia.day),
                    key=f"caldia_{dia.strftime('%Y%m%d')}",
                    use_container_width=True,
                ):
                    st.session_state.admin_jornada_pendiente = dia
                    st.rerun()


def render_respaldo_base_datos(clave: str = "respaldo") -> None:
    st.markdown('<div class="section-title">Respaldo de la base de datos</div>', unsafe_allow_html=True)
    st.caption(
        "En cada despliegue, Streamlit Cloud vuelve a bajar main y borra lo capturado en el disco. "
        "Con el secreto GITHUB_TOKEN la bitácora (fechas, pozos, actividades, insumos y estatus) "
        "se copia a la rama datos, sin redesplegar el sitio. Las fotografías y los PDF no se respaldan: "
        "viajan en la copia descargable y, al descargarla, se retiran del servidor."
    )
    ruta_xlsx = persistencia.resolver_ruta(persistencia.ARCHIVO_BITACORA)
    filas = persistencia.contar_filas_bitacora(ruta_xlsx)
    estado = persistencia.leer_estado()
    detalle = str(estado.get("mensaje") or "Todavía no hay un respaldo en esta ejecución.")
    st.caption(f"Registros en la bitácora de este servidor: {filas}. {detalle}")
    if not persistencia.credenciales_github():
        st.warning(
            "Falta GITHUB_TOKEN en los secretos de Streamlit. "
            "Créalo con permiso de escritura sobre el repositorio y guárdalo junto con la app. "
            "Puedes indicar GITHUB_REPO y GITHUB_RAMA_DATOS; si no, se usa la rama datos. "
            "Hasta configurarlo, descarga esta copia antes de cada git push a main."
        )
    firma = "vacio"
    if os.path.exists(ruta_xlsx):
        firma = f"{os.path.getmtime(ruta_xlsx)}:{os.path.getsize(ruta_xlsx)}"
    if st.session_state.get(f"{clave}_firma") != firma or f"{clave}_zip" not in st.session_state:
        try:
            datos, nombre, pesados = persistencia.construir_copia_seguridad()
        except OSError as exc:
            st.warning(f"No se pudo armar la copia de seguridad: {exc}")
        else:
            st.session_state[f"{clave}_firma"] = firma
            st.session_state[f"{clave}_zip"] = datos
            st.session_state[f"{clave}_nombre"] = nombre
            st.session_state[f"{clave}_pesados"] = tuple(pesados)
    if f"{clave}_zip" not in st.session_state:
        return
    st.download_button(
        "📥 Descargar Copia de Seguridad de la Base de Datos",
        data=st.session_state[f"{clave}_zip"],
        file_name=st.session_state[f"{clave}_nombre"],
        mime="application/zip",
        key=f"{clave}_descarga",
        use_container_width=True,
        on_click=_al_descargar_copia_respaldo,
        args=(tuple(st.session_state.get(f"{clave}_pesados") or ()),),
    )
    col_subir, col_bajar = st.columns(2)
    with col_subir:
        if st.button("Respaldar ahora en GitHub", key=f"{clave}_subir", use_container_width=True):
            resultado = persistencia.respaldar_cambios("Respaldo manual desde administración")
            if resultado.get("ok"):
                st.success(resultado.get("mensaje") or "Respaldo enviado.")
            else:
                st.warning(resultado.get("mensaje") or "No se pudo respaldar.")
    with col_bajar:
        if st.button("Recuperar respaldo de GitHub", key=f"{clave}_bajar", use_container_width=True):
            resultado = persistencia.restaurar_desde_github()
            for llave in ("catalogos", "catalogo_actividades", "permisos"):
                st.session_state.pop(llave, None)
            st.session_state.pop(f"{clave}_zip", None)
            if resultado.get("ok"):
                st.success(resultado.get("mensaje") or "Respaldo recuperado.")
            else:
                st.warning(resultado.get("mensaje") or "No se pudo recuperar.")
            st.rerun()


_COLUMNAS_RESUMEN_ASISTENCIA = {"realizadas", "total", "avance"}


def ruta_calendario_asistencia() -> str:
    if os.path.exists(CALENDARIO_ASISTENCIA_FILE):
        return CALENDARIO_ASISTENCIA_FILE
    if not os.path.isdir(DATABASE_DIR):
        return CALENDARIO_ASISTENCIA_FILE
    for nombre in sorted(os.listdir(DATABASE_DIR)):
        clave = nombre.casefold()
        if nombre.startswith("~$") or "calendario" not in clave:
            continue
        if clave.endswith((".xlsx", ".xlsm")):
            return os.path.join(DATABASE_DIR, nombre)
    return CALENDARIO_ASISTENCIA_FILE


def _texto_celda_calendario(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor == int(valor):
        return str(int(valor))
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    texto = str(valor).strip()
    if texto.startswith("=") or texto.lower() in {"nan", "none", "<na>"}:
        return ""
    return texto


def _valores_hoja_calendario(ws) -> dict[tuple[int, int], str]:
    """Copia el valor de cada celda combinada a todo su rango."""
    merges = list(ws.merged_cells.ranges)
    valores: dict[tuple[int, int], str] = {}
    for fila in range(1, (ws.max_row or 1) + 1):
        for col in range(1, (ws.max_column or 1) + 1):
            bruto = ws.cell(fila, col).value
            if bruto is None:
                for rango in merges:
                    if rango.min_row <= fila <= rango.max_row and rango.min_col <= col <= rango.max_col:
                        bruto = ws.cell(rango.min_row, rango.min_col).value
                        break
            texto = _texto_celda_calendario(bruto)
            if texto:
                valores[(fila, col)] = texto
    return valores


def _mes_anio_calendario(titulo: str) -> tuple[int, int]:
    texto = _sin_acentos(titulo).casefold()
    mes = next((i + 1 for i, nombre in enumerate(NOMBRES_MES) if _sin_acentos(nombre).casefold() in texto), 0)
    anio_m = re.search(r"(20\d{2})", titulo)
    anio = int(anio_m.group(1)) if anio_m else date.today().year
    return mes or date.today().month, anio


def _es_descanso_fo(valor: str) -> bool:
    clave = _sin_acentos(valor).casefold().replace(" ", "").replace(".", "")
    return clave in {"f/o", "fo", "descanso"}


def _numero_dia_calendario(valor: str) -> int:
    digitos = re.sub(r"\D", "", str(valor or ""))
    if not digitos:
        return 0
    try:
        dia = int(digitos)
    except ValueError:
        return 0
    return dia if 1 <= dia <= 31 else 0


def _leer_hoja_calendario(ws) -> dict | None:
    valores = _valores_hoja_calendario(ws)
    ancla = next(
        (
            (fila, col)
            for (fila, col), texto in valores.items()
            if _sin_acentos(texto).casefold().startswith("cuadrilla")
        ),
        None,
    )
    if ancla is None:
        return None
    fila_nombres, col_cuadrilla = ancla
    fila_dias = fila_nombres - 1
    fila_semanas = fila_nombres - 2
    titulo = next(
        (texto for texto in valores.values() if "calendario" in _sin_acentos(texto).casefold()),
        str(ws.title or "Calendario de asistencia"),
    )
    mes, anio = _mes_anio_calendario(f"{titulo} {ws.title}")
    dias = []
    for col in range(col_cuadrilla + 1, (ws.max_column or col_cuadrilla) + 1):
        encabezado = _sin_acentos(valores.get((fila_nombres, col), "")).casefold()
        if encabezado in _COLUMNAS_RESUMEN_ASISTENCIA:
            continue
        numero = _numero_dia_calendario(valores.get((fila_dias, col), ""))
        if not numero:
            continue
        try:
            fecha = date(anio, mes, numero)
        except ValueError:
            continue
        dias.append(
            {
                "fecha": fecha,
                "dia": numero,
                "nombre": NOMBRES_DIA[fecha.weekday()],
                "semana": valores.get((fila_semanas, col), "") or "Sin semana",
                "col": col,
            }
        )
    if not dias:
        return None
    filas = []
    vacios = 0
    for fila in range(fila_nombres + 1, (ws.max_row or fila_nombres) + 1):
        nombre = valores.get((fila, col_cuadrilla), "")
        if not nombre:
            vacios += 1
            if vacios >= 2:
                break
            continue
        vacios = 0
        if _sin_acentos(nombre).casefold().startswith("cuadrilla"):
            continue
        celdas = {dia["fecha"]: valores.get((fila, dia["col"]), "") for dia in dias}
        total_txt = ""
        for col in range(col_cuadrilla + 1, (ws.max_column or col_cuadrilla) + 1):
            if _sin_acentos(valores.get((fila_nombres, col), "")).casefold() == "total":
                total_txt = valores.get((fila, col), "")
                break
        try:
            total = int(float(total_txt)) if total_txt else 0
        except ValueError:
            total = 0
        filas.append(
            {
                "cuadrilla": etiqueta_cuadrilla_historial(nombre),
                "fila": fila,
                "celdas": celdas,
                "total": total,
            }
        )
    filas.sort(key=lambda fila: orden_cuadrilla_oficial(fila["cuadrilla"]))
    semanas = []
    for dia in dias:
        if dia["semana"] not in semanas:
            semanas.append(dia["semana"])
    return {
        "titulo": titulo,
        "etiqueta": f"{NOMBRES_MES[mes - 1]} {anio}",
        "anio": anio,
        "mes": mes,
        "semanas": semanas,
        "dias": dias,
        "filas": filas,
    }


@st.cache_data(ttl=60, show_spinner=False)
def leer_calendarios_asistencia(ruta: str, mtime: float) -> list[dict]:
    _ = mtime
    from openpyxl import load_workbook

    wb = load_workbook(_copia_excel_legible(ruta), data_only=False)
    try:
        calendarios = []
        for nombre in wb.sheetnames:
            calendario = _leer_hoja_calendario(wb[nombre])
            if calendario and calendario["filas"]:
                calendarios.append(calendario)
        return calendarios
    finally:
        wb.close()


def _casilla_calendario_vacia(valor: str) -> bool:
    texto = str(valor or "").strip()
    return texto.casefold() in {"", "—", "--", "-", "nan", "none", "<na>"}


def _clase_jornada_asistencia(valor: str) -> str:
    if _casilla_calendario_vacia(valor):
        return "vacio"
    if _es_descanso_fo(valor):
        return "fo"
    return "campo"


def _ficha_fila_bitacora(fila) -> str:
    for columna in ("PERSONAL", "MECANICO EN TURNO", "MECANICO TITULAR"):
        texto = str(fila.get(columna) or "").strip()
        if texto and texto.casefold() not in {"nan", "none", "<na>"}:
            return texto
    return ""


def jornadas_bitacora_por_cuadrilla(df: pd.DataFrame | None = None) -> dict[tuple[str, date], str]:
    """Fecha y cuadrilla con actividad. La ficha vacía queda como cadena vacía, no como ausencia."""
    base = bitacora_con_fechas(df if df is not None else cargar_bitacora())
    if base.empty or "CUADRILLA" not in base.columns or "_fecha" not in base.columns:
        return {}
    fichas: dict[tuple[str, date], list[str]] = {}
    for _, fila in base.iterrows():
        dia = fila.get("_fecha")
        if not isinstance(dia, date):
            continue
        clave = clave_cuadrilla_oficial(str(fila.get("CUADRILLA") or ""))
        if not clave:
            continue
        fichas.setdefault((clave, dia), [])
        ficha = _ficha_fila_bitacora(fila)
        if ficha and ficha not in fichas[(clave, dia)]:
            fichas[(clave, dia)].append(ficha)
    return {clave: " / ".join(valores) for clave, valores in fichas.items()}


def valor_casilla_desde_bitacora(
    cuadrilla: str,
    fecha: date,
    jornadas: dict[tuple[str, date], str],
    hoy: date | None = None,
) -> str:
    """Ficha si hubo actividad, F/O si el día ya pasó sin registro, vacío si es hoy o futuro."""
    hoy = hoy or date.today()
    clave = (clave_cuadrilla_oficial(cuadrilla), fecha)
    if clave in jornadas:
        return jornadas[clave] or "En campo"
    if fecha < hoy:
        return "F/O"
    return ""


def aplicar_bitacora_a_calendario(
    calendario: dict,
    jornadas: dict[tuple[str, date], str],
    hoy: date | None = None,
) -> dict:
    hoy = hoy or date.today()
    filas = []
    for fila in calendario.get("filas") or []:
        celdas = {
            dia["fecha"]: valor_casilla_desde_bitacora(fila["cuadrilla"], dia["fecha"], jornadas, hoy)
            for dia in calendario.get("dias") or []
        }
        filas.append({**fila, "celdas": celdas})
    salida = dict(calendario)
    salida["filas"] = filas
    return salida


def _casilla_calendario_equivalente(actual: str, deseado: str) -> bool:
    if _casilla_calendario_vacia(actual) and _casilla_calendario_vacia(deseado):
        return True
    if _es_descanso_fo(actual) and _es_descanso_fo(deseado):
        return True
    return str(actual or "").strip() == str(deseado or "").strip()


def _firma_sync_calendario(ruta_cal: str, hoy: date) -> tuple:
    return (_mtime_archivo(ruta_cal), _mtime_archivo(EXCEL_FILE), hoy.isoformat())


def _guardar_libro_calendario(wb, destino: str) -> None:
    import time

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    ultimo = None
    try:
        wb.save(tmp)
        for intento in range(4):
            try:
                shutil.copy2(tmp, destino)
                return
            except PermissionError as exc:
                ultimo = exc
                if intento < 3:
                    time.sleep(0.3 * (intento + 1))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
    raise PermissionError(
        "calendario_asistencia.xlsx está abierto o en uso. "
        "La tabla ya muestra la bitácora y el archivo se guardará cuando quede libre."
    ) from ultimo


def sincronizar_calendario_desde_bitacora(
    jornadas: dict[tuple[str, date], str] | None = None,
    hoy: date | None = None,
) -> dict:
    """Escribe en el calendario la ficha o F/O que salen de la bitácora. No toca fórmulas."""
    hoy = hoy or date.today()
    ruta = ruta_calendario_asistencia()
    if not os.path.exists(ruta):
        return {"ok": False, "cambios": 0, "mensaje": f"No se encontró {CALENDARIO_ASISTENCIA_FILE}."}
    firma = _firma_sync_calendario(ruta, hoy)
    try:
        if st.session_state.get("_cal_sync_firma") == firma:
            return {"ok": True, "cambios": 0, "mensaje": ""}
        fallo = st.session_state.get("_cal_sync_fallo")
        if isinstance(fallo, tuple) and len(fallo) == 3 and fallo[0] == firma and (datetime.now().timestamp() - fallo[1]) < 90:
            return {"ok": False, "cambios": 0, "mensaje": fallo[2]}
    except Exception:
        pass
    if jornadas is None:
        jornadas = jornadas_bitacora_por_cuadrilla()
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell

    wb = load_workbook(_copia_excel_legible(ruta), data_only=False)
    cambios = 0
    try:
        for nombre in wb.sheetnames:
            ws = wb[nombre]
            calendario = _leer_hoja_calendario(ws)
            if not calendario:
                continue
            for fila in calendario["filas"]:
                for dia in calendario["dias"]:
                    celda = ws.cell(fila["fila"], dia["col"])
                    if isinstance(celda, MergedCell):
                        continue
                    deseado = valor_casilla_desde_bitacora(fila["cuadrilla"], dia["fecha"], jornadas, hoy)
                    actual = _texto_celda_calendario(celda.value)
                    if _casilla_calendario_equivalente(actual, deseado):
                        continue
                    celda.value = deseado or None
                    cambios += 1
        if cambios:
            _guardar_libro_calendario(wb, ruta)
            try:
                leer_calendarios_asistencia.clear()
            except Exception:
                pass
            try:
                _leer_calendario_asistencia_cache.clear()
            except Exception:
                pass
            firma = _firma_sync_calendario(ruta, hoy)
    except PermissionError as exc:
        try:
            st.session_state["_cal_sync_fallo"] = (firma, datetime.now().timestamp(), str(exc))
        except Exception:
            pass
        return {"ok": False, "cambios": 0, "mensaje": str(exc)}
    finally:
        wb.close()
    try:
        st.session_state["_cal_sync_firma"] = firma
        st.session_state.pop("_cal_sync_fallo", None)
    except Exception:
        pass
    return {"ok": True, "cambios": cambios, "mensaje": ""}


def _html_calendario_asistencia(calendario: dict, semana: str) -> str:
    dias = [dia for dia in calendario["dias"] if semana == "Todas las semanas" or dia["semana"] == semana]
    grupos = []
    for dia in dias:
        if not grupos or grupos[-1][0] != dia["semana"]:
            grupos.append([dia["semana"], 1])
        else:
            grupos[-1][1] += 1
    semanas_html = "".join(
        f'<th class="semana" colspan="{n}">{html.escape(nombre)}</th>' for nombre, n in grupos
    )
    dias_html = "".join(
        f'<th class="dia"><span>{dia["dia"]:02d}</span><small>{html.escape(dia["nombre"][:3])}</small></th>'
        for dia in dias
    )
    cuerpo = []
    for fila in calendario["filas"]:
        celdas = []
        for dia in dias:
            valor = str(fila["celdas"].get(dia["fecha"]) or "")
            clase = _clase_jornada_asistencia(valor)
            texto = "F/O" if clase == "fo" else (valor or "--")
            celdas.append(f'<td class="{clase}">{html.escape(texto)}</td>')
        cuerpo.append(
            f'<tr><th class="crew">{html.escape(fila["cuadrilla"])}</th>{"".join(celdas)}</tr>'
        )
    return (
        '<div class="cal-asist-wrap"><table class="cal-asist">'
        f'<thead><tr><th class="crew" rowspan="2">Cuadrilla</th>{semanas_html}</tr>'
        f"<tr>{dias_html}</tr></thead><tbody>{''.join(cuerpo)}</tbody></table></div>"
    )


def _estilo_calendario_asistencia() -> str:
    return """
    <style>
    .cal-asist-wrap { overflow-x: auto; margin-top: 0.4rem; }
    .cal-asist { border-collapse: separate; border-spacing: 0; min-width: 100%; font-size: 0.78rem; }
    .cal-asist th, .cal-asist td {
        border: 1px solid #334155; padding: 0.35rem 0.4rem; text-align: center; white-space: nowrap;
    }
    .cal-asist th.semana { background: #1e3a5f; color: #e2e8f0; font-size: 0.72rem; letter-spacing: 0.04em; }
    .cal-asist th.dia { background: #0f172a; color: #cbd5e1; }
    .cal-asist th.dia span { display: block; color: #f8fafc; font-size: 0.85rem; }
    .cal-asist th.dia small { color: #94a3b8; font-weight: 600; }
    .cal-asist th.crew {
        position: sticky; left: 0; z-index: 1; text-align: left; background: #0f172a; color: #f8fafc;
        min-width: 8.5rem;
    }
    .cal-asist td.campo { background: rgba(20, 83, 45, 0.55); color: #bbf7d0; }
    .cal-asist td.fo { background: rgba(120, 53, 15, 0.55); color: #fde68a; font-weight: 700; }
    .cal-asist td.vacio { background: #0b1220; color: #475569; }
    </style>
    """


def render_calendario_asistencia() -> None:
    st.markdown(_estilo_calendario_asistencia(), unsafe_allow_html=True)
    ruta = ruta_calendario_asistencia()
    st.caption(
        "Sale de la bitácora operativa. Verde: ficha del personal en campo. "
        "Ámbar: F/O cuando la jornada ya pasó y esa cuadrilla no registró actividad. "
        "El guion es el día de hoy o una fecha futura."
    )
    if not os.path.exists(ruta):
        st.warning(f"No se encontró el calendario en {CALENDARIO_ASISTENCIA_FILE}.")
        return
    jornadas = jornadas_bitacora_por_cuadrilla()
    try:
        sync = sincronizar_calendario_desde_bitacora(jornadas)
    except Exception as exc:
        sync = {"ok": False, "mensaje": str(exc)}
    if not sync.get("ok") and sync.get("mensaje"):
        st.caption(sync["mensaje"])
    try:
        calendarios = leer_calendarios_asistencia(ruta, os.path.getmtime(ruta))
    except Exception as exc:
        st.error(f"No se pudo leer el calendario de asistencia: {exc}")
        return
    calendarios = [aplicar_bitacora_a_calendario(item, jornadas) for item in calendarios]
    if not calendarios:
        st.info("El archivo no tiene una matriz de cuadrillas reconocible.")
        return
    etiquetas = [item["etiqueta"] for item in calendarios]
    col_mes, col_sem = st.columns(2)
    with col_mes:
        etiqueta = st.selectbox("Mes", etiquetas, key="cal_asist_mes")
    calendario = next(item for item in calendarios if item["etiqueta"] == etiqueta)
    with col_sem:
        semana = st.selectbox(
            "Semana",
            ["Todas las semanas", *calendario["semanas"]],
            key="cal_asist_semana",
        )
    dias = [dia for dia in calendario["dias"] if semana == "Todas las semanas" or dia["semana"] == semana]
    fechas = {dia["fecha"] for dia in dias}
    salidas = descansos = 0
    for fila in calendario["filas"]:
        for fecha, valor in fila["celdas"].items():
            if fecha not in fechas:
                continue
            clase = _clase_jornada_asistencia(valor)
            if clase == "campo":
                salidas += 1
            elif clase == "fo":
                descansos += 1
    m1, m2, m3 = st.columns(3)
    m1.metric("Cuadrillas", len(calendario["filas"]))
    m2.metric("Salidas a campo", salidas)
    m3.metric("Descansos F/O", descansos)
    st.markdown(_html_calendario_asistencia(calendario, semana), unsafe_allow_html=True)
    st.caption(
        f"Sincronizado con la bitácora operativa. {calendario['titulo']}. "
        f"El archivo {os.path.basename(ruta)} se actualiza solo."
    )


def _texto_celda_admin(valor) -> str:
    return texto_visible(valor, "")


def _indice_pozos_planos() -> dict[str, dict]:
    ruta = UBICACIONES_FILE
    firma = os.path.getmtime(ruta) if os.path.exists(ruta) else 0
    cache = getattr(_indice_pozos_planos, "cache", None)
    if cache and cache[0] == firma:
        return cache[1]
    indice: dict[str, dict] = {}
    for ficha in (cargar_ubicaciones().get("fichas") or {}).values():
        if not isinstance(ficha, dict):
            continue
        for plana in _claves_ficha(ficha):
            indice.setdefault(plana, ficha)
    for plana, oficial in _indice_infraestructura().items():
        indice[plana] = oficial
    _indice_pozos_planos.cache = (firma, indice)
    return indice


def _nombre_pozo_canonico(texto: str) -> str:
    bruto = str(texto or "").strip()
    if not bruto or bruto.lower() in {"nan", "none"}:
        return ""
    plana = clave_pozo_plana(bruto)
    ficha = _indice_pozos_planos().get(plana)
    if ficha:
        return etiqueta_codigo_ubicacion(ficha.get("nombre")) or bruto
    if plana and re.search(r"\d", plana) and re.match(r"^[A-Z]+\d", plana):
        return plana
    return bruto


def _pozo_fila_admin(fila) -> str:
    sitio = formato_ubicacion_reporte(fila)
    if sitio:
        return sitio
    pozo = _texto_celda_admin(fila.get("POZO"))
    if sitios_operativos.es_vacio_sitio(pozo):
        ubicacion = _texto_celda_admin(fila.get("UBICACIÓN")) or _texto_celda_admin(
            fila.get("EQUIPO/UBICACIÓN")
        )
        partes = [
            parte.strip()
            for parte in ubicacion.split("·")
            if parte.strip() and not sitios_operativos.es_vacio_sitio(parte)
        ]
        pozo = partes[-1] if partes else ""
    return _nombre_pozo_canonico(pozo)


def _sector_fila_admin(fila) -> str:
    sector = _texto_celda_admin(fila.get("SECTOR"))
    if sector:
        return sector
    ubicacion = _texto_celda_admin(fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN"))
    partes = [parte.strip() for parte in ubicacion.split("·") if parte.strip()]
    if len(partes) >= 2:
        return partes[0]
    return ""


def _estatus_fila_admin(fila) -> str:
    return _texto_celda_admin(
        fila.get("EVALUACIÓN DEL ESTATUS") or fila.get("REVISIÓN") or fila.get("ESTATUS")
    )


def _cuadrilla_fila_admin(fila) -> str:
    bruto = _texto_celda_admin(fila.get("CUADRILLA"))
    return etiqueta_cuadrilla_historial(bruto) if bruto else ""


def _descripcion_fila_admin(fila) -> str:
    return _texto_celda_admin(
        fila.get("DESCRIPCIÓN DE LA ACTIVIDAD") or fila.get("DESCRIPCION DE LA ACTIVIDAD")
    )


def _ubicacion_fila_admin(fila) -> str:
    return etiqueta_ubicacion_fila(fila) or _texto_celda_admin(
        fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN")
    )


def _periodo_mes_calendario(ancla: date) -> tuple[date, date]:
    ultimo = calendar.monthrange(ancla.year, ancla.month)[1]
    return date(ancla.year, ancla.month, 1), date(ancla.year, ancla.month, ultimo)


def _bitacora_rango_admin(inicio: date, fin: date) -> pd.DataFrame:
    """Actividades cuya fecha cae en el periodo, inclusive. Fuera de rango no entran."""
    base = bitacora_con_fechas(cargar_bitacora())
    if base is None or base.empty or "_fecha" not in base.columns:
        return base.iloc[0:0] if base is not None else pd.DataFrame()
    if not isinstance(inicio, date) or not isinstance(fin, date) or inicio > fin:
        return base.iloc[0:0]
    return base.loc[
        base["_fecha"].map(lambda dia: isinstance(dia, date) and not isinstance(dia, datetime) and inicio <= dia <= fin)
    ].copy()


def _bitacora_mes_admin(jornada: date) -> pd.DataFrame:
    inicio, fin = _periodo_mes_calendario(jornada)
    return _bitacora_rango_admin(inicio, fin)


def _etiqueta_periodo_admin(inicio: date, fin: date) -> str:
    mes_completo = (
        inicio.year == fin.year
        and inicio.month == fin.month
        and inicio.day == 1
        and fin.day == calendar.monthrange(fin.year, fin.month)[1]
    )
    if mes_completo:
        return f"{NOMBRES_MES[inicio.month - 1]} {inicio.year}"
    return f"{inicio.strftime('%d/%m/%Y')} – {fin.strftime('%d/%m/%Y')}"


def _filtrar_vista_admin(df: pd.DataFrame, ot: str, sector: str, pozo: str, cuadrilla: str, estatus: str) -> pd.DataFrame:
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()
    ot = (ot or "").strip()
    conservar = []
    for posicion, (_, fila) in enumerate(df.iterrows()):
        if ot:
            secuencia = str(numero_secuencia_actividad(fila.get("SECUENCIA")) or _texto_celda_admin(fila.get("SECUENCIA")))
            if ot.casefold() not in secuencia.casefold():
                continue
        if sector and sector != "Todos" and _sector_fila_admin(fila).casefold() != sector.casefold():
            continue
        if pozo and pozo != "Todos" and _pozo_fila_admin(fila).casefold() != pozo.casefold():
            continue
        if cuadrilla and cuadrilla != "Todos" and _cuadrilla_fila_admin(fila).casefold() != cuadrilla.casefold():
            continue
        if estatus and estatus != "Todos" and _estatus_fila_admin(fila).casefold() != estatus.casefold():
            continue
        conservar.append(posicion)
    return df.iloc[conservar].copy()


def _opciones_admin(df: pd.DataFrame, extractor) -> list[str]:
    valores = set()
    if df is not None and not df.empty:
        for _, fila in df.iterrows():
            texto = extractor(fila).strip()
            if texto:
                valores.add(texto)
    return ["Todos"] + sorted(valores, key=str.casefold)


def _elegir_filtro_admin(etiqueta: str, opciones: list[str], key: str) -> str:
    if not opciones:
        opciones = ["Todos"]
    if st.session_state.get(key) not in opciones:
        st.session_state[key] = opciones[0]
    return st.selectbox(etiqueta, opciones, key=key)


def _clase_avance_admin(fila, programadas: set[str]) -> str:
    """Una actividad del filtro: programada o no programada, nunca un tercer residuo."""
    actividad = fila.get("TIPO DE ACTIVIDAD")
    if _clave_actividad_oficial(actividad) in {"", "otra", "nan", "none"}:
        actividad = ""
    destino = _destino_hoja_suministros(actividad, programadas, _descripcion_fila_admin(fila))
    if destino == "programado":
        return "Programado"
    if destino == "relleno":
        return "No programado"
    if clasificar_mantenimiento(fila) == "Programado":
        return "Programado"
    return "No programado"


def _tipo_mantenimiento_admin(fila) -> str:
    for columna in ("TIPO DE ACTIVIDAD", "ACTIVIDAD CORTA"):
        texto = _texto_celda_admin(fila.get(columna))
        if texto and _clave_actividad_oficial(texto) not in {"otra", "nan", "none"}:
            return texto
    return _descripcion_fila_admin(fila)


def _ot_fila_admin(fila) -> str:
    return str(numero_secuencia_actividad(fila.get("SECUENCIA")) or _texto_celda_admin(fila.get("SECUENCIA")))


COLUMNAS_AUDITORIA_PERIODO = [
    "Fecha",
    "OT",
    "Pozo",
    "Sector",
    "Cuadrilla",
    "Tipo de mantenimiento",
    "Programado o no programado",
]


def _tabla_auditoria_periodo(df: pd.DataFrame) -> pd.DataFrame:
    """Un renglón por actividad del filtro, con la misma clase que las tarjetas."""
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_AUDITORIA_PERIODO)
    programadas = _actividades_programadas_suministros()
    vista = df.sort_values("_fecha", kind="stable") if "_fecha" in df.columns else df
    filas = []
    for _, fila in vista.iterrows():
        dia = fila.get("_fecha")
        filas.append(
            {
                "Fecha": dia.strftime("%d/%m/%Y") if isinstance(dia, date) and not isinstance(dia, datetime) else "",
                "OT": _ot_fila_admin(fila),
                "Pozo": _pozo_fila_admin(fila),
                "Sector": _sector_fila_admin(fila),
                "Cuadrilla": _cuadrilla_fila_admin(fila),
                "Tipo de mantenimiento": _tipo_mantenimiento_admin(fila),
                "Programado o no programado": _clase_avance_admin(fila, programadas),
            }
        )
    return pd.DataFrame(filas, columns=COLUMNAS_AUDITORIA_PERIODO)


def _figura_insumos_vista_previa(detalle: pd.DataFrame):
    """Barras del periodo con la cantidad escrita arriba, sin depender del cursor."""
    resumen = (
        detalle.groupby("refaccion", as_index=False)["cantidad"]
        .sum()
        .sort_values("cantidad", ascending=False)
        .head(8)
    )
    resumen["etiqueta"] = [
        _texto_cantidad_refaccion(str(nombre), float(valor))
        for nombre, valor in zip(resumen["refaccion"], resumen["cantidad"])
    ]
    figura = px.bar(
        resumen,
        x="refaccion",
        y="cantidad",
        text="etiqueta",
        labels={"refaccion": "Insumo", "cantidad": "Cantidad"},
    )
    tope = float(resumen["cantidad"].max() or 0) * 1.28
    figura.update_traces(
        textposition="outside",
        textfont=dict(color="#e2e8f0", size=13),
        cliponaxis=False,
    )
    figura.update_layout(
        height=280,
        margin=dict(l=8, r=8, t=36, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#e2e8f0",
        title="Insumos consumidos",
        uniformtext=dict(minsize=11, mode="show"),
        yaxis=dict(rangemode="tozero", range=[0, tope] if tope else None),
    )
    return figura


def _resumen_auditoria_periodo(tabla: pd.DataFrame) -> tuple[int, int, int]:
    if tabla is None or tabla.empty:
        return 0, 0, 0
    programados = int((tabla["Programado o no programado"] == "Programado").sum())
    no_programados = int((tabla["Programado o no programado"] == "No programado").sum())
    return programados, no_programados, len(tabla)


def _contar_avance_admin(df: pd.DataFrame) -> tuple[int, int, int]:
    return _resumen_auditoria_periodo(_tabla_auditoria_periodo(df))


def _nombre_excel_programa(nombre: str) -> str:
    base = os.path.basename(str(nombre or "").replace("\\", "/"))
    if not base or base.startswith("~$") or ".." in base:
        return ""
    if not base.lower().endswith(EXTENSIONES_PROGRAMA_EXCEL):
        return ""
    return base


def _guardar_excel_programa_admin(archivo) -> None:
    if archivo is None:
        st.warning("Elige un archivo Excel de mantenimientos programados.")
        return
    nombre = _nombre_excel_programa(getattr(archivo, "name", ""))
    if not nombre:
        st.warning("El archivo tiene que ser un Excel (.xlsx, .xls o .xlsm).")
        return
    carpeta = directorio_programas_mtto()
    os.makedirs(carpeta, exist_ok=True)
    destino = os.path.join(carpeta, nombre)
    reemplazo = os.path.exists(destino)
    with open(destino, "wb") as salida:
        salida.write(archivo.getvalue())
    try:
        cargar_actividades_octubre.clear()
    except Exception:
        pass
    verbo = "Se reemplazó" if reemplazo else "Quedó cargado"
    st.success(
        f"{verbo} {nombre} en la carpeta de programas de mantenimiento. "
        "La bitácora y los catálogos siguen respaldándose en la rama datos."
    )


def _guardar_edicion_dia_admin(dia: pd.DataFrame, editada: pd.DataFrame) -> None:
    guardados = 0
    errores = []
    indices = list(dia.index)
    for posicion, (_, row) in enumerate(editada.iterrows()):
        bruto = row["_idx"] if "_idx" in row.index else None
        try:
            idx = int(bruto) if bruto is not None and pd.notna(bruto) else indices[posicion]
        except (TypeError, ValueError, IndexError):
            continue
        if idx not in dia.index:
            continue
        original = dia.loc[idx]
        if isinstance(original, pd.DataFrame):
            original = original.iloc[0]
        desc_nueva = str(row.get("Descripción") or "").strip()
        ubic_nueva = str(row.get("Ubicación") or "").strip()
        if desc_nueva == _descripcion_fila_admin(original) and ubic_nueva == _ubicacion_fila_admin(original):
            continue
        try:
            actualizar_registro_bitacora(
                idx,
                original,
                {
                    "fecha": parsear_fecha_valor(original.get("FECHA")),
                    "cuadrilla": str(original.get("CUADRILLA") or "").strip(),
                    "secuencia": numero_secuencia_actividad(original.get("SECUENCIA")),
                    "descripcion": desc_nueva,
                    "ubicacion": ubic_nueva,
                    "personal": _texto_celda_admin(original.get("PERSONAL")),
                    "comentarios": _texto_celda_admin(original.get("COMENTARIOS")),
                },
            )
            guardados += 1
        except Exception as exc:
            errores.append(str(exc))
    if guardados:
        despues_de_cambio_bitacora("edición de actividades de la jornada en el panel administrativo")
        st.session_state.adm_m_edicion_ok = guardados
    for error in errores:
        st.warning(error)
    if guardados and not errores:
        st.rerun()
    elif not guardados and not errores:
        st.info("La tabla ya coincide con lo guardado.")


def render_barra_herramientas_admin() -> None:
    aviso = st.session_state.pop("adm_barra_aviso", None)
    if isinstance(aviso, dict):
        if aviso.get("ok"):
            st.success(aviso.get("mensaje") or "Respaldo enviado a GitHub.")
        else:
            st.warning(aviso.get("mensaje") or "No se pudo respaldar.")
    estado = persistencia.leer_estado()
    filas = persistencia.contar_filas_bitacora(persistencia.resolver_ruta(persistencia.ARCHIVO_BITACORA))
    if not persistencia.credenciales_github():
        marca = "Sin token"
    elif estado.get("ok"):
        marca = "Sincronizado"
    else:
        marca = "Pendiente"
    detalle = str(estado.get("mensaje") or "Todavía no hay un respaldo en esta ejecución.")
    with st.container(border=True):
        estado_col, respaldo_col, copia_col, recuperar_col = st.columns([2.4, 1.5, 1.35, 1.35], vertical_alignment="center")
        with estado_col:
            st.markdown(
                f'<p class="adm-sync"><b>{html.escape(marca)}</b> · {filas} registros<br>{html.escape(detalle)}</p>',
                unsafe_allow_html=True,
            )
        with respaldo_col:
            if st.button("Respaldar datos en GitHub", type="primary", use_container_width=True, key="adm_barra_subir"):
                st.session_state.adm_barra_aviso = persistencia.respaldar_cambios(
                    "Respaldo manual desde administración"
                )
                st.rerun()
        with copia_col:
            _boton_descarga_respaldo_admin()
        with recuperar_col:
            if st.button("Recuperar respaldo", use_container_width=True, key="adm_barra_bajar"):
                resultado = persistencia.restaurar_desde_github()
                for llave in ("catalogos", "catalogo_actividades", "permisos"):
                    st.session_state.pop(llave, None)
                st.session_state.pop("adm_barra_zip", None)
                st.session_state.adm_barra_aviso = resultado
                st.rerun()


def _boton_descarga_respaldo_admin() -> None:
    ruta = persistencia.resolver_ruta(persistencia.ARCHIVO_BITACORA)
    firma = "vacio"
    if os.path.exists(ruta):
        firma = f"{os.path.getmtime(ruta)}:{os.path.getsize(ruta)}"
    if st.session_state.get("adm_barra_firma") != firma or "adm_barra_zip" not in st.session_state:
        try:
            datos, nombre, pesados = persistencia.construir_copia_seguridad()
        except OSError as exc:
            st.warning(f"No se pudo armar la copia de seguridad: {exc}")
        else:
            st.session_state.adm_barra_firma = firma
            st.session_state.adm_barra_zip = datos
            st.session_state.adm_barra_nombre = nombre
            st.session_state.adm_barra_pesados = tuple(pesados)
    if "adm_barra_zip" not in st.session_state:
        return
    st.download_button(
        "Descargar copia",
        data=st.session_state.adm_barra_zip,
        file_name=st.session_state.adm_barra_nombre,
        mime="application/zip",
        key="adm_barra_descarga",
        use_container_width=True,
        on_click=_al_descargar_copia_respaldo,
        args=(tuple(st.session_state.get("adm_barra_pesados") or ()),),
    )


def render_entrada_datos_admin() -> None:
    if st.button(
        "Registrar o editar actividad de bitácora",
        use_container_width=True,
        key="adm_m_registro",
    ):
        abrir_modulo_admin("registro_bitacora")
    archivo_programa = st.file_uploader(
        "Excel de mantenimientos programados",
        type=["xlsx", "xls", "xlsm"],
        key="adm_m_excel_prog",
    )
    if st.button("Cargar Excel al programa", use_container_width=True, key="adm_m_cargar_excel"):
        _guardar_excel_programa_admin(archivo_programa)
    if st.button("Mantenimientos programados", use_container_width=True, key="adm_m_programa_anual"):
        abrir_modulo_admin("mantenimientos_programados")
    alta_ubic, alta_eq, alta_ins = st.columns(3)
    with alta_ubic:
        if st.button("Ubicación", use_container_width=True, key="adm_m_alta_ub"):
            abrir_modulo_admin(
                "auditoria",
                admin_sub_auditoria="catalogos",
                admin_sub_catalogos="ubicaciones",
            )
    with alta_eq:
        if st.button("Equipo", use_container_width=True, key="adm_m_alta_eq"):
            abrir_modulo_admin("auditoria", admin_sub_auditoria="catalogos")
    with alta_ins:
        if st.button("Insumos", use_container_width=True, key="adm_m_alta_ins"):
            abrir_modulo_admin("almacen", admin_sub_almacen="gestion")


def render_edicion_captura_admin(jornada: date) -> None:
    aviso = st.session_state.pop("adm_m_edicion_ok", None)
    if aviso:
        st.success(f"Se actualizaron {aviso} actividad(es). El cambio quedó en la bitácora y en la rama datos.")
    dia = bitacora_del_dia(bitacora_con_fechas(cargar_bitacora()), jornada)
    if dia.empty:
        st.info(f"No hay actividades cargadas el {jornada.strftime('%d/%m/%Y')}.")
    else:
        tabla = pd.DataFrame(
            {
                "_idx": list(dia.index),
                "OT": [
                    str(numero_secuencia_actividad(fila.get("SECUENCIA")) or _texto_celda_admin(fila.get("SECUENCIA")))
                    for _, fila in dia.iterrows()
                ],
                "Cuadrilla": [_cuadrilla_fila_admin(fila) for _, fila in dia.iterrows()],
                "Ubicación": [_ubicacion_fila_admin(fila) for _, fila in dia.iterrows()],
                "Descripción": [_descripcion_fila_admin(fila) for _, fila in dia.iterrows()],
            }
        )
        editada = st.data_editor(
            tabla,
            num_rows="fixed",
            hide_index=True,
            use_container_width=True,
            key=f"adm_m_editor_{jornada.isoformat()}",
            disabled=["_idx", "OT", "Cuadrilla"],
            column_config={
                "_idx": None,
                "Descripción": st.column_config.TextColumn("Descripción", width="large"),
            },
        )
        if st.button(
            "Guardar actividades del día",
            type="primary",
            use_container_width=True,
            key="adm_m_guardar_dia",
        ):
            _guardar_edicion_dia_admin(dia, editada)
    ed_cat, ed_ins = st.columns(2)
    with ed_cat:
        if st.button("Editar equipos y ubicaciones", use_container_width=True, key="adm_m_edit_cat"):
            abrir_modulo_admin("auditoria", admin_sub_auditoria="catalogos")
    with ed_ins:
        if st.button("Editar insumos", use_container_width=True, key="adm_m_edit_ins"):
            abrir_modulo_admin("almacen", admin_sub_almacen="gestion")


def render_vista_previa_admin(jornada: date) -> None:
    inicio_mes, fin_mes = _periodo_mes_calendario(date.today())
    with st.container(border=True):
        desde_col, hasta_col = st.columns(2)
        with desde_col:
            desde = st.date_input(
                "Desde",
                value=inicio_mes,
                format="DD/MM/YYYY",
                key="adm_m_prev_desde",
            )
        with hasta_col:
            hasta = st.date_input(
                "Hasta",
                value=fin_mes,
                format="DD/MM/YYYY",
                key="adm_m_prev_hasta",
            )
    if not isinstance(desde, date) or not isinstance(hasta, date):
        st.warning("Elige un rango de fechas válido.")
        return
    if desde > hasta:
        st.warning("La fecha Desde tiene que ser anterior o igual a Hasta.")
        return
    periodo = _bitacora_rango_admin(desde, hasta)
    st.caption(
        f"{_etiqueta_periodo_admin(desde, hasta)} · {len(periodo)} actividad(es) en el periodo"
    )
    filtros = st.columns(2)
    with filtros[0]:
        ot = st.text_input("OT", key="adm_m_f_ot", placeholder="Secuencia")
        pozo = _elegir_filtro_admin("Pozo", _opciones_admin(periodo, _pozo_fila_admin), "adm_m_f_pozo")
        estatus = _elegir_filtro_admin("Estatus", _opciones_admin(periodo, _estatus_fila_admin), "adm_m_f_estatus")
    with filtros[1]:
        sector = _elegir_filtro_admin("Sector", _opciones_admin(periodo, _sector_fila_admin), "adm_m_f_sector")
        cuadrilla = _elegir_filtro_admin(
            "Cuadrilla", _opciones_admin(periodo, _cuadrilla_fila_admin), "adm_m_f_cuadrilla"
        )
    filtrada = _filtrar_vista_admin(periodo, ot, sector, pozo, cuadrilla, estatus)
    auditoria = _tabla_auditoria_periodo(filtrada)
    programados, imprevistos, total = _resumen_auditoria_periodo(auditoria)
    m1, m2, m3 = st.columns(3)
    m1.metric("Programados", programados)
    m2.metric("No programados", imprevistos)
    m3.metric("Total", total)
    avance = go.Figure(
        go.Bar(
            x=["Programados", "No programados", "Total"],
            y=[programados, imprevistos, total],
            marker_color=["#86efac", "#fbbf24", "#93c5fd"],
        )
    )
    avance.update_layout(
        height=190,
        margin=dict(l=8, r=8, t=28, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#e2e8f0",
        title="Avance de mantenimientos",
    )
    st.plotly_chart(avance, use_container_width=True, key="adm_m_chart_avance")
    detalle = detalle_insumos_suministrados(filtrada)
    if detalle.empty:
        st.caption("Sin insumos consumidos en el periodo seleccionado.")
    else:
        figura = _figura_insumos_vista_previa(detalle)
        st.plotly_chart(figura, use_container_width=True, key="adm_m_chart_insumos")
    with st.expander("Verificar actividades del periodo"):
        st.caption(
            f"{programados} programadas · {imprevistos} no programadas · {total} en total."
        )
        if auditoria.empty:
            st.info("No hay actividades en el periodo seleccionado.")
        else:
            st.dataframe(
                auditoria,
                hide_index=True,
                use_container_width=True,
                height=420,
            )


PLANTILLA_SUMINISTROS = os.path.join("assets", "plantilla_suministros.xlsm")


def _ruta_plantilla_suministros() -> str:
    if os.path.isfile(PLANTILLA_SUMINISTROS):
        return PLANTILLA_SUMINISTROS
    alternativa = PLANTILLA_SUMINISTROS + ".xlsm"
    if os.path.isfile(alternativa):
        return alternativa
    raise FileNotFoundError(
        "No está la plantilla de suministros en assets/plantilla_suministros.xlsm."
    )


def _hoja_programados_suministros(wb):
    for nombre in wb.sheetnames:
        if "aceite de manttos" in nombre.casefold():
            return wb[nombre]
    for nombre in wb.sheetnames:
        hoja = wb[nombre]
        pozo = str(hoja.cell(2, 1).value or "").strip().casefold()
        trabajo = str(hoja.cell(2, 3).value or "").strip().casefold()
        if pozo.startswith("pozo") and trabajo.startswith("trabajo"):
            return hoja
    return None


def _hoja_relleno_suministros(wb):
    for nombre in wb.sheetnames:
        if "aceite de relleno" in nombre.casefold():
            return wb[nombre]
    for nombre in wb.sheetnames:
        hoja = wb[nombre]
        mecanico = str(hoja.cell(1, 1).value or "").strip().casefold()
        pozo = str(hoja.cell(1, 2).value or "").strip().casefold()
        if mecanico.startswith("mecan") and pozo.startswith("pozo"):
            return hoja
    return None


def _rango_datos_formula(formula: str) -> tuple[int, int] | None:
    coincidencia = re.search(r"SUM\(\s*J(\d+)\s*:\s*J(\d+)\s*\)", str(formula or ""), re.IGNORECASE)
    if not coincidencia:
        return None
    return int(coincidencia.group(1)), int(coincidencia.group(2))


def _fila_datos_programados(hoja) -> tuple[int, int]:
    limite = min(hoja.max_row or 1, 2500)
    for fila in range(3, limite + 1):
        rango = _rango_datos_formula(hoja.cell(fila, 10).value)
        if rango:
            return rango
    return 3, 1278


def _cuadrilla_plantilla_suministros(nombre: str) -> str:
    clave = clave_cuadrilla_oficial(nombre)
    if clave.startswith("mecanico "):
        numero = clave.rsplit(" ", 1)[-1]
        if numero.isdigit():
            return f"MEC-{numero}"
    etiquetas = {
        "electricos": "ELECT",
        "instrumentos": "INSTR",
        "montaje": "MONTA",
        "soldadura": "SOLD",
        "retro excavadora": "RETRO",
    }
    return etiquetas.get(clave) or etiqueta_cuadrilla_historial(nombre)


def _equipo_fila_suministros(fila) -> str:
    equipo = _texto_celda_admin(fila.get("EQUIPO"))
    if equipo:
        return equipo
    ubicacion = _texto_celda_admin(fila.get("UBICACIÓN") or fila.get("EQUIPO/UBICACIÓN"))
    partes = [parte.strip() for parte in ubicacion.split("·") if parte.strip()]
    if len(partes) >= 3:
        return partes[2]
    return ""


def _ubm_y_motor(equipo: str) -> tuple[str, str]:
    texto = _sin_acentos(equipo).casefold()
    if any(token in texto for token in ("ubm", "lufkin", "wfo", "ici", "bomba")):
        return equipo, ""
    if any(token in texto for token in ("arrow", "vortec", "mci", "motor")):
        return "", equipo
    return "", ""


def _litros_celda(valor: float):
    redondo = round(float(valor), 2)
    if abs(redondo - round(redondo)) < 0.001:
        return int(round(redondo))
    return redondo


_ACTIVIDAD_RELLENO_SUMINISTROS = "inspeccion y suministro"


def _clave_actividad_oficial(valor) -> str:
    return " ".join(_sin_acentos(str(valor or "")).casefold().split())


def _es_mantenimiento_preventivo_catalogo(nombre: str) -> bool:
    """Bimestral, semestral, habilitado, puesta en marcha o mantenimiento preventivo."""
    clave = _clave_actividad_oficial(nombre)
    if not clave or clave in {"otra", _ACTIVIDAD_RELLENO_SUMINISTROS}:
        return False
    if any(token in clave for token in ("bimestral", "semestral", "habilitado", "puesta en marcha", "preventiv")):
        return True
    if "mantenimiento" in clave and not any(token in clave for token in ("restablec", "recuper", "correctiv")):
        return True
    return False


def _actividades_programadas_suministros() -> set[str]:
    datos = leer_json(ACTIVIDADES_FILE, catalogo_actividades_default())
    catalogo = normalizar_catalogo_actividades(datos)
    nombres: set[str] = set()
    for items in catalogo.values():
        for item in items or []:
            nombre = str((item or {}).get("nombre") or "")
            if actividad_esta_activa(item) and _es_mantenimiento_preventivo_catalogo(nombre):
                nombres.add(_clave_actividad_oficial(nombre))
    return nombres


def _destino_por_descripcion(descripcion: str) -> str:
    """Respaldo para bitácoras antiguas sin Actividad oficial."""
    clave = _clave_actividad_oficial(descripcion)
    if not clave:
        return ""
    if _ACTIVIDAD_RELLENO_SUMINISTROS in clave:
        return "relleno"
    if any(token in clave for token in ("bimestral", "semestral", "habilitado", "puesta en marcha", "preventiv")):
        return "programado"
    return ""


def _destino_hoja_suministros(actividad, programadas: set[str], descripcion: str = "") -> str:
    clave = _clave_actividad_oficial(actividad)
    if clave:
        if clave == _ACTIVIDAD_RELLENO_SUMINISTROS:
            return "relleno"
        if clave in programadas:
            return "programado"
        return ""
    return _destino_por_descripcion(descripcion)


_PREFIJOS_NOMBRE_LARGO = (
    ("PSN", "SOLEDAD NORTE"),
    ("PGA", "GALLO"),
    ("PPB", "PALO BLANCO"),
    ("PB", "PALO BLANCO"),
    ("PA", "ARAGON"),
    ("PS", "SOLEDAD"),
    ("PC", "COYOTE"),
)


def _norm_etiqueta_equivalencia(valor) -> str:
    return " ".join(_sin_acentos(_texto_celda(valor)).upper().split())


def _es_encabezado_corto(valor) -> bool:
    return "NOMBRE CORTO" in _norm_etiqueta_equivalencia(valor)


def _es_encabezado_largo(valor) -> bool:
    return "NOMBRE LARGO" in _norm_etiqueta_equivalencia(valor)


def _claves_nombre_excel(nombre: str) -> set[str]:
    texto = _norm_etiqueta_equivalencia(nombre)
    compacto = re.sub(r"[^A-Z0-9]", "", texto)
    claves = set()
    if texto:
        claves.add(texto)
    if compacto:
        claves.add(compacto)
    return claves


def _absorber_equivalencia(mapa: dict[str, str], cortos: list[str], corto, largo) -> None:
    corto_txt = " ".join(str(corto or "").split())
    largo_txt = " ".join(str(largo or "").split())
    if not corto_txt or not largo_txt:
        return
    if _es_encabezado_corto(corto_txt) or _es_encabezado_largo(largo_txt):
        return
    plana = clave_pozo_plana(corto_txt)
    if not plana or plana in mapa:
        return
    mapa[plana] = largo_txt
    cortos.append(corto_txt)


def _absorber_json_equivalencias(datos, mapa: dict[str, str], cortos: list[str]) -> None:
    if isinstance(datos, list):
        for item in datos:
            if isinstance(item, dict):
                corto = item.get("corto") or item.get("nombre_corto") or item.get("app") or item.get("codigo")
                largo = item.get("largo") or item.get("nombre_largo") or item.get("excel")
                _absorber_equivalencia(mapa, cortos, corto, largo)
            elif isinstance(item, (list, tuple)) and len(item) >= 2:
                _absorber_equivalencia(mapa, cortos, item[0], item[1])
        return
    if not isinstance(datos, dict):
        return
    for clave in ("equivalencias", "pares", "mapeo", "diccionario"):
        if clave in datos:
            _absorber_json_equivalencias(datos[clave], mapa, cortos)
            return
    for corto, largo in datos.items():
        if isinstance(largo, str):
            _absorber_equivalencia(mapa, cortos, corto, largo)


def _rejilla_hoja_equivalencias(ruta: str) -> list[list]:
    from openpyxl import load_workbook

    kwargs = {"read_only": True, "data_only": True}
    if str(ruta).lower().endswith(".xlsm"):
        kwargs["keep_vba"] = True
    wb = load_workbook(ruta, **kwargs)
    try:
        hoja = None
        for nombre in wb.sheetnames:
            if "UBICACION" in _norm_titulo(nombre):
                hoja = wb[nombre]
                break
        if hoja is None:
            return []
        return [list(fila) for fila in hoja.iter_rows(values_only=True)]
    finally:
        wb.close()


def _absorber_rejilla_equivalencias(grid: list[list], mapa: dict[str, str], cortos: list[str]) -> None:
    encabezados = [
        indice
        for indice, fila in enumerate(grid)
        if any(_es_encabezado_corto(valor) or _es_encabezado_largo(valor) for valor in fila)
    ]
    for posicion, fila_header in enumerate(encabezados):
        fila = grid[fila_header]
        cortos_col = [col for col, valor in enumerate(fila) if _es_encabezado_corto(valor)]
        largos_col = [col for col, valor in enumerate(fila) if _es_encabezado_largo(valor)]
        fin = encabezados[posicion + 1] if posicion + 1 < len(encabezados) else len(grid)
        for col_corto in cortos_col:
            if not largos_col:
                continue
            col_largo = min(largos_col, key=lambda col: (abs(col - col_corto), col))
            for fila_datos in range(fila_header + 1, fin):
                renglon = grid[fila_datos]
                corto = renglon[col_corto] if col_corto < len(renglon) else ""
                largo = renglon[col_largo] if col_largo < len(renglon) else ""
                _absorber_equivalencia(mapa, cortos, corto, largo)


def _rutas_catalogo_equivalencias() -> list[str]:
    rutas = []
    if os.path.isdir(UBICACIONES_EXCEL_DIR):
        for nombre in sorted(os.listdir(UBICACIONES_EXCEL_DIR)):
            if nombre.startswith("~$"):
                continue
            bajo = nombre.casefold()
            if "ubicacion" not in bajo or not bajo.endswith((".json", ".xlsx", ".xlsm")):
                continue
            rutas.append(os.path.join(UBICACIONES_EXCEL_DIR, nombre))
    if os.path.isfile(PLANTILLA_SUMINISTROS):
        rutas.append(PLANTILLA_SUMINISTROS)
    return rutas


def _equivalencias_pozo() -> tuple[dict[str, str], list[str]]:
    firmas = []
    rutas = _rutas_catalogo_equivalencias()
    for ruta in rutas:
        try:
            firmas.append((os.path.abspath(ruta), os.path.getmtime(ruta)))
        except OSError:
            continue
    firma = tuple(firmas)
    cache = getattr(_equivalencias_pozo, "cache", None)
    if cache and cache[0] == firma:
        return cache[1], cache[2]
    mapa: dict[str, str] = {}
    cortos: list[str] = []
    plantilla = os.path.abspath(PLANTILLA_SUMINISTROS) if os.path.isfile(PLANTILLA_SUMINISTROS) else ""
    for ruta in rutas:
        try:
            if os.path.abspath(ruta) == plantilla and mapa:
                continue
            if ruta.casefold().endswith(".json"):
                _absorber_json_equivalencias(leer_json(ruta, {}), mapa, cortos)
            else:
                _absorber_rejilla_equivalencias(_rejilla_hoja_equivalencias(ruta), mapa, cortos)
        except Exception:
            continue
    _equivalencias_pozo.cache = (firma, mapa, cortos)
    return mapa, cortos


def diccionario_equivalencias_pozo() -> dict[str, str]:
    """Nombre corto plano (PSN334, PA67) → nombre largo de la matriz."""
    return _equivalencias_pozo()[0]


def nombres_cortos_oficiales() -> list[str]:
    """Nombres cortos del catálogo oficial, en el orden del archivo."""
    return list(_equivalencias_pozo()[1])


def nombre_largo_excel(corto: str, diccionario: dict[str, str] | None = None) -> str:
    """Traduce el pozo de la bitácora al nombre largo con el que la matriz identifica la fila."""
    bruto = " ".join(str(corto or "").split())
    if not bruto:
        return ""
    if diccionario is None:
        diccionario = diccionario_equivalencias_pozo()
    plana = clave_pozo_plana(bruto)
    if plana in diccionario:
        return diccionario[plana]
    compacto = re.sub(r"[^A-Z0-9]", "", _norm_etiqueta_equivalencia(bruto))
    for largo in diccionario.values():
        if compacto and compacto == re.sub(r"[^A-Z0-9]", "", _norm_etiqueta_equivalencia(largo)):
            return largo
    for prefijo, zona in _PREFIJOS_NOMBRE_LARGO:
        if plana.startswith(prefijo) and len(plana) > len(prefijo) and plana[len(prefijo)].isdigit():
            return f"{zona} {plana[len(prefijo):]}"
    return bruto


def _matriz_dias(hoja) -> tuple[int, int, dict[int, int]] | None:
    """Encabezado, columna del pozo y columna de cada día. Con el pozo en B, el día 1 queda en C y el 10 en L."""
    limite_fila = min(int(hoja.max_row or 1), 6)
    limite_col = min(int(hoja.max_column or 1), 45)
    for fila in range(1, limite_fila + 1):
        dias: dict[int, int] = {}
        pozo_col = None
        for col in range(1, limite_col + 1):
            valor = hoja.cell(fila, col).value
            if isinstance(valor, (int, float)) and not isinstance(valor, bool) and float(valor).is_integer():
                dia = int(valor)
                if 1 <= dia <= 31:
                    dias[dia] = col
                continue
            if _norm_etiqueta_equivalencia(valor).startswith("POZO"):
                pozo_col = col
        if pozo_col and len(dias) >= 28 and dias.get(1) and dias.get(10):
            return fila, pozo_col, dias
    return None


def _indice_filas_pozo(hoja, col_pozo: int, fila_ini: int, fila_fin: int) -> dict[str, int]:
    indice: dict[str, int] = {}
    for fila in range(fila_ini, fila_fin + 1):
        valor = hoja.cell(fila, col_pozo).value
        if not isinstance(valor, str) or not valor.strip() or valor.startswith("="):
            continue
        if _norm_etiqueta_equivalencia(valor).startswith("TOTAL"):
            break
        for clave in _claves_nombre_excel(valor):
            indice.setdefault(clave, fila)
    return indice


def _claves_busqueda_excel(nombre: str) -> set[str]:
    """La matriz de relleno escribe COYOTES y el catálogo, COYOTE. Las dos formas localizan la misma fila."""
    texto = _norm_etiqueta_equivalencia(nombre)
    variantes = {texto} if texto else set()
    if "COYOTES" in texto:
        variantes.add(texto.replace("COYOTES", "COYOTE"))
    elif "COYOTE" in texto:
        variantes.add(texto.replace("COYOTE", "COYOTES"))
    claves: set[str] = set()
    for variante in variantes:
        claves |= _claves_nombre_excel(variante)
    return claves


def _buscar_fila_pozo(indice: dict[str, int], nombre: str) -> int | None:
    for clave in _claves_busqueda_excel(nombre):
        fila = indice.get(clave)
        if fila:
            return fila
    return None


def _anclas_combinadas(hoja) -> dict[tuple[int, int], tuple[int, int]]:
    """Celda esclava → esquina superior izquierda del rango combinado."""
    cache = getattr(hoja, "_pjp4_anclas", None)
    if cache is not None:
        return cache
    mapa: dict[tuple[int, int], tuple[int, int]] = {}
    for rango in hoja.merged_cells.ranges:
        ancla = (rango.min_row, rango.min_col)
        for fila in range(rango.min_row, rango.max_row + 1):
            for col in range(rango.min_col, rango.max_col + 1):
                if (fila, col) != ancla:
                    mapa[(fila, col)] = ancla
    hoja._pjp4_anclas = mapa
    return mapa


def _celda_escribible(celda) -> bool:
    from openpyxl.cell.cell import MergedCell

    if isinstance(celda, MergedCell):
        return False
    valor = celda.value
    return not (isinstance(valor, str) and valor.startswith("="))


def _destino_escritura(hoja, fila: int, col: int):
    """Celda donde se puede asignar value. En un rango combinado, solo la superior izquierda."""
    from openpyxl.cell.cell import MergedCell

    celda = hoja.cell(fila, col)
    if not isinstance(celda, MergedCell):
        return celda if _celda_escribible(celda) else None
    ancla = _anclas_combinadas(hoja).get((fila, col))
    if not ancla:
        return None
    principal = hoja.cell(ancla[0], ancla[1])
    if isinstance(principal, MergedCell) or not _celda_escribible(principal):
        return None
    if principal.row != fila:
        return None
    return principal


def _poner_valor_celda(hoja, fila: int, col: int, valor, formato: str = "") -> bool:
    celda = _destino_escritura(hoja, fila, col)
    if celda is None:
        return False
    celda.value = valor
    if formato:
        celda.number_format = formato
    return True


def _sumar_litros(hoja, fila: int, col: int, litros: float) -> bool:
    celda = _destino_escritura(hoja, fila, col)
    if celda is None:
        return False
    actual = celda.value
    base = float(actual) if isinstance(actual, (int, float)) and not isinstance(actual, bool) else 0.0
    celda.value = _litros_celda(base + float(litros))
    return True


def _limpiar_columnas_datos(hoja, fila_ini: int, fila_fin: int, columnas: list[int]) -> None:
    for fila in range(fila_ini, fila_fin + 1):
        for col in columnas:
            celda = _destino_escritura(hoja, fila, col)
            if celda is not None and celda.value is not None:
                celda.value = None


def _desunir_columnas(hoja, fila_ini: int, fila_fin: int, columnas: list[int]) -> None:
    buscadas = set(columnas)
    quitar = []
    for rango in hoja.merged_cells.ranges:
        if rango.min_row < fila_ini or rango.max_row > fila_fin:
            continue
        cols = set(range(rango.min_col, rango.max_col + 1))
        if cols and cols <= buscadas:
            quitar.append(str(rango))
    for referencia in quitar:
        hoja.unmerge_cells(referencia)
    if hasattr(hoja, "_pjp4_anclas"):
        del hoja._pjp4_anclas


def _limite_matriz_dias(hoja, col_dia: int, fila_ini: int) -> int:
    tope = min(int(hoja.max_row or fila_ini), fila_ini + 800)
    for fila in range(fila_ini, tope + 1):
        valor = hoja.cell(fila, col_dia).value
        if isinstance(valor, str) and valor.startswith("="):
            return fila - 1
    return tope


def _agrupar_litros_por_dia(items: list[dict]) -> dict[str, dict[int, float]]:
    acumulado: dict[str, dict[int, float]] = {}
    for item in items:
        fecha = item.get("fecha")
        if not isinstance(fecha, datetime):
            continue
        pozo = str(item.get("pozo") or "").strip()
        if not pozo:
            continue
        dias = acumulado.setdefault(pozo, {})
        dias[int(fecha.day)] = dias.get(int(fecha.day), 0) + float(item.get("litros") or 0)
    return acumulado


def _relleno_por_pozo(relleno: dict[tuple[str, str], dict[int, float]]) -> dict[str, dict[int, float]]:
    acumulado: dict[str, dict[int, float]] = {}
    for (_cuadrilla, pozo), dias in relleno.items():
        destino = acumulado.setdefault(str(pozo), {})
        for dia, litros in dias.items():
            destino[int(dia)] = destino.get(int(dia), 0) + float(litros)
    return acumulado


def _vaciar_matriz_dias(hoja, consumos: dict[str, dict[int, float]]) -> tuple[int, list[str]]:
    layout = _matriz_dias(hoja)
    if layout is None:
        return 0, list(consumos)
    _fila_header, col_pozo, dias = layout
    fila_ini = _fila_header + 1
    fila_fin = _limite_matriz_dias(hoja, dias[1], fila_ini)
    columnas = list(dias.values())
    _desunir_columnas(hoja, fila_ini, fila_fin, columnas)
    _limpiar_columnas_datos(hoja, fila_ini, fila_fin, columnas)
    indice = _indice_filas_pozo(hoja, col_pozo, fila_ini, fila_fin)
    puestos = 0
    faltantes: list[str] = []
    for pozo, por_dia in consumos.items():
        fila = _buscar_fila_pozo(indice, pozo)
        if not fila:
            faltantes.append(pozo)
            continue
        escribio = False
        for dia, litros in por_dia.items():
            columna = dias.get(int(dia))
            if not columna or litros <= 0:
                continue
            if _sumar_litros(hoja, fila, columna, litros):
                escribio = True
        if escribio:
            puestos += 1
    return puestos, faltantes


def _columna_aceite_lista(hoja) -> int:
    limite = min(int(hoja.max_column or 1), 40)
    for col in range(1, limite + 1):
        texto = _norm_etiqueta_equivalencia(hoja.cell(2, col).value)
        if "ACEITE SAE" in texto or texto.startswith("ACEITE "):
            return col
    return 10


def _grupos_filas_lista(hoja, fila_ini: int, fila_fin: int) -> dict[str, list[int]]:
    grupos: dict[str, list[int]] = {}
    for fila in range(fila_ini, fila_fin + 1):
        valor = hoja.cell(fila, 1).value
        if not isinstance(valor, str) or not valor.strip() or valor.startswith("="):
            continue
        nombre = _norm_etiqueta_equivalencia(valor)
        if nombre in {"MXN", "USD"} or nombre.startswith("TOTAL"):
            continue
        for clave in _claves_nombre_excel(valor):
            filas = grupos.setdefault(clave, [])
            if fila not in filas:
                filas.append(fila)
    return grupos


def _vaciar_lista_mantenimiento(hoja, programados: list[dict]) -> tuple[int, list[str]]:
    """Escribe el aceite de mantenimiento en la fila del nombre largo, sin borrar el esqueleto ni las sumas."""
    fila_ini, fila_fin = _fila_datos_programados(hoja)
    col_aceite = _columna_aceite_lista(hoja)
    columnas = [2, 3, 4, 5, 6, 7, col_aceite]
    _desunir_columnas(hoja, fila_ini, fila_fin, columnas)
    _limpiar_columnas_datos(hoja, fila_ini, fila_fin, columnas)
    grupos = _grupos_filas_lista(hoja, fila_ini, fila_fin)
    usados: set[int] = set()
    puestos = 0
    faltantes: list[str] = []
    for item in programados:
        pozo = str(item.get("pozo") or "").strip()
        fila = None
        for clave in _claves_busqueda_excel(pozo):
            for candidata in grupos.get(clave, []):
                if candidata not in usados:
                    fila = candidata
                    break
            if fila:
                break
        if fila is None:
            ultima = None
            for clave in _claves_busqueda_excel(pozo):
                if grupos.get(clave):
                    ultima = grupos[clave][-1]
                    break
            if ultima is None:
                faltantes.append(pozo)
                continue
            if _sumar_litros(hoja, ultima, col_aceite, item["litros"]):
                puestos += 1
            continue
        usados.add(fila)
        if item.get("ot"):
            _poner_valor_celda(hoja, fila, 2, item["ot"])
        if item.get("trabajo"):
            _poner_valor_celda(hoja, fila, 3, item["trabajo"])
        if item.get("fecha"):
            _poner_valor_celda(hoja, fila, 5, item["fecha"], "DD/MM/YYYY")
        if item.get("cuadrilla"):
            _poner_valor_celda(hoja, fila, 6, item["cuadrilla"])
        if item.get("ubm"):
            _poner_valor_celda(hoja, fila, 8, item["ubm"])
        if item.get("motor"):
            _poner_valor_celda(hoja, fila, 9, item["motor"])
        _sumar_litros(hoja, fila, col_aceite, item["litros"])
        puestos += 1
    return puestos, faltantes


def _hoja_matriz_por_nombre(wb, fragmento: str):
    for nombre in wb.sheetnames:
        if fragmento in nombre.casefold() and _matriz_dias(wb[nombre]):
            return wb[nombre]
    return None


def _aceite_en_rango(desde: date, hasta: date) -> tuple[list[dict], dict[tuple[str, str], dict[int, float]]]:
    inicio, fin = min(desde, hasta), max(desde, hasta)
    base = bitacora_con_fechas(cargar_bitacora())
    programados: list[dict] = []
    relleno: dict[tuple[str, str], dict[int, float]] = {}
    if base.empty or "_fecha" not in base.columns:
        return programados, relleno
    diccionario = diccionario_equivalencias_pozo()
    programadas = _actividades_programadas_suministros()
    periodo = base.loc[
        base["_fecha"].map(lambda dia: isinstance(dia, date) and inicio <= dia <= fin)
    ]
    for _, fila in periodo.iterrows():
        litros = _numero_insumo(fila.get("ACEITE (LTS)"))
        if litros <= 0:
            continue
        dia = fila.get("_fecha")
        if not isinstance(dia, date):
            continue
        destino = _destino_hoja_suministros(
            fila.get("TIPO DE ACTIVIDAD"),
            programadas,
            _descripcion_fila_admin(fila),
        )
        if not destino:
            continue
        tag = _ubicacion_fila_refaccion(fila)
        if not tag or tag == "Sin ubicación":
            tag = _texto_celda_admin(fila.get("POZO")) or _texto_celda_admin(fila.get("UBICACIÓN"))
        pozo = nombre_largo_excel(tag or "Sin pozo", diccionario)
        cuadrilla = _cuadrilla_plantilla_suministros(fila.get("CUADRILLA"))
        if destino == "programado":
            ubm, motor = _ubm_y_motor(_equipo_fila_suministros(fila))
            programados.append(
                {
                    "pozo": pozo,
                    "ot": numero_secuencia_actividad(fila.get("SECUENCIA")) or _texto_celda_admin(fila.get("SECUENCIA")),
                    "trabajo": _descripcion_fila_admin(fila) or _texto_celda_admin(fila.get("TIPO DE ACTIVIDAD")),
                    "fecha": datetime(dia.year, dia.month, dia.day),
                    "cuadrilla": cuadrilla,
                    "ubm": ubm,
                    "motor": motor,
                    "litros": litros,
                }
            )
            continue
        clave = (cuadrilla, pozo)
        dias = relleno.setdefault(clave, {})
        dias[dia.day] = dias.get(dia.day, 0) + litros
    programados.sort(key=lambda item: (item["fecha"], item["pozo"], str(item["ot"])))
    return programados, relleno


_MESES_EN_HOJA = (
    "SEPTIEMBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
    "FEBRERO",
    "OCTUBRE",
    "AGOSTO",
    "ENERO",
    "MARZO",
    "ABRIL",
    "JUNIO",
    "JULIO",
    "MAYO",
)


def _mes_anio_reporte(desde: date, hasta: date) -> tuple[int, int]:
    """Mes del rango. Si cruza dos meses, se queda el que tiene más días."""
    inicio, fin = min(desde, hasta), max(desde, hasta)
    conteo: dict[tuple[int, int], int] = {}
    dia = inicio
    while dia <= fin:
        clave = (dia.year, dia.month)
        conteo[clave] = conteo.get(clave, 0) + 1
        dia += timedelta(days=1)
    anio, mes = max(conteo, key=lambda clave: (conteo[clave], clave[0], clave[1]))
    return mes, anio


def _etiqueta_mes_hoja(mes: int, anio: int) -> str:
    return f"{_sin_acentos(NOMBRES_MES[mes - 1]).upper()} {anio}"


def _acortar_nombre_hoja(nombre: str) -> str:
    limpio = re.sub(r"\s+", " ", nombre).strip()
    while len(limpio) > 31 and " DE " in limpio:
        limpio = limpio.replace(" DE ", " ", 1)
    if len(limpio) > 31:
        limpio = limpio[:31].rstrip(" .,")
    return limpio


def _nombre_hoja_con_mes(actual: str, mes: int, anio: int) -> str:
    texto = _sin_acentos(str(actual or "")).upper()
    etiqueta = _etiqueta_mes_hoja(mes, anio)
    for nombre in _MESES_EN_HOJA:
        coincidencia = re.search(rf"\b{nombre}\b(?:\s+\d{{4}})?", texto)
        if not coincidencia:
            continue
        return _acortar_nombre_hoja(f"{texto[:coincidencia.start()]}{etiqueta}{texto[coincidencia.end():]}")
    return _acortar_nombre_hoja(f"{texto.strip(' .')} {etiqueta}")


def _reescribir_referencias_hoja(wb, anterior: str, nuevo: str) -> None:
    if not anterior or anterior == nuevo:
        return
    for hoja in wb.worksheets:
        for fila in hoja.iter_rows():
            for celda in fila:
                valor = celda.value
                if not isinstance(valor, str) or anterior not in valor:
                    continue
                celda.value = valor.replace(f"'{anterior}'", f"'{nuevo}'").replace(anterior, nuevo)
    for definicion in wb.defined_names.values():
        texto = str(getattr(definicion, "attr_text", "") or "")
        if anterior in texto:
            definicion.attr_text = texto.replace(f"'{anterior}'", f"'{nuevo}'").replace(anterior, nuevo)


def _renombrar_hojas_al_mes(wb, hojas: list, mes: int, anio: int) -> list[str]:
    titulos = []
    ocupados = set(wb.sheetnames)
    for hoja in hojas:
        if hoja is None:
            continue
        anterior = hoja.title
        nuevo = _nombre_hoja_con_mes(anterior, mes, anio)
        if nuevo in ocupados and nuevo != anterior:
            continue
        _reescribir_referencias_hoja(wb, anterior, nuevo)
        hoja.title = nuevo
        ocupados.discard(anterior)
        ocupados.add(nuevo)
        titulos.append(nuevo)
    return titulos


def generar_reporte_suministros_xlsm(desde: date, hasta: date) -> dict:
    """Llena la plantilla oficial sin tocar macros ni fórmulas.

    El nombre corto de la bitácora se traduce al nombre largo y la cantidad
    se suma en la fila de ese pozo. En la matriz diaria, la columna C es el día 1.
    Las hojas de proyecto y relleno quedan con el mes del rango.
    Incluye cada consumo del rango, tenga o no fotografías.
    """
    from openpyxl import load_workbook

    ruta = _ruta_plantilla_suministros()
    wb = load_workbook(ruta, keep_vba=True)
    try:
        if wb.vba_archive is None:
            raise ValueError("La plantilla perdió las macros. No se generó el archivo.")
        hoja_manttos = _hoja_matriz_por_nombre(wb, "aceite de manttos")
        hoja_relleno = _hoja_matriz_por_nombre(wb, "aceite de relleno")
        if hoja_relleno is None:
            candidata = _hoja_relleno_suministros(wb)
            if candidata is not None and _matriz_dias(candidata):
                hoja_relleno = candidata
        hoja_lista = None if hoja_manttos is not None else _hoja_programados_suministros(wb)
        if hoja_manttos is None and hoja_relleno is None and hoja_lista is None:
            raise ValueError("La plantilla no tiene las hojas de aceite de manttos y de relleno.")
        programados, relleno = _aceite_en_rango(desde, hasta)
        sin_fila: list[str] = []
        if hoja_manttos is not None:
            puestos_programados, faltan = _vaciar_matriz_dias(hoja_manttos, _agrupar_litros_por_dia(programados))
            sin_fila.extend(faltan)
        elif hoja_lista is not None:
            puestos_programados, faltan = _vaciar_lista_mantenimiento(hoja_lista, programados)
            sin_fila.extend(faltan)
        else:
            puestos_programados = 0
        if hoja_relleno is not None:
            puestos_relleno, faltan = _vaciar_matriz_dias(hoja_relleno, _relleno_por_pozo(relleno))
            sin_fila.extend(faltan)
        else:
            puestos_relleno = 0
        mes, anio = _mes_anio_reporte(desde, hasta)
        hojas_mes = []
        for hoja in (hoja_manttos, hoja_lista, hoja_relleno):
            if hoja is not None and hoja not in hojas_mes:
                hojas_mes.append(hoja)
        titulos_mes = _renombrar_hojas_al_mes(wb, hojas_mes, mes, anio)
        detalle = _insumos_rango_admin(desde, hasta)
        if "Detalle de insumos" in wb.sheetnames:
            del wb["Detalle de insumos"]
        detalle_hoja = wb.create_sheet("Detalle de insumos")
        for columna, titulo in enumerate(("Fecha", "Ubicación", "Rubro", "Insumo", "Cantidad"), start=1):
            detalle_hoja.cell(1, columna).value = titulo
        for desplazamiento, fila in enumerate(detalle.itertuples(index=False), start=2):
            detalle_hoja.cell(desplazamiento, 1).value = fila.fecha
            detalle_hoja.cell(desplazamiento, 2).value = fila.ubicacion
            detalle_hoja.cell(desplazamiento, 3).value = fila.rubro
            detalle_hoja.cell(desplazamiento, 4).value = fila.refaccion
            detalle_hoja.cell(desplazamiento, 5).value = fila.cantidad
        buffer = io.BytesIO()
        wb.save(buffer)
    finally:
        wb.close()
    inicio, fin = min(desde, hasta), max(desde, hasta)
    return {
        "contenido": buffer.getvalue(),
        "nombre": f"suministros_{inicio.strftime('%Y%m%d')}_{fin.strftime('%Y%m%d')}.xlsm",
        "programados": puestos_programados,
        "relleno": puestos_relleno,
        "insumos": int(len(detalle)),
        "sin_fila": len(list(dict.fromkeys(sin_fila))),
        "mes": _etiqueta_mes_hoja(mes, anio),
        "hojas": titulos_mes,
    }


def _insumos_rango_admin(desde: date, hasta: date) -> pd.DataFrame:
    if not isinstance(desde, date) or not isinstance(hasta, date):
        return pd.DataFrame(columns=["fecha", "ubicacion", "rubro", "refaccion", "cantidad"])
    inicio, fin = min(desde, hasta), max(desde, hasta)
    base = bitacora_con_fechas(cargar_bitacora())
    if base.empty or "_fecha" not in base.columns:
        return pd.DataFrame(columns=["fecha", "ubicacion", "rubro", "refaccion", "cantidad"])
    periodo = base.loc[
        base["_fecha"].map(lambda dia: isinstance(dia, date) and inicio <= dia <= fin)
    ].copy()
    detalle = detalle_insumos_suministrados(periodo)
    if detalle.empty:
        return detalle
    detalle = detalle.copy()
    detalle["fecha"] = detalle["fecha"].map(lambda dia: dia.strftime("%d/%m/%Y") if isinstance(dia, date) else dia)
    return detalle


def render_extraccion_evidencia_admin(jornada: date) -> None:
    st.caption(
        "El Excel usa la Actividad oficial. Si el registro antiguo no la trae, lee la descripción: "
        "«Inspección y suministro» va a relleno; bimestral, semestral, habilitado, puesta en marcha o preventivo va a programados. "
        "El nombre corto (PSN334, PA67) se traduce al nombre largo de la matriz (SOLEDAD NORTE 334, ARAGON 67) "
        "y la cantidad se suma en la fila de ese pozo. En la matriz diaria, C es el día 1 y L es el día 10. "
        "Las pestañas de proyecto y relleno toman el mes del rango, por ejemplo NOVIEMBRE 2026."
    )
    desde_col, hasta_col = st.columns(2)
    with desde_col:
        desde = st.date_input(
            "Desde",
            value=date(jornada.year, jornada.month, 1),
            format="DD/MM/YYYY",
            key="adm_m_ins_desde",
        )
    with hasta_col:
        hasta = st.date_input(
            "Hasta",
            value=jornada,
            format="DD/MM/YYYY",
            key="adm_m_ins_hasta",
        )
    if st.button("Generar reporte de insumos", use_container_width=True, key="adm_m_gen_insumos"):
        try:
            paquete = generar_reporte_suministros_xlsm(desde, hasta)
        except Exception as exc:
            st.error(f"No se pudo llenar la plantilla de suministros: {exc}")
        else:
            detalle = _insumos_rango_admin(desde, hasta)
            st.session_state.adm_m_ins_xlsx = paquete["contenido"]
            st.session_state.adm_m_ins_nombre = paquete["nombre"]
            st.session_state.adm_m_ins_csv = detalle.to_csv(index=False).encode("utf-8-sig")
            st.session_state.adm_m_ins_filas = int(paquete["programados"] + paquete["relleno"])
            st.session_state.adm_m_ins_resumen = (
                f"{paquete.get('mes', '')}: "
                f"{paquete['programados']} mantenimiento(s) programado(s) · "
                f"{paquete['relleno']} pozo(s) de relleno · "
                f"{paquete.get('insumos', 0)} consumo(s) de insumos."
            ).lstrip(": ")
            if paquete.get("sin_fila"):
                st.session_state.adm_m_ins_resumen += (
                    f" {paquete['sin_fila']} pozo(s) no están en la matriz."
                )
    if st.session_state.get("adm_m_ins_xlsx"):
        st.caption(st.session_state.get("adm_m_ins_resumen") or "Plantilla de suministros lista.")
        marca = st.session_state.get("adm_m_ins_nombre") or "suministros.xlsm"
        baja_xls, baja_csv = st.columns(2)
        with baja_xls:
            st.download_button(
                "Descargar Excel",
                data=st.session_state.adm_m_ins_xlsx,
                file_name=marca,
                mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                use_container_width=True,
                key="adm_m_dl_xlsx",
            )
        with baja_csv:
            nombre_csv = marca[:-5] + ".csv" if str(marca).endswith(".xlsm") else "insumos.csv"
            st.download_button(
                "Descargar CSV",
                data=st.session_state.adm_m_ins_csv,
                file_name=nombre_csv,
                mime="text/csv",
                use_container_width=True,
                key="adm_m_dl_csv",
            )
    with st.expander("Reporte fotográfico PDF", expanded=False):
        render_generador_reporte_fotografico(key_prefix="adm_matriz_foto")


def _desplazar_mes_admin(jornada: date, delta: int) -> date:
    indice = jornada.month - 1 + delta
    anio = jornada.year + indice // 12
    mes = indice % 12 + 1
    ultimo = calendar.monthrange(anio, mes)[1]
    return date(anio, mes, min(jornada.day, ultimo))


def render_calendario_jornada_admin(jornada: date, metricas: dict) -> None:
    with st.container(border=True):
        mes_txt = NOMBRES_MES[jornada.month - 1]
        anterior, titulo, siguiente = st.columns([1, 4, 1])
        with anterior:
            if st.button("‹", key="adm_mes_ant", use_container_width=True):
                st.session_state.admin_jornada_pendiente = _desplazar_mes_admin(jornada, -1)
                st.rerun()
        with titulo:
            st.markdown(
                f'<p class="cal-mes-admin">{html.escape(mes_txt)} {jornada.year}</p>',
                unsafe_allow_html=True,
            )
        with siguiente:
            if st.button("›", key="adm_mes_sig", use_container_width=True):
                st.session_state.admin_jornada_pendiente = _desplazar_mes_admin(jornada, 1)
                st.rerun()
        render_mini_calendario_admin(jornada, metricas["fechas_bit"], metricas["fechas_ot"])


def render_seccion_completa_admin(vista: str, jornada: date) -> None:
    titulos = {
        "entrada": "ENTRADA DE DATOS",
        "edicion": "EDICION DE CAPTURA",
        "previa": "VISTA PREVIA",
        "extraccion": "EXTRACCION DE EVIDENCIA",
    }
    if st.button(
        "← Volver al Menú Administrativo",
        type="primary",
        use_container_width=True,
        key="adm_volver_menu",
    ):
        st.session_state.admin_vista = None
        st.rerun()
    st.markdown(
        f'<h2 class="adm-seccion-titulo">{html.escape(titulos.get(vista, vista))}</h2>',
        unsafe_allow_html=True,
    )
    if vista == "entrada":
        render_entrada_datos_admin()
    elif vista == "edicion":
        render_edicion_captura_admin(jornada)
    elif vista == "previa":
        render_vista_previa_admin(jornada)
    elif vista == "extraccion":
        render_extraccion_evidencia_admin(jornada)


def render_explorador_admin() -> None:
    st.markdown(
        """
        <style>
            [class*="st-key-adm_card_"] button {
                min-height: 4.75rem;
                border-radius: 14px;
                border: 1px solid #334155 !important;
                background: #0f172a !important;
                color: #f8fafc !important;
                font-weight: 700;
                letter-spacing: 0.08em;
            }
            [data-testid="stHorizontalBlock"]:has(.cal-mes-admin) {
                align-items: center;
            }
            .cal-mes-admin {
                margin: 0.35rem 0 0.15rem;
                text-align: center;
                color: #f8fafc;
                font-weight: 700;
                font-size: 1.05rem;
            }
            .adm-seccion-titulo {
                margin: 0.2rem 0 1rem;
                text-align: center;
                color: #f8fafc;
                letter-spacing: 0.08em;
            }
            .adm-sync {
                margin: 0;
                color: #e2e8f0;
                font-size: 0.92rem;
                line-height: 1.35;
            }
            [data-testid="stAppViewContainer"] .block-container {
                padding-top: 1.1rem;
                padding-bottom: 0.6rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )
    pendiente = st.session_state.pop("admin_jornada_pendiente", None)
    if isinstance(pendiente, date):
        st.session_state.admin_jornada = pendiente
    if "admin_jornada" not in st.session_state or not isinstance(st.session_state.admin_jornada, date):
        st.session_state.admin_jornada = parsear_fecha_valor(st.session_state.get("admin_jornada")) or date.today()
    jornada = st.session_state.admin_jornada
    vista = st.session_state.get("admin_vista")
    if vista in {"entrada", "edicion", "previa", "extraccion"}:
        render_seccion_completa_admin(vista, jornada)
        return
    metricas = metricas_nucleo_admin(jornada)
    col_izq, col_centro, col_der = st.columns(
        [1.12, 1.35, 1.12],
        gap="medium",
        vertical_alignment="center",
    )
    with col_izq:
        if st.button("ENTRADA DE DATOS", use_container_width=True, key="adm_card_entrada"):
            st.session_state.admin_vista = "entrada"
            st.rerun()
        if st.button("EDICION DE CAPTURA", use_container_width=True, key="adm_card_edicion"):
            st.session_state.admin_vista = "edicion"
            st.rerun()
    with col_centro:
        render_calendario_jornada_admin(jornada, metricas)
    with col_der:
        if st.button("VISTA PREVIA", use_container_width=True, key="adm_card_previa"):
            st.session_state.admin_vista = "previa"
            st.rerun()
        if st.button("EXTRACCION DE EVIDENCIA", use_container_width=True, key="adm_card_extraccion"):
            st.session_state.admin_vista = "extraccion"
            st.rerun()

    pie_ub, pie_hist = st.columns(2)
    with pie_ub:
        if st.button("UBICACION Y EQUIPOS", use_container_width=True, key="adm_pie_ubicaciones"):
            abrir_modulo_admin("ubicaciones")
    with pie_hist:
        if st.button("HISTORIAL", use_container_width=True, key="adm_pie_historial"):
            abrir_modulo_admin("historial_base")
    pie_prog, pie_ejec = st.columns(2)
    with pie_prog:
        if st.button("MANTENIMIENTOS PROGRAMADOS", use_container_width=True, key="adm_pie_programados"):
            abrir_modulo_admin("mantenimientos_programados")
    with pie_ejec:
        if st.button("REPORTES EJECUTIVOS", use_container_width=True, key="adm_pie_ejecutivos"):
            abrir_modulo_admin("reportes_ejecutivos")


CRITERIOS_REPORTE_EJECUTIVO = (
    ("Bandas", "insumo", ("banda",), ()),
    ("Refrigerantes", "insumo", ("refrigerante", "anticongelante"), ()),
    ("Aceite", "insumo", ("aceite",), ("filtro",)),
    ("Grasa", "insumo", ("grasa",), ()),
    ("Filtros", "insumo", ("filtro",), ()),
    ("Bujías", "insumo", ("bujia",), ()),
    ("Mantenimientos mensuales y bimestrales", "mantenimiento", ("mensual", "bimestral"), ("semestral",)),
    ("Mantenimientos semestrales", "mantenimiento", ("semestral",), ()),
)


def _criterio_reporte_ejecutivo(nombre: str) -> tuple[str, str, tuple[str, ...], tuple[str, ...]]:
    for criterio in CRITERIOS_REPORTE_EJECUTIVO:
        if criterio[0] == nombre:
            return criterio
    return CRITERIOS_REPORTE_EJECUTIVO[0]


def _texto_coincide_criterio(texto: str, tokens: tuple[str, ...], excluir: tuple[str, ...]) -> bool:
    clave = _sin_acentos(texto).casefold()
    if any(pieza in clave for pieza in excluir):
        return False
    return any(pieza in clave for pieza in tokens)


def _descripcion_fila_reporte(fila) -> str:
    for columna in (
        "DESCRIPCION DE LA ACTIVIDAD",
        "DESCRIPCIÓN DE LA ACTIVIDAD",
        "ACTIVIDAD CORTA",
        "TIPO DE ACTIVIDAD",
    ):
        texto = texto_visible(fila.get(columna), "")
        if texto:
            return texto
    return ""


def _cuadrilla_personal_reporte(fila) -> str:
    cuadrilla = etiqueta_cuadrilla_historial(str(fila.get("CUADRILLA") or ""))
    personal = texto_visible(fila.get("PERSONAL") or fila.get("MECANICO EN TURNO"), "")
    if personal:
        return f"{cuadrilla} / {personal}"
    return cuadrilla


def _ot_fila_reporte(fila) -> str:
    numero = numero_secuencia_actividad(fila.get("SECUENCIA"))
    if numero is None:
        return ""
    return etiqueta_secuencia_consecutiva(numero)


def _insumos_fila_reporte(fila) -> list[tuple[str, float]]:
    columnas = []
    for columna in getattr(fila, "index", []):
        nombre = str(columna)
        if not _es_columna_cantidad_refaccion(nombre):
            continue
        origen = _ESPEJO_INSUMO_BITACORA.get(nombre)
        if origen and origen in fila.index:
            continue
        columnas.append(nombre)
    salida = []
    for columna in columnas:
        bruto = fila.get(columna)
        cantidad = cantidad_grasa_kg(bruto) if _es_grasa(columna) else _numero_insumo(bruto)
        if cantidad <= 0:
            continue
        salida.append((etiqueta_refaccion(columna), float(cantidad)))
    for nombre, cantidad in _partes_otros_insumos(fila.get("OTROS INSUMOS")):
        if cantidad > 0:
            salida.append((nombre, float(cantidad)))
    return salida


def filas_reporte_ejecutivo(df: pd.DataFrame | None, criterio: str) -> list[dict]:
    """Actividades del periodo que corresponden al criterio, una fila por insumo."""
    _etiqueta, tipo, tokens, excluir = _criterio_reporte_ejecutivo(criterio)
    if df is None or df.empty:
        return []
    filas = []
    for _, fila in df.iterrows():
        dia = fila.get("_fecha")
        if not isinstance(dia, date) or isinstance(dia, datetime):
            continue
        descripcion = _descripcion_fila_reporte(fila)
        insumos = _insumos_fila_reporte(fila)
        base = {
            "ot": _ot_fila_reporte(fila),
            "fecha": dia,
            "ubicacion": etiqueta_ubicacion_fila(fila) or formato_ubicacion_reporte(fila),
            "cuadrilla": _cuadrilla_personal_reporte(fila),
            "descripcion": sitios_operativos.quitar_ninguno(descripcion),
        }
        if tipo == "insumo":
            for nombre, cantidad in insumos:
                if _texto_coincide_criterio(nombre, tokens, excluir):
                    filas.append({**base, "insumo": nombre, "cantidad": cantidad})
            continue
        if not _texto_coincide_criterio(descripcion, tokens, excluir):
            continue
        if insumos:
            for nombre, cantidad in insumos:
                filas.append({**base, "insumo": nombre, "cantidad": cantidad})
        else:
            filas.append({**base, "insumo": "", "cantidad": None})
    filas.sort(key=lambda item: (item["fecha"], item["ot"], item["ubicacion"], item["insumo"]))
    return filas


def _tabla_reporte_ejecutivo(filas: list[dict]) -> pd.DataFrame:
    columnas = [
        "N° de OT",
        "Fecha",
        "Ubicación/Pozo",
        "Cuadrilla/Personal asignado",
        "Descripción de la actividad",
        "Insumo/Refacción",
        "Cantidad",
    ]
    if not filas:
        return pd.DataFrame(columns=columnas)
    registros = []
    for fila in filas:
        cantidad = fila.get("cantidad")
        registros.append(
            {
                "N° de OT": fila.get("ot") or "",
                "Fecha": fila["fecha"].strftime("%d/%m/%Y") if isinstance(fila.get("fecha"), date) else "",
                "Ubicación/Pozo": fila.get("ubicacion") or "",
                "Cuadrilla/Personal asignado": fila.get("cuadrilla") or "",
                "Descripción de la actividad": fila.get("descripcion") or "",
                "Insumo/Refacción": fila.get("insumo") or "",
                "Cantidad": "" if cantidad in (None, "") else _texto_cantidad_refaccion(str(fila.get("insumo") or ""), float(cantidad)),
            }
        )
    return pd.DataFrame(registros, columns=columnas)


def render_reportes_ejecutivos() -> None:
    import reporte_ejecutivo

    st.markdown('<div class="section-title">Reportes ejecutivos</div>', unsafe_allow_html=True)
    st.caption("Cruza la bitácora con la plantilla formal para presentarla ante OPS.")
    inicio_mes, fin_mes = _periodo_mes_calendario(date.today())
    col_desde, col_hasta, col_criterio = st.columns([1, 1, 1.45])
    with col_desde:
        inicio = st.date_input("Desde", value=inicio_mes, key="ejec_desde")
    with col_hasta:
        fin = st.date_input("Hasta", value=fin_mes, key="ejec_hasta")
    with col_criterio:
        criterio = st.selectbox(
            "Criterio",
            [item[0] for item in CRITERIOS_REPORTE_EJECUTIVO],
            key="ejec_criterio",
        )
    if not isinstance(inicio, date) or not isinstance(fin, date) or isinstance(inicio, datetime) or isinstance(fin, datetime):
        st.warning("Selecciona un rango de fechas válido.")
        return
    if inicio > fin:
        st.warning("La fecha inicial queda después de la final.")
        return
    periodo = _etiqueta_periodo_admin(inicio, fin)
    filas = filas_reporte_ejecutivo(_bitacora_rango_admin(inicio, fin), criterio)
    st.caption(f"{len(filas)} registro(s) · {periodo}")
    tabla = _tabla_reporte_ejecutivo(filas)
    if tabla.empty:
        st.info("No hay registros de bitácora para ese periodo y criterio.")
    else:
        st.dataframe(tabla, use_container_width=True, hide_index=True)
    try:
        excel = reporte_ejecutivo.generar_excel_reporte_ejecutivo(filas, criterio, periodo)
        pdf = reporte_ejecutivo.generar_pdf_reporte_ejecutivo(filas, criterio, periodo)
    except Exception as exc:
        st.error(f"No se pudo armar el reporte: {exc}")
        return
    slug = re.sub(r"[^a-z0-9]+", "_", _sin_acentos(criterio).casefold()).strip("_") or "reporte"
    nombre = f"reporte_ejecutivo_{slug}_{inicio.strftime('%Y%m%d')}_{fin.strftime('%Y%m%d')}"
    col_excel, col_pdf = st.columns(2)
    with col_excel:
        st.download_button(
            "Descargar Excel",
            data=excel,
            file_name=f"{nombre}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="ejec_descarga_excel",
        )
    with col_pdf:
        st.download_button(
            "Descargar PDF",
            data=pdf,
            file_name=f"{nombre}.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="ejec_descarga_pdf",
        )


def render_historial_base_admin() -> None:
    st.markdown('<div class="section-title">Historial</div>', unsafe_allow_html=True)
    seccion = st.radio(
        "Sección",
        [
            "Historial de OTs",
            "Bitácora General",
            "Calendario de Asistencia de Cuadrillas",
        ],
        horizontal=True,
        key="historial_base_seccion",
    )
    if seccion == "Historial de OTs":
        render_historial_diario_ot()
    elif seccion == "Bitácora General":
        render_revision_bitacoras()
    else:
        render_calendario_asistencia()


def render_barra_modulo_admin(titulo: str) -> None:
    st.markdown(
        f'<div class="explorer-crumb">Panel de Control / <b>{html.escape(titulo)}</b></div>',
        unsafe_allow_html=True,
    )
    if st.button("← Volver al panel de control", key="admin_volver_explorar"):
        st.session_state.admin_modulo = None
        reset_submodulos_admin()
        st.rerun()


def render_auditoria_accesos() -> None:
    st.markdown('<div class="section-title">Auditoría y accesos</div>', unsafe_allow_html=True)
    st.caption("Revisión de bitácoras de campo, catálogos operativos y registro local de permisos.")
    sub = st.session_state.get("admin_sub_auditoria")
    if not sub:
        render_cuadricula_explorador(
            SUBMENUS_AUDITORIA,
            "admin_sub_auditoria",
            "Elige la sección de auditoría o accesos que quieres abrir.",
        )
        return
    render_barra_volver(
        f'Auditoría y Accesos / <b>{html.escape(titulo_submenu(SUBMENUS_AUDITORIA, sub))}</b>',
        "admin_sub_auditoria",
        "← Volver a Auditoría y Accesos",
    )
    if sub == "revision":
        render_revision_bitacoras()
    elif sub == "catalogos":
        render_gestion_catalogos()
    elif sub == "permisos":
        render_control_permisos()


def etiquetas_mes_programa(cruce: pd.DataFrame) -> list[str]:
    if cruce is None or cruce.empty or "MES" not in cruce.columns:
        return []
    orden = cruce.sort_values(by="FECHA", key=lambda serie: serie.map(lambda valor: valor or date.max))
    return [etiqueta for etiqueta in dict.fromkeys(orden["MES"].tolist()) if etiqueta]


def mes_prioridad_programa(etiquetas: list[str], hoy: date | None = None) -> str:
    """El mes en curso, o el cierre más reciente si ese mes todavía no está en el programa."""
    hoy = hoy or date.today()
    actual = _etiqueta_mes_programa(hoy)
    if actual in etiquetas:
        return actual
    fechados = []
    for etiqueta in etiquetas:
        partes = str(etiqueta).split()
        if len(partes) != 2 or partes[0] not in NOMBRES_MES:
            continue
        try:
            anio = int(partes[1])
        except ValueError:
            continue
        fechados.append((date(anio, NOMBRES_MES.index(partes[0]) + 1, 1), etiqueta))
    if not fechados:
        return ""
    inicio_mes = date(hoy.year, hoy.month, 1)
    previos = [item for item in fechados if item[0] <= inicio_mes]
    if previos:
        return max(previos, key=lambda item: item[0])[1]
    return min(fechados, key=lambda item: item[0])[1]


def _marca_archivo_programa(ruta: str) -> str:
    if not os.path.isfile(ruta):
        return "Aún no cargado"
    momento = datetime.fromtimestamp(os.path.getmtime(ruta))
    return f"Actualizado {momento.strftime('%d/%m/%Y %H:%M')}"


def render_mantenimientos_programados() -> None:
    st.markdown(
        '<div class="section-title">Mantenimientos programados</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Carga por separado, o a la vez, el programa de Zona Norte y el de Zona Sur "
        "(por ejemplo Programa de Mantto UBM 2026). Quedan en database/programa_norte.xlsx "
        "y database/programa_sur.xlsx. Las alertas de campo y el avance usan el consolidado de las dos zonas."
    )
    aviso = st.session_state.pop("prog_mtto_aviso", None)
    if aviso:
        st.success(aviso)
    col_norte, col_sur = st.columns(2)
    with col_norte:
        st.markdown("**Zona Norte**")
        st.caption(f"`database/programa_norte.xlsx` · {_marca_archivo_programa(PROGRAMA_NORTE_FILE)}")
        archivo_norte = st.file_uploader(
            "Programa Zona Norte",
            type=["xlsx", "xlsm", "xls", "csv"],
            key="prog_mtto_norte",
        )
    with col_sur:
        st.markdown("**Zona Sur**")
        st.caption(f"`database/programa_sur.xlsx` · {_marca_archivo_programa(PROGRAMA_SUR_FILE)}")
        archivo_sur = st.file_uploader(
            "Programa Zona Sur",
            type=["xlsx", "xlsm", "xls", "csv"],
            key="prog_mtto_sur",
        )
    if st.button("Actualizar programas", type="primary", key="prog_mtto_guardar"):
        if archivo_norte is None and archivo_sur is None:
            st.warning("Elige el archivo de Zona Norte, el de Zona Sur, o los dos.")
        else:
            mensajes = []
            errores = []
            for zona, archivo in (("norte", archivo_norte), ("sur", archivo_sur)):
                if archivo is None:
                    continue
                etiqueta = ZONAS_PROGRAMA[zona]["etiqueta"]
                try:
                    total = importar_programa_zona(zona, archivo.getvalue(), archivo.name)
                except Exception as exc:
                    errores.append(f"{etiqueta}: {exc}")
                else:
                    mensajes.append(f"{etiqueta}: {total} actividad(es)")
            for error in errores:
                st.error(error)
            if mensajes:
                st.session_state.prog_mtto_aviso = "Programa actualizado. " + " · ".join(mensajes)
                st.rerun()

    programa = cargar_programa_mantenimiento()
    if programa.empty:
        st.info("Todavía no hay un programa de Norte ni de Sur. El avance aparece en cuanto se cargue al menos una zona.")
        return

    norte_n = int((programa["ZONA"] == "Norte").sum()) if "ZONA" in programa.columns else 0
    sur_n = int((programa["ZONA"] == "Sur").sum()) if "ZONA" in programa.columns else 0
    st.caption(f"Consolidado: {norte_n} actividades de Norte y {sur_n} de Sur.")

    cruce = cruce_cumplimiento_programa(programa, _bitacora_para_programa())
    meses_historial = etiquetas_mes_programa(cruce)
    mes_actual = mes_prioridad_programa(meses_historial)
    meses = ["Todos", *meses_historial]
    zonas_filtro = ["Todas"]
    for etiqueta in ("Norte", "Sur"):
        if etiqueta in set(cruce["ZONA"].astype(str)):
            zonas_filtro.append(etiqueta)
    filtro_zona, filtro_mes = st.columns(2)
    with filtro_zona:
        zona_sel = st.selectbox("Zona", zonas_filtro, key="prog_mtto_zona")
    with filtro_mes:
        indice_mes = meses.index(mes_actual) if mes_actual in meses else 0
        mes_sel = st.selectbox(
            "Mes del programa",
            meses,
            index=indice_mes,
            key="prog_mtto_mes",
            help="Abre en el mes en curso. El histórico de otros años y los meses siguientes siguen disponibles.",
        )
    base = cruce if zona_sel == "Todas" else cruce[cruce["ZONA"] == zona_sel]
    if mes_sel == "Todos":
        vista = base
        foco = mes_actual
    else:
        vista = base[base["MES"] == mes_sel]
        foco = mes_sel
    planeados = len(vista)
    realizados = int((vista["ESTATUS"] == "Realizado").sum()) if planeados else 0
    pendientes = planeados - realizados
    porcentaje = round(100.0 * realizados / planeados, 1) if planeados else 0.0
    if mes_sel == "Todos":
        texto_corte = "Histórico completo"
    elif mes_sel == mes_actual:
        texto_corte = f"Mes en curso: {mes_sel}"
    else:
        texto_corte = f"Consulta: {mes_sel}"
    if zona_sel == "Todas":
        texto_corte += " · Norte y Sur"
    st.caption(texto_corte)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Planeados", planeados)
    m2.metric("Realizados", realizados)
    m3.metric("Pendientes", pendientes)
    m4.metric("Cumplimiento", f"{porcentaje}%")
    st.progress(realizados / planeados if planeados else 0.0)

    resumen = (
        base.groupby("MES", dropna=False)
        .agg(
            Planeados=("ID", "size"),
            Realizados=("ESTATUS", lambda serie: int((serie == "Realizado").sum())),
        )
        .reset_index()
    )
    if not resumen.empty:
        orden_mes = (
            base.groupby("MES")["FECHA"].min().sort_values(key=lambda serie: serie.map(lambda valor: valor or date.max))
        )
        resumen["MES"] = pd.Categorical(resumen["MES"], categories=list(orden_mes.index), ordered=True)
        resumen = resumen.sort_values("MES")
        resumen["Pendientes"] = resumen["Planeados"] - resumen["Realizados"]
        resumen["Detalle"] = [
            "Mes en curso" if str(mes) == mes_actual else ("En consulta" if str(mes) == foco and foco != mes_actual else "")
            for mes in resumen["MES"]
        ]
        colores_realizados = ["#f59e0b" if str(mes) == foco else "#64748b" for mes in resumen["MES"]]
        colores_pendientes = ["#fde68a" if str(mes) == foco else "#334155" for mes in resumen["MES"]]
        fig = go.Figure(
            data=[
                go.Bar(
                    name="Realizados",
                    x=resumen["MES"].astype(str),
                    y=resumen["Realizados"],
                    marker_color=colores_realizados,
                ),
                go.Bar(
                    name="Pendientes",
                    x=resumen["MES"].astype(str),
                    y=resumen["Pendientes"],
                    marker_color=colores_pendientes,
                ),
            ]
        )
        if not foco:
            titulo_grafica = "Histórico"
        elif mes_sel == "Todos":
            titulo_grafica = f"Histórico · {foco} resaltado"
        elif foco == mes_actual:
            titulo_grafica = f"Mes en curso: {foco}"
        else:
            titulo_grafica = f"Consulta: {foco}"
        fig.update_layout(
            barmode="stack",
            height=320,
            margin=dict(l=16, r=16, t=36, b=16),
            legend=dict(orientation="h"),
            yaxis_title="Actividades",
            title=titulo_grafica,
        )
        st.plotly_chart(fig, use_container_width=True)
        tabla_resumen = resumen.rename(columns={"MES": "Mes"})[
            ["Mes", "Planeados", "Realizados", "Pendientes", "Detalle"]
        ]

        def _resaltar_mes_programa(fila: pd.Series) -> list[str]:
            if str(fila.get("Mes")) == foco:
                return ["background-color: #78350f; color: #fffbeb; font-weight: 700"] * len(fila)
            return [""] * len(fila)

        st.dataframe(
            tabla_resumen.style.apply(_resaltar_mes_programa, axis=1),
            use_container_width=True,
            hide_index=True,
        )

    tabla = vista.rename(
        columns={
            "ZONA": "Zona",
            "UBICACION": "Ubicación",
            "TIPO DE MANTENIMIENTO": "Tipo de mantenimiento",
            "MES": "Mes",
            "ESTATUS": "Estatus",
            "FECHA REALIZADA": "Fecha realizada",
        }
    )
    st.dataframe(
        tabla[["Zona", "Ubicación", "Tipo de mantenimiento", "Mes", "Estatus", "Fecha realizada"]],
        use_container_width=True,
        hide_index=True,
    )


def render_panel_administrativo() -> None:
    modulo_id = st.session_state.get("admin_modulo")
    if modulo_id == "combustible":
        modulo_id = "vehiculos"
        st.session_state.admin_modulo = "vehiculos"
    render_barra_herramientas_admin()
    if not modulo_id:
        render_explorador_admin()
        return

    modulo = next((m for m in MODULOS_ADMIN if m["id"] == modulo_id), None)
    if not modulo:
        st.session_state.admin_modulo = None
        render_explorador_admin()
        return

    render_encabezado(
        "🗂️ Panel Administrativo",
        "Centro de control — Soledad · Base Álamo",
    )
    render_barra_modulo_admin(modulo["titulo"])
    if modulo_id == "registro_bitacora":
        render_captura_bitacora_admin()
    elif modulo_id == "dashboard":
        render_dashboard_ejecutivo()
    elif modulo_id == "insumos_refacciones":
        render_control_insumos_refacciones()
    elif modulo_id == "ot_diario":
        render_historial_diario_ot()
    elif modulo_id == "asignaciones":
        render_editor_asignaciones_diarias()
    elif modulo_id == "almacen":
        render_control_almacen()
    elif modulo_id == "ubicaciones":
        render_panel_ubicaciones(encabezado=False)
    elif modulo_id == "historial_base":
        render_historial_base_admin()
    elif modulo_id == "programados":
        render_control_programados()
    elif modulo_id == "programas":
        render_consulta_programas_mensuales()
    elif modulo_id == "programa_octubre":
        render_programa_octubre()
    elif modulo_id == "mantenimientos_programados":
        render_mantenimientos_programados()
    elif modulo_id == "reportes_ejecutivos":
        render_reportes_ejecutivos()
    elif modulo_id == "auditoria":
        render_auditoria_accesos()
    elif modulo_id == "gestionar_actividades":
        render_gestionar_actividades()
    elif modulo_id == "calendario_asistencia":
        render_calendario_asistencia()
    elif modulo_id == "cuadrillas":
        render_historial_cuadrillas(encabezado=False)
    elif modulo_id == "soporte":
        render_soporte_refacciones()
    elif modulo_id == "vehiculos":
        render_modulo_vehiculos()
    else:
        render_explorador_admin()


def exigir_contrasena() -> None:
    """Bloquea la aplicación hasta que la clave coincida con APP_PASSWORD."""
    if st.session_state.get("autenticado"):
        return

    st.markdown(
        """
        <style>
            section[data-testid="stSidebar"],
            [data-testid="stSidebarCollapsedControl"] { display: none; }
        </style>
        <div class="panel-header">
            <h1>Soledad</h1>
            <p>Ingresa la contraseña para continuar.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _, centro, _ = st.columns([1, 1.15, 1])
    with centro:
        with st.form("acceso_app"):
            clave = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Entrar")
        if entrar:
            try:
                esperada = str(st.secrets["APP_PASSWORD"])
            except Exception:
                st.error(
                    "Falta el secreto APP_PASSWORD. "
                    "Configúralo en Streamlit Cloud o en .streamlit/secrets.toml."
                )
                st.stop()
            if hmac.compare_digest(clave, esperada):
                st.session_state.autenticado = True
                st.rerun()
            st.error("Contraseña incorrecta.")
    st.stop()


# ── Arranque ───────────────────────────────────────────────────────────────────
exigir_contrasena()
persistencia.restaurar_al_arrancar()
inicializar_estado()
rol = render_selector_rol()

if rol in CLAVES_POR_ROL and not rol_esta_desbloqueado(rol):
    render_acceso_rol(rol)
elif rol is None:
    render_bienvenida()
elif rol == ROL_OPERATIVO:
    render_panel_operativo()
elif rol == ROL_ADMIN:
    render_panel_administrativo()
elif rol == MODULO_UBICACIONES:
    render_panel_ubicaciones()
elif rol == MODULO_CUADRILLAS:
    render_historial_cuadrillas()
else:
    render_bienvenida()

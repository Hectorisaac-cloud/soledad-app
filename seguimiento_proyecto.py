"""Seguimiento de residencia a partir del cronograma y de los registros ya cargados.

Solo copia valores que están en la bitácora, en las fichas de ubicación o en los
programados. Si un dato no existe, la celda queda vacía: no se estiman cifras
ni se rellenan ceros.
"""

from __future__ import annotations

import io
import os
import re
import unicodedata
from datetime import date, datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

RAIZ = os.path.dirname(os.path.abspath(__file__))
RUTA_CRONOGRAMA = os.path.join(RAIZ, "data", "residencia", "CRONOGRAMA DE ACTIVIDADES.xlsx")
RUTA_BITACORA = os.path.join(RAIZ, "database", "bitacora_operativa_soledad.xlsx")
RUTA_UBICACIONES = os.path.join(RAIZ, "ubicaciones_equipos.json")
RUTA_PROGRAMADOS = os.path.join(RAIZ, "programados_ops.json")
RUTA_ALMACEN = os.path.join(RAIZ, "almacen_insumos.json")
FILA_ENCABEZADO = 4
FILA_DATOS = 5

HOJAS = (
    {
        "id": "01_diagnostico_registros",
        "etiqueta": "01 · Diagnóstico",
        "nombre": "Diagnóstico de registros",
    },
    {
        "id": "02_flota_ubicaciones",
        "etiqueta": "02 · Flota",
        "nombre": "Flota y ubicaciones",
    },
    {
        "id": "03_indicadores_kpi",
        "etiqueta": "03 · Indicadores",
        "nombre": "Indicadores KPI",
    },
    {
        "id": "04_consumo_insumos",
        "etiqueta": "04 · Consumos",
        "nombre": "Consumo de insumos",
    },
    {
        "id": "05_caso_critico",
        "etiqueta": "05 · Caso crítico",
        "nombre": "Caso crítico",
    },
    {
        "id": "06_propuesta_plan_preventivo",
        "etiqueta": "06 · Plan preventivo",
        "nombre": "Propuesta de plan preventivo",
    },
    {
        "id": "07_cierre_evaluacion",
        "etiqueta": "07 · Cierre",
        "nombre": "Cierre y evaluación",
    },
)

# Cada insumo sale de una columna de la bitácora. `grupo` junta columnas que
# guardan el mismo consumo; si en la fila traen la misma cantidad, se emite una sola.
INSUMOS = (
    {"columna": "ACEITE SAE 50 PJP4", "grupo": "aceite", "nombre": "Aceite SAE 50", "unidad": "L", "proveedor": "PJP4"},
    {"columna": "ACEITE (LTS)", "grupo": "aceite", "nombre": "Aceite SAE 50", "unidad": "L", "proveedor": ""},
    {"columna": "GRASA OPS", "grupo": "grasa", "nombre": "Grasa", "unidad": "Kg", "proveedor": "OPS"},
    {"columna": "GRASA (KG)", "grupo": "grasa", "nombre": "Grasa", "unidad": "Kg", "proveedor": ""},
    {"columna": "REFRIGERANTE OPS", "grupo": "refrigerante", "nombre": "Refrigerante", "unidad": "L", "proveedor": "OPS"},
    {"columna": "REFRIGERANTE (LTS)", "grupo": "refrigerante", "nombre": "Refrigerante", "unidad": "L", "proveedor": ""},
    {"columna": "BANDAS C144 PJP4", "grupo": "banda", "nombre": "Banda C-144", "unidad": "Pzas", "proveedor": "PJP4"},
    {"columna": "BANDAS CANTIDAD", "grupo": "banda", "nombre": "Bandas", "unidad": "Pzas", "proveedor": ""},
    {"columna": "FILTRO DE ACEITE PJP4", "grupo": "filtro", "nombre": "Filtro de aceite", "unidad": "Pzas", "proveedor": "PJP4"},
    {"columna": "FILTROS (PZAS)", "grupo": "filtro", "nombre": "Filtro de aceite", "unidad": "Pzas", "proveedor": ""},
    {"columna": "BUJIAS PJP4", "grupo": "bujia", "nombre": "Bujías", "unidad": "Pzas", "proveedor": "PJP4"},
    {"columna": "BUJÍAS (PZAS)", "grupo": "bujia", "nombre": "Bujías", "unidad": "Pzas", "proveedor": ""},
    {"columna": "BOMBA DE AGUA REPARADA PJP4", "grupo": "", "nombre": "Bomba de agua reparada", "unidad": "Pzas", "proveedor": "PJP4"},
    {"columna": "JUEGO DE CABLES PARA BUJIAS", "grupo": "", "nombre": "Juego de cables para bujías", "unidad": "", "proveedor": ""},
    {"columna": "ALTERNADOR CANTIDAD", "grupo": "", "nombre": "Alternador", "unidad": "Pzas", "proveedor": ""},
    {"columna": "OTROS INSUMOS", "grupo": "", "nombre": "Otros insumos", "unidad": "", "proveedor": ""},
    {"columna": "PIEZAS ROTATIVAS", "grupo": "", "nombre": "Piezas rotativas", "unidad": "", "proveedor": ""},
    {"columna": 'POLEA 6"Ø', "grupo": "", "nombre": 'Polea 6"Ø', "unidad": "", "proveedor": ""},
    {"columna": "2 MANGUERAS DE 1 PULGADA", "grupo": "", "nombre": "2 mangueras de 1 pulgada", "unidad": "", "proveedor": ""},
    {"columna": "4 MANGUERAS RECUPERADAS", "grupo": "", "nombre": "4 mangueras recuperadas", "unidad": "", "proveedor": ""},
    {"columna": "1 ALTRONIC RECUPERADO", "grupo": "", "nombre": "Altronic recuperado", "unidad": "", "proveedor": ""},
    {"columna": "RADIADOR", "grupo": "", "nombre": "Radiador", "unidad": "", "proveedor": ""},
)

RAZON_VACIA = {
    "01_diagnostico_registros": {
        "ID_REGISTRO": "La bitácora no trae ID_REGISTRO en los registros leídos.",
        "FECHA": "Ningún registro trae una fecha reconocible.",
        "MES_AÑO": "No hay fechas de las que se pueda tomar el mes.",
        "ESPECIALIDAD": "Ninguna cuadrilla registrada trae en el nombre la especialidad (mecánico, instrumentos, eléctricos, soldadura o montaje).",
        "ID_UBICACION": "Los registros no traen ubicación, pozo ni equipo/ubicación.",
        "CUADRILLA": "Los registros no traen cuadrilla.",
        "TIPO_REGISTRO": "Los registros no traen tipo de actividad.",
        "NATURALEZA_ACTIVIDAD": "La columna PROGRAMADO de la bitácora no trae valores. No se infiere «programado» cuando la celda está vacía.",
        "ESTATUS": "La columna EVALUACIÓN DEL ESTATUS no trae valores. No se usó el estatus de revisión ni se completó el resto.",
    },
    "02_flota_ubicaciones": {
        "ID_UBICACION": "Las fichas de ubicación no traen identificador.",
        "NOMBRE_UBICACION": "Las fichas de ubicación no traen nombre.",
        "ZONA_SECTOR": "Las fichas no traen sector ni zona.",
        "TIPO_UNIDAD": "Ninguna ficha tiene motor MCI ni unidad de bombeo capturados. No se asignó un tipo por el nombre del pozo.",
        "ID_EQUIPO": "Las fichas guardan el nombre del equipo, no un identificador de equipo distinto.",
        "FAMILIA_MCI": "Ningún motor capturado declara familia Arrow, Weichai, Vortec o Cummins.",
        "ACTIVA_EN_PERIODO": "Ninguna ubicación de la flota tiene registros en la bitácora. No se escribió «No» ni cero.",
        "MODELO_MCI": "El catálogo no guarda un modelo aparte de la familia. Solo se copia V6 o V8 cuando el nombre del motor lo dice.",
        "REGIMEN_OPERACION": "No hay un régimen de operación capturado. El estatus «Operativo» viene por defecto al dar de alta la ficha y no se usó.",
        "TOTAL_INTERVENCIONES_PERIODO": "Ninguna ubicación tiene intervenciones en la bitácora. Las que no tienen se dejan vacías, sin cero.",
    },
    "03_indicadores_kpi": {
        "PERIODO": "No hay periodos con familia de motor identificada ni programados atribuibles a una familia.",
        "FAMILIA_MCI": "Ni la bitácora ni los programados declaran una familia de MCI en un registro utilizable.",
        "MODELO_MCI": "No hay un modelo V6 o V8 común y explícito dentro de cada periodo y familia. No se eligió el modelo más frecuente.",
        "UNIDADES_EVALUADAS": "No hay ubicaciones distintas con registro de bitácora para esas familias. No se puso cero.",
        "MANT_PROGRAMADOS_TOTAL": "Los programados cargados no traen familia de MCI y mes a la vez, o no hay líneas reales (las marcadas como demo no se usan).",
        "MANT_REALIZADOS": "No hay actividades cuyo tipo diga «mantenimiento» para esa familia y periodo. No se contaron inspecciones ni restablecimientos como mantenimiento, ni se puso cero.",
        "CUMPLIMIENTO_PCT": "El plan OPS y los mantenimientos de bitácora no son el mismo conjunto, así que el porcentaje no se puede verificar. Queda vacío.",
        "PAROS_NO_PROGRAMADOS": "No existe un campo de paro. «PROGRAMADO = No» es actividad no programada, no un paro, y no se usó.",
        "FRECUENCIA_PAROS_POR_UNIDAD": "Depende del conteo de paros, que no está registrado.",
        "SUBSISTEMA_FALLA_PRINCIPAL": "La bitácora no clasifica el subsistema de la falla.",
        "HORAS_FUERA_SERVICIO": "HORAS DE EJECUCION mide el tiempo de la actividad, no las horas fuera de servicio. No se copió.",
        "COSTO_INSUMOS_REFACCIONES": "La bitácora y el almacén no registran el costo de insumos o refacciones de estos grupos.",
        "MONEDA": "No hay un costo que traiga moneda. No se asumió MXN.",
    },
    "04_consumo_insumos": {
        "ID_CONSUMO": "No hay consumos con cantidad mayor a cero en la bitácora ni movimientos en el almacén.",
        "FECHA": "Los consumos encontrados no traen fecha reconocible.",
        "PERIODO": "Los consumos encontrados no traen fecha de la que se pueda tomar el mes.",
        "ID_UBICACION": "Los consumos encontrados no traen ubicación.",
        "FAMILIA_MCI": "Esos registros no declaran familia de MCI y la ubicación no tiene motor MCI en la ficha.",
        "MODELO_MCI": "No hay modelo V6 o V8 explícito en el equipo del registro ni en la ficha.",
        "ACTIVIDAD_CATALOGO": "Esos registros no traen tipo de actividad ni actividad corta.",
        "INSUMO": "No hay consumos con cantidad mayor a cero.",
        "CANTIDAD": "No hay cantidades mayores a cero. Los ceros de la bitácora no se copiaron.",
        "UNIDAD": "La columna de origen y el catálogo de almacén no indican la unidad de esos consumos.",
        "PROVEEDOR_SUMINISTRO": "Esas columnas no dicen si el insumo lo dio PJP4 u OPS, y el campo de origen viene vacío.",
        "COSTO_UNITARIO": "Ni la bitácora ni los movimientos de almacén guardan costo unitario.",
        "COSTO_TOTAL": "Ni la bitácora ni los movimientos de almacén guardan costo total.",
        "EXISTENCIA_REFERENCIA": "El almacén no tiene un saldo de existencia por insumo. No se calculó ni se puso cero.",
    },
    "05_caso_critico": {
        "ID_EVALUACION": "No hay evaluaciones de criticidad guardadas en la aplicación.",
        "FAMILIA_O_MODELO": "Elegir el caso crítico sería una calificación, y no está registrada.",
        "PERIODO_EVALUADO": "No hay un periodo de evaluación de criticidad capturado.",
        "CRITERIO_EVALUACION": "Los criterios de criticidad no están definidos en los registros.",
        "VALOR_OBTENIDO": "No hay un valor de evaluación ya calculado y guardado.",
        "UNIDAD_MEDIDA": "No hay una evaluación de la que tomar la unidad.",
        "PUNTAJE": "El puntaje no está capturado. No se asignó uno.",
        "PONDERACION": "La ponderación no está capturada.",
        "PUNTAJE_PONDERADO": "No se calculó un puntaje ponderado.",
        "RANKING_CRITICIDAD": "No hay un ranking registrado. No se ordenó a las familias por un criterio propio.",
        "ES_CASO_CRITICO": "Qué familia es el caso crítico no está declarado en los datos.",
        "JUSTIFICACION_TECNICA": "No hay una justificación técnica redactada en el sistema.",
        "FUENTE_DATO": "No hay una ficha de evaluación que cite su fuente.",
    },
    "06_propuesta_plan_preventivo": {
        "ID_MEJORA": "No hay propuestas de mejora guardadas.",
        "FAMILIA_O_EQUIPO": "La propuesta no está registrada, así que no se asignó una familia.",
        "PROBLEMA_DETECTADO": "Los problemas a corregir no están redactados en el sistema.",
        "EVIDENCIA_PROBLEMA": "No hay una evidencia ligada a una propuesta.",
        "ACCION_PROPUESTA": "Las acciones del plan preventivo se capturan en esta hoja; la app no las tiene.",
        "DESCRIPCION_MEJORA": "No hay una descripción de mejora registrada.",
        "INSUMO_O_COMPONENTE_CLAVE": "No hay un componente clave definido para una propuesta.",
        "IMPACTO_ESPERADO": "El impacto esperado no está registrado. No se estimó.",
        "INDICADOR_VERIFICACION": "No hay un indicador de verificación de la propuesta.",
        "META_OBJETIVO": "No hay una meta capturada.",
        "RESPONSABLE": "No hay un responsable asignado a una propuesta.",
        "FRECUENCIA_REVISION": "No hay una frecuencia de revisión registrada.",
        "ESTATUS_PROPUESTA": "No hay estatus de propuesta. No se puso «pendiente» por defecto.",
    },
    "07_cierre_evaluacion": {
        "ID_RESUMEN": "No hay un cierre de indicadores guardado.",
        "INDICADOR_RESUMEN": "Los indicadores de cierre no están definidos en los registros.",
        "VALOR_INICIAL_O_LINEA_BASE": "No hay una línea base capturada. Sin ella no se compara el cierre.",
        "VALOR_ACTUAL_CIERRE": "No hay un valor de cierre guardado. No se reutilizó un conteo de otra hoja como si fuera el cierre.",
        "UNIDAD_MEDIDA": "No hay un indicador de cierre del que tomar la unidad.",
        "VARIACION": "Sin línea base y sin valor de cierre no hay variación verificable. No se calculó.",
        "ESTADO_ACTIVIDAD": "El estado de cierre de cada actividad no está capturado.",
        "FUENTE_VERIFICACION": "No hay una fuente de verificación del cierre registrada.",
        "NOTAS_CIERRE": "No hay notas de cierre en el sistema.",
    },
}


def _sin_acentos(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return "".join(caracter for caracter in texto if not unicodedata.combining(caracter))


def _plano(valor) -> str:
    return re.sub(r"[^A-Z0-9]", "", _sin_acentos(valor).upper())


def _texto(valor) -> str:
    if valor is None or isinstance(valor, bool):
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    if isinstance(valor, (datetime, date)):
        return ""
    texto = str(valor).strip()
    if texto.casefold() in {"", "nan", "none", "nat", "<na>"}:
        return ""
    return texto


def _fecha(valor) -> date | None:
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, float) and pd.isna(valor):
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
            convertido = pd.to_datetime(float(valor), unit="D", origin="1899-12-30", errors="coerce")
            if pd.isna(convertido):
                return None
            return convertido.date()
    texto = str(valor).strip()
    if not texto or texto.casefold() in {"nan", "none", "nat", "<na>"}:
        return None
    if re.match(r"^\d{4}-\d{2}-\d{2}", texto):
        convertido = pd.to_datetime(texto, errors="coerce")
    else:
        convertido = pd.to_datetime(texto, errors="coerce", dayfirst=True)
    if pd.isna(convertido):
        return None
    return convertido.date()


def _fecha_visible(valor) -> str:
    dia = _fecha(valor)
    if dia is not None:
        return dia.strftime("%d/%m/%Y")
    return _texto(valor)


def _mes(valor) -> str:
    dia = _fecha(valor)
    if dia is None:
        return ""
    return f"{dia.year:04d}-{dia.month:02d}"


def _cantidad(valor):
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        numero = float(valor)
    else:
        texto = str(valor).strip().replace(",", "")
        if not texto or texto.casefold() in {"nan", "none", "nat", "<na>"} or texto.startswith("="):
            return None
        try:
            numero = float(texto)
        except ValueError:
            return None
    if numero <= 0:
        return None
    if numero.is_integer():
        return int(numero)
    return numero


def _familia_mci(valor) -> str:
    plano = _sin_acentos(valor).upper()
    if "WEICHAI" in plano:
        return "WEICHAI"
    if "CUMMIN" in plano:
        return "CUMMINS"
    if "VORTEC" in plano:
        return "VORTEC"
    if "ARROW" in plano or "AROW" in plano:
        return "ARROW"
    return ""


def _modelo_mci(valor) -> str:
    plano = _sin_acentos(valor).upper()
    if "VORTEC" not in plano:
        return ""
    if re.search(r"\bV\s*6\b", plano):
        return "V6"
    if re.search(r"\bV\s*8\b", plano):
        return "V8"
    return ""


def _especialidad(cuadrilla: str) -> str:
    plano = _sin_acentos(cuadrilla).upper()
    if "INSTRUMENT" in plano:
        return "Instrumentación"
    if "ELECTRIC" in plano:
        return "Eléctrico"
    if "SOLDAD" in plano:
        return "Soldadura"
    if "MONTAJ" in plano:
        return "Montaje"
    if "MECANIC" in plano:
        return "Mecánico"
    return ""


def _naturaleza(valor) -> str:
    texto = _texto(valor)
    if not texto:
        return ""
    if texto.casefold() == "no":
        return "No programado"
    return texto


def _tipo_unidad(ficha: dict) -> str:
    partes = []
    if _texto(ficha.get("equipo_mci")):
        partes.append("MCI")
    ub = _sin_acentos(ficha.get("equipo_ub")).upper()
    if ub:
        partes.append("HLU" if "HLU" in ub or "HIDRAUL" in ub else "UBM")
    return " / ".join(partes)


def _leer_bytes(ruta: str) -> bytes:
    try:
        with open(ruta, "rb") as archivo:
            return archivo.read()
    except PermissionError:
        import control_costos_mci

        return control_costos_mci._leer_bytes_compartido(ruta)


def _leer_json(ruta: str, vacio):
    if not os.path.isfile(ruta):
        return vacio
    import json

    with open(ruta, encoding="utf-8") as archivo:
        datos = json.load(archivo)
    return datos


def leer_plantilla(ruta: str | None = None) -> list[dict]:
    """Lee títulos (fila 2) y encabezados reales (fila 4, desde la columna B)."""
    ruta = ruta or RUTA_CRONOGRAMA
    libro = load_workbook(io.BytesIO(_leer_bytes(ruta)), data_only=True, read_only=True)
    try:
        hojas = []
        for nombre in libro.sheetnames:
            hoja = libro[nombre]
            titulo = hoja.cell(2, 2).value
            columnas = []
            for columna in range(2, (hoja.max_column or 1) + 1):
                valor = hoja.cell(FILA_ENCABEZADO, columna).value
                if valor is None or not str(valor).strip():
                    break
                columnas.append(str(valor).strip())
            hojas.append(
                {
                    "id": nombre,
                    "titulo": _texto(titulo) or str(titulo or "").strip(),
                    "columnas": columnas,
                }
            )
        return hojas
    finally:
        libro.close()


_CLASES_INFRA = {"GDO", "CBN", "OLD", "OLG"}
_IGNORAR_UBICACION = {
    "ninguno",
    "sin equipo",
    "no especificado en catalogo",
}


def _alias_infra(codigo) -> str:
    """El catálogo guarda CBN-TDSGF-IQ-SIN-13 y la bitácora a veces trae solo TDSGF-IQ-SIN-13."""
    texto = str(codigo or "").strip().upper()
    clase, separador, resto = texto.partition("-")
    if separador and clase in _CLASES_INFRA and resto:
        return resto
    return ""


def _indexar_fichas(fichas: dict) -> dict[str, dict]:
    indice: dict[str, dict] = {}
    for ficha in fichas.values():
        textos = [ficha.get("id"), ficha.get("nombre"), _alias_infra(ficha.get("id"))]
        alias = ficha.get("aliases") or []
        if isinstance(alias, list):
            textos.extend(alias)
        for texto in textos:
            clave = _plano(texto)
            if clave:
                indice.setdefault(clave, ficha)
    return indice


def _candidatos_ubicacion(valor) -> list[str]:
    bruto = _texto(valor)
    if not bruto:
        return []
    salida: list[str] = []

    def agregar(texto) -> None:
        limpio = _texto(texto)
        if not limpio or _sin_acentos(limpio).casefold() in _IGNORAR_UBICACION:
            return
        if limpio not in salida:
            salida.append(limpio)

    agregar(bruto)
    for segmento in re.split(r"\s*[·•]\s*", bruto):
        agregar(segmento)
        for parentesis in re.findall(r"\(([^)]+)\)", segmento):
            agregar(parentesis)
        plano = _sin_acentos(segmento).upper()
        for pozo in re.findall(r"\b(?:PSN|PPB|PGA|PS|PA|PC|PB|PG)\s*-?\s*\d+[A-Z]?\b", plano):
            agregar(pozo)
    return salida


def _ficha_de(indice: dict[str, dict], *textos) -> dict | None:
    for texto in textos:
        for candidato in _candidatos_ubicacion(texto):
            clave = _plano(candidato)
            if clave and clave in indice:
                return indice[clave]
    return None


def _ficha_desde_fila(indice: dict[str, dict], fila) -> dict | None:
    columnas = getattr(fila, "index", [])
    textos = [fila.get(columna) for columna in ("UBICACIÓN", "POZO", "EQUIPO/UBICACIÓN") if columna in columnas]
    return _ficha_de(indice, *textos)


def _ubicacion_fila(fila) -> str:
    for columna in ("UBICACIÓN", "POZO", "EQUIPO/UBICACIÓN", "SECTOR"):
        if columna not in fila.index:
            continue
        texto = _texto(fila.get(columna))
        if texto:
            return texto
    return ""


def _equipo_mci_fila(fila, ficha: dict | None) -> str:
    """El equipo escrito en el registro manda. La ficha solo completa si el registro no trae equipo."""
    columnas = getattr(fila, "index", [])
    equipo = _texto(fila.get("EQUIPO")) if "EQUIPO" in columnas else ""
    if equipo:
        return equipo if _familia_mci(equipo) else ""
    for columna in ("UBICACIÓN", "EQUIPO/UBICACIÓN"):
        if columna in columnas and _familia_mci(fila.get(columna)):
            return _texto(fila.get(columna))
    if ficha:
        return _texto(ficha.get("equipo_mci"))
    return ""


def _tabla(filas: list[dict], columnas: list[str]) -> pd.DataFrame:
    if not filas:
        return pd.DataFrame(columns=columnas)
    tabla = pd.DataFrame(filas)
    for columna in columnas:
        if columna not in tabla.columns:
            tabla[columna] = ""
    tabla = tabla[columnas]
    return tabla.fillna("")


def _diagnostico(hoja_id: str, columnas: list[str], tabla: pd.DataFrame, razones_llenas: dict[str, str]) -> pd.DataFrame:
    vacias = RAZON_VACIA.get(hoja_id, {})
    filas = []
    for columna in columnas:
        if tabla.empty or columna not in tabla.columns:
            celdas = 0
        else:
            serie = tabla[columna].map(_texto) if tabla[columna].dtype == object else tabla[columna]
            if tabla[columna].dtype == object:
                celdas = int(serie.ne("").sum())
            else:
                celdas = int(tabla[columna].map(lambda valor: _texto(valor) != "" and not (isinstance(valor, float) and pd.isna(valor))).sum())
            # Las cantidades numéricas no pasan por _texto si el dtype es numérico.
            if celdas == 0:
                celdas = int(
                    tabla[columna].map(
                        lambda valor: valor not in ("", None) and not (isinstance(valor, float) and pd.isna(valor))
                    ).sum()
                )
        if celdas:
            estado = "Con datos de la app"
            detalle = razones_llenas.get(columna, "Valor tomado de un registro existente.")
        else:
            estado = "En blanco"
            detalle = vacias.get(columna, "No hay un dato directo en la bitácora, la flota, los programas ni los consumos.")
        filas.append(
            {
                "Columna": columna,
                "Estado": estado,
                "Celdas con valor": celdas,
                "Detalle": detalle,
            }
        )
    return pd.DataFrame(filas)


def _cargar_bitacora() -> pd.DataFrame:
    if not os.path.isfile(RUTA_BITACORA):
        return pd.DataFrame()
    return pd.read_excel(io.BytesIO(_leer_bytes(RUTA_BITACORA)), sheet_name="Bitacora")


def _hoja_vacia(columnas: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columnas)


def _construir_diagnostico(bitacora: pd.DataFrame, columnas: list[str]) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    filas = []
    sin_fecha = 0
    for _, fila in bitacora.iterrows():
        if "FECHA" in bitacora.columns and _fecha(fila.get("FECHA")) is None:
            sin_fecha += 1
        filas.append(
            {
                "ID_REGISTRO": _texto(fila.get("ID_REGISTRO")),
                "FECHA": _fecha_visible(fila.get("FECHA")),
                "MES_AÑO": _mes(fila.get("FECHA")),
                "ESPECIALIDAD": _especialidad(_texto(fila.get("CUADRILLA"))),
                "ID_UBICACION": _ubicacion_fila(fila),
                "CUADRILLA": _texto(fila.get("CUADRILLA")),
                "TIPO_REGISTRO": _texto(fila.get("TIPO DE ACTIVIDAD")),
                "NATURALEZA_ACTIVIDAD": _naturaleza(fila.get("PROGRAMADO")),
                "ESTATUS": _texto(fila.get("EVALUACIÓN DEL ESTATUS")),
            }
        )
    razones = {
        "ID_REGISTRO": "Copiado de ID_REGISTRO en cada fila de la bitácora.",
        "FECHA": "Copiada de FECHA. Si ya es fecha se muestra día/mes/año; si el texto no es una fecha, se deja el texto original.",
        "MES_AÑO": "Año y mes tomados de la misma FECHA, solo cuando la fecha se reconoce.",
        "ESPECIALIDAD": "Especialidad que ya viene en el nombre de la cuadrilla (mecánico, instrumentos, eléctricos, soldadura o montaje).",
        "ID_UBICACION": "Ubicación del registro. Si viene vacía, se usa el pozo y, si también falta, equipo/ubicación.",
        "CUADRILLA": "Copiada de CUADRILLA.",
        "TIPO_REGISTRO": "Copiado de TIPO DE ACTIVIDAD.",
        "NATURALEZA_ACTIVIDAD": "Solo donde PROGRAMADO trae un valor. «No» se escribe como «No programado». La celda vacía no se interpretó como programado.",
        "ESTATUS": "Copiado de EVALUACIÓN DEL ESTATUS solo donde esa celda tiene texto.",
    }
    notas = [
        f"Una fila por cada registro de la bitácora ({len(filas)}). No se eliminaron filas ni columnas de la plantilla.",
    ]
    if sin_fecha:
        notas.append(
            f"{sin_fecha} registros no tienen una fecha reconocible: MES_AÑO queda vacío y FECHA conserva el texto solo cuando hay texto."
        )
    return _tabla(filas, columnas), razones, notas


def _construir_flota(
    fichas: dict,
    zonas: list,
    bitacora: pd.DataFrame,
    indice: dict[str, dict],
    columnas: list[str],
) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    nombres_zona = {str(zona.get("id")): str(zona.get("nombre") or "") for zona in zonas if isinstance(zona, dict)}
    conteo: dict[str, int] = {}
    sin_ficha = 0
    for _, fila in bitacora.iterrows():
        ficha = _ficha_desde_fila(indice, fila)
        if not ficha:
            sin_ficha += 1
            continue
        llave = str(ficha.get("id") or "")
        conteo[llave] = conteo.get(llave, 0) + 1
    filas = []
    for ficha in fichas.values():
        sector = _texto(ficha.get("sector"))
        if not sector:
            sector = nombres_zona.get(str(ficha.get("zona_id")), "")
        intervenciones = conteo.get(str(ficha.get("id") or ""), 0)
        motor = _texto(ficha.get("equipo_mci"))
        filas.append(
            {
                "ID_UBICACION": _texto(ficha.get("id")),
                "NOMBRE_UBICACION": _texto(ficha.get("nombre")),
                "ZONA_SECTOR": sector,
                "TIPO_UNIDAD": _tipo_unidad(ficha),
                "ID_EQUIPO": "",
                "FAMILIA_MCI": _familia_mci(motor),
                "ACTIVA_EN_PERIODO": "Sí" if intervenciones else "",
                "MODELO_MCI": _modelo_mci(motor),
                "REGIMEN_OPERACION": "",
                "TOTAL_INTERVENCIONES_PERIODO": intervenciones if intervenciones else "",
            }
        )
    filas.sort(key=lambda fila: _plano(fila.get("ID_UBICACION")))
    razones = {
        "ID_UBICACION": "Identificador de la ficha en ubicaciones y equipos.",
        "NOMBRE_UBICACION": "Nombre de la ficha.",
        "ZONA_SECTOR": "Sector guardado en la ficha. Si no tiene, el nombre de su zona.",
        "TIPO_UNIDAD": "MCI, UBM o HLU según el equipo que la ficha tiene capturado. Si no hay equipo, queda vacía.",
        "FAMILIA_MCI": "Familia leída del nombre del motor MCI (Arrow, Weichai, Vortec o Cummins).",
        "ACTIVA_EN_PERIODO": "«Sí» solo si esa ubicación tiene al menos un registro en la bitácora. Si no tiene, queda vacía.",
        "MODELO_MCI": "V6 o V8 solo cuando el nombre del motor Vortec lo dice. Arrow, Weichai y Cummins no traen modelo.",
        "TOTAL_INTERVENCIONES_PERIODO": "Registros de bitácora ligados a esa ficha. Si no hay, la celda queda vacía en lugar de cero.",
    }
    notas = [
        f"Una fila por cada ubicación del catálogo ({len(filas)}). El periodo es toda la bitácora cargada: la plantilla no fija un corte.",
        f"{sin_ficha} registros de bitácora no coincidieron con una ficha y no se sumaron a ninguna ubicación.",
        "El estatus «Operativo» está en todas las fichas desde el alta. No se usó como régimen de operación.",
    ]
    return _tabla(filas, columnas), razones, notas


def _construir_kpi(
    bitacora: pd.DataFrame,
    indice: dict[str, dict],
    programados: list[dict],
    columnas: list[str],
) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    grupos: dict[tuple[str, str], dict] = {}

    def grupo(periodo: str, familia: str) -> dict:
        llave = (periodo, familia)
        if llave not in grupos:
            grupos[llave] = {
                "modelos": set(),
                "unidades": set(),
                "realizados": 0,
                "programados": 0,
            }
        return grupos[llave]

    for _, fila in bitacora.iterrows():
        periodo = _mes(fila.get("FECHA"))
        ficha = _ficha_desde_fila(indice, fila)
        actividad_mayus = _sin_acentos(fila.get("TIPO DE ACTIVIDAD")).upper()
        motor = _equipo_mci_fila(fila, None)
        familia = _familia_mci(motor)
        if not familia and ficha and "MCI" in actividad_mayus:
            familia = _familia_mci(ficha.get("equipo_mci"))
            motor = _texto(ficha.get("equipo_mci"))
        if not periodo or not familia:
            continue
        actual = grupo(periodo, familia)
        modelo = _modelo_mci(motor)
        if modelo:
            actual["modelos"].add(modelo)
        ubicacion = _texto(ficha.get("id")) if ficha else _ubicacion_fila(fila)
        if ubicacion:
            actual["unidades"].add(ubicacion)
        actividad = _sin_acentos(fila.get("TIPO DE ACTIVIDAD")).casefold()
        if "mantenimiento" in actividad:
            actual["realizados"] += 1

    reales = []
    omitidos = []
    demo = 0
    for item in programados:
        if not isinstance(item, dict):
            continue
        if item.get("demo"):
            demo += 1
            continue
        reales.append(item)
        familia = _familia_mci(item.get("equipo"))
        periodo = _texto(item.get("mes"))
        if not familia or not re.fullmatch(r"\d{4}-\d{2}", periodo):
            omitidos.append(item)
            continue
        actual = grupo(periodo, familia)
        actual["programados"] += 1
        modelo = _modelo_mci(item.get("equipo"))
        if modelo:
            actual["modelos"].add(modelo)

    filas = []
    for (periodo, familia), actual in sorted(grupos.items()):
        modelo = ""
        if len(actual["modelos"]) == 1:
            modelo = next(iter(actual["modelos"]))
        filas.append(
            {
                "PERIODO": periodo,
                "FAMILIA_MCI": familia,
                "MODELO_MCI": modelo,
                "UNIDADES_EVALUADAS": len(actual["unidades"]) if actual["unidades"] else "",
                "MANT_PROGRAMADOS_TOTAL": actual["programados"] if actual["programados"] else "",
                "MANT_REALIZADOS": actual["realizados"] if actual["realizados"] else "",
                "CUMPLIMIENTO_PCT": "",
                "PAROS_NO_PROGRAMADOS": "",
                "FRECUENCIA_PAROS_POR_UNIDAD": "",
                "SUBSISTEMA_FALLA_PRINCIPAL": "",
                "HORAS_FUERA_SERVICIO": "",
                "COSTO_INSUMOS_REFACCIONES": "",
                "MONEDA": "",
            }
        )
    razones = {
        "PERIODO": "Mes de la fecha de bitácora, o mes del programado, en formato año-mes.",
        "FAMILIA_MCI": "Familia declarada en el equipo del registro. Si ese campo viene vacío y la actividad menciona MCI, se usa el motor de la ficha. En programados, la familia del campo equipo.",
        "MODELO_MCI": "V6 o V8 solo si todos los registros de ese periodo y familia dicen el mismo modelo.",
        "UNIDADES_EVALUADAS": "Ubicaciones distintas con al menos un registro de bitácora de esa familia en el periodo. No es un censo externo.",
        "MANT_PROGRAMADOS_TOTAL": "Líneas de programados_ops.json, sin las marcadas como demo, cuyo equipo indica esa familia y cuyo mes es el periodo.",
        "MANT_REALIZADOS": "Registros de bitácora de esa familia y periodo cuyo tipo de actividad contiene la palabra «mantenimiento».",
    }
    notas = [
        "Cada fila es un periodo con una familia de MCI identificada. Las columnas de paro, costo, frecuencia y cumplimiento quedan vacías porque no hay un dato que las respalde.",
        f"Programados reales considerados: {len(reales)}. Marcados como demo y excluidos: {demo}.",
    ]
    if omitidos:
        notas.append(
            f"{len(omitidos)} programados reales no entraron a la tabla porque el equipo no es una familia MCI o el mes no tiene formato año-mes."
        )
    return _tabla(filas, columnas), razones, notas


def _consumos_de_fila(fila, indice: dict[str, dict]) -> tuple[list[dict], int]:
    """Devuelve consumos con cantidad mayor a cero y cuántas columnas repetidas se omitieron."""
    por_grupo: dict[str, list[tuple[dict, object]]] = {}
    sueltos: list[tuple[dict, object]] = []
    for spec in INSUMOS:
        if spec["columna"] not in fila.index:
            continue
        cantidad = _cantidad(fila.get(spec["columna"]))
        if cantidad is None:
            continue
        if spec["grupo"]:
            por_grupo.setdefault(spec["grupo"], []).append((spec, cantidad))
        else:
            sueltos.append((spec, cantidad))

    elegidos: list[tuple[dict, object]] = []
    omitidos = 0
    for candidatos in por_grupo.values():
        cantidades = {valor for _, valor in candidatos}
        if len(cantidades) == 1:
            elegidos.append(candidatos[0])
            omitidos += len(candidatos) - 1
        else:
            elegidos.extend(candidatos)
    elegidos.extend(sueltos)

    ficha = _ficha_de(indice, _ubicacion_fila(fila), fila.get("POZO"), fila.get("UBICACIÓN"))
    motor = _equipo_mci_fila(fila, ficha)
    origen_bandas = _texto(fila.get("BANDAS PROPORCIONADO POR")) if "BANDAS PROPORCIONADO POR" in fila.index else ""
    descripcion_bandas = _texto(fila.get("BANDAS DESCRIPCIÓN")) if "BANDAS DESCRIPCIÓN" in fila.index else ""
    origen_alternador = _texto(fila.get("ALTERNADOR ESTADO/ORIGEN")) if "ALTERNADOR ESTADO/ORIGEN" in fila.index else ""
    descripcion_alternador = _texto(fila.get("ALTERNADOR DESCRIPCIÓN")) if "ALTERNADOR DESCRIPCIÓN" in fila.index else ""
    identificador = _texto(fila.get("ID_REGISTRO"))
    salida = []
    for spec, cantidad in elegidos:
        nombre = spec["nombre"]
        proveedor = spec["proveedor"]
        if spec["grupo"] == "banda":
            if descripcion_bandas:
                nombre = descripcion_bandas
            if origen_bandas:
                proveedor = origen_bandas
        if spec["columna"] == "ALTERNADOR CANTIDAD":
            if descripcion_alternador:
                nombre = descripcion_alternador
            if "PJP4" in origen_alternador.upper() or "OPS" in origen_alternador.upper():
                proveedor = origen_alternador
        salida.append(
            {
                "ID_CONSUMO": f"{identificador}|{spec['columna']}" if identificador else "",
                "FECHA": _fecha_visible(fila.get("FECHA")),
                "PERIODO": _mes(fila.get("FECHA")),
                "ID_UBICACION": _ubicacion_fila(fila),
                "FAMILIA_MCI": _familia_mci(motor),
                "MODELO_MCI": _modelo_mci(motor),
                "ACTIVIDAD_CATALOGO": _texto(fila.get("TIPO DE ACTIVIDAD")) or _texto(fila.get("ACTIVIDAD CORTA")),
                "INSUMO": nombre,
                "CANTIDAD": cantidad,
                "UNIDAD": spec["unidad"],
                "PROVEEDOR_SUMINISTRO": proveedor,
                "COSTO_UNITARIO": "",
                "COSTO_TOTAL": "",
                "EXISTENCIA_REFERENCIA": "",
            }
        )
    return salida, omitidos


def _construir_consumos(
    bitacora: pd.DataFrame,
    indice: dict[str, dict],
    movimientos: list,
    columnas: list[str],
) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    filas = []
    repetidos = 0
    for _, fila in bitacora.iterrows():
        nuevos, omitidos = _consumos_de_fila(fila, indice)
        filas.extend(nuevos)
        repetidos += omitidos
    sin_id = sum(1 for fila in filas if not fila["ID_CONSUMO"])
    razones = {
        "ID_CONSUMO": "ID_REGISTRO de la bitácora más la columna del insumo. Si el registro no trae ID, queda vacío.",
        "FECHA": "Fecha del mismo registro de bitácora.",
        "PERIODO": "Año y mes de esa fecha, cuando la fecha se reconoce.",
        "ID_UBICACION": "Ubicación del registro. Si falta, pozo o equipo/ubicación.",
        "FAMILIA_MCI": "Familia del equipo anotado en el registro o, si no viene, del motor MCI de la ficha de esa ubicación.",
        "MODELO_MCI": "V6 o V8 solo si el nombre del motor lo dice.",
        "ACTIVIDAD_CATALOGO": "Tipo de actividad del registro. Si falta, la actividad corta.",
        "INSUMO": "Nombre de la columna de consumo. En bandas y alternador se usa la descripción capturada cuando existe.",
        "CANTIDAD": "Cantidad registrada, solo si es mayor a cero.",
        "UNIDAD": "Unidad que ya viene en el encabezado de la columna (litros, kilos o piezas) o, en bandas, bomba y alternador, la unidad del catálogo de almacén.",
        "PROVEEDOR_SUMINISTRO": "PJP4 u OPS cuando el nombre de la columna o el campo «proporcionado por» lo dice. Si no lo dice, queda vacío.",
    }
    notas = [
        f"Solo filas con cantidad mayor a cero ({len(filas)}). Los ceros capturados en la bitácora no se copiaron.",
        f"Se omitieron {repetidos} columnas que repetían, en la misma fila, la misma cantidad de un insumo ya contado.",
        f"Movimientos de almacén leídos: {len(movimientos)}. No hay existencia ni costo en esos movimientos, así que esas columnas quedan vacías.",
    ]
    if sin_id:
        notas.append(
            f"{sin_id} consumos vienen de registros sin ID_REGISTRO. El identificador quedó vacío en lugar de inventar uno."
        )
    return _tabla(filas, columnas), razones, notas


def construir(ruta_cronograma: str | None = None) -> dict:
    plantilla = leer_plantilla(ruta_cronograma)
    por_id = {hoja["id"]: hoja for hoja in plantilla}
    bitacora = _cargar_bitacora()
    ubicaciones = _leer_json(RUTA_UBICACIONES, {"zonas": [], "fichas": {}})
    fichas = ubicaciones.get("fichas") if isinstance(ubicaciones, dict) else {}
    if not isinstance(fichas, dict):
        fichas = {}
    zonas = ubicaciones.get("zonas") if isinstance(ubicaciones, dict) else []
    if not isinstance(zonas, list):
        zonas = []
    indice = _indexar_fichas(fichas)
    programados = _leer_json(RUTA_PROGRAMADOS, [])
    if isinstance(programados, dict):
        plano = []
        for mes, filas in programados.items():
            if isinstance(filas, list):
                for fila in filas:
                    if isinstance(fila, dict):
                        item = dict(fila)
                        item.setdefault("mes", mes)
                        plano.append(item)
        programados = plano
    if not isinstance(programados, list):
        programados = []
    almacen = _leer_json(RUTA_ALMACEN, {"movimientos": []})
    movimientos = almacen.get("movimientos") if isinstance(almacen, dict) else []
    if not isinstance(movimientos, list):
        movimientos = []

    columnas = {hoja["id"]: hoja["columnas"] for hoja in plantilla}
    bloques = {
        "01_diagnostico_registros": _construir_diagnostico(bitacora, columnas.get("01_diagnostico_registros", [])),
        "02_flota_ubicaciones": _construir_flota(fichas, zonas, bitacora, indice, columnas.get("02_flota_ubicaciones", [])),
        "03_indicadores_kpi": _construir_kpi(bitacora, indice, programados, columnas.get("03_indicadores_kpi", [])),
        "04_consumo_insumos": _construir_consumos(bitacora, indice, movimientos, columnas.get("04_consumo_insumos", [])),
    }
    hojas = []
    for meta in HOJAS:
        hoja = por_id.get(meta["id"], {"titulo": "", "columnas": []})
        columnas_hoja = list(hoja.get("columnas") or [])
        if meta["id"] in bloques:
            tabla, razones, notas = bloques[meta["id"]]
        else:
            tabla, razones, notas = _hoja_vacia(columnas_hoja), {}, [
                "Esta actividad pide una evaluación, una propuesta o un cierre que la aplicación no tiene guardados. La estructura queda lista para captura manual.",
            ]
        diagnostico = _diagnostico(meta["id"], columnas_hoja, tabla, razones)
        hojas.append(
            {
                "id": meta["id"],
                "etiqueta": meta["etiqueta"],
                "nombre": meta["nombre"],
                "titulo": hoja.get("titulo") or "",
                "columnas": columnas_hoja,
                "tabla": tabla,
                "diagnostico": diagnostico,
                "notas": notas,
            }
        )
    demo = sum(1 for item in programados if isinstance(item, dict) and item.get("demo"))
    return {
        "hojas": hojas,
        "fuentes": {
            "bitacora": int(len(bitacora)),
            "ubicaciones": int(len(fichas)),
            "programados_reales": int(sum(1 for item in programados if isinstance(item, dict) and not item.get("demo"))),
            "programados_demo": int(demo),
            "movimientos_almacen": int(len(movimientos)),
        },
    }


def _valor_excel(valor):
    if valor is None or valor == "":
        return None
    if isinstance(valor, float) and pd.isna(valor):
        return None
    if isinstance(valor, str) and not valor.strip():
        return None
    return valor


def exportar_excel(resultado: dict, ruta_cronograma: str | None = None) -> bytes:
    """Copia la plantilla y escribe los valores desde la fila 5, sin quitar columnas."""
    ruta = ruta_cronograma or RUTA_CRONOGRAMA
    libro = load_workbook(io.BytesIO(_leer_bytes(ruta)))
    por_id = {hoja["id"]: hoja for hoja in resultado["hojas"]}
    for nombre in libro.sheetnames:
        hoja = por_id.get(nombre)
        if not hoja:
            continue
        ws = libro[nombre]
        tabla = hoja["tabla"]
        columnas = hoja["columnas"]
        for desplazamiento, columna in enumerate(columnas):
            serie = tabla[columna] if columna in tabla.columns else []
            valores = list(serie) if len(tabla) else []
            for indice, valor in enumerate(valores):
                escrito = _valor_excel(valor)
                if escrito is None:
                    continue
                ws.cell(FILA_DATOS + indice, 2 + desplazamiento, escrito)
    buffer = io.BytesIO()
    libro.save(buffer)
    libro.close()
    return buffer.getvalue()


def _vista(tabla: pd.DataFrame) -> pd.DataFrame:
    if tabla.empty:
        return tabla
    vista = tabla.copy()
    for columna in vista.columns:
        vista[columna] = vista[columna].map(lambda valor: "" if valor is None or (isinstance(valor, float) and pd.isna(valor)) else valor)
    return vista


@st.fragment
def render(*, embebido: bool = False) -> None:
    if embebido:
        st.markdown("#### Seguimiento de proyecto")
    else:
        st.markdown('<div class="section-title">Seguimiento de Proyecto</div>', unsafe_allow_html=True)
    st.caption(
        "Lee el cronograma de residencia y completa solo las celdas que salen de la bitácora, "
        "las ubicaciones, los programados cargados y los consumos. Lo demás queda en blanco para captura manual."
    )
    if not os.path.isfile(RUTA_CRONOGRAMA):
        st.error("No está el archivo data/residencia/CRONOGRAMA DE ACTIVIDADES.xlsx.")
        return
    try:
        resultado = construir()
        archivo = exportar_excel(resultado)
    except Exception as exc:
        st.error(f"No se pudo armar el seguimiento: {exc}")
        return

    fuentes = resultado["fuentes"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Registros de bitácora", f"{fuentes['bitacora']:,}")
    c2.metric("Ubicaciones", f"{fuentes['ubicaciones']:,}")
    c3.metric("Programados reales", f"{fuentes['programados_reales']:,}")
    c4.metric("Movimientos de almacén", f"{fuentes['movimientos_almacen']:,}")
    if fuentes["programados_demo"]:
        st.caption(f"Se dejaron fuera {fuentes['programados_demo']} programados marcados como demo.")
    st.download_button(
        "Descargar Excel consolidado",
        data=archivo,
        file_name="seguimiento_proyecto_residencia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="descarga_seguimiento_proyecto",
    )

    etiquetas = [f"{hoja['etiqueta']}  ·  {hoja['nombre']}" for hoja in resultado["hojas"]]
    elegida = st.selectbox("Actividad del cronograma", etiquetas, key="seguimiento_actividad")
    hoja = resultado["hojas"][etiquetas.index(elegida)]
    st.markdown(f"#### {hoja['nombre']}")
    if hoja["titulo"]:
        st.caption(hoja["titulo"])
    for nota in hoja["notas"]:
        st.caption(nota)
    diagnostico = hoja["diagnostico"]
    if not diagnostico.empty:
        con_datos = int((diagnostico["Estado"] == "Con datos de la app").sum())
        en_blanco = int((diagnostico["Estado"] == "En blanco").sum())
        st.markdown(f"**Diagnóstico:** {con_datos} columnas con datos de la app · {en_blanco} columnas en blanco.")
        st.dataframe(diagnostico, use_container_width=True, hide_index=True)
    tabla = _vista(hoja["tabla"])
    st.markdown(f"**Tabla** · {len(tabla):,} filas · {len(hoja['columnas'])} columnas")
    st.dataframe(tabla, use_container_width=True, hide_index=True, height=460)

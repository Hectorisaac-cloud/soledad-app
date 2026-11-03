"""Análisis del control de costos y rendimiento por equipo MCI."""

from __future__ import annotations

import html
import io
import os
import re
import unicodedata
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook

RAIZ = os.path.dirname(os.path.abspath(__file__))
RUTA = os.path.join(RAIZ, "data", "control_costos_mci.xlsx")
HOJA = "ANALISIS DE MCI"
COLUMNAS = (
    "fecha",
    "ubicacion",
    "modelo",
    "tipo",
    "descripcion",
    "insumos",
    "costo",
    "inactividad",
    "inactividad_texto",
)
_VACIO = {
    "modelo": "Sin modelo",
    "tipo": "Sin tipo",
    "ubicacion": "Sin ubicación",
}
_RE_MILES = re.compile(r"\d{1,3}(?:,\d{3})+")


def _clave(valor) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = "".join(caracter for caracter in texto if not unicodedata.combining(caracter))
    return re.sub(r"[^a-z0-9]+", " ", texto.casefold()).strip()


def _campo(clave: str) -> str | None:
    if not clave:
        return None
    if "modelo" in clave:
        return "modelo"
    if clave == "fecha" or clave.startswith("fecha "):
        return "fecha"
    if "ubicacion" in clave:
        return "ubicacion"
    if "mantenimiento" in clave:
        return "tipo"
    if "descripcion" in clave or clave.startswith("falla"):
        return "descripcion"
    if "insumo" in clave or "refaccion" in clave:
        return "insumos"
    if "costo" in clave:
        return "costo"
    if "inactividad" in clave or "horas de operacion" in clave:
        return "inactividad"
    return None


def _texto(valor) -> str:
    if valor is None or isinstance(valor, bool):
        return ""
    if isinstance(valor, float) and pd.isna(valor):
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    texto = str(valor).strip()
    if texto.casefold() in {"nan", "none"} or texto.startswith("="):
        return ""
    return texto


def _es_total(valor) -> bool:
    return _clave(valor) in {"total", "totales", "subtotal"}


def _fecha(valor) -> datetime | None:
    if valor is None or valor == "" or isinstance(valor, bool):
        return None
    if isinstance(valor, datetime):
        return valor.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(valor, date):
        return datetime(valor.year, valor.month, valor.day)
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if isinstance(valor, float) and pd.isna(valor):
            return None
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(valor))
        except (OverflowError, OSError, ValueError):
            return None
    texto = str(valor).strip()
    if not texto or texto.startswith("=") or _es_total(texto):
        return None
    for formato in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%Y/%m/%d"):
        try:
            return datetime.strptime(texto[:10], formato)
        except ValueError:
            continue
    return None


def _numero(valor) -> float | None:
    if valor is None or valor == "" or isinstance(valor, bool):
        return None
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and pd.isna(valor):
            return None
        return float(valor)
    texto = str(valor).strip()
    if not texto or texto.startswith("="):
        return None
    negativo = texto.startswith("-") or texto.startswith("(")
    token = re.sub(r"[^\d,.-]", "", texto).replace("-", "")
    if not token:
        return None
    if "," in token and "." in token:
        if token.rfind(",") > token.rfind("."):
            token = token.replace(".", "").replace(",", ".")
        else:
            token = token.replace(",", "")
    elif "," in token:
        token = token.replace(",", "") if _RE_MILES.fullmatch(token) else token.replace(",", ".")
    elif token.count(".") > 1:
        token = token.replace(".", "")
    try:
        numero = float(token)
    except ValueError:
        return None
    return -numero if negativo else numero


def _leer_bytes_compartido(ruta: str) -> bytes:
    """Lee el Excel aunque esté abierto en otro programa (bloqueo compartido de Windows)."""
    import ctypes
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.CreateFileW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    kernel.CreateFileW.restype = wintypes.HANDLE
    kernel.ReadFile.argtypes = [
        wintypes.HANDLE,
        wintypes.LPVOID,
        wintypes.DWORD,
        ctypes.POINTER(wintypes.DWORD),
        wintypes.LPVOID,
    ]
    kernel.ReadFile.restype = wintypes.BOOL
    kernel.GetFileSizeEx.argtypes = [wintypes.HANDLE, ctypes.POINTER(ctypes.c_longlong)]
    kernel.GetFileSizeEx.restype = wintypes.BOOL
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel.CloseHandle.restype = wintypes.BOOL

    handle = kernel.CreateFileW(ruta, 0x80000000, 0x1 | 0x2 | 0x4, None, 3, 0x80, None)
    if handle == wintypes.HANDLE(-1).value:
        codigo = ctypes.get_last_error()
        raise PermissionError(codigo, f"No se pudo leer {os.path.basename(ruta)} (código {codigo}).")
    try:
        tamano = ctypes.c_longlong()
        if not kernel.GetFileSizeEx(handle, ctypes.byref(tamano)):
            raise OSError(f"No se pudo medir {os.path.basename(ruta)}.")
        restante = tamano.value
        partes = []
        while restante > 0:
            bloque = min(restante, 1024 * 1024)
            buffer = ctypes.create_string_buffer(bloque)
            leidos = wintypes.DWORD()
            if not kernel.ReadFile(handle, buffer, bloque, ctypes.byref(leidos), None):
                raise OSError(f"No se pudo leer {os.path.basename(ruta)}.")
            if not leidos.value:
                break
            partes.append(buffer.raw[: leidos.value])
            restante -= leidos.value
        return b"".join(partes)
    finally:
        kernel.CloseHandle(handle)


def _abrir(ruta: str):
    try:
        return load_workbook(ruta, read_only=True, data_only=False, keep_links=False)
    except PermissionError:
        if os.name != "nt":
            raise
        return load_workbook(
            io.BytesIO(_leer_bytes_compartido(ruta)),
            read_only=True,
            data_only=False,
            keep_links=False,
        )


def _encabezado(filas: list[tuple]) -> tuple[int, dict[int, str]] | None:
    for indice, fila in enumerate(filas):
        mapa: dict[int, str] = {}
        for columna, valor in enumerate(fila):
            campo = _campo(_clave(valor))
            if campo and campo not in mapa.values():
                mapa[columna] = campo
        if "modelo" in mapa.values() and "costo" in mapa.values():
            return indice, mapa
    return None


def _fila_vacia(registro: dict) -> bool:
    if registro.get("inactividad_texto"):
        return False
    return not any(registro.get(columna) not in (None, "") for columna in COLUMNAS[:-1])


def _leer_hoja(filas: list[tuple]) -> pd.DataFrame | None:
    encontrado = _encabezado(filas)
    if encontrado is None:
        return None
    indice, mapa = encontrado
    registros = []
    for fila in filas[indice + 1 :]:
        crudo = {campo: fila[columna] if columna < len(fila) else None for columna, campo in mapa.items()}
        if any(_es_total(valor) for valor in crudo.values()):
            continue
        inactividad = _numero(crudo.get("inactividad"))
        texto_inactividad = "" if inactividad is not None else _texto(crudo.get("inactividad"))
        registro = {
            "fecha": _fecha(crudo.get("fecha")),
            "ubicacion": _texto(crudo.get("ubicacion")),
            "modelo": _texto(crudo.get("modelo")),
            "tipo": _texto(crudo.get("tipo")),
            "descripcion": _texto(crudo.get("descripcion")),
            "insumos": _texto(crudo.get("insumos")),
            "costo": _numero(crudo.get("costo")),
            "inactividad": inactividad,
            "inactividad_texto": texto_inactividad,
        }
        if _fila_vacia(registro):
            continue
        registros.append(registro)
    tabla = pd.DataFrame(registros, columns=COLUMNAS)
    tabla["fecha"] = pd.to_datetime(tabla["fecha"], errors="coerce")
    return tabla


def cargar(ruta: str | None = None) -> tuple[pd.DataFrame, str]:
    """Lee la tabla de intervenciones. Devuelve el cuadro y un aviso, si hace falta."""
    ruta = ruta or RUTA
    if not os.path.isfile(ruta):
        raise FileNotFoundError(ruta)
    libro = None
    try:
        libro = _abrir(ruta)
        hojas = [HOJA] if HOJA in libro.sheetnames else []
        hojas.extend(nombre for nombre in libro.sheetnames if nombre not in hojas)
        for nombre in hojas:
            tabla = _leer_hoja(list(libro[nombre].iter_rows(values_only=True)))
            if tabla is not None:
                return tabla, ""
        return pd.DataFrame(columns=COLUMNAS), (
            "No encontré la tabla con MODELO DE MCI y COSTO TOTAL DE LA INTERVENCION."
        )
    finally:
        if libro is not None:
            libro.close()


def marca_archivo(ruta: str | None = None) -> tuple[float, int]:
    ruta = ruta or RUTA
    if not os.path.isfile(ruta):
        return 0.0, 0
    estado = os.stat(ruta)
    return estado.st_mtime, estado.st_size


@st.cache_data(show_spinner=False)
def _cargar_cache(ruta: str, modificado: float, peso: int) -> tuple[pd.DataFrame, str]:
    del modificado, peso
    return cargar(ruta)


def leer(ruta: str | None = None) -> tuple[pd.DataFrame, str]:
    ruta = ruta or RUTA
    modificado, peso = marca_archivo(ruta)
    return _cargar_cache(ruta, modificado, peso)


def filtrar_modelo(tabla: pd.DataFrame, modelo: str) -> pd.DataFrame:
    if tabla.empty or not modelo or modelo == "Todos los modelos":
        return tabla
    return tabla.loc[tabla["modelo"].eq(modelo)]


def _suma(serie: pd.Series) -> float:
    numeros = pd.to_numeric(serie, errors="coerce")
    if not numeros.notna().any():
        return 0.0
    return float(numeros.sum())


def resumir(tabla: pd.DataFrame) -> dict[str, float | int]:
    if tabla.empty:
        return {"intervenciones": 0, "costo": 0.0, "inactividad": 0.0, "modelos": 0}
    modelos = tabla["modelo"].map(lambda valor: str(valor).strip()).replace("", pd.NA)
    return {
        "intervenciones": int(len(tabla)),
        "costo": _suma(tabla["costo"]),
        "inactividad": _suma(tabla["inactividad"]),
        "modelos": int(modelos.nunique(dropna=True)),
    }


def totales_por(tabla: pd.DataFrame, columna: str) -> pd.DataFrame:
    vacio = pd.DataFrame(columns=[columna, "intervenciones", "costo", "inactividad"])
    if tabla.empty or columna not in tabla.columns:
        return vacio
    etiqueta = _VACIO.get(columna, "Sin dato")
    trabajo = tabla.loc[:, [columna, "costo", "inactividad"]].copy()
    trabajo[columna] = trabajo[columna].map(lambda valor: str(valor or "").strip() or etiqueta)
    trabajo["costo"] = pd.to_numeric(trabajo["costo"], errors="coerce").fillna(0.0)
    trabajo["inactividad"] = pd.to_numeric(trabajo["inactividad"], errors="coerce").fillna(0.0)
    resumen = (
        trabajo.groupby(columna, dropna=False, sort=False)
        .agg(intervenciones=(columna, "size"), costo=("costo", "sum"), inactividad=("inactividad", "sum"))
        .reset_index()
    )
    if resumen.empty:
        return vacio
    return resumen.sort_values(["costo", "inactividad"], ascending=False, ignore_index=True)


def formato_moneda(valor: float) -> str:
    return f"${valor:,.2f}"


def formato_horas(valor: float) -> str:
    if abs(valor - round(valor)) < 0.05:
        return f"{int(round(valor)):,} h"
    return f"{valor:,.1f} h"


def _kpi(titulo: str, valor: str, detalle: str = "", tono: str = "") -> str:
    clase = "kpi-card " + tono if tono else "kpi-card"
    extra = f'<p class="s">{html.escape(detalle)}</p>' if detalle else ""
    return (
        f'<div class="{clase}"><div class="k">{html.escape(titulo)}</div>'
        f"<h3>{html.escape(valor)}</h3>{extra}</div>"
    )


def _estilo(figura, titulo: str) -> None:
    figura.update_layout(
        title=dict(text=titulo, font=dict(size=15, color="#f8fafc"), x=0.01, xanchor="left"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(15,23,42,0.35)",
        font=dict(color="#cbd5e1", family="Segoe UI, system-ui, sans-serif", size=12),
        margin=dict(l=12, r=12, t=52, b=8),
        height=380,
        showlegend=False,
        xaxis=dict(gridcolor="#1e293b", zerolinecolor="#334155", color="#94a3b8"),
        yaxis=dict(gridcolor="#1e293b", zerolinecolor="#334155", color="#94a3b8"),
        hoverlabel=dict(bgcolor="#0f172a", font_color="#e2e8f0", bordercolor="#334155"),
    )


def _barras(resumen: pd.DataFrame, columna: str, medida: str, titulo: str, color: str, moneda: bool) -> None:
    figura = px.bar(resumen, x=columna, y=medida, color_discrete_sequence=[color])
    if moneda:
        figura.update_traces(hovertemplate="%{x}<br>$%{y:,.2f}<extra></extra>")
        figura.update_yaxes(tickprefix="$", tickformat=",.0f", title_text="")
    else:
        figura.update_traces(hovertemplate="%{x}<br>%{y:,.1f} h<extra></extra>")
        figura.update_yaxes(ticksuffix=" h", title_text="")
    figura.update_xaxes(categoryorder="total descending", title_text="")
    _estilo(figura, titulo)
    st.plotly_chart(figura, width="stretch", config={"displayModeBar": False})


def _tabla_visible(tabla: pd.DataFrame) -> pd.DataFrame:
    if tabla.empty:
        return pd.DataFrame(
            columns=[
                "Fecha",
                "Ubicación",
                "Modelo de MCI",
                "Tipo de mantenimiento",
                "Descripción",
                "Insumos / refacciones",
                "Costo total",
                "Horas de inactividad",
            ]
        )
    ordenada = tabla.sort_values("fecha", ascending=False, na_position="last")

    def horas(fila) -> str:
        if pd.notna(fila["inactividad"]):
            return formato_horas(float(fila["inactividad"]))
        return str(fila.get("inactividad_texto") or "")

    return pd.DataFrame(
        {
            "Fecha": ordenada["fecha"].map(lambda valor: valor.strftime("%d/%m/%Y") if pd.notna(valor) else ""),
            "Ubicación": ordenada["ubicacion"],
            "Modelo de MCI": ordenada["modelo"],
            "Tipo de mantenimiento": ordenada["tipo"],
            "Descripción": ordenada["descripcion"],
            "Insumos / refacciones": ordenada["insumos"],
            "Costo total": ordenada["costo"].map(lambda valor: formato_moneda(float(valor)) if pd.notna(valor) else ""),
            "Horas de inactividad": [horas(fila) for _, fila in ordenada.iterrows()],
        }
    )


@st.fragment
def render() -> None:
    st.markdown('<div class="section-title">Control de costos MCI</div>', unsafe_allow_html=True)
    st.caption(
        "Lee `data/control_costos_mci.xlsx`. "
        "El costo total y las horas de inactividad corresponden al modelo seleccionado."
    )
    if not os.path.isfile(RUTA):
        st.error("No está el archivo data/control_costos_mci.xlsx.")
        return
    momento = datetime.fromtimestamp(os.path.getmtime(RUTA))
    st.caption(f"Archivo actualizado el {momento.strftime('%d/%m/%Y %H:%M')}.")
    try:
        tabla, aviso = leer(RUTA)
    except Exception as exc:
        st.error(f"No se pudo leer el control de costos: {exc}")
        return
    if aviso:
        st.warning(aviso)

    modelos = sorted(
        {str(valor).strip() for valor in tabla["modelo"].tolist() if str(valor).strip()},
        key=str.casefold,
    )
    if modelos:
        opciones = ["Todos los modelos", *modelos]
        elegido = st.session_state.get("costos_mci_modelo")
        if elegido is not None and elegido not in opciones:
            st.session_state.costos_mci_modelo = "Todos los modelos"
        modelo = st.selectbox("Modelo de equipo", opciones, key="costos_mci_modelo")
    else:
        modelo = "Todos los modelos"
        st.caption("Todavía no hay modelos para filtrar.")
    vista = filtrar_modelo(tabla, modelo)
    numeros = resumir(vista)
    costo = float(numeros["costo"])
    horas = float(numeros["inactividad"])
    intervenciones = int(numeros["intervenciones"])
    promedio = costo / intervenciones if intervenciones else 0.0
    if modelo == "Todos los modelos":
        detalle_corte = f"{int(numeros['modelos'])} modelo(s) en el archivo"
    else:
        detalle_corte = modelo

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(_kpi("Intervenciones", f"{intervenciones:,}", detalle_corte, "info"), unsafe_allow_html=True)
    c2.markdown(_kpi("Costo total", formato_moneda(costo), f"Promedio {formato_moneda(promedio)}", "ok"), unsafe_allow_html=True)
    c3.markdown(_kpi("Inactividad", formato_horas(horas), "Suma de horas fuera de operación", "warn"), unsafe_allow_html=True)
    por_hora = formato_moneda(costo / horas) if horas else "—"
    c4.markdown(_kpi("Costo por hora de paro", por_hora, "Costo total / horas de inactividad", "danger"), unsafe_allow_html=True)

    if tabla.empty:
        st.info(
            "La hoja «ANALISIS DE MCI» ya tiene la tabla de control, pero todavía no hay intervenciones capturadas. "
            "Al registrar fecha, modelo, costo y horas de inactividad, los totales y las gráficas se actualizan con el archivo."
        )
        return

    if vista.empty:
        st.warning("No hay intervenciones para el modelo seleccionado.")
        return

    por_modelo = totales_por(vista, "modelo")
    st.markdown("**Totales por modelo**")
    st.dataframe(
        pd.DataFrame(
            {
                "Modelo de MCI": por_modelo["modelo"],
                "Intervenciones": por_modelo["intervenciones"],
                "Costo total": por_modelo["costo"].map(formato_moneda),
                "Horas de inactividad": por_modelo["inactividad"].map(formato_horas),
            }
        ),
        width="stretch",
        hide_index=True,
    )

    grafica_izq, grafica_der = st.columns(2)
    if modelo == "Todos los modelos":
        with grafica_izq:
            _barras(por_modelo, "modelo", "costo", "Costo por modelo", "#38bdf8", True)
        with grafica_der:
            _barras(por_modelo, "modelo", "inactividad", "Inactividad por modelo", "#fbbf24", False)
    else:
        por_tipo = totales_por(vista, "tipo")
        por_ubicacion = totales_por(vista, "ubicacion")
        with grafica_izq:
            _barras(por_tipo, "tipo", "costo", f"Costo por tipo · {modelo}", "#38bdf8", True)
        with grafica_der:
            _barras(por_ubicacion, "ubicacion", "inactividad", f"Inactividad por ubicación · {modelo}", "#fbbf24", False)

    st.markdown("**Intervenciones**")
    st.dataframe(_tabla_visible(vista), width="stretch", hide_index=True)

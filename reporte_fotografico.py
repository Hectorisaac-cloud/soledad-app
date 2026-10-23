"""Reporte fotográfico consolidado a partir del machote oficial."""

from __future__ import annotations

import io
import os
import re
import unicodedata
import warnings
import zipfile
from datetime import date, datetime

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from PIL import Image as PILImage
from PIL import ImageOps
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

VERSION_DISENO_PDF = 8
PLANTILLA_XLSX = os.path.join("database", "machote_fotografico_template.xlsx")
FOTOS_POR_FILA = 3
FOTOS_POR_HOJA = 9
RANURAS = (
    (1, 10, 4, 23),
    (6, 10, 9, 23),
    (11, 10, 14, 23),
    (1, 25, 4, 38),
    (6, 25, 9, 38),
    (11, 25, 14, 38),
    (1, 40, 4, 53),
    (6, 40, 9, 53),
    (11, 40, 14, 53),
)
FIRMA_DERECHA = TwoCellAnchor(
    _from=AnchorMarker(col=9, colOff=243840, row=62, rowOff=38100),
    to=AnchorMarker(col=11, colOff=647700, row=66, rowOff=617220),
)
RECURSOS_PLANTILLA = (
    (
        "xl/media/image3.png",
        TwoCellAnchor(
            _from=AnchorMarker(col=0, colOff=243840, row=0, rowOff=114300),
            to=AnchorMarker(col=2, colOff=388620, row=0, rowOff=944880),
        ),
        ("logo_ops.png", "logo_izquierdo.png", "logo_soledad.png"),
    ),
    (
        "xl/media/image4.png",
        TwoCellAnchor(
            _from=AnchorMarker(col=12, colOff=350520, row=0, rowOff=76200),
            to=AnchorMarker(col=13, colOff=586740, row=0, rowOff=960120),
        ),
        ("logo_pjp4.png", "logo_derecho.png"),
    ),
    (
        "xl/media/image1.png",
        TwoCellAnchor(
            _from=AnchorMarker(col=2, colOff=883920, row=62, rowOff=76200),
            to=AnchorMarker(col=4, colOff=38100, row=66, rowOff=609600),
        ),
        ("firma_izquierda.png", "firma_luis.png"),
    ),
)
FIRMAS_ESTATICAS = ("firma_ops.png", "firma_derecha.png", "firma_gerardo.png")
_MARGEN_RANURA = 8
_FUENTE_PDF = "ReporteFoto"
_FUENTE_PDF_NEGRITA = "ReporteFoto-Bold"
CABECERA_PDF = (
    ("N° OT", "A4", 0.12),
    ("Ubicación", "C4", 0.13),
    ("Fecha", "D4", 0.12),
    ("Descripción de la Actividad", "E3", 0.37),
    ("% Avance", "K4", 0.10),
    ("Cuadrilla", "L4", 0.16),
)
FIRMANTES_PDF = (
    ("LUIS CARMELO CAMPOS BRITO", "COORDINADOR Y/O GERENTE DE PJP4"),
    ("GERARDO BERNABE RICARDEZ", "COORDINADOR DE OPS"),
)
COLOR_MARCO_FOTO = colors.HexColor("#B8862E")
COLOR_BARRA_FIRMAS = colors.HexColor("#E5A200")


def _ancho_columna_px(ws: Worksheet, columna: int) -> float:
    letra = get_column_letter(columna)
    ancho = ws.column_dimensions[letra].width
    if not ancho:
        ancho = ws.sheet_format.defaultColWidth or 8.43
    return ((256 * float(ancho) + int(128 / 7)) / 256) * 7


def _alto_fila_px(ws: Worksheet, fila: int) -> float:
    alto = ws.row_dimensions[fila].height
    if not alto:
        alto = 15
    return float(alto) * 96 / 72


def _origen_px(ws: Worksheet, columna: int, fila: int) -> tuple[float, float]:
    x = sum(_ancho_columna_px(ws, c) for c in range(1, columna))
    y = sum(_alto_fila_px(ws, r) for r in range(1, fila))
    return x, y


def _tamano_rango_px(ws: Worksheet, col_ini: int, fila_ini: int, col_fin: int, fila_fin: int) -> tuple[float, float]:
    ancho = sum(_ancho_columna_px(ws, c) for c in range(col_ini, col_fin + 1))
    alto = sum(_alto_fila_px(ws, r) for r in range(fila_ini, fila_fin + 1))
    return ancho, alto


def _ajustar_en_caja(ancho: int, alto: int, caja_ancho: float, caja_alto: float) -> tuple[int, int, int, int]:
    """Escala la imagen dentro de la caja y la centra. Conserva la proporción."""
    util_ancho = max(1, int(caja_ancho) - (2 * _MARGEN_RANURA))
    util_alto = max(1, int(caja_alto) - (2 * _MARGEN_RANURA))
    escala = min(util_ancho / max(ancho, 1), util_alto / max(alto, 1))
    nuevo_ancho = max(1, int(round(ancho * escala)))
    nuevo_alto = max(1, int(round(alto * escala)))
    desplazamiento_x = max(_MARGEN_RANURA, int((caja_ancho - nuevo_ancho) / 2))
    desplazamiento_y = max(_MARGEN_RANURA, int((caja_alto - nuevo_alto) / 2))
    return nuevo_ancho, nuevo_alto, desplazamiento_x, desplazamiento_y


def _preparar_imagen(ruta: str, max_ancho: int, max_alto: int) -> tuple[PILImage.Image, int, int]:
    with PILImage.open(ruta) as origen:
        imagen = ImageOps.exif_transpose(origen)
        if imagen.mode not in ("RGB", "RGBA"):
            imagen = imagen.convert("RGB")
        else:
            imagen = imagen.copy()
    ancho, alto = imagen.size
    nuevo_ancho, nuevo_alto, _, _ = _ajustar_en_caja(ancho, alto, max_ancho + (2 * _MARGEN_RANURA), max_alto + (2 * _MARGEN_RANURA))
    if (nuevo_ancho, nuevo_alto) != (ancho, alto):
        imagen = imagen.resize((nuevo_ancho, nuevo_alto), PILImage.Resampling.LANCZOS)
    if imagen.mode == "RGBA":
        fondo = PILImage.new("RGB", imagen.size, "white")
        fondo.paste(imagen, mask=imagen.getchannel("A"))
        imagen = fondo
    return imagen, imagen.size[0], imagen.size[1]


def _es_metaarchivo(datos: bytes, nombre: str = "") -> bool:
    """WMF y EMF no tienen cargador en el Pillow de Linux."""
    extension = os.path.splitext(str(nombre or ""))[1].lower()
    if extension in {".wmf", ".emf"}:
        return True
    if len(datos) >= 4 and datos[:4] == b"\xd7\xcd\xc6\x9a":
        return True
    if len(datos) >= 44 and datos[40:44] == b" EMF":
        return True
    if len(datos) >= 4 and datos[:4] in {b"\x01\x00\x09\x00", b"\x02\x00\x09\x00"}:
        return True
    return False


def _bytes_raster(datos: bytes, nombre: str = "") -> bytes | None:
    """Devuelve PNG o JPEG. Si Pillow no puede leer el WMF/EMF, regresa None."""
    if not datos:
        return None
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            with PILImage.open(io.BytesIO(datos)) as imagen:
                imagen.load()
                formato = str(imagen.format or "").upper()
                if formato in {"PNG", "JPEG", "JPG", "GIF"} and not _es_metaarchivo(datos, nombre):
                    return datos
                buffer = io.BytesIO()
                imagen.convert("RGB").save(buffer, format="PNG")
                return buffer.getvalue()
    except Exception:
        return None


def _archivo_estatico(*nombres: str) -> bytes | None:
    carpetas = (
        os.path.dirname(PLANTILLA_XLSX) or "database",
        os.path.join("database", "logos"),
        "static",
        os.path.join("static", "logos"),
    )
    for carpeta in carpetas:
        for nombre in nombres:
            ruta = os.path.join(carpeta, nombre)
            if not os.path.isfile(ruta):
                continue
            try:
                with open(ruta, "rb") as archivo:
                    datos = archivo.read()
            except OSError:
                continue
            raster = _bytes_raster(datos, ruta)
            if raster:
                return raster
    return None


def _png_minimo() -> bytes:
    buffer = io.BytesIO()
    PILImage.new("RGBA", (1, 1), (0, 0, 0, 0)).save(buffer, format="PNG")
    return buffer.getvalue()


def _plantilla_sin_vectores(ruta: str) -> io.BytesIO:
    """Sustituye WMF/EMF por un PNG vacío para que openpyxl pueda abrir el libro."""
    reemplazo = _png_minimo()
    salida = io.BytesIO()
    with zipfile.ZipFile(ruta) as origen, zipfile.ZipFile(salida, "w") as destino:
        for info in origen.infolist():
            datos = origen.read(info.filename)
            if _es_metaarchivo(datos, info.filename):
                datos = reemplazo
            destino.writestr(info.filename, datos)
    salida.seek(0)
    return salida


def _buffer_png(datos: bytes, nombre: str) -> io.BytesIO:
    buffer = io.BytesIO(datos)
    buffer.name = nombre
    return buffer


def _firma_derecha_png(ruta_plantilla: str) -> io.BytesIO | None:
    estatica = _archivo_estatico(*FIRMAS_ESTATICAS)
    if estatica:
        return _buffer_png(estatica, "firma_derecha.png")
    datos = b""
    try:
        with zipfile.ZipFile(ruta_plantilla) as paquete:
            for nombre in paquete.namelist():
                if os.path.splitext(nombre)[1].lower() in {".emf", ".wmf"}:
                    datos = paquete.read(nombre)
                    break
    except (FileNotFoundError, zipfile.BadZipFile, KeyError):
        datos = b""
    raster = _bytes_raster(datos, "firma.emf") if datos else None
    if not raster:
        return None
    return _buffer_png(raster, "firma_derecha.png")


def _anclar_foto(ws: Worksheet, imagen: PILImage.Image, ranura: tuple[int, int, int, int]) -> dict:
    col_ini, fila_ini, col_fin, fila_fin = ranura
    caja_ancho, caja_alto = _tamano_rango_px(ws, col_ini, fila_ini, col_fin, fila_fin)
    ancho, alto, dx, dy = _ajustar_en_caja(imagen.size[0], imagen.size[1], caja_ancho, caja_alto)
    if (ancho, alto) != imagen.size:
        imagen = imagen.resize((ancho, alto), PILImage.Resampling.LANCZOS)
    buffer = io.BytesIO()
    imagen.save(buffer, format="JPEG", quality=85)
    buffer.seek(0)
    buffer.name = "evidencia.jpg"
    dibujo = XLImage(buffer)
    dibujo.width = ancho
    dibujo.height = alto
    dibujo.anchor = OneCellAnchor(
        _from=AnchorMarker(
            col=col_ini - 1,
            colOff=pixels_to_EMU(dx),
            row=fila_ini - 1,
            rowOff=pixels_to_EMU(dy),
        ),
        ext=XDRPositiveSize2D(pixels_to_EMU(ancho), pixels_to_EMU(alto)),
    )
    ws.add_image(dibujo)
    origen_x, origen_y = _origen_px(ws, col_ini, fila_ini)
    return {
        "x": origen_x + dx,
        "y": origen_y + dy,
        "ancho": ancho,
        "alto": alto,
        "ranura": (
            origen_x,
            origen_y,
            caja_ancho,
            caja_alto,
        ),
        "proporcion": ancho / alto,
    }


def _quitar_metaarchivos(hoja: Worksheet) -> None:
    conservadas = []
    for imagen in list(getattr(hoja, "_images", []) or []):
        formato = str(getattr(imagen, "format", "") or "").lower()
        ruta = str(getattr(imagen, "path", "") or "")
        if formato in {"wmf", "emf"} or _es_metaarchivo(b"", ruta):
            continue
        conservadas.append(imagen)
    hoja._images = conservadas


def _recursos_oficiales(ruta_plantilla: str) -> list[tuple[bytes, TwoCellAnchor]]:
    recursos = []
    try:
        paquete = zipfile.ZipFile(ruta_plantilla)
    except (FileNotFoundError, zipfile.BadZipFile):
        paquete = None
    for nombre, ancla, respaldos in RECURSOS_PLANTILLA:
        raster = None
        if paquete is not None:
            try:
                raster = _bytes_raster(paquete.read(nombre), nombre)
            except KeyError:
                raster = None
        if raster is None:
            raster = _archivo_estatico(*respaldos)
        if raster:
            recursos.append((raster, ancla))
    if paquete is not None:
        paquete.close()
    return recursos


def _clonar_ancla(ancla: TwoCellAnchor) -> TwoCellAnchor:
    return TwoCellAnchor(
        _from=AnchorMarker(
            col=ancla._from.col,
            colOff=ancla._from.colOff,
            row=ancla._from.row,
            rowOff=ancla._from.rowOff,
        ),
        to=AnchorMarker(
            col=ancla.to.col,
            colOff=ancla.to.colOff,
            row=ancla.to.row,
            rowOff=ancla.to.rowOff,
        ),
    )


def _colocar_recurso(ws: Worksheet, datos: bytes, ancla: TwoCellAnchor, retenidas: list, nombre: str) -> None:
    raster = _bytes_raster(datos, nombre)
    if not raster:
        return
    try:
        buffer = _buffer_png(raster, nombre if nombre.lower().endswith((".png", ".jpg", ".jpeg")) else "recurso.png")
        imagen = XLImage(buffer)
        imagen.anchor = _clonar_ancla(ancla)
        ws.add_image(imagen)
        retenidas.append(buffer)
    except Exception:
        return


def _agregar_firma(ws: Worksheet, firma: io.BytesIO | None, retenidas: list) -> None:
    if firma is None:
        return
    try:
        firma.seek(0)
        copia = io.BytesIO(firma.read())
        copia.name = "firma_derecha.png"
        imagen = XLImage(copia)
        imagen.anchor = _clonar_ancla(FIRMA_DERECHA)
        ws.add_image(imagen)
        retenidas.append(copia)
    except Exception:
        return


def _escribir_encabezado(ws: Worksheet, actividad: dict, continuacion: str) -> None:
    descripcion = str(actividad.get("descripcion") or "").strip()
    if continuacion:
        descripcion = f"{descripcion} {continuacion}".strip()
    ws["A4"] = str(actividad.get("ot") or "")
    ws["C4"] = str(actividad.get("ubicacion") or "")
    fecha = actividad.get("fecha")
    if isinstance(fecha, datetime):
        fecha = fecha.date()
    ws["D4"] = fecha if isinstance(fecha, date) else str(fecha or "")
    if isinstance(fecha, date):
        ws["D4"].number_format = "DD/MM/YYYY"
    ws["E3"] = descripcion
    ws["K4"] = str(actividad.get("avance") or "")
    ws["L4"] = str(actividad.get("cuadrilla") or "")
    for coordenada in ("C4", "E3", "L4"):
        ws[coordenada].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _nombre_hoja(indice: int, actividad: dict, parte: int, usados: set[str]) -> str:
    fecha = actividad.get("fecha")
    dia = fecha.strftime("%d%m") if isinstance(fecha, date) else ""
    cuadrilla = re.sub(r"[:\\/?*\[\]]+", " ", str(actividad.get("cuadrilla") or "Cuadrilla"))
    cuadrilla = re.sub(r"\s+", " ", cuadrilla).strip()
    sufijo = f" {parte}" if parte else ""
    base = f"{indice:02d} {cuadrilla} {dia}{sufijo}".strip()
    base = base[:31] or f"Hoja {indice}"
    nombre = base
    extra = 2
    while nombre.casefold() in usados:
        recorte = base[: 31 - len(str(extra)) - 1]
        nombre = f"{recorte} {extra}"
        extra += 1
    usados.add(nombre.casefold())
    return nombre


def _preparar_libro(ruta_plantilla: str) -> tuple[Workbook, Worksheet, io.BytesIO | None]:
    if not os.path.isfile(ruta_plantilla):
        raise FileNotFoundError(
            f"No está la plantilla del reporte fotográfico en {ruta_plantilla}."
        )
    with warnings.catch_warnings():
        warnings.filterwarnings(
            "ignore",
            message="wmf image format is not supported",
            category=UserWarning,
        )
        warnings.filterwarnings(
            "ignore",
            message="will be removed because it cannot be read",
            category=UserWarning,
        )
        try:
            libro = load_workbook(ruta_plantilla)
        except Exception:
            libro = load_workbook(_plantilla_sin_vectores(ruta_plantilla))
            for hoja in libro.worksheets:
                hoja._images = [
                    imagen
                    for imagen in list(getattr(hoja, "_images", []) or [])
                    if not (getattr(imagen, "width", None) == 1 and getattr(imagen, "height", None) == 1)
                ]
    libro._external_links = []
    for hoja in libro.worksheets:
        _quitar_metaarchivos(hoja)
    plantilla = libro.active
    plantilla["K4"] = plantilla["K4"].value if not isinstance(plantilla["K4"].value, str) or not str(plantilla["K4"].value).startswith("=") else ""
    if isinstance(plantilla["Q21"].value, str) and str(plantilla["Q21"].value).startswith("="):
        plantilla["Q21"] = None
    firma = _firma_derecha_png(ruta_plantilla)
    return libro, plantilla, firma


def _hoja_a_pdf(
    pdf: canvas.Canvas,
    ws: Worksheet,
    fotos: list[tuple[PILImage.Image, dict]],
    recursos: list[tuple[bytes, TwoCellAnchor]],
) -> None:
    ancho_hoja, alto_hoja = letter
    margen = 8
    ancho_contenido = sum(_ancho_columna_px(ws, c) for c in range(1, 15)) * 72 / 96
    alto_contenido = sum((ws.row_dimensions[r].height or 15) for r in range(1, 70))
    escala = min(
        (ancho_hoja - 2 * margen) / max(ancho_contenido, 1),
        (alto_hoja - 2 * margen) / max(alto_contenido, 1),
    )
    fuente, fuente_negrita = _fuentes_pdf()

    def x_pt(columna: int, extra: float = 0) -> float:
        return margen + (sum(_ancho_columna_px(ws, c) for c in range(1, columna)) + extra) * 72 / 96 * escala

    def y_pt(fila: int, extra: float = 0) -> float:
        acumulado = sum((ws.row_dimensions[r].height or 15) for r in range(1, fila)) + extra * 72 / 96
        return alto_hoja - margen - acumulado * escala

    def alto_pt(fila_ini: int, fila_fin: int) -> float:
        return sum((ws.row_dimensions[r].height or 15) for r in range(fila_ini, fila_fin + 1)) * escala

    def ancho_pt(col_ini: int, col_fin: int) -> float:
        return sum(_ancho_columna_px(ws, c) for c in range(col_ini, col_fin + 1)) * 72 / 96 * escala

    pdf.setFillColorRGB(1, 1, 1)
    pdf.rect(0, 0, ancho_hoja, alto_hoja, stroke=0, fill=1)

    for datos, ancla in recursos:
        try:
            lector = ImageReader(io.BytesIO(datos))
        except Exception:
            continue
        x0 = x_pt(ancla._from.col + 1, ancla._from.colOff / 9525)
        y_sup = y_pt(ancla._from.row + 1, ancla._from.rowOff / 9525)
        x1 = x_pt(ancla.to.col + 1, ancla.to.colOff / 9525)
        y_inf = y_pt(ancla.to.row + 1, ancla.to.rowOff / 9525)
        _dibujar_proporcion(pdf, lector, x0, y_inf, max(1, x1 - x0), max(1, y_sup - y_inf))

    pdf.setFillColorRGB(0.08, 0.12, 0.1)
    pdf.setFont(fuente_negrita, 11)
    pdf.drawCentredString(x_pt(4) + ancho_pt(4, 12) / 2, y_pt(1) - 28, "REPORTE DE EVIDENCIA FOTOGRAFICA")
    tabla_x = x_pt(1)
    tabla_ancho = ancho_pt(1, 14)
    tabla_arriba = y_pt(3)
    tabla_abajo = y_pt(8)
    alto_etiquetas = min(16.0, (tabla_arriba - tabla_abajo) * 0.3)
    alto_valores = tabla_arriba - tabla_abajo - alto_etiquetas
    anchos = [tabla_ancho * proporcion for _etiqueta, _celda, proporcion in CABECERA_PDF]
    padding = 2
    fila_etiquetas = [
        _parrafo_ajustado(etiqueta, fuente_negrita, 7.5, ancho - 2 * padding, alto_etiquetas - 2)
        for (etiqueta, _celda, _p), ancho in zip(CABECERA_PDF, anchos)
    ]
    fila_valores = [
        _parrafo_ajustado(
            _valor_tabla_pdf(celda, ws[celda].value, ws["C4"].value),
            fuente_negrita if celda == "E3" else fuente,
            7 if celda == "E3" else 8,
            ancho - 2 * padding,
            alto_valores - 2 * padding,
        )
        for (_etiqueta, celda, _p), ancho in zip(CABECERA_PDF, anchos)
    ]
    tabla = Table([fila_etiquetas, fila_valores], colWidths=anchos, rowHeights=[alto_etiquetas, alto_valores])
    tabla.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 1.2, colors.black),
                ("INNERGRID", (0, 0), (-1, -1), 0.8, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E6E6E6")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), padding),
                ("RIGHTPADDING", (0, 0), (-1, -1), padding),
                ("TOPPADDING", (0, 0), (-1, -1), 1),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ]
        )
    )
    tabla.wrapOn(pdf, tabla_ancho, tabla_arriba - tabla_abajo)
    tabla.drawOn(pdf, tabla_x, tabla_abajo)

    pdf.setStrokeColor(COLOR_MARCO_FOTO)
    pdf.setLineWidth(1.5)
    for imagen, caja in fotos:
        col_ini, fila_ini, col_fin, fila_fin = _ranura_de_caja(ws, caja)
        buffer = io.BytesIO()
        imagen.save(buffer, format="JPEG", quality=85)
        buffer.seek(0)
        dibujado = _dibujar_proporcion(
            pdf,
            ImageReader(buffer),
            x_pt(col_ini) + caja["dx"] * 72 / 96 * escala,
            y_pt(fila_ini, caja["dy"] + caja["alto"]),
            caja["ancho"] * 72 / 96 * escala,
            caja["alto"] * 72 / 96 * escala,
        )
        if dibujado:
            pdf.rect(*dibujado, stroke=1, fill=0)

    caja_arriba = y_pt(62) + 4
    caja_abajo = y_pt(70)
    mitad = tabla_x + tabla_ancho / 2
    pdf.setFillColor(COLOR_BARRA_FIRMAS)
    pdf.rect(tabla_x, caja_arriba + 6, tabla_ancho, 7, stroke=0, fill=1)
    pdf.setStrokeColorRGB(0, 0, 0)
    pdf.setLineWidth(1)
    pdf.rect(tabla_x, caja_abajo, tabla_ancho, caja_arriba - caja_abajo, stroke=1, fill=0)
    pdf.line(mitad, caja_arriba, mitad, caja_abajo)
    pdf.setLineWidth(0.6)
    for indice, (nombre, cargo) in enumerate(FIRMANTES_PDF):
        x_columna = tabla_x + indice * tabla_ancho / 2
        centro = x_columna + tabla_ancho / 4
        pdf.line(x_columna + tabla_ancho * 0.08, y_pt(68) + 1, x_columna + tabla_ancho * 0.42, y_pt(68) + 1)
        pdf.setFillColorRGB(0.1, 0.14, 0.12)
        pdf.setFont(fuente, 7)
        pdf.drawCentredString(centro, y_pt(62) - 6, _texto_celda(ws["A62" if indice == 0 else "H62"].value) or "APROBADO:")
        pdf.setFont(fuente_negrita, 8)
        pdf.drawCentredString(centro, y_pt(68) - 10, nombre)
        pdf.setFont(fuente, 7)
        pdf.drawCentredString(centro, y_pt(69) - 9, cargo)
    pdf.showPage()


def _ranura_de_caja(ws: Worksheet, caja: dict) -> tuple[int, int, int, int]:
    return caja["ranura_idx"]


def _dibujar_proporcion(pdf, lector, x, y, ancho, alto) -> tuple[float, float, float, float] | None:
    try:
        imagen_ancho, imagen_alto = lector.getSize()
    except Exception:
        pdf.drawImage(lector, x, y, width=ancho, height=alto, preserveAspectRatio=True, mask="auto", anchor="c")
        return None
    if not imagen_ancho or not imagen_alto:
        return None
    escala = min(ancho / imagen_ancho, alto / imagen_alto)
    dibujo_ancho = imagen_ancho * escala
    dibujo_alto = imagen_alto * escala
    x_dibujo = x + (ancho - dibujo_ancho) / 2
    y_dibujo = y + (alto - dibujo_alto) / 2
    pdf.drawImage(
        lector,
        x_dibujo,
        y_dibujo,
        width=dibujo_ancho,
        height=dibujo_alto,
        preserveAspectRatio=True,
        mask="auto",
    )
    return x_dibujo, y_dibujo, dibujo_ancho, dibujo_alto


def _texto_celda(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()


def _sin_acentos(texto: str) -> str:
    normal = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(c for c in normal if unicodedata.category(c) != "Mn")


_TAGS_CUADRILLA_PDF = {
    "electricos": "ELECT",
    "electrico": "ELECT",
    "instrumentos": "INSTR",
    "instrumentista": "INSTR",
    "mecanico 1": "MECA-1",
    "mecanico 2": "MECA-2",
    "mecanico 3": "MECA-3",
    "mecanico 4": "MECA-4",
    "montaje": "MONTA",
    "soldadura": "SOLD",
    "retro": "RETRO",
    "retroexcavadora": "RETRO",
    "retro excavadora": "RETRO",
}


def _es_sector_zona(texto: str) -> bool:
    """El primer tramo de la ubicación de campo es el sector, no el pozo."""
    limpio = str(texto or "").strip()
    if not limpio:
        return False
    if "(" in limpio and ")" in limpio:
        return True
    clave = re.sub(r"\s+", " ", _sin_acentos(limpio).casefold()).strip()
    return clave in {"base pjp4"} or clave.startswith(("almacen ", "bateria ", "estacion ", "sector "))


def _ubicacion_tabla_pdf(valor) -> str:
    """Deja solo el pozo. Quita el sector o la zona y el equipo que vengan pegados."""
    texto = _texto_celda(valor)
    partes = [parte.strip() for parte in texto.split("·") if parte.strip()]
    if not partes:
        return texto
    if len(partes) >= 2 and _es_sector_zona(partes[0]):
        return partes[1]
    return partes[0]


def _tag_cuadrilla_pdf(valor) -> str:
    texto = _texto_celda(valor)
    clave = re.sub(r"\s+", " ", _sin_acentos(texto).casefold()).strip()
    if clave in _TAGS_CUADRILLA_PDF:
        return _TAGS_CUADRILLA_PDF[clave]
    compacta = clave.replace(" ", "")
    if compacta in _TAGS_CUADRILLA_PDF:
        return _TAGS_CUADRILLA_PDF[compacta]
    return texto.upper()


def _descripcion_tabla_pdf(descripcion, ubicacion) -> str:
    """Agrega el pozo al final de la descripción. Solo se usa al dibujar el PDF."""
    texto = _texto_celda(descripcion)
    pozo = _ubicacion_tabla_pdf(ubicacion)
    if not pozo:
        return texto
    sufijo = f"EN {pozo}"
    if texto.endswith(sufijo):
        return texto
    if not texto:
        return sufijo
    return f"{texto} {sufijo}"


def _valor_tabla_pdf(celda: str, valor, ubicacion=None) -> str:
    if celda == "E3":
        return _descripcion_tabla_pdf(valor, ubicacion)
    if celda == "C4":
        return _ubicacion_tabla_pdf(valor)
    if celda == "L4":
        return _tag_cuadrilla_pdf(valor)
    return _texto_celda(valor)


def _parrafo_ajustado(texto: str, fuente: str, tamano: float, ancho: float, alto: float) -> Paragraph:
    contenido = escape(texto or "")
    while True:
        estilo = ParagraphStyle(
            "celda",
            fontName=fuente,
            fontSize=tamano,
            leading=tamano + 1.5,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#141F1A"),
        )
        parrafo = Paragraph(contenido, estilo)
        _ancho, alto_texto = parrafo.wrap(ancho, alto)
        if alto_texto <= alto or tamano <= 5:
            return parrafo
        tamano -= 0.5


def _fuentes_pdf() -> tuple[str, str]:
    if _FUENTE_PDF in pdfmetrics.getRegisteredFontNames():
        return _FUENTE_PDF, _FUENTE_PDF_NEGRITA
    windir = os.environ.get("WINDIR", r"C:\Windows")
    pares = (
        (
            os.path.join(windir, "Fonts", "arial.ttf"),
            os.path.join(windir, "Fonts", "arialbd.ttf"),
        ),
        (
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ),
        (
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        ),
    )
    for regular, negrita in pares:
        if os.path.isfile(regular) and os.path.isfile(negrita):
            pdfmetrics.registerFont(TTFont(_FUENTE_PDF, regular))
            pdfmetrics.registerFont(TTFont(_FUENTE_PDF_NEGRITA, negrita))
            return _FUENTE_PDF, _FUENTE_PDF_NEGRITA
    return "Helvetica", "Helvetica-Bold"


def _ordenar(actividades: list[dict]) -> list[dict]:
    def clave(item: dict):
        fecha = item.get("fecha") or date.min
        if isinstance(fecha, datetime):
            fecha = fecha.date()
        orden = item.get("orden")
        return (fecha, 99 if orden is None else int(orden), int(item.get("secuencia") or 0))

    return sorted(actividades, key=clave)


def generar_reporte_fotografico(
    actividades: list[dict],
    ruta_plantilla: str = PLANTILLA_XLSX,
) -> tuple[bytes, bytes, dict]:
    """Devuelve el Excel, el PDF y un resumen de hojas, fotos y omisiones."""
    if not actividades:
        raise ValueError("No hay actividades en el periodo seleccionado.")
    libro, plantilla, firma = _preparar_libro(ruta_plantilla)
    recursos = _recursos_oficiales(ruta_plantilla)
    retenidas: list = []
    if firma is not None:
        retenidas.append(firma)
    usados: set[str] = set()
    paginas_pdf: list[tuple[Worksheet, list]] = []
    fotos_insertadas = 0
    fotos_omitidas = 0
    hojas = 0
    for indice, actividad in enumerate(_ordenar(actividades), start=1):
        rutas = [str(ruta) for ruta in (actividad.get("fotos") or []) if str(ruta).strip()]
        bloques = [rutas[i : i + FOTOS_POR_HOJA] for i in range(0, len(rutas), FOTOS_POR_HOJA)] or [[]]
        for parte, bloque in enumerate(bloques):
            hoja = libro.copy_worksheet(plantilla)
            _quitar_metaarchivos(hoja)
            hoja.title = _nombre_hoja(indice, actividad, parte + 1 if len(bloques) > 1 else 0, usados)
            _escribir_encabezado(
                hoja,
                actividad,
                "(continuación)" if parte else "",
            )
            hoja.page_setup.orientation = "portrait"
            hoja.page_setup.fitToWidth = 1
            hoja.page_setup.fitToHeight = 1
            hoja.page_setup.paperSize = hoja.PAPERSIZE_LETTER
            hoja.sheet_properties.pageSetUpPr.fitToPage = True
            hoja.print_area = "A1:N69"
            for indice_recurso, (datos, ancla) in enumerate(recursos):
                _colocar_recurso(hoja, datos, ancla, retenidas, f"recurso_{indice_recurso}.png")
            _agregar_firma(hoja, firma, retenidas)
            fotos_hoja = []
            for posicion, ruta in enumerate(bloque[:FOTOS_POR_FILA * 3]):
                if not os.path.isfile(ruta):
                    fotos_omitidas += 1
                    continue
                ranura = RANURAS[posicion]
                caja_ancho, caja_alto = _tamano_rango_px(hoja, *ranura)
                try:
                    imagen, _, _ = _preparar_imagen(
                        ruta,
                        max(1, int(caja_ancho) - 2 * _MARGEN_RANURA),
                        max(1, int(caja_alto) - 2 * _MARGEN_RANURA),
                    )
                except Exception:
                    fotos_omitidas += 1
                    continue
                caja = _anclar_foto(hoja, imagen, ranura)
                caja["dx"] = caja["x"] - caja["ranura"][0]
                caja["dy"] = caja["y"] - caja["ranura"][1]
                caja["ranura_idx"] = ranura
                caja["origen"] = (imagen.size[0], imagen.size[1])
                fotos_hoja.append((imagen, caja))
                fotos_insertadas += 1
            paginas_pdf.append((hoja, fotos_hoja))
            hojas += 1
    libro.remove(plantilla)
    recursos_pdf = list(recursos)
    if firma is not None:
        firma.seek(0)
        recursos_pdf.append((firma.read(), FIRMA_DERECHA))
    pdf_buffer = io.BytesIO()
    pdf = canvas.Canvas(pdf_buffer, pagesize=letter)
    for hoja, fotos_hoja in paginas_pdf:
        _hoja_a_pdf(pdf, hoja, fotos_hoja, recursos_pdf)
    pdf.save()
    for hoja in libro.worksheets:
        _quitar_metaarchivos(hoja)
    excel_buffer = io.BytesIO()
    libro.save(excel_buffer)
    resumen = {
        "hojas": hojas,
        "actividades": len(actividades),
        "fotos": fotos_insertadas,
        "omitidas": fotos_omitidas,
    }
    return excel_buffer.getvalue(), pdf_buffer.getvalue(), resumen


def cajas_desde_libro(contenido: bytes) -> list[dict]:
    """Lee las fotos de evidencia ancladas en la cuadrícula 3xN."""
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning)
        libro = load_workbook(io.BytesIO(contenido))
    cajas = []
    for hoja in libro.worksheets:
        for dibujo in hoja._images:
            ancla = dibujo.anchor
            if not isinstance(ancla, OneCellAnchor):
                continue
            fila = ancla._from.row + 1
            if fila < 10 or fila > 53:
                continue
            columna = ancla._from.col + 1
            dx = ancla._from.colOff / 9525
            dy = ancla._from.rowOff / 9525
            ancho = ancla.ext.cx / 9525
            alto = ancla.ext.cy / 9525
            origen_x, origen_y = _origen_px(hoja, columna, fila)
            ranura = next(
                item
                for item in RANURAS
                if item[0] == columna and item[1] == fila
            )
            caja_x, caja_y = _origen_px(hoja, ranura[0], ranura[1])
            caja_ancho, caja_alto = _tamano_rango_px(hoja, *ranura)
            cajas.append(
                {
                    "hoja": hoja.title,
                    "x": origen_x + dx,
                    "y": origen_y + dy,
                    "ancho": ancho,
                    "alto": alto,
                    "ranura": (caja_x, caja_y, caja_ancho, caja_alto),
                    "pixeles": (dibujo.width, dibujo.height),
                }
            )
    return cajas

"""Sitios de campo que no son pozos: baterías, estaciones, módulos y áreas."""

from __future__ import annotations

import re
import unicodedata

SITIOS_VACIOS = frozenset(
    {
        "",
        "nan",
        "none",
        "ninguno",
        "n/a",
        "na",
        "-",
        "—",
        "sin pozo",
        "sin ubicacion",
        "sin ubicación",
        "seleccionar...",
    }
)

# Clave compacta (sin espacios ni acentos) -> etiqueta oficial del reporte.
SITIOS_OPERATIVOS = {
    "BPB": "BPB",
    "GDOBPB": "BPB",
    "BATERIAPALOBLANCO": "BPB",
    "BATERIAPALOBLANCOBPB": "BPB",
    "BSI": "BSI",
    "GDOBSI": "BSI",
    "BATERIADESSEPARACIONSOLEDADI": "BSI",
    "BATERIADESSEPARACIONSOLEDADIBSI": "BSI",
    "BATERIASOLEDADI": "BSI",
    "BSII": "BSII",
    "GDOBSII": "BSII",
    "BATERIADESSEPARACIONSOLEDADII": "BSII",
    "BATERIADESSEPARACIONSOLEDADIIBSII": "BSII",
    "BATERIASOLEDADII": "BSII",
    "BSNI": "BSNI",
    "GDOBSNI": "BSNI",
    "BSSN1": "BSNI",
    "BSN1": "BSNI",
    "BSN": "BSNI",
    "BATERIADESSEPARACIONSOLEDADNORTE": "BSNI",
    "BATERIADESSEPARACIONSOLEDADNORTEBSNI": "BSNI",
    "BATERIADESSEPARACIONSOLEDADNORTEI": "BSNI",
    "BATERIASOLEDADNORTE": "BSNI",
    "ECSII": "ECSII",
    "GDOECSII": "ECSII",
    "ECS2": "ECSII",
    "ESTACIONDECOMPRESIONSOLEDADII": "ECSII",
    "ESTACIONDECOMPRESIONSOLEDADIICSII": "ECSII",
    "ESTACIONDECOMPRESIONSOLEDAD2": "ECSII",
    "ESTACIONDECOMPRESIONSOLEDADNORTEECSII": "ECSII",
    "ECSN": "ECSN",
    "GDOECSN": "ECSN",
    "ECOSLN": "ECSN",
    "ESTACIONDECOMPRESIONSOLEDADNORTE": "ECSN",
    "ESTACIONDECOMPRESIONSOLEDADNORTEECSN": "ECSN",
    "MODI": "MOD I",
    "MOD1": "MOD I",
    "MODULOI": "MOD I",
    "MODULOISOLEDADNORTE": "MOD I",
    "MODULOISOLEDADNORTEMODI": "MOD I",
    "MODII": "MOD II",
    "MOD2": "MOD II",
    "MODULOII": "MOD II",
    "MODIII": "MOD III",
    "MOD3": "MOD III",
    "MODULOIII": "MOD III",
    "MODIV": "MOD IV",
    "MOD4": "MOD IV",
    "MODULOIV": "MOD IV",
    "AOPS": "AOPS",
    "ALMACENOPS": "AOPS",
    "ALMACENOPSAOPS": "AOPS",
    "ALMACENOPERACIONES": "AOPS",
    "BASEPJP4": "BASE PJP4",
    "BSEPJP4": "BSE PJP4",
    "BASESOLEDAD": "BSE PJP4",
    "BATSOLNORTE": "BSNI",
    "BATSOLI": "BSI",
    "BATSOLII": "BSII",
    "BATPALOBLANCO": "BPB",
    "ESTACIONDECOMPRESORASSOLEDADNORTE": "ECSN",
    "ESTACIONDECOMPRESORAS2SUR": "ECSII",
    "ESTACIONDECOMPRESORASSUR": "ECSII",
    "TANQUES": "TANQUES",
    "TANQUE": "TANQUES",
    "AREADEMAQUINAS": "AREA DE MAQUINAS",
    "AREAMAQUINAS": "AREA DE MAQUINAS",
    "CUARTODEMAQUINAS": "AREA DE MAQUINAS",
}

_TOKENS = tuple(sorted(SITIOS_OPERATIVOS, key=len, reverse=True))
_RE_VACIO_TOKEN = re.compile(r"\b(ninguno|none|nan|n/?a)\b", re.I)
_RE_PARENTESIS = re.compile(r"\(([A-Za-z][A-Za-z0-9 /_-]{1,24})\)")
_RE_TOKEN_SUELTO = re.compile(
    r"\b(BSNI|BSSN-?1|BSN-?1|BSII|BSI|BPB|BSE\s*PJP4|ECSII|ECS-?2|ECSN|ECO\s*SLN|"
    r"MOD(?:ULO)?\s*-?\s*(?:IV|III|II|I|4|3|2|1)|"
    r"AOPS|TANQUES?)\b",
    re.I,
)
_RE_NOMBRES_LARGOS = (
    (re.compile(r"BATER[IÍ]A.{0,48}PALO\s*BLANCO|\bPALO\s*BLANCO\b.{0,12}\bBPB\b", re.I), "BPB"),
    (re.compile(r"BATER[IÍ]A.{0,48}SOLEDAD\s*(?:II|2)\b", re.I), "BSII"),
    (re.compile(r"BATER[IÍ]A.{0,48}SOLEDAD\s*NORTE", re.I), "BSNI"),
    (re.compile(r"BATER[IÍ]A.{0,48}SOLEDAD\s*I\b", re.I), "BSI"),
    (re.compile(r"ESTACI[OÓ]N.{0,40}COMPRESORAS?\s*-?\s*2\s*SUR", re.I), "ECSII"),
    (re.compile(r"ESTACI[OÓ]N.{0,40}COMPRESORAS?.{0,40}SOLEDAD\s*NORTE", re.I), "ECSN"),
    (re.compile(r"ESTACI[OÓ]N.{0,40}COMPRESI[OÓ]N.{0,40}SOLEDAD\s*(?:II|2)\b", re.I), "ECSII"),
    (re.compile(r"ESTACI[OÓ]N.{0,40}COMPRESI[OÓ]N.{0,40}SOLEDAD\s*NORTE", re.I), "ECSN"),
    (re.compile(r"\bBAT\.?\s*SOL\.?\s*NORTE\b", re.I), "BSNI"),
    (re.compile(r"\bBAT\.?\s*SOL\.?\s*II\b", re.I), "BSII"),
    (re.compile(r"\bBAT\.?\s*SOL\.?\s*I\b", re.I), "BSI"),
    (re.compile(r"\bBAT\.?\s*PALO\s*BLANCO\b", re.I), "BPB"),
    (re.compile(r"\bBASE\s+SOLEDAD\b", re.I), "BSE PJP4"),
    (re.compile(r"M[OÓ]DULO\s*(?:IV|4)\b", re.I), "MOD IV"),
    (re.compile(r"M[OÓ]DULO\s*(?:III|3)\b", re.I), "MOD III"),
    (re.compile(r"M[OÓ]DULO\s*(?:II|2)\b", re.I), "MOD II"),
    (re.compile(r"M[OÓ]DULO\s*(?:I|1)\b", re.I), "MOD I"),
    (re.compile(r"ALMAC[EÉ]N\s*(?:OPS|OPERACIONES)", re.I), "AOPS"),
    (re.compile(r"\b(?:AREA|CUARTO)\s+DE\s+M[AÁ]QUINAS\b", re.I), "AREA DE MAQUINAS"),
)


def _sin_acentos(texto: str) -> str:
    normal = unicodedata.normalize("NFD", str(texto or ""))
    return "".join(c for c in normal if unicodedata.category(c) != "Mn")


def compactar_sitio(texto: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", _sin_acentos(texto).upper())


def _reconstruir_tokens() -> None:
    global _TOKENS
    _TOKENS = tuple(sorted(SITIOS_OPERATIVOS, key=len, reverse=True))


def registrar_par_catalogo(largo, corto) -> str:
    """Registra Nombre Largo (Excel) → Nombre Corto (App) del catálogo maestro."""
    clave = clave_sitio(corto) or clave_sitio(largo)
    if not clave:
        plano_corto = compactar_sitio(corto)
        clave = SITIOS_OPERATIVOS.get(plano_corto) or str(corto or "").strip()
    if not clave or es_vacio_sitio(clave):
        return ""
    for texto in (corto, largo):
        plano = compactar_sitio(texto)
        if plano:
            SITIOS_OPERATIVOS[plano] = clave
    if clave not in NOMBRES_SITIO:
        largo_txt = " ".join(str(largo or "").split())
        NOMBRES_SITIO[clave] = f"{largo_txt} ({clave})" if largo_txt else clave
    _reconstruir_tokens()
    return clave


def es_vacio_sitio(texto) -> bool:
    plano = re.sub(r"\s+", " ", _sin_acentos(str(texto or "")).casefold()).strip()
    return plano in SITIOS_VACIOS


def _canon_compacto(plano: str) -> str:
    if not plano:
        return ""
    directo = SITIOS_OPERATIVOS.get(plano)
    if directo:
        return directo
    for token in _TOKENS:
        if plano == token or plano.endswith(token) or token in plano:
            if token in {"BSI", "BSN"} and any(otro in plano for otro in ("BSNI", "BSII", "BSSN")):
                continue
            if token == "BATSOLI" and "BATSOLII" in plano:
                continue
            if token == "MODI" and any(otro in plano for otro in ("MODII", "MODIII", "MODIV")):
                continue
            if token == "MODII" and "MODIII" in plano:
                continue
            if token == "ECSN" and "ECSII" in plano:
                continue
            if token == "TANQUE" and "TANQUES" in plano:
                return "TANQUES"
            return SITIOS_OPERATIVOS[token]
    return ""


def _clave_en_texto(parte: str) -> str:
    limpio = _RE_VACIO_TOKEN.sub(" ", parte)
    plano = compactar_sitio(limpio)
    hallada = SITIOS_OPERATIVOS.get(plano) or _canon_compacto(plano)
    if hallada:
        return hallada
    par = _RE_PARENTESIS.search(parte)
    if par:
        hallada = SITIOS_OPERATIVOS.get(compactar_sitio(par.group(1))) or _canon_compacto(
            compactar_sitio(par.group(1))
        )
        if hallada:
            return hallada
    suelto = _RE_TOKEN_SUELTO.search(parte)
    if suelto:
        hallada = SITIOS_OPERATIVOS.get(compactar_sitio(suelto.group(1)))
        if hallada:
            return hallada
    for patron, clave in _RE_NOMBRES_LARGOS:
        if patron.search(parte):
            return clave
    return ""


def clave_sitio(texto) -> str:
    """Devuelve BPB, BSNI, MOD I, etc. si el texto contiene la clave o el nombre largo."""
    bruto = " ".join(str(texto or "").replace('"', "").replace("'", "").split())
    if es_vacio_sitio(bruto):
        return ""
    partes = [p.strip() for p in re.split(r"[·|;,/]+", bruto) if p.strip() and not es_vacio_sitio(p)]
    candidatos = partes or [bruto]
    for parte in candidatos:
        hallada = _clave_en_texto(parte)
        if hallada:
            return hallada
    return _clave_en_texto(bruto) or _canon_compacto(compactar_sitio(_RE_VACIO_TOKEN.sub(" ", bruto)))


def clave_sitio_campos(*textos) -> str:
    """Busca la clave en todos los textos del registro (ubicación, pozo, descripción)."""
    vistos: list[str] = []
    for texto in textos:
        bruto = " ".join(str(texto or "").replace('"', "").replace("'", "").split())
        if not bruto:
            continue
        hallada = clave_sitio(bruto)
        if hallada:
            return hallada
        if not es_vacio_sitio(bruto) and bruto not in vistos:
            vistos.append(bruto)
    if vistos:
        return clave_sitio(" · ".join(vistos))
    return ""


COMPONENTES_INSTALACION = (
    "Separador",
    "Compresor",
    "Quemador",
    "Header",
    "Manifold",
    "Skid",
    "Tanque",
    "Bomba",
    "Valvula",
    "PCV",
    "LCV",
    "Scrubber",
)

NOMBRES_SITIO = {
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
    "BASE PJP4": "Base PJP4",
    "BSE PJP4": "Base Soledad (BSE PJP4)",
    "TANQUES": "Tanques",
    "AREA DE MAQUINAS": "Area de Maquinas",
}

_RE_EQUIPO_EXPLICITO = re.compile(
    r"\b(MCI|UBM|UB\b|HLU|ICI|ARROW|WEICHAI|VORTEC|CUMMINS|LUFKIN|WEATHERFORD|GR-?20)\b",
    re.I,
)
_RE_CODIGO_POZO = re.compile(r"^(PSN|PGA|PPB|PS|PA|PB|PC)-?\d", re.I)


def nombre_oficial_sitio(clave: str) -> str:
    return NOMBRES_SITIO.get(str(clave or "").strip(), str(clave or "").strip())


def opciones_tag_instalacion(clave: str) -> list[str]:
    vistos: list[str] = []
    for item in (str(clave or "").strip(), *COMPONENTES_INSTALACION):
        if item and item not in vistos:
            vistos.append(item)
    return vistos


def extraer_tag_instalacion(partes: list[str], clave: str) -> str:
    """TAG o componente capturado. Si no hay, usa la clave de la instalación (nunca Ninguno)."""
    nombre = nombre_oficial_sitio(clave)
    for parte in partes:
        limpio = quitar_ninguno(parte)
        if not limpio or es_vacio_sitio(limpio):
            continue
        if limpio.casefold() == nombre.casefold():
            continue
        return limpio
    return str(clave or "").strip()


def componer_etiqueta_instalacion(ubicacion, tag="") -> str:
    clave = clave_sitio(ubicacion) or clave_sitio(tag)
    if not clave:
        return quitar_ninguno(" · ".join(p for p in (ubicacion, tag) if str(p or "").strip()))
    nombre = nombre_oficial_sitio(clave)
    etiqueta = quitar_ninguno(tag) or extraer_tag_instalacion(partes_ubicacion_visibles(ubicacion, tag), clave)
    if not etiqueta or es_vacio_sitio(etiqueta):
        etiqueta = clave
    if etiqueta.casefold() == nombre.casefold():
        return f"{nombre} · {clave}"
    return f"{nombre} · {etiqueta}"


def es_instalacion(texto) -> bool:
    return bool(clave_sitio(texto))


def es_equipo_explicito(texto) -> bool:
    bruto = str(texto or "").strip()
    if not bruto or es_vacio_sitio(bruto) or clave_sitio(bruto):
        return False
    if _RE_CODIGO_POZO.match(compactar_sitio(bruto)) or _RE_CODIGO_POZO.match(bruto.replace(" ", "")):
        return False
    return bool(_RE_EQUIPO_EXPLICITO.search(bruto))


_RE_EN_NINGUNO = re.compile(r"\s*\ben\s+ninguno\b", re.I)
_RE_PARTICULA_NINGUNO = re.compile(
    r"(?:\s*[·•|,;/\-]\s*|\s+)?(?:en\s+)?ninguno\b",
    re.I,
)


def quitar_ninguno(texto) -> str:
    """Quita 'Ninguno', '· Ninguno' y 'EN NINGUNO' de cualquier etiqueta o descripción."""
    bruto = str(texto or "")
    if not bruto.strip():
        return ""
    bruto = _RE_EN_NINGUNO.sub(" ", bruto)
    bruto = _RE_PARTICULA_NINGUNO.sub(" ", bruto)
    partes = []
    for parte in re.split(r"[·•|;]+", bruto):
        limpio = " ".join(parte.replace('"', "").replace("'", "").split())
        if not limpio or es_vacio_sitio(limpio):
            continue
        partes.append(limpio)
    return " · ".join(partes)


def partes_ubicacion_visibles(*textos) -> list[str]:
    vistos: list[str] = []
    for texto in textos:
        limpio = quitar_ninguno(texto)
        for parte in re.split(r"[·|;]+", limpio):
            item = " ".join(parte.split())
            if not item or es_vacio_sitio(item):
                continue
            if item not in vistos:
                vistos.append(item)
    return vistos


def texto_ubicacion_final(*textos) -> str:
    """Batería/estación + TAG capturado. Nunca 'Ninguno'."""
    partes = partes_ubicacion_visibles(*textos)
    clave = clave_sitio_campos(*partes, *[str(t or "") for t in textos])
    if clave:
        return componer_etiqueta_instalacion(nombre_oficial_sitio(clave), extraer_tag_instalacion(partes, clave))
    unidos = quitar_ninguno(" · ".join(partes))
    if es_vacio_sitio(unidos) or unidos.upper() == "NINGUNO":
        return ""
    return unidos


def limpiar_etiqueta_ubicacion(*textos) -> str:
    """Nombre de batería/estación sin '· Ninguno' ni pozo inventado."""
    return texto_ubicacion_final(*textos)


def clave_o_nombre_sitio(texto) -> str:
    """TAG para 'EN BSNI' / 'EN separador'. Vacío si el único valor era Ninguno."""
    limpio = quitar_ninguno(texto)
    if not limpio or es_vacio_sitio(limpio) or _RE_VACIO_TOKEN.search(limpio):
        return ""
    partes = partes_ubicacion_visibles(limpio)
    clave = clave_sitio(limpio)
    if clave:
        tag = extraer_tag_instalacion(partes, clave) or clave
        if not tag or es_vacio_sitio(tag) or _RE_VACIO_TOKEN.search(tag):
            return clave
        return tag
    compacto = " ".join(limpio.split())
    if compacto.upper() == "NINGUNO" or len(compacto) > 32:
        return ""
    return compacto

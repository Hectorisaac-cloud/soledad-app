"""Respaldo de la bitácora y las evidencias fuera del disco efímero.

Streamlit Cloud reconstruye el disco con el contenido de `main` en cada
despliegue. Los registros capturados en la nube se conservan en la rama
Git `datos`, que no dispara ese despliegue. Al arrancar, la app se queda
con la bitácora que tenga más registros y descarga las fotos que falten.
"""

from __future__ import annotations

import base64
import hashlib
import io
import json
import os
import tempfile
import threading
import urllib.error
import urllib.request
import zipfile
from datetime import datetime

RAIZ = os.path.dirname(os.path.abspath(__file__))
RAMA_DATOS_DEFAULT = "datos"
REPO_DEFAULT = "Hectorisaac-cloud/soledad-app"
ARCHIVO_BITACORA = "database/bitacora_operativa_soledad.xlsx"
ARCHIVO_ESTADO = "database/persistencia_estado.json"
ARCHIVO_EVIDENCIAS_NUEVAS = "database/evidencias_pendientes.txt"
ARCHIVOS_DATOS = (
    ARCHIVO_BITACORA,
    "baul_evidencias.json",
    "evidencia_fotografica.json",
    "almacen_insumos.json",
    "asignaciones_diarias.json",
    "catalogos.json",
    "ubicaciones_equipos.json",
    "control_programa_octubre.json",
    "programados_ops.json",
    "actividades_cuadrillas.json",
    "personal_cuadrillas.json",
    "vehiculos_ot.json",
    "soporte_refacciones.json",
)
CARPETAS_EVIDENCIA = (
    "evidencia_fotografica_octubre",
    "Baul_Evidencias",
    "evidencias",
    "Evidencia Fotografica",
    os.path.join("taller_vehicular", "evidencias"),
)
EXTENSIONES_EVIDENCIA = {".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"}
MAX_ARCHIVOS_POR_RESPALDO = 25
MAX_BYTES_POR_RESPALDO = 20 * 1024 * 1024
MAX_BYTES_ZIP = 25 * 1024 * 1024
MAX_DESCARGAS_ARRANQUE = 30
API = "https://api.github.com"
_CANDADO = threading.Lock()


def elegir_bitacora(filas_locales: int, filas_remotas: int) -> str:
    """La copia con más registros gana. Un empate conserva la local."""
    if filas_remotas > filas_locales:
        return "remota"
    return "local"


def debe_bloquear_vaciado(filas_actuales: int, filas_nuevas: int) -> bool:
    if filas_nuevas > 0:
        return False
    return filas_actuales != 0


def git_blob_sha(contenido: bytes) -> str:
    encabezado = f"blob {len(contenido)}\0".encode("ascii")
    return hashlib.sha1(encabezado + contenido).hexdigest()


def ruta_segura(relativa: str) -> str | None:
    texto = str(relativa or "").replace("\\", "/").strip()
    if texto.startswith("./"):
        texto = texto[2:]
    texto = texto.lstrip("/")
    if not texto or texto.startswith("../"):
        return None
    partes = [parte for parte in texto.split("/") if parte and parte != "."]
    if any(parte == ".." for parte in partes):
        return None
    return "/".join(partes)


def resolver_ruta(relativa: str) -> str:
    segura = ruta_segura(relativa) or ""
    directa = segura.replace("/", os.sep)
    if directa and os.path.exists(directa):
        return directa
    return os.path.join(RAIZ, directa) if directa else os.path.join(RAIZ, ARCHIVO_BITACORA)


def _secreto(nombre: str) -> str:
    entorno = os.environ.get(nombre, "").strip()
    if entorno:
        return entorno
    try:
        import streamlit as st

        valor = st.secrets.get(nombre, "")
    except Exception:
        return ""
    return str(valor or "").strip()


def credenciales_github() -> dict | None:
    token = _secreto("GITHUB_TOKEN") or _secreto("GH_TOKEN")
    if not token:
        return None
    repo = _secreto("GITHUB_REPO") or REPO_DEFAULT
    rama = _secreto("GITHUB_RAMA_DATOS") or RAMA_DATOS_DEFAULT
    if "/" not in repo or not rama or rama in {"main", "master"}:
        return None
    return {"token": token, "repo": repo, "rama": rama}


def contar_filas_bitacora(ruta: str | None) -> int:
    if not ruta or not os.path.exists(ruta) or os.path.getsize(ruta) <= 0:
        return 0
    try:
        from openpyxl import load_workbook
    except Exception:
        return 0
    wb = load_workbook(ruta, read_only=True, data_only=True)
    try:
        nombres = {nombre.casefold(): nombre for nombre in wb.sheetnames}
        hoja = None
        for preferida in ("bitacora", "hoja5", "hoja2"):
            if preferida in nombres:
                hoja = nombres[preferida]
                break
        ws = wb[hoja or wb.sheetnames[0]]
        col_fecha = None
        fila_header = None
        total = 0
        for indice, fila in enumerate(ws.iter_rows(values_only=True)):
            if col_fecha is None and indice < 12:
                for columna, valor in enumerate(fila):
                    if str(valor or "").strip().casefold() == "fecha":
                        col_fecha = columna
                        fila_header = indice
                        break
                continue
            if col_fecha is None or fila_header is None or indice <= fila_header:
                continue
            if col_fecha < len(fila) and fila[col_fecha] not in (None, ""):
                total += 1
        return total
    except Exception:
        return -1
    finally:
        wb.close()


def _guardar_estado(estado: dict) -> None:
    destino = resolver_ruta(ARCHIVO_ESTADO)
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    temporal = destino + ".tmp"
    with open(temporal, "w", encoding="utf-8") as archivo:
        json.dump(estado, archivo, ensure_ascii=False, indent=2)
    os.replace(temporal, destino)


def leer_estado() -> dict:
    ruta = resolver_ruta(ARCHIVO_ESTADO)
    if not os.path.exists(ruta):
        return {}
    try:
        with open(ruta, encoding="utf-8") as archivo:
            datos = json.load(archivo)
    except Exception:
        return {}
    return datos if isinstance(datos, dict) else {}


def _api(metodo: str, url: str, token: str, payload: dict | None = None, crudo: bool = False):
    datos = None if payload is None else json.dumps(payload).encode("utf-8")
    peticion = urllib.request.Request(url, data=datos, method=metodo)
    peticion.add_header("Authorization", f"Bearer {token}")
    peticion.add_header("User-Agent", "soledad-app-persistencia")
    peticion.add_header("X-GitHub-Api-Version", "2022-11-28")
    peticion.add_header("Accept", "application/vnd.github.raw" if crudo else "application/vnd.github+json")
    if payload is not None:
        peticion.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(peticion, timeout=45) as respuesta:
            cuerpo = respuesta.read()
    except urllib.error.HTTPError as exc:
        detalle = exc.read().decode("utf-8", errors="replace")[:500]
        raise RuntimeError(f"GitHub respondió {exc.code}: {detalle}") from exc
    if crudo:
        return cuerpo
    if not cuerpo:
        return {}
    return json.loads(cuerpo.decode("utf-8"))


def _obtener_ref(cred: dict, rama: str) -> dict | None:
    url = f"{API}/repos/{cred['repo']}/git/ref/heads/{rama}"
    try:
        return _api("GET", url, cred["token"])
    except RuntimeError as exc:
        if "404" in str(exc):
            return None
        raise


def _asegurar_rama(cred: dict) -> str:
    ref = _obtener_ref(cred, cred["rama"])
    if ref:
        return ref["object"]["sha"]
    base = _obtener_ref(cred, "main") or _obtener_ref(cred, "master")
    if not base:
        raise RuntimeError("No se encontró la rama main para crear el respaldo.")
    creado = _api(
        "POST",
        f"{API}/repos/{cred['repo']}/git/refs",
        cred["token"],
        {"ref": f"refs/heads/{cred['rama']}", "sha": base["object"]["sha"]},
    )
    return creado["object"]["sha"]


def _arbol_remoto(cred: dict, commit_sha: str) -> dict[str, str]:
    commit = _api("GET", f"{API}/repos/{cred['repo']}/git/commits/{commit_sha}", cred["token"])
    arbol = _api(
        "GET",
        f"{API}/repos/{cred['repo']}/git/trees/{commit['tree']['sha']}?recursive=1",
        cred["token"],
    )
    mapa = {}
    for item in arbol.get("tree") or []:
        if item.get("type") == "blob" and item.get("path") and item.get("sha"):
            mapa[item["path"]] = item["sha"]
    return mapa


def _descargar_blob(cred: dict, sha: str) -> bytes:
    return _api("GET", f"{API}/repos/{cred['repo']}/git/blobs/{sha}", cred["token"], crudo=True)


def _escribir_atomico(destino: str, contenido: bytes) -> None:
    os.makedirs(os.path.dirname(destino) or ".", exist_ok=True)
    fd, temporal = tempfile.mkstemp(suffix=".tmp", dir=os.path.dirname(destino) or None)
    os.close(fd)
    try:
        with open(temporal, "wb") as archivo:
            archivo.write(contenido)
        os.replace(temporal, destino)
    finally:
        if os.path.exists(temporal):
            try:
                os.remove(temporal)
            except OSError:
                pass


def _prefijo_evidencia(ruta: str) -> bool:
    normal = ruta.replace("\\", "/")
    for carpeta in CARPETAS_EVIDENCIA:
        prefijo = carpeta.replace("\\", "/").rstrip("/") + "/"
        if normal == carpeta.replace("\\", "/") or normal.startswith(prefijo):
            return True
    return False


def restaurar_desde_github() -> dict:
    cred = credenciales_github()
    if not cred:
        return {"ok": False, "motivo": "sin_token", "mensaje": "No hay GITHUB_TOKEN configurado."}
    if not _CANDADO.acquire(blocking=False):
        return {"ok": False, "motivo": "ocupado", "mensaje": "Ya hay una sincronización en curso."}
    try:
        ref = _obtener_ref(cred, cred["rama"])
        if not ref:
            estado = {
                "ok": True,
                "cuando": datetime.now().isoformat(timespec="seconds"),
                "mensaje": f"La rama {cred['rama']} todavía no existe. Se creará con el próximo registro.",
                "filas_locales": contar_filas_bitacora(resolver_ruta(ARCHIVO_BITACORA)),
            }
            _guardar_estado(estado)
            return estado
        arbol = _arbol_remoto(cred, ref["object"]["sha"])
        locales = contar_filas_bitacora(resolver_ruta(ARCHIVO_BITACORA))
        remotas = 0
        adopto_remota = False
        blob_bitacora = arbol.get(ARCHIVO_BITACORA)
        if blob_bitacora:
            contenido = _descargar_blob(cred, blob_bitacora)
            fd, temporal = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            try:
                with open(temporal, "wb") as archivo:
                    archivo.write(contenido)
                remotas = contar_filas_bitacora(temporal)
                if locales >= 0 and remotas >= 0 and elegir_bitacora(locales, remotas) == "remota":
                    _escribir_atomico(resolver_ruta(ARCHIVO_BITACORA), contenido)
                    locales = remotas
                    adopto_remota = True
            finally:
                if os.path.exists(temporal):
                    os.remove(temporal)
        recuperados = 0
        if adopto_remota:
            for relativa in ARCHIVOS_DATOS:
                if relativa == ARCHIVO_BITACORA or relativa not in arbol:
                    continue
                _escribir_atomico(resolver_ruta(relativa), _descargar_blob(cred, arbol[relativa]))
                recuperados += 1
        else:
            for relativa in ARCHIVOS_DATOS:
                if relativa == ARCHIVO_BITACORA or relativa not in arbol:
                    continue
                destino = resolver_ruta(relativa)
                if os.path.exists(destino) and os.path.getsize(destino) > 0:
                    continue
                _escribir_atomico(destino, _descargar_blob(cred, arbol[relativa]))
                recuperados += 1
        faltantes = []
        for ruta, sha in arbol.items():
            segura = ruta_segura(ruta)
            if not segura or not _prefijo_evidencia(segura):
                continue
            destino = resolver_ruta(segura)
            if os.path.exists(destino) and os.path.getsize(destino) > 0:
                continue
            faltantes.append((segura, sha))
        for segura, sha in faltantes[:MAX_DESCARGAS_ARRANQUE]:
            _escribir_atomico(resolver_ruta(segura), _descargar_blob(cred, sha))
            recuperados += 1
        pendientes = max(0, len(faltantes) - MAX_DESCARGAS_ARRANQUE)
        mensaje = (
            f"Bitácora local: {locales} registros. "
            f"Respaldo en {cred['rama']}: {remotas} registros. "
            f"Se recuperaron {recuperados} archivo(s)."
        )
        if pendientes:
            mensaje += f" Quedan {pendientes} evidencias por descargar en el siguiente arranque."
        estado = {
            "ok": True,
            "cuando": datetime.now().isoformat(timespec="seconds"),
            "mensaje": mensaje,
            "filas_locales": locales,
            "filas_remotas": remotas,
            "recuperados": recuperados,
            "pendientes": pendientes,
        }
        _guardar_estado(estado)
        return estado
    except Exception as exc:
        estado = {
            "ok": False,
            "cuando": datetime.now().isoformat(timespec="seconds"),
            "mensaje": f"No se pudo recuperar el respaldo: {exc}",
        }
        try:
            _guardar_estado(estado)
        except Exception:
            pass
        return estado
    finally:
        _CANDADO.release()


def restaurar_al_arrancar() -> dict:
    try:
        import streamlit as st

        if st.session_state.get("_persistencia_restaurada"):
            return {"ok": True, "omitido": True}
        st.session_state["_persistencia_restaurada"] = True
    except Exception:
        pass
    return restaurar_desde_github()


def marcar_evidencia_nueva(rutas: list[str] | None) -> None:
    limpias = []
    for ruta in rutas or []:
        texto = str(ruta or "").replace("\\", "/").strip()
        if not texto:
            continue
        if os.path.isabs(texto):
            try:
                texto = os.path.relpath(texto, RAIZ).replace("\\", "/")
            except ValueError:
                continue
        segura = ruta_segura(texto)
        if segura and _prefijo_evidencia(segura):
            limpias.append(segura)
    if not limpias:
        return
    destino = resolver_ruta(ARCHIVO_EVIDENCIAS_NUEVAS)
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    previas = []
    if os.path.exists(destino):
        with open(destino, encoding="utf-8") as archivo:
            previas = [linea.strip() for linea in archivo if linea.strip()]
    vistos = set(previas)
    with open(destino, "a", encoding="utf-8") as archivo:
        for ruta in limpias:
            if ruta not in vistos:
                archivo.write(ruta + "\n")
                vistos.add(ruta)


def _evidencias_marcadas() -> list[str]:
    ruta = resolver_ruta(ARCHIVO_EVIDENCIAS_NUEVAS)
    if not os.path.exists(ruta):
        return []
    with open(ruta, encoding="utf-8") as archivo:
        return [linea.strip().replace("\\", "/") for linea in archivo if linea.strip()]


def _archivos_a_respaldar(arbol: dict[str, str] | None = None, *, solo_copia_local: bool = False) -> list[str]:
    elegidos = []
    for relativa in ARCHIVOS_DATOS:
        if os.path.exists(resolver_ruta(relativa)):
            elegidos.append(relativa.replace("\\", "/"))
    if solo_copia_local:
        for relativa in _evidencias_marcadas():
            if relativa not in elegidos and os.path.exists(resolver_ruta(relativa)):
                elegidos.append(relativa)
        return elegidos
    arbol = arbol or {}
    for carpeta in CARPETAS_EVIDENCIA:
        base = resolver_ruta(carpeta)
        if not os.path.isdir(base):
            continue
        for raiz, _, nombres in os.walk(base):
            for nombre in nombres:
                if nombre.startswith("~$") or nombre.startswith("."):
                    continue
                if os.path.splitext(nombre)[1].lower() not in EXTENSIONES_EVIDENCIA:
                    continue
                absoluto = os.path.join(raiz, nombre)
                relativa = ruta_segura(os.path.relpath(absoluto, RAIZ).replace("\\", "/"))
                if not relativa or relativa in arbol or relativa in elegidos:
                    continue
                elegidos.append(relativa)
    return elegidos


def respaldar_cambios(motivo: str) -> dict:
    cred = credenciales_github()
    if not cred:
        estado = {
            "ok": False,
            "motivo": "sin_token",
            "cuando": datetime.now().isoformat(timespec="seconds"),
            "mensaje": "El registro quedó en el servidor, sin copia en GitHub. Configura GITHUB_TOKEN para que sobreviva a un redespliegue.",
            "filas_locales": contar_filas_bitacora(resolver_ruta(ARCHIVO_BITACORA)),
        }
        try:
            _guardar_estado(estado)
        except Exception:
            pass
        return estado
    if not _CANDADO.acquire(blocking=False):
        return {"ok": False, "motivo": "ocupado", "mensaje": "Ya hay una sincronización en curso."}
    try:
        commit_sha = _asegurar_rama(cred)
        arbol = _arbol_remoto(cred, commit_sha)
        locales = contar_filas_bitacora(resolver_ruta(ARCHIVO_BITACORA))
        blob_remoto = arbol.get(ARCHIVO_BITACORA)
        if blob_remoto:
            fd, temporal = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            try:
                with open(temporal, "wb") as archivo:
                    archivo.write(_descargar_blob(cred, blob_remoto))
                remotas = contar_filas_bitacora(temporal)
            finally:
                if os.path.exists(temporal):
                    os.remove(temporal)
            if locales < 0 or remotas < 0:
                estado = {
                    "ok": False,
                    "motivo": "conteo",
                    "cuando": datetime.now().isoformat(timespec="seconds"),
                    "mensaje": "No se subió la bitácora porque no se pudo contar una de las dos copias.",
                    "filas_locales": locales,
                    "filas_remotas": remotas,
                }
                _guardar_estado(estado)
                return estado
            if remotas > locales:
                estado = {
                    "ok": False,
                    "motivo": "remoto_mas_nuevo",
                    "cuando": datetime.now().isoformat(timespec="seconds"),
                    "mensaje": (
                        f"No se subió la bitácora: la rama {cred['rama']} tiene {remotas} registros "
                        f"y esta copia tiene {locales}. Recupera el respaldo antes de seguir capturando."
                    ),
                    "filas_locales": locales,
                    "filas_remotas": remotas,
                }
                _guardar_estado(estado)
                return estado
        cambios = []
        bytes_usados = 0
        omitidos = 0
        for relativa in _archivos_a_respaldar(arbol):
            if relativa == ARCHIVO_BITACORA and locales <= 0:
                continue
            absoluto = resolver_ruta(relativa)
            try:
                with open(absoluto, "rb") as archivo:
                    contenido = archivo.read()
            except OSError:
                continue
            if not contenido:
                continue
            firma = git_blob_sha(contenido)
            if arbol.get(relativa) == firma:
                continue
            if len(cambios) >= MAX_ARCHIVOS_POR_RESPALDO or bytes_usados + len(contenido) > MAX_BYTES_POR_RESPALDO:
                omitidos += 1
                continue
            blob = _api(
                "POST",
                f"{API}/repos/{cred['repo']}/git/blobs",
                cred["token"],
                {"content": base64.b64encode(contenido).decode("ascii"), "encoding": "base64"},
            )
            cambios.append(
                {"path": relativa, "mode": "100644", "type": "blob", "sha": blob["sha"]}
            )
            bytes_usados += len(contenido)
        if not cambios:
            estado = {
                "ok": True,
                "cuando": datetime.now().isoformat(timespec="seconds"),
                "mensaje": f"La rama {cred['rama']} ya tiene esta bitácora ({locales} registros).",
                "filas_locales": locales,
                "subidos": 0,
            }
            _guardar_estado(estado)
            return estado
        commit = _api("GET", f"{API}/repos/{cred['repo']}/git/commits/{commit_sha}", cred["token"])
        arbol_nuevo = _api(
            "POST",
            f"{API}/repos/{cred['repo']}/git/trees",
            cred["token"],
            {"base_tree": commit["tree"]["sha"], "tree": cambios},
        )
        nuevo = _api(
            "POST",
            f"{API}/repos/{cred['repo']}/git/commits",
            cred["token"],
            {
                "message": f"Respaldo operativo Soledad\n\n{motivo}. Registros: {locales}.",
                "tree": arbol_nuevo["sha"],
                "parents": [commit_sha],
            },
        )
        _api(
            "PATCH",
            f"{API}/repos/{cred['repo']}/git/refs/heads/{cred['rama']}",
            cred["token"],
            {"sha": nuevo["sha"]},
        )
        mensaje = f"Respaldo enviado a la rama {cred['rama']}: {locales} registros y {len(cambios)} archivo(s)."
        if omitidos:
            mensaje += f" Quedaron {omitidos} archivo(s) para el siguiente respaldo."
        estado = {
            "ok": True,
            "cuando": datetime.now().isoformat(timespec="seconds"),
            "mensaje": mensaje,
            "filas_locales": locales,
            "subidos": len(cambios),
            "omitidos": omitidos,
            "commit": nuevo.get("sha", ""),
        }
        _guardar_estado(estado)
        return estado
    except Exception as exc:
        estado = {
            "ok": False,
            "cuando": datetime.now().isoformat(timespec="seconds"),
            "mensaje": f"No se pudo respaldar en GitHub: {exc}",
            "filas_locales": contar_filas_bitacora(resolver_ruta(ARCHIVO_BITACORA)),
        }
        try:
            _guardar_estado(estado)
        except Exception:
            pass
        return estado
    finally:
        _CANDADO.release()


def construir_copia_seguridad() -> tuple[bytes, str]:
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"respaldo_bitacora_{sello}.zip"
    buffer = io.BytesIO()
    incluidos = []
    omitidos = []
    usados = 0
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as paquete:
        for relativa in _archivos_a_respaldar(solo_copia_local=True):
            absoluto = resolver_ruta(relativa)
            try:
                tamano = os.path.getsize(absoluto)
            except OSError:
                continue
            if usados + tamano > MAX_BYTES_ZIP:
                omitidos.append(relativa)
                continue
            paquete.write(absoluto, relativa.replace("\\", "/"))
            incluidos.append(relativa.replace("\\", "/"))
            usados += tamano
        nota = [
            "Copia de seguridad de Soledad",
            f"Generada: {datetime.now().isoformat(timespec='seconds')}",
            f"Registros en la bitácora: {contar_filas_bitacora(resolver_ruta(ARCHIVO_BITACORA))}",
            "",
            "Incluye database/bitacora_operativa_soledad.xlsx, los JSON operativos",
            "y las evidencias capturadas en esta ejecución que aún están en el disco.",
            "Las fotos históricas ya versionadas siguen en el repositorio.",
            "La rama Git de respaldo es datos. No la conectes al despliegue de Streamlit.",
            "",
            "Archivos incluidos:",
            *incluidos,
        ]
        if omitidos:
            nota.extend(["", "No cupieron en esta copia:", *omitidos[:200]])
        paquete.writestr("LEEME_RESPALDO.txt", "\n".join(nota))
    return buffer.getvalue(), nombre

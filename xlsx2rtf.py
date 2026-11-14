#!/usr/bin/env python3
"""
xlsx2rtf: convierte automáticamente archivos de Excel (.xlsx / .xlsm) a RTF.

- Cada hoja visible del libro se escribe en el RTF con su título.
- Los números se muestran con el formato de cada celda (decimales, miles,
  porcentajes, textos como "m."), igual que en Excel.
- Las columnas vacías se eliminan y los bloques separados por filas vacías
  se escriben como tablas independientes.
- La primera fila de un bloque se marca como encabezado solo si tiene al
  menos dos celdas y todas son texto.
- La orientación es vertical salvo que la tabla sea demasiado ancha.

Uso:
  python xlsx2rtf.py informe.xlsx
  python xlsx2rtf.py carpeta_entrada --salida carpeta_rtf
  python xlsx2rtf.py carpeta_entrada --salida carpeta_rtf --vigilar
  python xlsx2rtf.py carpeta_entrada --horizontal --vigilar
"""

import argparse
import datetime as dt
import logging
import os
import re
import sys
import time
from pathlib import Path

import openpyxl

EXTENSIONES = {".xlsx", ".xlsm"}

# Medidas en twips (1 pulgada = 1440 twips)
PAGINA_ANCHO = 12240   # carta, vertical
PAGINA_ALTO = 15840
MARGEN = 1440
UTIL_VERTICAL = PAGINA_ANCHO - 2 * MARGEN
UTIL_HORIZONTAL = PAGINA_ALTO - 2 * MARGEN
RELLENO_CELDA = 108
TWIPS_POR_CARACTER = 130   # margen de seguridad para Arial 10 pt (negrita incluida)
ANCHO_MIN_COLUMNA = 600
ANCHO_MAX_COLUMNA = 5000
UMBRAL_COMPRESION = 0.6    # por debajo de esto se avisa al usuario

log = logging.getLogger("xlsx2rtf")

# Marcador numérico de un formato de Excel: 0, #,##0, 0.00, #,##0.000 ...
_RE_NUMERO = re.compile(r"[0#][0#,]*(?:\.[0#]+)?")


# ---------------------------------------------------------------------------
# Formato de celdas
# ---------------------------------------------------------------------------

def _limpiar_seccion(seccion: str) -> str:
    seccion = re.sub(r"\[[^\]]*\]", "", seccion)   # [Red], [$€-409]...
    seccion = re.sub(r"_.", " ", seccion)          # _) deja un espacio
    seccion = re.sub(r"\\(.)", r"\1", seccion)      # \x -> x
    return seccion


def _formatear_numero(valor, formato: str) -> str:
    fmt = (formato or "General").strip()
    secciones = fmt.split(";")
    if valor < 0 and len(secciones) > 1:
        seccion, valor = secciones[1], -valor   # formato para negativos
    else:
        seccion = secciones[0]
    seccion = _limpiar_seccion(seccion)

    if fmt.lower() == "general" or not _RE_NUMERO.search(seccion):
        if isinstance(valor, float):
            if valor.is_integer() and abs(valor) < 1e15:
                return str(int(valor))
            return f"{valor:.10g}"
        return str(valor)

    if "%" in seccion:
        valor = valor * 100
    m = _RE_NUMERO.search(seccion)
    numero = m.group(0)
    decimales = len(numero.split(".")[1]) if "." in numero else 0
    texto = f"{valor:,.{decimales}f}" if "," in numero else f"{valor:.{decimales}f}"
    antes = seccion[:m.start()].replace('"', "")
    despues = seccion[m.end():].replace('"', "")
    return f"{antes}{texto}{despues}"


def formatear(valor, formato: str = "General") -> str:
    """Convierte una celda de Excel a texto tal como se vería en Excel."""
    if valor is None:
        return ""
    if isinstance(valor, bool):
        return "Sí" if valor else "No"
    if isinstance(valor, dt.datetime):
        if valor.time() == dt.time(0, 0):
            return valor.strftime("%d/%m/%Y")
        return valor.strftime("%d/%m/%Y %H:%M")
    if isinstance(valor, dt.date):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, dt.time):
        return valor.strftime("%H:%M")
    if isinstance(valor, (int, float)):
        return _formatear_numero(valor, formato)
    return str(valor)


# La fuente Symbol dibuja letras latinas como griegas (b -> β, f -> φ...).
_SIMBOLO = dict(zip(
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "αβχδεφγηιϕκλμνοπθρστυϖωξψζΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ",
))


def _texto_celda(celda, formato: str) -> str:
    texto = formatear(celda.value, formato)
    fuente = getattr(getattr(celda, "font", None), "name", None)
    if fuente == "Symbol" and isinstance(celda.value, str):
        texto = "".join(_SIMBOLO.get(ch, ch) for ch in texto)
    return texto


def _vacia(fila) -> bool:
    return not any(c.strip() for c in fila)


_RE_RANGO = re.compile(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?")


def _area_de_impresion(ws):
    """
    Devuelve (min_col, min_row, max_col, max_row) del área de impresión de la
    hoja, o None si no tiene. Si hay varias áreas, usa el rectángulo que las abarca.
    """
    texto = ws.print_area
    if not texto:
        return None
    if isinstance(texto, (list, tuple)):
        texto = ",".join(texto)
    from openpyxl.utils import column_index_from_string as col_idx
    cols, filas = [], []
    for ref in texto.split(","):
        for m in _RE_RANGO.finditer(ref.split("!")[-1]):
            c1, r1 = col_idx(m.group(1)), int(m.group(2))
            c2 = col_idx(m.group(3)) if m.group(3) else c1
            r2 = int(m.group(4)) if m.group(4) else r1
            cols += [c1, c2]
            filas += [r1, r2]
    if not cols:
        return None
    return min(cols), min(filas), max(cols), max(filas)


def _parece_numero(texto: str) -> bool:
    try:
        float(texto.replace(",", ""))
        return True
    except ValueError:
        return False


# ---------------------------------------------------------------------------
# Lectura del Excel
# ---------------------------------------------------------------------------

def leer_hojas(ruta: Path):
    """Devuelve una lista de (titulo_hoja, filas) con el texto ya formateado."""
    libro = openpyxl.load_workbook(ruta, data_only=True)
    try:
        resultado = []
        for ws in libro.worksheets:
            if getattr(ws, "sheet_state", "visible") != "visible":
                continue
            limites = _area_de_impresion(ws)
            if limites:
                min_c, min_r, max_c, max_r = limites
                celdas = ws.iter_rows(min_row=min_r, max_row=max_r,
                                      min_col=min_c, max_col=max_c)
            else:
                celdas = ws.iter_rows()
            filas = [
                [_texto_celda(c, getattr(c, "number_format", "General")) for c in fila]
                for fila in celdas
            ]
            while filas and _vacia(filas[0]):
                filas.pop(0)
            while filas and _vacia(filas[-1]):
                filas.pop()
            if not filas:
                continue
            ncols = 0
            for fila in filas:
                for i, texto in enumerate(fila):
                    if texto.strip():
                        ncols = max(ncols, i + 1)
            filas = [(f + [""] * ncols)[:ncols] for f in filas]
            resultado.append((ws.title, filas))
        return resultado
    finally:
        libro.close()


# ---------------------------------------------------------------------------
# Organización en bloques y columnas
# ---------------------------------------------------------------------------

def _segmentos(filas):
    """Separa las filas en ('tabla', filas) y ('espacio', número_de_filas_vacías)."""
    bloques, actual, vacias = [], [], 0
    for fila in filas:
        if _vacia(fila):
            if actual:
                bloques.append(("tabla", actual))
                actual = []
            vacias += 1
        else:
            if vacias:
                bloques.append(("espacio", vacias))
                vacias = 0
            actual.append(fila)
    if actual:
        bloques.append(("tabla", actual))
    return bloques


def _columnas_usadas(filas):
    ncols = len(filas[0])
    return [j for j in range(ncols) if any(f[j].strip() for f in filas)]


def _es_encabezado(fila) -> bool:
    llenas = [c for c in fila if c.strip()]
    return len(llenas) >= 2 and not any(_parece_numero(c) for c in llenas)


def _posiciones(proy, util):
    """Calcula las posiciones acumuladas (cellx) de cada columna, en twips."""
    n = len(proy[0])
    anchos = []
    for j in range(n):
        largo = max((len(linea) for f in proy for linea in f[j].split("\n")), default=0)
        ancho = largo * TWIPS_POR_CARACTER + 2 * RELLENO_CELDA + 120  # 120: holgura
        anchos.append(max(ANCHO_MIN_COLUMNA, min(ancho, ANCHO_MAX_COLUMNA)))
    natural = sum(anchos)
    factor = min(1.0, util / natural)
    anchos = [max(1, int(a * factor)) for a in anchos]
    acumulado, posiciones = 0, []
    for a in anchos:
        acumulado += a
        posiciones.append(acumulado)
    return posiciones, natural


def _planificar(hojas, util):
    planes = []
    for titulo, filas in hojas:
        bloques = []
        for tipo, datos in _segmentos(filas):
            if tipo == "espacio":
                bloques.append(("espacio", datos))
                continue
            usadas = _columnas_usadas(datos)
            proy = [[f[j] for j in usadas] for f in datos]
            posiciones, natural = _posiciones(proy, util)
            bloques.append(("tabla", proy, posiciones, _es_encabezado(datos[0]), natural))
        planes.append((titulo, bloques))
    return planes


def _ancho_natural_max(planes) -> int:
    return max((b[4] for _, bloques in planes for b in bloques if b[0] == "tabla"),
               default=0)


# ---------------------------------------------------------------------------
# Generación de RTF
# ---------------------------------------------------------------------------

def escapar(texto: str) -> str:
    """Escapa texto para RTF. Todo lo no ASCII se codifica como \\uN."""
    partes = []
    for ch in texto:
        codigo = ord(ch)
        if ch == "\\":
            partes.append("\\\\")
        elif ch == "{":
            partes.append("\\{")
        elif ch == "}":
            partes.append("\\}")
        elif ch == "\n":
            partes.append("\\line ")
        elif ch == "\t":
            partes.append("\\tab ")
        elif codigo < 32 or ch == "\r":
            continue
        elif codigo < 128:
            partes.append(ch)
        else:
            # RTF usa enteros con signo de 16 bits (UTF-16); \uc1 indica
            # que el carácter de reemplazo '?' va después de cada escape.
            if codigo > 0xFFFF:
                c = codigo - 0x10000
                unidades = [0xD800 + (c >> 10), 0xDC00 + (c & 0x3FF)]
            else:
                unidades = [codigo]
            for u in unidades:
                if u > 32767:
                    u -= 65536
                partes.append(f"\\u{u}?")
    return "".join(partes)


def _fila_rtf(celdas, posiciones, encabezado: bool) -> str:
    bordes = "\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs"
    sombra = "\\clcbpat1" if encabezado else ""
    out = [f"\\trowd\\trgaph{RELLENO_CELDA}\\trleft0"]
    if encabezado:
        out.append("\\trhdr")  # repetir el encabezado en cada página
    for pos in posiciones:
        out.append(f"{bordes}{sombra}\\cellx{pos}")
    for texto in celdas:
        contenido = escapar(texto)
        if encabezado and contenido:
            contenido = f"\\b {contenido}\\b0"
        out.append(f"\\pard\\intbl {contenido}\\cell")
    out.append("\\row")
    return "".join(out)


def generar_rtf(hojas, horizontal=None) -> str:
    """
    horizontal=None: la orientación se elige sola (horizontal solo si una
    tabla no cabe en vertical). True/False fuerzan la orientación.
    """
    planes = _planificar(hojas, UTIL_VERTICAL)
    if horizontal is None:
        horizontal = _ancho_natural_max(planes) > UTIL_VERTICAL

    if horizontal:
        util = UTIL_HORIZONTAL
        planes = _planificar(hojas, util)
        ancho_pag, alto_pag, orient = PAGINA_ALTO, PAGINA_ANCHO, "\\landscape"
    else:
        util = UTIL_VERTICAL
        ancho_pag, alto_pag, orient = PAGINA_ANCHO, PAGINA_ALTO, ""

    partes = [
        "{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1",
        "{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}}",
        "{\\colortbl;\\red221\\green230\\blue240;}",  # color de encabezado
        f"\\paperw{ancho_pag}\\paperh{alto_pag}{orient}"
        f"\\margl{MARGEN}\\margr{MARGEN}\\margt{MARGEN}\\margb{MARGEN}",
        "\\viewkind4\\pard\\plain\\f0\\fs20",
    ]

    for titulo, bloques in planes:
        partes.append(
            f"\\pard\\keepn\\sb200\\sa120\\b\\fs28 {escapar(titulo)}\\b0\\fs20\\par"
        )
        for b in bloques:
            if b[0] == "espacio":
                alto = min(b[1], 3) * 160
                partes.append(f"\\pard\\sa{alto}\\par")
                continue
            _, proy, posiciones, enc, natural = b
            if natural > util / UMBRAL_COMPRESION:
                log.warning("Hoja '%s': una tabla es demasiado ancha y sus columnas "
                            "se redujeron; conviene dividirla en partes.", titulo)
            for i, fila in enumerate(proy):
                partes.append(_fila_rtf(fila, posiciones, encabezado=(enc and i == 0)))
            partes.append("\\pard\\par")

    partes.append("}")
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# Conversión de archivos
# ---------------------------------------------------------------------------

def convertir(origen: Path, carpeta_salida: Path, horizontal=None) -> Path:
    hojas = leer_hojas(origen)
    if not hojas:
        raise ValueError("el libro no tiene hojas visibles con datos")
    rtf = generar_rtf(hojas, horizontal)

    destino = carpeta_salida / (origen.stem + ".rtf")
    tmp = destino.with_name(destino.name + ".tmp")
    # El RTF generado es ASCII puro (todo lo demás se escapó).
    tmp.write_text(rtf, encoding="ascii", newline="\n")
    os.replace(tmp, destino)  # escritura atómica: nunca queda un RTF a medias
    return destino


def procesar_carpeta(entrada: Path, salida: Path, horizontal, estado: dict) -> int:
    """Convierte los archivos nuevos o modificados de 'entrada'. Devuelve cuántos."""
    convertidos = 0
    for ruta in sorted(entrada.iterdir()):
        if not ruta.is_file() or ruta.suffix.lower() not in EXTENSIONES:
            continue
        if ruta.name.startswith("~$"):  # archivo de bloqueo de Excel
            continue
        marca = ruta.stat().st_mtime_ns
        if estado.get(ruta) == marca:
            continue
        try:
            destino = convertir(ruta, salida, horizontal)
            log.info("%s -> %s", ruta.name, destino.name)
            convertidos += 1
        except Exception as exc:  # p. ej. archivo corrupto o aún copiándose
            log.error("No se pudo convertir %s: %s", ruta.name, exc)
        # Se recuerda la marca aunque haya fallado: se reintentará
        # automáticamente cuando el archivo cambie de nuevo.
        estado[ruta] = marca
    return convertidos


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Convierte archivos Excel (.xlsx/.xlsm) a RTF, de forma automática."
    )
    ap.add_argument("fuente", type=Path,
                    help="archivo de Excel o carpeta a procesar")
    ap.add_argument("--salida", type=Path, default=None,
                    help="carpeta donde guardar los RTF (por defecto, la misma fuente)")
    ap.add_argument("--vigilar", action="store_true",
                    help="seguir ejecutándose y convertir archivos nuevos o modificados (solo carpetas)")
    ap.add_argument("--intervalo", type=float, default=5.0,
                    help="segundos entre revisiones en modo --vigilar (por defecto 5)")
    ap.add_argument("--horizontal", action="store_true",
                    help="forzar páginas en orientación horizontal (por defecto se elige sola)")
    args = ap.parse_args(argv)
    horizontal = True if args.horizontal else None

    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")

    fuente = args.fuente.expanduser().resolve()

    if fuente.is_file():
        if fuente.suffix.lower() not in EXTENSIONES:
            log.error("Formato no soportado: %s (use .xlsx o .xlsm)", fuente.suffix)
            return 1
        salida = (args.salida or fuente.parent).expanduser().resolve()
        salida.mkdir(parents=True, exist_ok=True)
        try:
            destino = convertir(fuente, salida, horizontal)
        except Exception as exc:
            log.error("No se pudo convertir %s: %s", fuente.name, exc)
            return 1
        log.info("Listo: %s", destino)
        return 0

    if not fuente.is_dir():
        log.error("No existe: %s", fuente)
        return 1

    salida = (args.salida or fuente).expanduser().resolve()
    salida.mkdir(parents=True, exist_ok=True)
    estado: dict = {}

    if not args.vigilar:
        n = procesar_carpeta(fuente, salida, horizontal, estado)
        log.info("Archivos convertidos: %d", n)
        return 0

    log.info("Vigilando %s -> %s (Ctrl+C para salir)", fuente, salida)
    try:
        while True:
            procesar_carpeta(fuente, salida, horizontal, estado)
            time.sleep(args.intervalo)
    except KeyboardInterrupt:
        log.info("Detenido.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
"""
xlsx2rtf: convierte automáticamente archivos de Excel (.xlsx / .xlsm) a RTF.

- Cada hoja visible del libro se escribe como una tabla con bordes.
- La primera fila de cada hoja se trata como encabezado (negrita y sombreado).
- Puede convertir un archivo, una carpeta completa, o quedarse vigilando una
  carpeta y convertir cada archivo nuevo o modificado que aparezca en ella.

Uso:
  python xlsx2rtf.py informe.xlsx
  python xlsx2rtf.py carpeta_entrada --salida carpeta_rtf
  python xlsx2rtf.py carpeta_entrada --salida carpeta_rtf --vigilar
  python xlsx2rtf.py carpeta_entrada --vigilar --horizontal --intervalo 3
"""

import argparse
import datetime as dt
import logging
import os
import sys
import time
from pathlib import Path

import openpyxl

EXTENSIONES = {".xlsx", ".xlsm"}

# Medidas en twips (1 pulgada = 1440 twips)
PAGINA_ANCHO = 12240   # carta, vertical
PAGINA_ALTO = 15840
MARGEN = 1440          # 1 pulgada
RELLENO_CELDA = 108    # margen interno de cada celda
TWIPS_POR_CARACTER = 110  # aprox. para Arial 10 pt
ANCHO_MAX_COLUMNA = 5000

log = logging.getLogger("xlsx2rtf")


# ---------------------------------------------------------------------------
# Lectura del Excel
# ---------------------------------------------------------------------------

def formatear(valor) -> str:
    """Convierte una celda de Excel a texto legible."""
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
    if isinstance(valor, float):
        if valor.is_integer() and abs(valor) < 1e15:
            return str(int(valor))
        return f"{valor:.10g}"
    return str(valor)


def _vacia(fila) -> bool:
    return not any(c.strip() for c in fila)


def leer_hojas(ruta: Path):
    """Devuelve una lista de (titulo_hoja, filas, numero_columnas)."""
    libro = openpyxl.load_workbook(ruta, data_only=True)
    try:
        resultado = []
        for ws in libro.worksheets:
            if getattr(ws, "sheet_state", "visible") != "visible":
                continue
            filas = [
                [formatear(c) for c in fila]
                for fila in ws.iter_rows(values_only=True)
            ]
            # quitar filas vacías al inicio y al final
            while filas and _vacia(filas[0]):
                filas.pop(0)
            while filas and _vacia(filas[-1]):
                filas.pop()
            if not filas:
                continue
            # número de columnas útiles (hasta la última con contenido)
            ncols = 0
            for fila in filas:
                for i, c in enumerate(fila):
                    if c.strip():
                        ncols = max(ncols, i + 1)
            filas = [(f + [""] * ncols)[:ncols] for f in filas]
            resultado.append((ws.title, filas, ncols))
        return resultado
    finally:
        libro.close()


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
            continue  # caracteres de control: se descartan
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


def _anchos_columnas(filas, ncols, ancho_util):
    """Calcula posiciones acumuladas (cellx) de cada columna, en twips."""
    anchos = []
    for j in range(ncols):
        largo = 0
        for fila in filas:
            for linea in fila[j].split("\n"):
                largo = max(largo, len(linea))
        largo = max(largo, 4)
        anchos.append(min(largo * TWIPS_POR_CARACTER + 2 * RELLENO_CELDA,
                          ANCHO_MAX_COLUMNA))
    total = sum(anchos)
    if total > ancho_util:  # encajar en la página
        anchos = [int(a * ancho_util / total) for a in anchos]
    acumulado, posiciones = 0, []
    for a in anchos:
        acumulado += a
        posiciones.append(acumulado)
    return posiciones


def _fila_rtf(celdas, posiciones, encabezado: bool) -> str:
    bordes = "\\clbrdrt\\brdrs\\clbrdrl\\brdrs\\clbrdrb\\brdrs\\clbrdrr\\brdrs"
    sombra = "\\clcbpat1" if encabezado else ""
    out = [f"\\trowd\\trgaph{RELLENO_CELDA}\\trleft0"]
    if encabezado:
        out.append("\\trhdr")  # repetir encabezado en cada página
    for pos in posiciones:
        out.append(f"{bordes}{sombra}\\cellx{pos}")
    for texto in celdas:
        contenido = escapar(texto)
        if encabezado and contenido:
            contenido = f"\\b {contenido}\\b0"
        out.append(f"\\pard\\intbl {contenido}\\cell")
    out.append("\\row")
    return "".join(out)


def generar_rtf(hojas, horizontal: bool = False) -> str:
    if horizontal:
        ancho_pag, alto_pag, orient = PAGINA_ALTO, PAGINA_ANCHO, "\\landscape"
    else:
        ancho_pag, alto_pag, orient = PAGINA_ANCHO, PAGINA_ALTO, ""
    ancho_util = ancho_pag - 2 * MARGEN

    partes = [
        "{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1",
        "{\\fonttbl{\\f0\\fswiss\\fcharset0 Arial;}}",
        "{\\colortbl;\\red221\\green230\\blue240;}",  # color de encabezado
        f"\\paperw{ancho_pag}\\paperh{alto_pag}{orient}"
        f"\\margl{MARGEN}\\margr{MARGEN}\\margt{MARGEN}\\margb{MARGEN}",
        "\\viewkind4\\pard\\plain\\f0\\fs20",
    ]

    for titulo, filas, ncols in hojas:
        partes.append(
            f"\\pard\\keepn\\sb200\\sa120\\b\\fs28 {escapar(titulo)}\\b0\\fs20\\par"
        )
        posiciones = _anchos_columnas(filas, ncols, ancho_util)
        for i, fila in enumerate(filas):
            partes.append(_fila_rtf(fila, posiciones, encabezado=(i == 0)))
        partes.append("\\pard\\par")

    partes.append("}")
    return "\n".join(partes)


# ---------------------------------------------------------------------------
# Conversión de archivos
# ---------------------------------------------------------------------------

def convertir(origen: Path, carpeta_salida: Path, horizontal: bool = False) -> Path:
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


def procesar_carpeta(entrada: Path, salida: Path, horizontal: bool, estado: dict) -> int:
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
                    help="páginas en orientación horizontal")
    args = ap.parse_args(argv)

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
            destino = convertir(fuente, salida, args.horizontal)
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
        n = procesar_carpeta(fuente, salida, args.horizontal, estado)
        log.info("Archivos convertidos: %d", n)
        return 0

    log.info("Vigilando %s -> %s (Ctrl+C para salir)", fuente, salida)
    try:
        while True:
            procesar_carpeta(fuente, salida, args.horizontal, estado)
            time.sleep(args.intervalo)
    except KeyboardInterrupt:
        log.info("Detenido.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

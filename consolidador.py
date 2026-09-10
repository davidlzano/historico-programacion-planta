"""
Consolidador de programación de planta.

Lee todos los respaldos diarios y produce una base única con dos capas:

    HISTORICO   Una fila cada vez que un registro nace o cambia. Para saber
                cómo estaba algo en una fecha, se toma su última fila con
                snapshot menor o igual a esa fecha.
    VIGENTE     El estado actual de cada registro. Es la tabla para consultar
                y relacionar.

Y una tercera hoja, AVISOS, que es la que hay que revisar: lista las hojas que
no se pudieron leer y los archivos con problemas. Un consolidado silencioso que
perdió tres días de datos es peor que uno que falla ruidosamente.

Uso:
    python generar_snapshots.py    # crea los respaldos de ejemplo
    python consolidador.py         # produce BD_PROGRAMACION.xlsx
"""

import re
import warnings
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore")

DIR_SNAPSHOTS = Path("snapshots")
ARCHIVO_SALIDA = Path("BD_PROGRAMACION.xlsx")

# 'CAMBIOS'  guarda una fila solo cuando el registro nace o cambia
# 'COMPLETO' guarda la foto entera de cada día
MODO = "CAMBIOS"

# -----------------------------------------------------------------------------
# MAPEO DE COLUMNAS
# -----------------------------------------------------------------------------
# El archivo origen lo edita a mano el área de producción todos los días. Las
# columnas se mueven, se renombran y aparecen nuevas. Leer por posición
# funciona hasta el primer día en que alguien inserta una columna, y a partir
# de ahí carga datos en el campo equivocado sin error visible.
#
# Por eso el mapeo es por nombre de encabezado, con una lista de sinónimos por
# campo. Un encabezado desconocido se ignora en lugar de romper la carga.

SINONIMOS = {
    "op": ["# OP", "OP", "N OP", "NO OP"],
    "cliente": ["CLIENTE"],
    "referencia": ["REFERENCIA", "REFER", "CODIGO"],
    "descripcion": ["DESCRIPCION", "DESCRIPCIÓN"],
    "cantidad": ["TIROS", "CANTIDAD", "CANTIDAD PEDIDA"],
    "inicio": ["DIA Y HORA INICIO"],
    "fin": ["DIA Y HORA FIN"],
    "alistamiento": ["TIEMPO ALISTAMIENTO", "TIEMPO ALISTO"],
    "observaciones": ["OBSERVACIONES", "OBS", "OBSERV"],
    "medida": ["MEDIDA"],
}

INVERSO = {alias: canonico for canonico, alias in
           ((c, a) for c, alts in SINONIMOS.items() for a in alts)}

# Campos cuyo cambio significa que el registro cambió. Los de trazabilidad
# (snapshot, archivo) quedan fuera a propósito: si estuvieran, cada día
# generaría una versión nueva de todo y el modo de cambios no serviría.
CAMPOS_ESTADO = [
    "op", "cliente", "referencia", "descripcion",
    "cantidad", "inicio", "fin", "alistamiento", "observaciones",
]

# Filas que no son producción sino eventos de calendario o de máquina.
PATRON_EVENTO = re.compile(
    r"MANTENIMIENTO|LAVAD|INICIO DE TURNO|FIN DE TURNO|DOMINGO|FESTIVO|"
    r"PARO|DAÑO|CALIBRA|SIN PROGRAMA",
    re.IGNORECASE,
)


def normalizar(valor) -> str:
    if valor is None:
        return ""
    return " ".join(str(valor).replace("\n", " ").split()).upper().strip()


def fecha_de_archivo(ruta: Path) -> str | None:
    """Extrae la fecha del nombre PROGRAMACION DDMMAAAA.xlsx."""
    m = re.search(r"(\d{2})(\d{2})(\d{4})", ruta.stem)
    if not m:
        return None
    dia, mes, anio = m.groups()
    try:
        return datetime(int(anio), int(mes), int(dia)).strftime("%Y-%m-%d")
    except ValueError:
        return None


def buscar_encabezado(ws, max_filas: int = 12) -> tuple[int, dict] | tuple[None, None]:
    """
    Localiza la fila de encabezados y mapea columna -> campo canónico.

    El encabezado no siempre está en la fila 1: a veces hay títulos, logos o
    filas en blanco encima. Se busca la primera fila que contenga al menos dos
    encabezados reconocibles.
    """
    for fila in range(1, max_filas + 1):
        mapeo = {}
        for col in range(1, ws.max_column + 1):
            texto = normalizar(ws.cell(fila, col).value)
            if texto in INVERSO:
                mapeo[INVERSO[texto]] = col
        if len(mapeo) >= 2:
            return fila, mapeo
    return None, None


def extraer(ruta: Path) -> tuple[list[dict], list[dict], set[str]]:
    """
    Lee un respaldo y devuelve (registros, avisos, hojas_leidas).

    hojas_leidas importa tanto como los registros: distingue una hoja que se
    leyó y no tenía cierta fila, de una hoja que no se pudo abrir. Sin esa
    distinción, un archivo con una hoja corrupta borraría del estado vigente
    toda la programación de esa máquina.
    """
    fecha = fecha_de_archivo(ruta)
    if not fecha:
        return [], [{"archivo": ruta.name, "hoja": "", "motivo": "Nombre sin fecha reconocible"}], set()

    registros, avisos, leidas = [], [], set()

    try:
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    except Exception as e:
        return [], [{"archivo": ruta.name, "hoja": "", "motivo": f"No se pudo abrir: {e}"}], set()

    try:
        for hoja in wb.sheetnames:
            ws = wb[hoja]
            fila_enc, mapeo = buscar_encabezado(ws)

            if not mapeo:
                avisos.append({
                    "archivo": ruta.name,
                    "hoja": hoja,
                    "motivo": "Sin encabezados reconocibles",
                })
                continue

            leidas.add(hoja)

            for idx, fila in enumerate(
                ws.iter_rows(min_row=fila_enc + 1, values_only=True), start=fila_enc + 1
            ):
                registro = {
                    campo: normalizar(fila[col - 1]) if col - 1 < len(fila) else ""
                    for campo, col in mapeo.items()
                }

                if not any(registro.values()):
                    continue

                descripcion = registro.get("descripcion", "")
                registro["es_evento"] = bool(PATRON_EVENTO.search(descripcion))

                # La llave identifica la posición del tramo en la hoja. Es lo
                # que permite seguir un mismo renglón a través de los días.
                registro["llave"] = f"{hoja}|{idx}"
                registro["maquina"] = hoja
                registro["fila_origen"] = idx
                registro["snapshot"] = fecha
                registro["archivo"] = ruta.name
                registros.append(registro)
    finally:
        wb.close()

    return registros, avisos, leidas


def solo_cambios(registros: list[dict]) -> list[dict]:
    """
    Conserva una fila solo cuando el registro nace o cambia.

    El histórico resultante es equivalente al completo: para reconstruir cómo
    estaba un registro en cualquier fecha, se toma su última versión con
    snapshot menor o igual a esa fecha. La diferencia es el espacio.
    """
    ultimo: dict[str, tuple] = {}
    salida = []
    for r in sorted(registros, key=lambda x: (x["snapshot"], x["llave"])):
        estado = tuple(str(r.get(c, "")) for c in CAMPOS_ESTADO)
        if ultimo.get(r["llave"]) != estado:
            ultimo[r["llave"]] = estado
            salida.append(r)
    return salida


def reconstruir_vigente(registros: list[dict], leidas_por_dia: dict[str, set]) -> list[dict]:
    """
    Estado actual de cada registro, avanzando día por día.

    La sutileza está en las desapariciones. Si una llave deja de aparecer en un
    día, puede ser porque borraron la fila, o porque esa hoja no se pudo leer.
    Solo se purgan las llaves de hojas que **sí** se leyeron ese día; las que
    faltaron se dejan intactas.
    """
    por_dia = defaultdict(list)
    for r in registros:
        por_dia[r["snapshot"]].append(r)

    vigente: dict[str, dict] = {}

    for dia in sorted(por_dia):
        presentes = {r["llave"] for r in por_dia[dia]}
        hojas_leidas = leidas_por_dia.get(dia, set())

        # Purga solo dentro de las hojas efectivamente leídas
        for llave in list(vigente):
            hoja = llave.split("|")[0]
            if hoja in hojas_leidas and llave not in presentes:
                del vigente[llave]

        for r in por_dia[dia]:
            vigente[r["llave"]] = r

    return list(vigente.values())


def escribir(historico: list[dict], vigente: list[dict], avisos: list[dict]) -> None:
    columnas = [
        "snapshot", "maquina", "op", "cliente", "referencia", "descripcion",
        "cantidad", "inicio", "fin", "alistamiento", "observaciones",
        "es_evento", "fila_origen", "archivo",
    ]

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "VIGENTE"
    ws.append(columnas)
    for r in sorted(vigente, key=lambda x: (x["maquina"], x["fila_origen"])):
        ws.append([r.get(c, "") for c in columnas])

    ws = wb.create_sheet("HISTORICO")
    ws.append(columnas)
    for r in sorted(historico, key=lambda x: (x["snapshot"], x["llave"])):
        ws.append([r.get(c, "") for c in columnas])

    ws = wb.create_sheet("AVISOS")
    ws.append(["archivo", "hoja", "motivo"])
    for a in avisos:
        ws.append([a["archivo"], a["hoja"], a["motivo"]])

    for hoja in wb.worksheets:
        hoja.freeze_panes = "A2"

    wb.save(ARCHIVO_SALIDA)


def main() -> None:
    archivos = sorted(DIR_SNAPSHOTS.glob("*.xlsx"))
    if not archivos:
        raise SystemExit(
            f"No hay respaldos en {DIR_SNAPSHOTS}. Ejecuta primero: python generar_snapshots.py"
        )

    print(f"Procesando {len(archivos)} respaldos...")

    todos, avisos = [], []
    leidas_por_dia: dict[str, set] = {}

    for ruta in archivos:
        registros, avs, leidas = extraer(ruta)
        todos.extend(registros)
        avisos.extend(avs)
        fecha = fecha_de_archivo(ruta)
        if fecha:
            leidas_por_dia[fecha] = leidas

    if not todos:
        raise SystemExit("No se extrajo ningun registro. Revisa los avisos.")

    completo = len(todos)
    historico = solo_cambios(todos) if MODO == "CAMBIOS" else todos
    vigente = reconstruir_vigente(todos, leidas_por_dia)

    escribir(historico, vigente, avisos)

    reduccion = (1 - len(historico) / completo) * 100 if completo else 0
    eventos = sum(1 for r in vigente if r.get("es_evento"))

    print(f"\n{'Filas leidas':<24}{completo:>9}")
    print(f"{'Historico (' + MODO + ')':<24}{len(historico):>9}  ({reduccion:.0f}% menos)")
    print(f"{'Registros vigentes':<24}{len(vigente):>9}  ({eventos} son eventos, no produccion)")
    print(f"{'Avisos de lectura':<24}{len(avisos):>9}")

    if avisos:
        print("\nHojas que no se pudieron leer:")
        for a in avisos[:5]:
            print(f"  {a['archivo']} / {a['hoja']}: {a['motivo']}")
        if len(avisos) > 5:
            print(f"  ... y {len(avisos) - 5} mas (ver hoja AVISOS)")

    print(f"\nSalida: {ARCHIVO_SALIDA.resolve()}")


if __name__ == "__main__":
    main()

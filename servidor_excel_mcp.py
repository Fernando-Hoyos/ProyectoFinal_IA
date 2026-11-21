from mcp.server.mcpserver import MCPServer
from openpyxl import Workbook, load_workbook
import os
import re

mcp = MCPServer("ExcelServer")


# ---------------------------------------------------------------
# FUNCIONES AUXILIARES (no son herramientas, el modelo no las ve)
# ---------------------------------------------------------------

def _normalizar_ruta(ruta: str) -> str:
    """Agrega la extensión .xlsx si falta."""
    if not ruta.lower().endswith(".xlsx"):
        ruta += ".xlsx"
    return ruta


def _abrir_libro(ruta: str):
    """Abre un libro existente o lanza un error claro."""
    ruta = _normalizar_ruta(ruta)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"El archivo '{ruta}' no existe.")
    return load_workbook(ruta), ruta


def _obtener_hoja(libro, nombre_hoja: str):
    """Devuelve la hoja pedida, o la hoja activa si nombre_hoja está vacío."""
    if not nombre_hoja:
        return libro.active
    if nombre_hoja not in libro.sheetnames:
        raise ValueError(
            f"La hoja '{nombre_hoja}' no existe. Hojas disponibles: {libro.sheetnames}"
        )
    return libro[nombre_hoja]


def _convertir_valor(valor: str):
    """Convierte '25' en 25 y '3.5' en 3.5 para que Excel los trate como números."""
    texto = str(valor).strip()
    try:
        return int(texto)
    except ValueError:
        pass
    try:
        return float(texto.replace(",", "."))
    except ValueError:
        return valor


def _guardar(libro, ruta: str):
    """Guarda el libro, avisando si el archivo está abierto en Excel."""
    try:
        libro.save(ruta)
    except PermissionError:
        raise PermissionError(
            f"No se pudo guardar '{ruta}'. Ciérralo en Excel e inténtalo de nuevo."
        )


# ---------------------------------------------------------------
# HERRAMIENTAS MCP
# ---------------------------------------------------------------

@mcp.tool()
def crear_excel(ruta: str, nombre_hoja: str = "Hoja1") -> str:
    """Crea un archivo de Excel (.xlsx) nuevo y vacío con una hoja inicial.
    Usa esta herramienta antes de escribir datos en un Excel que todavía no existe."""
    try:
        ruta = _normalizar_ruta(ruta)
        if os.path.exists(ruta):
            return f"Error: el archivo '{ruta}' ya existe. Usa otro nombre o escribe en él."
        libro = Workbook()
        libro.active.title = nombre_hoja
        _guardar(libro, ruta)
        return f"Excel creado correctamente en: {os.path.abspath(ruta)}"
    except Exception as e:
        return f"Error al crear el Excel: {e}"


@mcp.tool()
def listar_hojas(ruta: str) -> list[str]:
    """Lista los nombres de las hojas de un archivo de Excel (.xlsx)."""
    try:
        libro, _ = _abrir_libro(ruta)
        return libro.sheetnames
    except Exception as e:
        return [f"Error: {e}"]


@mcp.tool()
def crear_hoja(ruta: str, nombre_hoja: str) -> str:
    """Agrega una hoja nueva a un archivo de Excel existente."""
    try:
        libro, ruta = _abrir_libro(ruta)
        if nombre_hoja in libro.sheetnames:
            return f"Error: la hoja '{nombre_hoja}' ya existe."
        libro.create_sheet(nombre_hoja)
        _guardar(libro, ruta)
        return f"Hoja '{nombre_hoja}' creada en {ruta}."
    except Exception as e:
        return f"Error al crear la hoja: {e}"


@mcp.tool()
def leer_excel(ruta: str, nombre_hoja: str = "") -> str:
    """Lee todo el contenido de una hoja de Excel y lo devuelve fila por fila.
    Si no se indica nombre_hoja, lee la hoja activa."""
    try:
        libro, ruta = _abrir_libro(ruta)
        hoja = _obtener_hoja(libro, nombre_hoja)
        filas = []
        for numero, fila in enumerate(hoja.iter_rows(values_only=True), start=1):
            valores = ["" if v is None else str(v) for v in fila]
            if any(valores):
                filas.append(f"Fila {numero}: " + " | ".join(valores))
        if not filas:
            return f"La hoja '{hoja.title}' de {ruta} está vacía."
        return f"Contenido de la hoja '{hoja.title}':\n" + "\n".join(filas)
    except Exception as e:
        return f"Error al leer el Excel: {e}"


@mcp.tool()
def escribir_celda(ruta: str, celda: str, valor: str, nombre_hoja: str = "") -> str:
    """Escribe un valor en una celda específica de un Excel (por ejemplo 'B3').
    Si no se indica nombre_hoja, escribe en la hoja activa."""
    try:
        celda = celda.strip().upper()
        if not re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", celda):
            return f"Error: '{celda}' no es una celda válida. Usa el formato letra+número, por ejemplo 'A1'."
        libro, ruta = _abrir_libro(ruta)
        hoja = _obtener_hoja(libro, nombre_hoja)
        hoja[celda] = _convertir_valor(valor)
        _guardar(libro, ruta)
        return f"Valor '{valor}' escrito en {celda} de la hoja '{hoja.title}'."
    except Exception as e:
        return f"Error al escribir la celda: {e}"


@mcp.tool()
def agregar_filas(ruta: str, filas: list[list[str]], nombre_hoja: str = "") -> str:
    """Agrega una o varias filas al final de una hoja de Excel.
    Cada fila es una lista de valores, por ejemplo [["Nombre", "Edad"], ["Ana", "25"]].
    Úsala también para escribir los encabezados de una tabla."""
    try:
        libro, ruta = _abrir_libro(ruta)
        hoja = _obtener_hoja(libro, nombre_hoja)

        # Si la hoja está vacía, la primera fila se escribe en la fila 1
        # (al revisar la fila 1, openpyxl la marca como usada y append empezaría en la 2)
        hoja_vacia = hoja.max_row == 1 and all(c.value is None for c in hoja[1])
        for i, fila in enumerate(filas):
            valores = [_convertir_valor(v) for v in fila]
            if hoja_vacia and i == 0:
                for col, v in enumerate(valores, start=1):
                    hoja.cell(row=1, column=col, value=v)
            else:
                hoja.append(valores)

        _guardar(libro, ruta)
        return f"{len(filas)} fila(s) agregada(s) a la hoja '{hoja.title}' de {ruta}."
    except Exception as e:
        return f"Error al agregar filas: {e}"


if __name__ == "__main__":
    mcp.run()
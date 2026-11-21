import asyncio
import os
import sys
from google import genai
from google.genai import types
from mcp import ClientSession
from mcp.client.stdio import (stdio_client, StdioServerParameters)

# Evita que model_copy falle al intentar clonar las tareas asíncronas de la sesión MCP
ClientSession.__deepcopy__ = lambda self, memo: self

# Rutas fijas: los servidores se buscan junto a este archivo
# y los dos trabajan dentro de archivos_prueba
CARPETA_PROYECTO = os.path.dirname(os.path.abspath(__file__))
CARPETA_TRABAJO = os.path.join(CARPETA_PROYECTO, "archivos_prueba")
os.makedirs(CARPETA_TRABAJO, exist_ok=True)


# FUNCIÓN PRINCIPAL
async def generate():

    client = genai.Client(api_key="Remplazar Llave")  # Reemplaza con tu API Key de Google Gemini

    # Servidor 1: archivos y directorios
    params_archivos = StdioServerParameters(
        command=sys.executable,
        args=[os.path.join(CARPETA_PROYECTO, "mi_servidor_mcp.py")],
        cwd=CARPETA_TRABAJO,
    )

    # Servidor 2: Excel
    params_excel = StdioServerParameters(
        command=sys.executable,
        args=[os.path.join(CARPETA_PROYECTO, "servidor_excel_mcp.py")],
        cwd=CARPETA_TRABAJO,
    )

    # CONEXIÓN MCP: una conexión y una sesión por cada servidor
    async with stdio_client(params_archivos) as (read_a, write_a), \
               stdio_client(params_excel) as (read_e, write_e):

        async with ClientSession(read_a, write_a) as session_archivos, \
                   ClientSession(read_e, write_e) as session_excel:

            await session_archivos.initialize()  # Inicializar servidor de archivos
            await session_excel.initialize()     # Inicializar servidor de Excel

            # instrucciones del sistema
            instruciones_sistema = (
                "Eres un asistente técnico experto con dos grupos de herramientas MCP. "
                "Herramientas de archivos: listar, leer, crear y modificar archivos de texto (.txt) "
                "y consultar o cambiar el directorio. "
                "Herramientas de Excel: crear archivos .xlsx, listar y crear hojas, leer su contenido, "
                "escribir en celdas y agregar filas. "
                "Usa las herramientas siempre que el usuario quiera trabajar con archivos, carpetas "
                "o Excel, y combínalas cuando la tarea lo requiera; por ejemplo, lee archivos de texto "
                "con las herramientas de archivos y pasa su información a Excel con las herramientas de Excel. "
                "Al pasar datos a Excel, organízalos como tabla: primero una fila de encabezados "
                "y luego una fila por cada registro. "
                "No cambies de directorio a menos que el usuario lo pida, para que ambos grupos "
                "de herramientas trabajen en la misma carpeta."
            )

            # configuración del chat, se le pasan las dos sesiones MCP
            configuracion = types.GenerateContentConfig(
                system_instruction=instruciones_sistema,
                tools=[session_archivos, session_excel],
                thinking_config=types.ThinkingConfig(
                    thinking_level="low",
                    include_thoughts=False
                ),
                temperature=0.8,
                # Permite más llamadas seguidas a herramientas (por defecto son 10)
                automatic_function_calling=types.AutomaticFunctionCallingConfig(
                    maximum_remote_calls=25
                ),
            )
            # modelo a usar
            modelo = "gemini-3.1-flash-lite"
            chat = client.aio.chats.create(model=modelo, config=configuracion)

            print("====================================")
            print("   CHAT MCP + GEMINI (2 SERVIDORES)")
            print("====================================")
            print(f"Carpeta de trabajo: {CARPETA_TRABAJO}")
            print("Escribe 'salir' para terminar")

            while True:

                pregunta = input(" Tú: ")
                if pregunta.lower() in ["salir", "exit"]:
                    break
                print("\n Pensando...\n")

                try:
                    response = await chat.send_message(
                        pregunta
                    )
                    print(response.text)

                except Exception as e:

                    print("\n Error:")
                    print(str(e))


if __name__ == "__main__":
    asyncio.run(generate())

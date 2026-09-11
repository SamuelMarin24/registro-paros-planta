"""
Configuración centralizada de la aplicación.

Todos los valores sensibles (credenciales, servidores, rutas de red) se leen
desde variables de entorno, nunca se escriben en el código. Copia el archivo
`.env.example` a `.env` y completa los valores de tu entorno.
"""
import os
from dotenv import load_dotenv

load_dotenv()

# ── Conexión a la base de datos (fuente de las órdenes de producción) ────────
SQL_SERVER = os.getenv("SQL_SERVER", "localhost\\SQLEXPRESS")
SQL_DATABASE = os.getenv("SQL_DATABASE", "produccion")
SQL_USER = os.getenv("SQL_USER", "")
SQL_PASSWORD = os.getenv("SQL_PASSWORD", "")
SQL_DRIVER = os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server")
SQL_TABLA_ORDENES = os.getenv("SQL_TABLA_ORDENES", "ordenes_produccion")

# ── Rutas de los archivos maestros y del consolidado ─────────────────────────
# En producción apuntan a una carpeta compartida en red; en desarrollo pueden
# apuntar a la carpeta local `datos/`.
RUTA_MAESTRAS = os.getenv("RUTA_MAESTRAS", "datos/maestras.xlsx")
RUTA_RAZONES = os.getenv("RUTA_RAZONES", "datos/razones_de_paro.xlsx")
RUTA_CONSOLIDADO = os.getenv("RUTA_CONSOLIDADO", "datos/registros.xlsx")

# ── Nombres de hojas y columnas ──────────────────────────────────────────────
HOJA_OPERARIOS = os.getenv("HOJA_OPERARIOS", "OPERARIOS")
HOJA_SUBEVENTOS = "SUBEVENTOS"
COL_COD_EVENTO = "ID Razon Parada"
COL_EVENTO = "Razon de parada"
COL_ID_OPERARIO = "ID"
COL_OPERARIO = "OPERARIO"
COL_AREA_OPERARIO = "PROCESO"

# ── Control de acceso ────────────────────────────────────────────────────────
# Solo el personal cuya area figure en esta lista puede iniciar sesion como
# inspector. Se configura por variable de entorno, separado por comas.
AREAS_LOGIN = [
    a.strip()
    for a in os.getenv("AREAS_LOGIN", "ANALITICA,SUPERVISION").split(",")
    if a.strip()
]

# ── Parámetros de operación ──────────────────────────────────────────────────
INTERVALO_RECARGA_MIN = int(os.getenv("INTERVALO_RECARGA_MIN", "60"))
PUERTO_APP = int(os.getenv("PUERTO_APP", "5002"))

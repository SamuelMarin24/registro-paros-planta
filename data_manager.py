import openpyxl
import pyodbc
from datetime import datetime
import threading
import os
import logging
import unicodedata

logging.basicConfig(level=logging.INFO, format='%(asctime)s [DataManager] %(message)s')

from config import (
    SQL_SERVER, SQL_DATABASE, SQL_USER, SQL_PASSWORD, SQL_DRIVER,
    SQL_TABLA_ORDENES, RUTA_MAESTRAS, RUTA_RAZONES, RUTA_CONSOLIDADO,
    HOJA_OPERARIOS, HOJA_SUBEVENTOS, COL_COD_EVENTO, COL_EVENTO,
    COL_ID_OPERARIO, COL_OPERARIO, COL_AREA_OPERARIO,
    AREAS_LOGIN, INTERVALO_RECARGA_MIN,
)

# Consulta que trae las ordenes de produccion vigentes desde la base de datos.
SQL_QUERY_OP = f"""
    select op
          ,cliente
          ,id_cliente
          ,referencia
          ,cantidad_unica
          ,estado
          ,elaboro
          ,centro_operacion
          ,fecha_creacion
          ,fecha_compromiso
          ,tipo_trabajo
          ,op_repeticion
          ,vendedor
          ,oc_cliente
          ,cod_cliente
          ,num_cotiza
    from {SQL_TABLA_ORDENES}
"""

# Los sub eventos (paradas cortas dentro de una razon de paro) van en una hoja
# aparte del consolidado. Es relacional: no repite maquina/OP/operario, solo
# guarda el sub evento y lo conecta a la razon de paro por ID FILA.
HOJA_SUB_EVENTOS = "Sheet2"
ENCABEZADOS_SUB_EVENTO = [
    'ID FILA', 'COD PARO', 'SUB EVENTO', 'FECHA HORA INICIO', 'FECHA HORA FIN',
    'DURACION', 'OBSERVACIONES'
]

VERSION = "1.0.0"

NOMBRE_HOJA_PRINCIPAL = "Sheet"
ENCABEZADOS_PRINCIPAL = [
    'ID FILA', 'FECHA', 'MAQUINA', 'PLANTA', 'OP', 'COD REFERENCIA', 'REFERENCIA',
    'COD PARO', 'RAZON PARO', 'FECHA INICIO', 'HORA INICIO', 'FECHA FIN', 'HORA FIN',
    'CANTIDAD (TIROS/UNDS)', 'ID OPERARIO', 'OPERARIO', 'INSPECCIONO', 'OBSERVACIONES',
    'PARTE'
]


def _normalizar(texto):
    """Devuelve el texto en MAYÚSCULAS, sin tildes y sin espacios de más.
    Sirve para comparar áreas y valores sin importar cómo estén escritos."""
    t = (str(texto) if texto is not None else '').strip().upper()
    t = unicodedata.normalize('NFD', t)
    return ''.join(c for c in t if unicodedata.category(c) != 'Mn')


def _norm_cedula(valor):
    """Normaliza una cédula/ID leído del Excel a texto (por si viene como número
    o como número con '.0' al final)."""
    s = str(valor).strip() if valor is not None else ''
    if s.endswith('.0'):
        s = s[:-2]
    return s


_excel_lock = threading.Lock()


class LecturaInvalida(ValueError):
    """La lectura del contador digitada no es válida (por ejemplo, es menor o
    igual a la última lectura registrada de ese mismo operario en esa OP y
    máquina el mismo día). Se traduce a un mensaje para el inspector."""
    pass


def _miles(n):
    """12500 → '12.500' (formato colombiano), para los mensajes al inspector."""
    return f"{int(n):,}".replace(',', '.')


# ── ETAPA DEL TURNO ──────────────────────────────────────────────────────────
# El inspector la elige en una ventana que sale después del login y antes de
# iniciar el turno, y queda fija hasta que cierre el turno. Lo ÚNICO que define
# es si la OP es obligatoria:
#
#   'productiva'    → la OP es OBLIGATORIA.
#   'mantenimiento' → la OP es OPCIONAL (la máquina está parada; si el registro
#                     corresponde a una OP, se puede escribir, pero no se exige).
#
# Los eventos y el campo de cantidad están disponibles en LAS DOS etapas: en
# mantenimiento también se hacen tirajes de prueba y hay que registrarlos.
#
# Esto reemplazó a la regla anterior, que adivinaba por el nombre del evento y
# por lo tanto dependía de cómo estuvieran escritos en el archivo de maestras.
# La etapa NO se guarda en el consolidado: viaja en el JSON solo para validar.
def etapa_es_mantenimiento(etapa):
    """True solo si la etapa recibida es explícitamente la de mantenimiento.

    Cualquier otro valor —incluido vacío o ausente— se trata como etapa
    productiva, que es la más estricta (exige OP y calcula producción). Así, si
    llega una petición vieja o incompleta, el servidor peca de exigente y no de
    permisivo.
    """
    return _normalizar(etapa) == 'MANTENIMIENTO'


class DataManager:

    def __init__(self):
        self.maestra_op        = {}
        self.maestra_eventos   = {}
        self.maestra_sub_eventos = {}   # motivos de sub evento (hoja SUBEVENTOS)
        self.maestra_operarios = []
        self._ultima_carga     = None

        # Carga inicial
        self._recargar()

        # Hilo de recarga automática en background
        self._iniciar_recarga_automatica()

    # ── Carga de maestras ────────────────────────────────────────────────────

    def _recargar(self):
        logging.info("Recargando maestras desde red...")
        nuevas_op        = self._cargar_op()
        nuevos_eventos   = self._cargar_eventos()
        nuevos_sub_ev    = self._cargar_sub_eventos()
        nuevos_operarios = self._cargar_operarios()

        with _excel_lock:
            # NUNCA reemplazar una maestra que YA tenía datos por una vacía.
            #
            # Este era el bug que "dañaba" la app después de una hora: cada 60 min
            # el hilo de recarga vuelve a leer SQL Server y el Excel de la red. Si
            # en ese instante la fuente falla un momento (SQL Server no responde,
            # el recurso \\el servidor de archivos se cae, la sesión de red expiró), _cargar_op /
            # _cargar_eventos / _cargar_operarios atrapan el error y devuelven algo
            # VACÍO. La versión anterior pisaba los datos buenos con ese vacío, así
            # que a partir de la recarga la app se quedaba sin OPs (la "consulta a
            # SQL" dejaba de encontrar nada) o sin eventos/operarios (los campos y
            # el login dejaban de funcionar) hasta la siguiente recarga buena.
            #
            # En producción estas tres listas NUNCA están legítimamente vacías, así
            # que un resultado vacío solo puede ser una falla puntual: se conserva
            # lo que ya había en memoria y la app sigue trabajando. En el arranque
            # sí se aceptan vacíos (todavía no hay nada que perder).
            if nuevas_op:
                self.maestra_op = nuevas_op
            elif self.maestra_op:
                logging.warning(f"Recarga de OPs vino vacía: se conservan las {len(self.maestra_op)} anteriores (¿SQL Server caído?)")

            if nuevos_eventos:
                self.maestra_eventos = nuevos_eventos
            elif self.maestra_eventos:
                logging.warning(f"Recarga de eventos vino vacía: se conservan los {len(self.maestra_eventos)} anteriores (¿red caída?)")

            if nuevos_sub_ev:
                self.maestra_sub_eventos = nuevos_sub_ev
            elif self.maestra_sub_eventos:
                logging.warning(f"Recarga de sub eventos vino vacía: se conservan los {len(self.maestra_sub_eventos)} anteriores (¿red caída?)")

            if nuevos_operarios:
                self.maestra_operarios = nuevos_operarios
            elif self.maestra_operarios:
                logging.warning(f"Recarga de operarios vino vacía: se conservan los {len(self.maestra_operarios)} anteriores (¿red caída?)")

            self._ultima_carga = datetime.now()

        logging.info(f"  → {len(self.maestra_op)} OPs | {len(self.maestra_eventos)} razones | {len(self.maestra_sub_eventos)} sub eventos | {len(self.maestra_operarios)} operarios en memoria")

    def _cargar_op(self):
        data = {}
        conn = None
        try:
            # timeout=10 → si SQL Server no contesta el login en 10 s, corta en vez
            # de dejar colgado el hilo de recarga indefinidamente. conn.timeout
            # hace lo mismo con la consulta. Sin esto, un SQL Server ocupado o una
            # red lenta podían dejar la recarga esperando para siempre.
            conn = pyodbc.connect(
                f"DRIVER={{{SQL_DRIVER}}};"
                f"SERVER={SQL_SERVER};"
                f"DATABASE={SQL_DATABASE};"
                f"UID={SQL_USER};"
                f"PWD={SQL_PASSWORD};"
                f"Encrypt=yes;"
                f"TrustServerCertificate=yes;",
                timeout=10,
            )
            conn.timeout = 30
            cursor = conn.cursor()
            cursor.execute(SQL_QUERY_OP)
            for row in cursor.fetchall():
                if row.op:
                    data[str(row.op).strip()] = {
                        'cod_referencia': row.cod_cliente if row.cod_cliente is not None else '',
                        'referencia':     row.referencia if row.referencia is not None else '',
                    }
        except Exception as e:
            # Devuelve {} para que _recargar CONSERVE las OPs que ya tenía en
            # memoria (ver el porqué en _recargar). Nunca deja la app sin OPs.
            logging.error(f"Error leyendo OP desde SQL Server: {e}")
            return {}
        finally:
            # Cerrar SIEMPRE la conexión, incluso si la consulta falló a la mitad.
            # Antes el conn.close() estaba dentro del try, después del bucle: si
            # execute()/fetchall() lanzaba, la conexión quedaba sin cerrar y se
            # iban acumulando conexiones colgadas contra SQL Server.
            if conn is not None:
                try:
                    conn.close()
                except Exception:
                    pass
        return data

    def _leer_lista_codigo_nombre(self, ws):
        """Lee una hoja con dos columnas (código + nombre) y devuelve {cod: nombre}.
        Busca los encabezados por nombre en las primeras filas; si no aparecen,
        cae a la posición: columna A = código, columna B = nombre."""
        data = {}
        idx_cod, idx_nombre, header_row_idx = None, None, None
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
            row_vals = [str(v).strip() if v is not None else '' for v in row]
            if COL_COD_EVENTO in row_vals and COL_EVENTO in row_vals:
                idx_cod = row_vals.index(COL_COD_EVENTO)
                idx_nombre = row_vals.index(COL_EVENTO)
                header_row_idx = i
                break

        if header_row_idx is None:
            logging.warning(f"Hoja '{ws.title}': no se hallaron los encabezados por nombre; se usa posición A/B.")
            idx_cod, idx_nombre, header_row_idx = 0, 1, 1

        for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
            if len(row) > max(idx_cod, idx_nombre) and row[idx_cod] is not None and str(row[idx_cod]).strip():
                data[str(row[idx_cod]).strip()] = row[idx_nombre]
        return data

    def _cargar_eventos(self):
        # Carga las RAZONES DE PARO desde el archivo de razones de paro (reemplazan a los
        # "eventos" de la version anterior). Internamente se siguen manejando como 'eventos'
        # para no reescribir el resto de la app; lo único que cambia es la fuente.
        data = {}
        if not os.path.exists(RUTA_RAZONES):
            logging.warning(f"No se encontró el archivo de razones de paro en: {RUTA_RAZONES}")
            return data
        try:
            wb = openpyxl.load_workbook(RUTA_RAZONES, data_only=True, read_only=True)
            # Las razones van en la primera hoja que NO sea la de SUBEVENTOS (ese
            # archivo tiene dos hojas: la de razones y la de motivos de sub evento).
            hojas = [h for h in wb.sheetnames if h.strip().upper() != HOJA_SUBEVENTOS.upper()]
            ws = wb[hojas[0]] if hojas else wb[wb.sheetnames[0]]
            data = self._leer_lista_codigo_nombre(ws)
            wb.close()
        except Exception as e:
            logging.error(f"Error leyendo las razones de paro de el archivo de razones de paro: {e}")
        return data

    def _cargar_sub_eventos(self):
        """Carga los MOTIVOS DE SUB EVENTO desde la hoja 'SUBEVENTOS' del mismo
        el archivo de razones de paro. Es la lista que se despliega cuando el inspector
        registra un sub evento (antes se escribía a mano)."""
        data = {}
        if not os.path.exists(RUTA_RAZONES):
            logging.warning(f"No se encontró el archivo de razones de paro en: {RUTA_RAZONES}")
            return data
        try:
            wb = openpyxl.load_workbook(RUTA_RAZONES, data_only=True, read_only=True)
            # Busca la hoja SUBEVENTOS sin importar mayúsculas/espacios.
            nombre = next((h for h in wb.sheetnames if h.strip().upper() == HOJA_SUBEVENTOS.upper()), None)
            if nombre is None:
                logging.warning(f"No se encontró la hoja '{HOJA_SUBEVENTOS}'. Hojas: {wb.sheetnames}")
                wb.close()
                return data
            data = self._leer_lista_codigo_nombre(wb[nombre])
            wb.close()
        except Exception as e:
            logging.error(f"Error leyendo la hoja SUBEVENTOS: {e}")
        return data

    # ── Recarga automática en background ────────────────────────────────────

    def _iniciar_recarga_automatica(self):
        def loop():
            while True:
                threading.Event().wait(INTERVALO_RECARGA_MIN * 60)
                try:
                    self._recargar()
                except Exception as e:
                    logging.error(f"Error en recarga automática: {e}")

        t = threading.Thread(target=loop, daemon=True)
        t.start()
        logging.info(f"Recarga automática configurada cada {INTERVALO_RECARGA_MIN} minutos")

    def _cargar_operarios(self):
        data = []
        if not os.path.exists(RUTA_MAESTRAS):
            return data
        try:
            wb = openpyxl.load_workbook(RUTA_MAESTRAS, data_only=True, read_only=True)
            if HOJA_OPERARIOS not in wb.sheetnames:
                logging.warning(f"Hoja '{HOJA_OPERARIOS}' no encontrada.")
                wb.close()
                return data
            ws = wb[HOJA_OPERARIOS]
            header_row_idx = None
            headers = []
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
                row_vals = [str(v).strip() if v else '' for v in row]
                if COL_ID_OPERARIO in row_vals and COL_OPERARIO in row_vals:
                    header_row_idx = i
                    headers = row_vals
                    break
            if header_row_idx is None:
                logging.warning(f"No se encontraron columnas '{COL_ID_OPERARIO}', '{COL_OPERARIO}' en MAESTRA OPERARIOS")
                wb.close()
                return data
            idx_id  = headers.index(COL_ID_OPERARIO)
            idx_nom = headers.index(COL_OPERARIO)

            # Columna del área/proceso. Primero se busca por nombre (tolerante a
            # tildes/mayúsculas); si no aparece, se usa la columna que está justo a
            # la derecha de OPERARIO (que es donde hoy está el área en la maestra).
            idx_area = None
            candidatos = {_normalizar(COL_AREA_OPERARIO), 'PROCESO', 'AREA',
                          'PROCESO PRODUCTIVO', 'SECCION', 'DEPARTAMENTO'}
            for j, h in enumerate(headers):
                if h and _normalizar(h) in candidatos:
                    idx_area = j
                    break
            if idx_area is None and idx_nom + 1 < len(headers):
                idx_area = idx_nom + 1

            for row in ws.iter_rows(min_row=header_row_idx + 1, values_only=True):
                if row[idx_id] and row[idx_nom]:
                    area = ''
                    if idx_area is not None and idx_area < len(row) and row[idx_area]:
                        area = str(row[idx_area]).strip()
                    data.append({
                        'id':     _norm_cedula(row[idx_id]),
                        'nombre': str(row[idx_nom]).strip(),
                        'area':   area,
                    })
            wb.close()
        except Exception as e:
            logging.error(f"Error leyendo operarios: {e}")
        return data

    # ── API pública ──────────────────────────────────────────────────────────

    def buscar_op(self, op):
        return self.maestra_op.get(str(op).strip(), None)

    def buscar_evento(self, cod):
        return self.maestra_eventos.get(str(cod).strip(), None)

    def get_operarios(self):
        # Para el autocompletado de "Operario" del formulario: se EXCLUYE al
        # personal de las áreas de login (supervisores/inspectores, p. ej.
        # 'ANALITICA'), para que ese campo solo muestre a quienes operan
        # la máquina. El id que devuelve es la cédula (columna ID de la maestra).
        areas_login = {_normalizar(a) for a in AREAS_LOGIN}
        return [op for op in self.maestra_operarios
                if _normalizar(op.get('area', '')) not in areas_login]

    def validar_login(self, cedula):
        """Valida el ingreso del inspector por CÉDULA (que es la contraseña).

        Busca la cédula en MAESTRA OPERARIOS y solo permite el ingreso si esa
        persona pertenece a un área autorizada (AREAS_LOGIN, p. ej.
        'ANALITICA' o 'SUPERVISION'). Devuelve el nombre tal como está en
        la maestra si el ingreso es válido, o None si la cédula no existe o no
        tiene permiso.
        """
        cedula_in = _norm_cedula(cedula)
        if not cedula_in:
            return None
        areas_ok = {_normalizar(a) for a in AREAS_LOGIN}
        for op in self.maestra_operarios:
            if op.get('id', '') == cedula_in and _normalizar(op.get('area', '')) in areas_ok:
                return op.get('nombre', '')
        return None

    def estado(self):
        return {
            'ultima_carga':      self._ultima_carga.strftime('%Y-%m-%d %H:%M:%S') if self._ultima_carga else 'nunca',
            'total_ops':         len(self.maestra_op),
            'version':           VERSION,
            'total_eventos':     len(self.maestra_eventos),
            'total_sub_eventos': len(self.maestra_sub_eventos),
            'total_operarios':   len(self.maestra_operarios),
            'ruta_op':           f"{SQL_SERVER} / {SQL_DATABASE}",
            'ruta_maestras':     RUTA_MAESTRAS,
            'proxima_recarga':   f"en ~{INTERVALO_RECARGA_MIN} min desde última carga",
        }

    def guardar_registro(self, data):
        with _excel_lock:
            # La hoja principal lleva ID FILA como primera columna (consecutivo por
            # evento), seguida de las columnas de siempre.
            encabezados = ENCABEZADOS_PRINCIPAL

            if not os.path.exists(RUTA_CONSOLIDADO):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = NOMBRE_HOJA_PRINCIPAL   # nombre fijo, no dejarlo al azar
                ws.append(encabezados)
                wb.save(RUTA_CONSOLIDADO)

            # El formulario envía estos 4 campos capturados con los botones
            # "Iniciar evento" / "Terminar evento". Si faltan (llamada antigua o
            # incompleta), se usa la fecha/hora del servidor como respaldo.
            ahora = datetime.now()
            fecha_inicio = str(data.get('fecha_inicio', '')).strip() or ahora.strftime('%Y-%m-%d')
            hora_inicio  = str(data.get('hora_inicio', '')).strip()
            fecha_fin    = str(data.get('fecha_fin', '')).strip() or ahora.strftime('%Y-%m-%d')
            hora_fin     = str(data.get('hora_fin', '')).strip() or ahora.strftime('%H:%M')
            fecha_hoy = fecha_inicio  # se usa como "FECHA" del registro y para el acumulado del día
            evento   = data.get('evento', '')
            op       = str(data.get('op', '')).strip()
            maquina  = str(data.get('maquina', '')).strip().upper()
            op_id    = str(data.get('operario_id', '')).strip()

            # ── Cálculo de la PRODUCCIÓN por evento (lectura acumulada del contador) ──
            # El operario digita la lectura total del contador. La producción de este
            # evento = lectura − (suma de lo ya guardado de ESTE operario + OP + máquina
            # + día). El primer registro del operario (base) no tiene nada previo, así
            # que se guarda la lectura tal cual. Al cambiar operario, OP o turno, el
            # acumulado vuelve a 0 → el siguiente registro es una base nueva.
            #
            # La lectura se puede digitar en CUALQUIER evento y en CUALQUIER etapa: es
            # el mismo contador físico, que nunca se reinicia. En mantenimiento también
            # se hacen tirajes de prueba, y eso hay que registrarlo. Cuando la OP va
            # vacía, esas lecturas se acumulan entre sí (ver _acumulado_operario).
            # Si el contador no se movió, el inspector deja el campo vacío y se guarda 0.
            try:
                lectura = int(data.get('cantidad', 0) or 0)
            except (ValueError, TypeError):
                lectura = 0

            # PARTE (A / B / vacío): una misma OP puede traer dos trabajos
            # distintos. Al pasar de la parte A a la B, el contador de la máquina
            # SE REINICIA EN CERO, así que cada parte lleva su PROPIO acumulado:
            # si se mezclaran, la primera lectura de la parte B se restaría contra
            # lo de la parte A y la producción saldría mal (p. ej. 12.275 − 10.600).
            parte = (data.get('parte') or '').strip().upper()
            if parte not in ('A', 'B'):
                parte = ''

            acumulado = self._acumulado_operario(op, maquina, op_id, fecha_hoy, parte)

            # ── Regla: la lectura del contador solo puede subir ─────────────
            # El contador es acumulativo y nunca arranca en cero, así que una
            # lectura menor o igual a la última de ESE operario (misma OP,
            # máquina y día) solo puede ser un error de digitación: antes se
            # guardaba como producción negativa.
            # Solo se valida cuando hay contra qué comparar (acumulado > 0) y
            # cuando el inspector realmente digitó una lectura (lectura > 0):
            # si el contador no se movió, deja el campo vacío y se guarda 0.
            if acumulado > 0 and lectura > 0 and lectura <= acumulado:
                raise LecturaInvalida(
                    f"La lectura {_miles(lectura)} no puede ser menor ni igual a la última "
                    f"registrada de este operario ({_miles(acumulado)}). "
                    f"Si el contador no se movió, deja el campo vacío."
                )

            if lectura <= 0:
                # Campo vacío: el contador no avanzó en este evento (un paro,
                # por ejemplo). Antes esto guardaba 0 − acumulado = negativo.
                cantidad_guardar = 0
            else:
                cantidad_guardar = lectura - acumulado   # base: acumulado=0 → guarda la lectura

            wb = openpyxl.load_workbook(RUTA_CONSOLIDADO)
            # La hoja principal se identifica POR NOMBRE ("Sheet"), nunca por
            # wb.active. wb.active depende de qué pestaña quedó seleccionada la
            # última vez que alguien guardó el archivo en Excel (con Autoguardado
            # activado, basta con hacer clic en la pestaña Sheet2 para que quede
            # marcada como activa) — si se usara wb.active, el evento principal
            # podría terminar escrito en Sheet2 por accidente, como ya pasó.
            ws = wb[NOMBRE_HOJA_PRINCIPAL] if NOMBRE_HOJA_PRINCIPAL in wb.sheetnames else wb.worksheets[0]

            # ID FILA: consecutivo por evento. Es la llave a la que se pegan los
            # sub eventos. Se calcula como el mayor ID ya presente + 1 (robusto
            # aunque haya filas de más), y va como PRIMERA columna de la fila.
            id_fila = self._siguiente_id_fila(ws)

            # Columna PARTE: si el Excel todavía no tiene ese encabezado, se agrega solo.
            col_parte = self._asegurar_encabezado(ws, 'PARTE')

            # Fila en el orden de ENCABEZADOS_PRINCIPAL (ID FILA primero).
            ws.append([
                id_fila,
                fecha_hoy,
                data.get('maquina', ''),
                data.get('planta', ''),
                data.get('op', ''),
                data.get('cod_referencia', ''),
                data.get('referencia', ''),
                str(data.get('cod_evento', '') or ''),   # código, siempre como texto
                evento,
                fecha_inicio,
                hora_inicio,
                fecha_fin,
                hora_fin,
                cantidad_guardar,
                data.get('operario_id', ''),
                data.get('operario_nombre', ''),
                data.get('inspeccionó', ''),
                data.get('observaciones', ''),
            ])
            # PARTE en su columna real (opcional: 'A', 'B' o vacío). Va aparte y al
            # final para no descuadrar las filas creadas con el layout anterior.
            ws.cell(row=ws.max_row, column=col_parte, value=parte)

            # Sub eventos del evento: llegan en data['sub_eventos'] (se acumulan en
            # el celular durante el evento y se mandan al cerrarlo). Cada uno se
            # guarda en la hoja Sheet2 con ESTE id_fila, así quedan pegados al
            # evento (el mismo ID replicado en todas sus filas de sub evento).
            self._guardar_sub_eventos(wb, id_fila, data.get('sub_eventos'))

            wb.save(RUTA_CONSOLIDADO)
            n_sub = len(data.get('sub_eventos') or [])
            logging.info(f"Registro guardado → ID FILA {id_fila} | {data.get('maquina')} | OP {op} | "
                         f"{evento} | {hora_inicio}-{hora_fin} | lectura={lectura} guardado={cantidad_guardar}"
                         + (f" | {n_sub} sub eventos" if n_sub else ""))

    def _asegurar_encabezado(self, ws, nombre):
        """Devuelve la columna (1-based) cuyo encabezado (fila 1) es 'nombre'.
        Si no existe, lo agrega al final de la fila de encabezados y devuelve su
        nueva columna.

        La comparación NO distingue mayúsculas, tildes ni espacios de más: si el
        encabezado del Excel quedó como 'Cod Paro ' o 'COD  PARO', igual lo
        reconoce en vez de crear una columna repetida al final."""
        def _norm(v):
            if v is None:
                return ''
            txt = str(v).strip().upper()
            txt = unicodedata.normalize('NFD', txt)
            txt = ''.join(c for c in txt if unicodedata.category(c) != 'Mn')
            return ' '.join(txt.split())   # colapsa espacios repetidos

        objetivo = _norm(nombre)
        fila1 = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        for i, valor in enumerate(fila1, start=1):
            if _norm(valor) == objetivo:
                return i
        col = len(fila1) + 1
        ws.cell(row=1, column=col, value=nombre)
        return col

    def _siguiente_id_fila(self, ws):
        """Mayor ID FILA (columna A) ya presente en la hoja principal, + 1.
        Si la hoja solo tiene el encabezado, arranca en 1."""
        max_id = 0
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            v = row[0]
            if isinstance(v, int):
                max_id = max(max_id, v)
            elif isinstance(v, str) and v.strip().isdigit():
                max_id = max(max_id, int(v.strip()))
        return max_id + 1

    def _codigo_de_sub_evento(self, nombre):
        """Busca el CÓDIGO de un sub evento a partir de su NOMBRE, usando la
        maestra de la hoja SUBEVENTOS. Sirve de respaldo para que la columna
        COD PARO quede llena aunque el celular no haya enviado el código.

        Compara ignorando mayúsculas, tildes y espacios de más, porque el nombre
        viaja como texto y puede venir con pequeñas diferencias."""
        if not nombre:
            return ''
        objetivo = _normalizar(str(nombre))
        for cod, nom in (self.maestra_sub_eventos or {}).items():
            if _normalizar(str(nom or '')) == objetivo:
                return str(cod).strip()
        return ''

    def _guardar_sub_eventos(self, wb, id_fila, sub_eventos):
        """Guarda en Sheet2 los sub eventos de un evento, todos con el mismo
        id_fila (la llave al evento principal). No repite máquina/OP/operario:
        esos se leen del evento vía ID FILA.

        Cada dato se escribe buscando SU COLUMNA POR NOMBRE en la fila 1, no por
        posición fija: así el orden de las columnas del Excel puede cambiar (o
        alguien puede insertar una en medio, como pasó con COD PARO) sin que los
        datos se corran de columna."""
        if not sub_eventos:
            return
        if HOJA_SUB_EVENTOS in wb.sheetnames:
            ws2 = wb[HOJA_SUB_EVENTOS]
            if ws2.max_row < 1 or all(c.value is None for c in ws2[1]):
                ws2.append(ENCABEZADOS_SUB_EVENTO)
        else:
            ws2 = wb.create_sheet(HOJA_SUB_EVENTOS)
            ws2.append(ENCABEZADOS_SUB_EVENTO)

        # Columnas que faltaran en la hoja (p. ej. COD PARO en archivos viejos)
        # se agregan al final para no perder el dato.
        columnas = {}
        for nombre in ENCABEZADOS_SUB_EVENTO:
            columnas[nombre] = self._asegurar_encabezado(ws2, nombre)

        for se in sub_eventos:
            # Duración medida por el cronómetro del celular (exacta al segundo).
            # Se guarda como fracción de día con formato [mm]:ss para que la
            # columna DURACION se pueda SUMAR en Excel (p. ej. 03:45).
            try:
                seg = int(se.get('duracion_segundos', 0) or 0)
            except (ValueError, TypeError):
                seg = 0
            if seg < 0:
                seg = 0

            nombre_se = se.get('sub_evento', '') or ''
            cod_se = str(se.get('cod_sub_evento', '') or '').strip()
            if not cod_se:
                # El celular no mandó el código (p. ej. tiene cargada una versión
                # anterior del formulario, o el motivo se escribió a mano). Se
                # deduce buscando el NOMBRE en la maestra de sub eventos, para que
                # la columna COD PARO nunca quede vacía y no toque llenarla a mano.
                cod_se = self._codigo_de_sub_evento(nombre_se)

            valores = {
                'ID FILA':           id_fila,                          # mismo ID del evento (se repite por sub evento)
                'COD PARO':          cod_se,                           # código del motivo, siempre como texto
                'SUB EVENTO':        nombre_se,                        # nombre del motivo
                'FECHA HORA INICIO': se.get('fecha_hora_inicio', ''),  # "YYYY-MM-DD HH:MM:SS"
                'FECHA HORA FIN':    se.get('fecha_hora_fin', ''),
                'DURACION':          seg / 86400.0,                    # DURACION como tiempo real
                'OBSERVACIONES':     se.get('observaciones', ''),
            }

            fila = ws2.max_row + 1
            for nombre, valor in valores.items():
                ws2.cell(row=fila, column=columnas[nombre], value=valor)
            ws2.cell(row=fila, column=columnas['DURACION']).number_format = '[mm]:ss'

    def _acumulado_operario(self, op, maquina, op_id, fecha, parte=''):
        """Suma lo ya guardado en CANTIDAD para un operario + OP + máquina + día
        + PARTE. Sirve de base para calcular la producción del siguiente registro.

        La PARTE entra en el grupo porque una misma OP puede traer dos trabajos
        (A y B) y el contador de la máquina se REINICIA al pasar de uno al otro:
        si se mezclaran, la primera lectura de la parte B se restaría contra lo
        de la parte A. Con la parte separada, cada una arranca su propia base.

        La OP vacía es un grupo válido, no un descarte: en etapa de mantenimiento
        se registran tirajes de prueba sin OP, y esas lecturas también tienen que
        acumularse entre sí. Si estas filas se saltaran, cada lectura se guardaría
        completa (500, 700, 900) en vez de la diferencia (500, 200, 200).
        """
        if not os.path.exists(RUTA_CONSOLIDADO):
            return 0
        op = str(op).strip()
        parte = (parte or '').strip().upper()
        if parte not in ('A', 'B'):
            parte = ''
        total = 0
        try:
            wb = openpyxl.load_workbook(RUTA_CONSOLIDADO, data_only=True, read_only=True)
            # Por nombre, no por pestaña activa — mismo motivo que en guardar_registro.
            ws = wb[NOMBRE_HOJA_PRINCIPAL] if NOMBRE_HOJA_PRINCIPAL in wb.sheetnames else wb.worksheets[0]
            filas = ws.iter_rows(min_row=1, values_only=True)
            headers = next(filas, None)
            if not headers:
                wb.close()
                return 0
            headers = [str(h).strip() if h else '' for h in headers]

            def col(nombre):
                for i, h in enumerate(headers):
                    if h == nombre:
                        return i
                return None

            i_fecha = col('FECHA')
            i_maq   = col('MAQUINA')
            i_op    = col('OP')
            i_cant  = col('CANTIDAD (TIROS/UNDS)')
            i_idop  = col('ID OPERARIO')
            i_parte = col('PARTE')

            if i_op is None or i_idop is None:
                wb.close()
                return 0

            def _fecha_str(v):
                if v is None:
                    return ''
                if hasattr(v, 'strftime'):
                    return v.strftime('%Y-%m-%d')
                return str(v).strip()[:10]

            def _celda(row, i):
                return str(row[i]).strip() if (i is not None and i < len(row) and row[i] is not None) else ''

            for row in filas:
                r_id = _celda(row, i_idop)
                if not r_id:
                    continue          # fila sin operario: no es un registro válido
                r_op  = _celda(row, i_op)
                r_fec = _fecha_str(row[i_fecha]) if (i_fecha is not None and i_fecha < len(row)) else ''
                r_maq = _celda(row, i_maq).upper()
                # PARTE de la fila: en archivos viejos la columna no existe y
                # cuenta como parte vacía, igual que un registro sin parte marcada.
                r_parte = _celda(row, i_parte).upper()
                if r_parte not in ('A', 'B'):
                    r_parte = ''
                if (r_op == op and r_id == op_id and r_fec == fecha
                        and r_maq == maquina and r_parte == parte):
                    try:
                        total += int(row[i_cant]) if (i_cant is not None and row[i_cant] is not None) else 0
                    except (ValueError, TypeError):
                        pass
            wb.close()
        except Exception as e:
            logging.error(f"Error calculando acumulado: {e}")
            return 0
        return total

    def historial_op(self, op, maquina, fecha=None):
        """Devuelve los registros del consolidado para una OP, máquina y día dados.

        Filtra por OP + MAQUINA + FECHA (por defecto hoy) y retorna la lista de
        registros (hora, evento, cantidad) más el total acumulado del día.
        """
        if fecha is None:
            fecha = datetime.now().strftime('%Y-%m-%d')
        op      = str(op).strip()
        maquina = str(maquina).strip().upper()

        if not os.path.exists(RUTA_CONSOLIDADO):
            return {'registros': [], 'total': 0}

        def _fecha_str(v):
            if v is None:
                return ''
            if hasattr(v, 'strftime'):
                return v.strftime('%Y-%m-%d')
            return str(v).strip()[:10]

        def _hora_str(v):
            if v is None:
                return ''
            if hasattr(v, 'strftime'):
                return v.strftime('%H:%M')
            return str(v).strip()[:5]

        registros = []
        total = 0
        try:
            with _excel_lock:
                wb = openpyxl.load_workbook(RUTA_CONSOLIDADO, data_only=True, read_only=True)
                # Por nombre, no por pestaña activa — mismo motivo que en guardar_registro.
                ws = wb[NOMBRE_HOJA_PRINCIPAL] if NOMBRE_HOJA_PRINCIPAL in wb.sheetnames else wb.worksheets[0]
                filas = ws.iter_rows(min_row=1, values_only=True)
                headers = next(filas, None)
                if not headers:
                    wb.close()
                    return {'registros': [], 'total': 0}
                headers = [str(h).strip() if h else '' for h in headers]

                def col(*nombres):
                    """Devuelve el índice de la primera columna que coincida con
                    alguno de los nombres dados. Acepta varios porque esta app
                    renombró COD EVENTO→COD PARO y EVENTO→RAZON PARO: así el
                    historial funciona con el nombre nuevo y también con el viejo
                    (por si quedan archivos con el encabezado anterior)."""
                    for nombre in nombres:
                        for i, h in enumerate(headers):
                            if h == nombre:
                                return i
                    return None

                i_fecha    = col('FECHA')
                i_maq      = col('MAQUINA')
                i_op       = col('OP')
                # Antes existía una sola columna 'HORA'; ahora se lee el rango
                # de inicio/fin de cada evento desde las 2 columnas nuevas.
                i_hora_ini = col('HORA INICIO')
                i_hora_fin = col('HORA FIN')
                i_cant     = col('CANTIDAD (TIROS/UNDS)')
                i_evt      = col('RAZON PARO', 'EVENTO')
                i_parte    = col('PARTE')
                i_idop     = col('ID OPERARIO')
                i_opnom    = col('OPERARIO')


                if i_op is None:
                    wb.close()
                    return {'registros': [], 'total': 0}

                for row in filas:
                    # La OP vacía es un grupo válido (mantenimiento sin OP), no
                    # una fila a descartar. Se filtra por operario para no contar
                    # renglones en blanco del Excel.
                    r_idop = (str(row[i_idop]).strip()
                              if (i_idop is not None and i_idop < len(row) and row[i_idop] is not None) else '')
                    if not r_idop:
                        continue
                    r_op = (str(row[i_op]).strip()
                            if (i_op < len(row) and row[i_op] is not None) else '')
                    r_fecha = _fecha_str(row[i_fecha]) if i_fecha is not None else ''
                    r_maq = (str(row[i_maq]).strip().upper()
                             if (i_maq is not None and row[i_maq] is not None) else '')
                    if r_op == op and r_fecha == fecha and r_maq == maquina:
                        try:
                            cant = int(row[i_cant]) if (i_cant is not None and row[i_cant] is not None) else 0
                        except (ValueError, TypeError):
                            cant = 0
                        registros.append({
                            # Reemplaza al antiguo campo único 'hora': ahora se
                            # devuelven inicio y fin para mostrar el rango en el historial.
                            'hora_inicio': _hora_str(row[i_hora_ini]) if i_hora_ini is not None else '',
                            'hora_fin':    _hora_str(row[i_hora_fin]) if i_hora_fin is not None else '',
                            'evento':   (str(row[i_evt]) if (i_evt is not None and row[i_evt] is not None) else ''),
                            # PARTE (A / B / vacío): el celular la necesita para
                            # calcular la base del contador por separado en cada parte.
                            'parte':    (str(row[i_parte]).strip().upper()
                                         if (i_parte is not None and i_parte < len(row)
                                             and row[i_parte] is not None) else ''),
                            'cantidad': cant,
                            'operario_id':     r_idop,
                            'operario_nombre': (str(row[i_opnom]).strip() if (i_opnom is not None and row[i_opnom] is not None) else ''),
                        })
                        total += cant
                wb.close()
        except Exception as e:
            logging.error(f"Error leyendo historial: {e}")
            return {'registros': [], 'total': 0}

        return {'registros': registros, 'total': total}

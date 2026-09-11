import logging

from flask import Flask, render_template, request, jsonify
from data_manager import DataManager, LecturaInvalida, etapa_es_mantenimiento
from config import PUERTO_APP

app = Flask(__name__)

# Al correr con debug=False, Flask compila form.html UNA sola vez y lo deja en
# memoria: si se reemplaza el archivo, la app sigue sirviendo el viejo hasta que
# se reinicie. Con esto vuelve a leerlo del disco cuando cambia.
app.config['TEMPLATES_AUTO_RELOAD'] = True

dm = DataManager()

@app.after_request
def sin_cache(resp):
    """Evita que el celular se quede con una versión vieja del formulario.
    Los navegadores móviles guardan la página en caché, así que tras actualizar
    la app el inspector podía seguir viendo el formulario anterior (y con él,
    errores ya corregidos). Con estas cabeceras siempre pide la última."""
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/buscar_op')
def buscar_op():
    op = request.args.get('op', '').strip()
    result = dm.buscar_op(op)
    return jsonify(result if result else {})

@app.route('/buscar_evento')
def buscar_evento():
    cod = request.args.get('cod', '').strip()
    evento = dm.buscar_evento(cod)
    return jsonify({'evento': evento})

@app.route('/guardar', methods=['POST'])
def guardar():
    data = request.json

    # La ETAPA del turno (la elige el inspector en la ventana que sale tras el
    # login) define si la OP es obligatoria. Viaja en el JSON pero NO se guarda
    # en el Excel: solo sirve para validar aquí. Las cantidades se registran en
    # las dos etapas (en mantenimiento también hacen tirajes de prueba).
    if not etapa_es_mantenimiento(data.get('etapa')) and not data.get('op'):
        return jsonify({'error': 'Debe ingresar una OP para trabajar en etapa productiva'}), 400
    if not data.get('maquina'):
        return jsonify({'error': 'Debe seleccionar una máquina'}), 400
    # El formulario solo llama a /guardar desde "Terminar evento", que siempre
    # manda hora_inicio y hora_fin. Si faltan, es que no se pasó por el flujo
    # de Iniciar/Terminar (llamada incompleta o directa a la API).
    if not data.get('hora_inicio') or not data.get('hora_fin'):
        return jsonify({'error': 'Debe iniciar y terminar el evento antes de guardar'}), 400

    try:
        dm.guardar_registro(data)
    except LecturaInvalida as e:
        # Error del inspector (lectura menor o igual a la anterior): el mensaje
        # se muestra tal cual en el banner del formulario.
        return jsonify({'error': str(e)}), 400
    except Exception:
        # Falla al escribir el consolidado (Excel abierto en la planta, red caída).
        # Antes reventaba en un 500 y el formulario solo decía "Error de conexión".
        logging.exception("Error guardando el registro")
        return jsonify({'error': 'No se pudo guardar en el consolidado. Revisa que '
                                 'el archivo consolidado no esté abierto y vuelve a intentar.'}), 500

    return jsonify({'status': 'ok'})

@app.route('/lista_eventos')
def lista_eventos():
    """Devuelve todos los eventos para el autocomplete."""
    return jsonify([{'cod': k, 'nombre': v} for k, v in dm.maestra_eventos.items()])

@app.route('/lista_sub_eventos')
def lista_sub_eventos():
    """Devuelve los motivos de SUB EVENTO (hoja SUBEVENTOS) para el desplegable
    que aparece al registrar un sub evento."""
    return jsonify([{'cod': k, 'nombre': v} for k, v in dm.maestra_sub_eventos.items()])

@app.route('/lista_ops')
def lista_ops():
    """Devuelve todas las OPs para el autocomplete del formulario."""
    return jsonify(list(dm.maestra_op.keys()))

@app.route('/lista_operarios')
def lista_operarios():
    return jsonify(dm.get_operarios())

@app.route('/login', methods=['POST'])
def login():
    """Valida el ingreso del inspector por codigo de empleado.
    El nombre se trae de MAESTRA OPERARIOS, filtrando por área autorizada."""
    data = request.json or {}
    nombre = dm.validar_login(data.get('cedula', ''))
    if nombre:
        return jsonify({'ok': True, 'nombre': nombre})
    return jsonify({'ok': False, 'error': 'Codigo de empleado no válida o sin permiso de ingreso'}), 401

@app.route('/estado')
def estado():
    """Endpoint de diagnóstico: muestra cuándo se cargaron las maestras y cuántos registros hay."""
    return jsonify(dm.estado())

@app.route('/recargar', methods=['POST'])
def recargar():
    """Recarga manual forzada (útil si acaban de actualizar un Excel)."""
    try:
        dm._recargar()
        return jsonify({'status': 'ok', 'info': dm.estado()})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/historial_op')
def historial_op():
    """Historial de registros de una OP para hoy y la máquina indicada."""
    op = request.args.get('op', '').strip()
    maquina = request.args.get('maquina', '').strip()
    return jsonify(dm.historial_op(op, maquina))

if __name__ == '__main__':
    # Puerto 5002: en este mismo PC ya corren del mismo tipo y otras apps.
    # Cada app necesita su propio puerto; si dos compartieran, la que arranca
    # segunda falla en silencio (pythonw no muestra consola) y el navegador
    # seguiría viendo a la primera.
    #
    # threaded=True → atiende varias peticiones a la vez. Sin esto, el servidor
    # procesa UNA sola a la vez, así que mientras un inspector guarda un registro
    # (que abre y reescribe el Excel completo por red, y puede tardar segundos si
    # la red va lenta), TODOS los demás celulares se quedan esperando y la app
    # "se pausa". Es seguro: toda la escritura/lectura del Excel ya está
    # serializada con _excel_lock, así que dos guardados no se pisan; y las
    # consultas que solo leen memoria (buscar OP, login, listas) responden al
    # instante sin esperar al que está guardando.
    app.run(host='0.0.0.0', port=PUERTO_APP, debug=False, threaded=True)

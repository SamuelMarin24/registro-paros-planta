# Registro de Paros en Planta

Aplicación web para el registro en tiempo real de paradas de máquina en planta de producción. Reemplaza el registro manual en papel por un formulario accesible desde el celular de los operarios, con cálculo automático de producción y trazabilidad de cada evento.

## El problema

En una planta de producción, cada minuto que una máquina está detenida tiene un costo. El registro de esas paradas se hacía en papel: el inspector anotaba la máquina, la hora, el motivo y la cantidad producida, y esos formatos se digitaban después a mano en Excel. Eso significaba:

- Información con horas o días de retraso.
- Errores de digitación en las lecturas del contador.
- Imposibilidad de hacer seguimiento en el momento.
- Sin trazabilidad entre una parada y sus causas específicas.

## La solución

Una aplicación Flask que corre en la red local de la planta y se abre desde el navegador del celular. El inspector inicia sesión con su código de empleado, selecciona la máquina, y registra cada evento con sus tiempos exactos medidos por cronómetro.

### Funcionalidades principales

- **Registro por evento con cronómetro:** botones de iniciar/terminar que capturan fecha y hora exactas, sin depender de que el operario las escriba.
- **Cálculo automático de producción:** el operario digita la lectura acumulada del contador de la máquina y el sistema calcula la producción real del evento restando lo ya registrado (por operario, orden de producción, máquina, día y parte del trabajo).
- **Validación de lecturas:** el contador es acumulativo y nunca retrocede, así que el sistema rechaza lecturas menores o iguales a la anterior y avisa al inspector en el momento, evitando producciones negativas por errores de digitación.
- **Sub eventos relacionales:** dentro de una parada larga se pueden registrar múltiples paradas cortas con su propio motivo y duración, enlazadas al evento principal mediante un ID.
- **Autocompletado desde la base de datos:** las órdenes de producción se consultan directamente desde SQL Server; los operarios y los motivos de parada se leen desde archivos maestros en red.
- **Control de acceso por área:** solo el personal de las áreas autorizadas puede iniciar sesión como inspector.
- **Recarga automática de maestras:** cada hora se refrescan los datos desde las fuentes, con protección ante fallos de red (si una fuente no responde, conserva los datos en memoria en vez de quedarse vacía).

## Stack técnico

| Componente | Tecnología |
|---|---|
| Backend | Python 3, Flask |
| Base de datos | SQL Server (vía `pyodbc`) |
| Persistencia de registros | Excel (`openpyxl`) sobre carpeta de red |
| Frontend | HTML, CSS y JavaScript sin frameworks |
| Concurrencia | `threading` con locks para escritura segura |

## Decisiones técnicas destacadas

**Escritura concurrente segura.** Varios inspectores usan la app al mismo tiempo desde distintos celulares. Toda la lectura y escritura del consolidado está serializada con un lock (`threading.Lock`), y el servidor corre con `threaded=True` para que las consultas que solo leen memoria (buscar OP, login, listas) respondan al instante sin esperar a quien está guardando.

**Tolerancia a fallos de red.** La recarga automática de maestras nunca reemplaza datos buenos por vacíos: si SQL Server o la carpeta de red fallan momentáneamente, la app conserva lo que ya tenía en memoria y sigue operando.

**Identificación de hojas por nombre, no por posición.** El consolidado se abre buscando la hoja por su nombre en vez de usar la pestaña activa, porque la pestaña activa depende de dónde quedó el cursor la última vez que alguien abrió el archivo en Excel.

**Columnas localizadas por encabezado.** Los datos se escriben buscando cada columna por su nombre en la fila de encabezados, tolerando diferencias de mayúsculas, tildes y espacios. Así el orden de las columnas puede cambiar sin que los datos se corran.

**Manejo de partes A/B.** Una misma orden puede traer dos trabajos distintos, y el contador de la máquina se reinicia al pasar de uno al otro. Cada parte lleva su propio acumulado para que la producción se calcule correctamente.

## Instalación

```bash
# 1. Clonar el repositorio
git clone https://github.com/tu-usuario/registro-paros-planta.git
cd registro-paros-planta

# 2. Crear entorno virtual e instalar dependencias
python -m venv venv
source venv/bin/activate      # En Windows: venv\Scripts\activate
pip install -r requirements.txt

# 3. Configurar el entorno
cp .env.example .env
# Editar .env con los datos de tu servidor y rutas

# 4. Ejecutar
python app.py
```

La aplicación queda disponible en `http://<ip-del-servidor>:5002`.

> **Requisito adicional en Windows:** ODBC Driver 17 for SQL Server.

## Configuración

Toda la configuración sensible se maneja mediante variables de entorno en un archivo `.env` (ver `.env.example`). El repositorio no contiene credenciales ni rutas reales de ningún entorno de producción.

## Estructura

```
registro-paros-planta/
├── app.py               # Rutas Flask y endpoints de la API
├── data_manager.py      # Lógica de negocio, acceso a datos y cálculos
├── config.py            # Configuración centralizada vía variables de entorno
├── requirements.txt
├── .env.example         # Plantilla de configuración
└── templates/
    └── form.html        # Interfaz del formulario (responsive, mobile-first)
```

## Notas

Este proyecto fue desarrollado como solución interna para un entorno de producción industrial. El código publicado aquí está desprovisto de credenciales, rutas y datos reales; los nombres de máquinas, áreas y archivos son genéricos.

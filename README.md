# Hikvision Hours Processor

Aplicacion interna en Python para procesar reportes de horas de Hikvision/iVMS-4200.

## Flujo

1. Leer archivo Hikvision (.xls HTML o .xlsx).
2. Validar columnas obligatorias.
3. Ordenar por ID de persona, Fecha, Registro de entrada y Horario.
4. Calcular minutos reales por tramo (salida - entrada).
5. Redondear cada tramo a bloques de 30 minutos.
6. Consolidar hojas Diario y Mensual.
7. Registrar inconsistencias sin detener el proceso.
8. Exportar Excel final con hojas Diario, Mensual e Inconsistencias.
9. Generar hoja `Liquidar` para estudio contable (matriz dia x empleado con colores por estado).

## Excepciones

- Si existe `date.xlsx` con hoja `Ausencias`, esas ausencias se cargan automaticamente al procesar.
  - Requiere columnas: `Legajo`, `Tipo ausencia`, `Fecha desde`.
  - Opcionales: `Fecha hasta`, `Estado`, `Observación`.
  - Los rangos `desde/hasta` se expanden por dia.
  - Estados `CANCELADO` y `RECHAZADO` se ignoran.
- Se pueden cargar desde archivo `.csv/.xls/.xlsx` con columnas:
  - `ID de persona` (opcional, vacio = aplica a todos)
  - `Fecha` (obligatorio, formato recomendado `YYYY-MM-DD`)
  - `Tipo` (obligatorio; por ejemplo `Feriado`, `Vacaciones`, `Enfermedad`, `Permiso`)
  - `Detalle` (opcional)
- La app agrega/normaliza la columna `Manual`. Cuando guardas o procesas, las lineas manuales se anexan al archivo de excepciones con `Manual = Si`.
- Tambien se pueden cargar manualmente desde la UI con formato por linea:
  - `ID|YYYY-MM-DD|TIPO|DETALLE`

## Estructura

- `main.py`: punto de entrada UI.
- `services/parser.py`: lectura y normalizacion de entrada.
- `services/calculator.py`: calculo de horas y deteccion de inconsistencias.
- `services/exporter.py`: exportacion de Excel.
- `services/processor.py`: orquestacion de proceso.
- `services/schedule_info.py`: lectura de horarios desde `date.xlsx`/`info.xlsx`.
- `ui/app.py`: interfaz desktop.
- `tests/`: pruebas unitarias.

## Plantillas de personal

- Para horarios y listado de empleados se prioriza `date.xlsx` (hoja `Empleados`) si existe.
- Si `date.xlsx` no existe, se usa `info.xlsx`.
- Se respetan los dias laborales por empleado (columna `Dias`).
  - Ejemplo: si un empleado trabaja `lunes miercoles y viernes`, martes/jueves sin fichada no se marcan como `Ausente`.

## Hoja Liquidar

- Se agrega automaticamente en cada reporte final.
- Estructura: `Fecha`, `Dia`, `Dia #` + una columna por empleado.
- Cada celda empleado/dia muestra solo horas normales (`(Minutos redondeados - Minutos extra) / 60`).
- Las horas extra no se incluyen en esta hoja.
- Se pintan celdas por estado (vacaciones, feriado, enfermedad, tardanza, etc.) respetando la paleta configurada.

## Ejecutar

```bash
pip install -r requirements.txt
python main.py
```

## Tests

```bash
python -m pytest -q
```

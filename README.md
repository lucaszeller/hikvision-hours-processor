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

## Excepciones

- Por defecto la app usa `feriados_nacionales_argentina_2026.xlsx` (en la carpeta del proyecto).
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
- `ui/app.py`: interfaz desktop.
- `tests/`: pruebas unitarias.

## Ejecutar

```bash
pip install -r requirements.txt
python main.py
```

## Tests

```bash
python -m pytest -q
```

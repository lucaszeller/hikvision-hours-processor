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

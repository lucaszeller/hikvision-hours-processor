# Hikvision Hours Processor

Aplicación de escritorio en Python para procesar fichadas exportadas desde relojes Hikvision y generar reportes automáticos de horas trabajadas.

## Qué resuelve
 codex/create-desktop-app-for-attendance-processing-18do52

- Toma el Excel exportado por Hikvision.
- Detecta columnas aunque cambien levemente los nombres.
- Agrupa fichadas por empleado y día.
- Empareja entrada/salida en secuencia.
- Calcula tramos y total diario.
- Genera resumen mensual por empleado.
- Lista inconsistencias para revisión administrativa.

## Estructura del proyecto

```text
hikvision-hours-processor/
├─ main.py
├─ requirements.txt
├─ hikvision_hours_processor.spec
├─ build_windows.bat
├─ domain/
│  ├─ column_aliases.py
│  └─ models.py
├─ services/
│  ├─ parser.py
│  ├─ calculator.py
│  ├─ exporter.py
│  └─ processor.py
├─ ui/
│  └─ app.py
├─ tests/
│  └─ test_calculator.py
└─ samples/
   ├─ generate_samples.py
   ├─ sample_input.xlsx
   └─ sample_output.xlsx
```

## Decisiones de diseño

1. **UI desacoplada del motor**: la interfaz (`ui/`) solo invoca `ProcessorService`.
2. **Motor reutilizable**: lógica de parsing/cálculo/exportación en `services/`, para reutilizar en una futura API web.
3. **Detección flexible de columnas**: matching por alias normalizados (sin tildes, case-insensitive).
4. **Inconsistencias visibles**: separación explícita en hoja dedicada.
5. **Salida trazable**: nombre de archivo con timestamp para no sobreescribir reportes anteriores.

## Requisitos

- Python **3.11+**
- Windows, macOS o Linux

## Instalación paso a paso

1. Clonar o descargar este repositorio.
2. Crear entorno virtual:

```bash
python -m venv .venv
```

3. Activarlo:

- Windows (PowerShell):

```powershell
.\.venv\Scripts\Activate.ps1
```

- Linux/macOS:

```bash
source .venv/bin/activate
```

4. Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Uso de la app

1. Ejecutar:

```bash
python main.py
```

2. Hacer clic en **Seleccionar archivo** y elegir el Excel de Hikvision.
3. Hacer clic en **Procesar**.
4. Al finalizar, la app muestra la ruta del Excel generado.

## Formato de salida

Se genera un Excel con 3 hojas (campos y títulos en español):

1. **Horas diarias**
   - Legajo
   - Nombre
   - Fecha
   - Cantidad de fichadas
   - Marcaciones (I/S)
   - Tramos trabajados
   - Horas totales
   - Minutos totales

2. **Resumen mensual**
   - Legajo
   - Nombre
   - Minutos totales
   - Horas mensuales

3. **Inconsistencias**
   - Legajo
   - Nombre
   - Fecha
   - Tipo de inconsistencia
   - Detalle

## Inconsistencias detectadas

- Cantidad impar de fichadas en el día.
- Duplicados cercanos (por defecto, <= 2 minutos).
- Orden inválido (salida anterior o igual a entrada).
- Datos incompletos (filas inválidas se descartan y si no queda nada útil se informa error).

## Ejemplos incluidos

Para regenerar ejemplos:

```bash
python samples/generate_samples.py
```

Archivos que se generan localmente (no versionados en Git):

- `samples/sample_input.xlsx`
- `samples/sample_output.xlsx`

## Tests

Ejecutar:

```bash
pytest -q
```

Cobertura básica incluida:

- cálculo de horas diarias y mensuales
- detección de cantidad impar de fichadas

## Preparación para ejecutable Windows

Ya se incluye configuración inicial de **PyInstaller**:

- `hikvision_hours_processor.spec`
- `build_windows.bat`

Para compilar en Windows:

```bat
build_windows.bat
```

El ejecutable se genera en `dist/`.

## Limitaciones conocidas de la v1

- Se toma la primera hoja del Excel (`read_excel` por defecto).
- No se interpreta lógica de turnos nocturnos cruzando medianoche.
- No hay configuración de reglas desde UI aún.


## Etiquetado de ingreso/salida

- En la hoja **Horas diarias**, la columna **Marcaciones (I/S)** indica explícitamente cada fichada como **Ingreso** o **Salida** en orden cronológico.
- En la columna **Tramos trabajados** también se muestra cada par como `Ingreso HH:MM - Salida HH:MM`.
=======

- Toma el Excel exportado por Hikvision.
- Detecta columnas aunque cambien levemente los nombres.
- Agrupa fichadas por empleado y día.
- Empareja entrada/salida en secuencia.
- Calcula tramos y total diario.
- Genera resumen mensual por empleado.
- Lista inconsistencias para revisión administrativa.

## Estructura del proyecto

```text
hikvision-hours-processor/
├─ main.py
├─ requirements.txt
├─ hikvision_hours_processor.spec
├─ build_windows.bat
├─ domain/
│  ├─ column_aliases.py
│  └─ models.py
├─ services/
│  ├─ parser.py
│  ├─ calculator.py
│  ├─ exporter.py
│  └─ processor.py
├─ ui/
│  └─ app.py
├─ tests/
│  └─ test_calculator.py
└─ samples/
   ├─ generate_samples.py
   ├─ sample_input.xlsx
   └─ sample_output.xlsx
```

## Decisiones de diseño

1. **UI desacoplada del motor**: la interfaz (`ui/`) solo invoca `ProcessorService`.
2. **Motor reutilizable**: lógica de parsing/cálculo/exportación en `services/`, para reutilizar en una futura API web.
3. **Detección flexible de columnas**: matching por alias normalizados (sin tildes, case-insensitive).
4. **Inconsistencias visibles**: separación explícita en hoja dedicada.
5. **Salida trazable**: nombre de archivo con timestamp para no sobreescribir reportes anteriores.

## Requisitos

- Python **3.11+**
- Windows, macOS o Linux

## Instalación paso a paso

1. Clonar o descargar este repositorio.
2. Crear entorno virtual:

```bash
python -m venv .venv
```

3. Activarlo:

- Windows (PowerShell):

```powershell
.\.venv\Scripts\Activate.ps1
```

- Linux/macOS:

```bash
source .venv/bin/activate
```

4. Instalar dependencias:

```bash
pip install -r requirements.txt
```

## Uso de la app

1. Ejecutar:

```bash
python main.py
```

2. Hacer clic en **Seleccionar archivo** y elegir el Excel de Hikvision.
3. Hacer clic en **Procesar**.
4. Al finalizar, la app muestra la ruta del Excel generado.

## Formato de salida

Se genera un Excel con 3 hojas:

1. **Horas diarias**
   - employee_id
   - employee_name
   - work_date
   - punch_count
   - segments
   - total_hours
   - total_minutes

2. **Resumen mensual**
   - employee_id
   - employee_name
   - total_minutes
   - monthly_total_hours

3. **Inconsistencias**
   - employee_id
   - employee_name
   - work_date
   - issue_type
   - details

## Inconsistencias detectadas

- Cantidad impar de fichadas en el día.
- Duplicados cercanos (por defecto, <= 2 minutos).
- Orden inválido (salida anterior o igual a entrada).
- Datos incompletos (filas inválidas se descartan y si no queda nada útil se informa error).

## Ejemplos incluidos

Para regenerar ejemplos:

```bash
python samples/generate_samples.py
```

Archivos esperados:

- `samples/sample_input.xlsx`
- `samples/sample_output.xlsx`

## Tests

Ejecutar:

```bash
pytest -q
```

Cobertura básica incluida:

- cálculo de horas diarias y mensuales
- detección de cantidad impar de fichadas

## Preparación para ejecutable Windows

Ya se incluye configuración inicial de **PyInstaller**:

- `hikvision_hours_processor.spec`
- `build_windows.bat`

Para compilar en Windows:

```bat
build_windows.bat
```

El ejecutable se genera en `dist/`.

## Limitaciones conocidas de la v1

- Se toma la primera hoja del Excel (`read_excel` por defecto).
- No se interpreta lógica de turnos nocturnos cruzando medianoche.
- No hay configuración de reglas desde UI aún.
 main

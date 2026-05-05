# AGENTS.md

Guía para agentes que trabajen en este repositorio (`hikvision-hours-processor`).

---

## Objetivo del proyecto

Aplicación interna en Python para procesar reportes de horas exportados desde Hikvision / iVMS-4200, calcular horas trabajadas por empleado, detectar inconsistencias y exportar resultados limpios en Excel.

El sistema está pensado para uso interno en empresa, con interfaz gráfica de usuario (UI).

---

## Formato de entrada

El archivo de entrada principal es un Excel exportado desde Hikvision.

El reporte siempre mantiene las mismas columnas:

- Índice
- ID de persona
- Nombre
- Departamento
- Posición
- Género
- Fecha
- Semana
- Horario
- Registro de entrada
- Registro de salida
- Trabajo
- Horas extra
- Asistió
- Entrada con retraso
- Temprano
- Ausente
- Permiso laboral

Ejemplo de fila:

ID de persona: 20  
Nombre: De Carli Gonzalo  
Fecha: 2026-04-01  
Horario: Mañana(07:30:00-12:00:00)  
Registro de entrada: 07:30:56  
Registro de salida: 12:02:12  

Cada fila representa un tramo de trabajo.

La app debe soportar empleados con:
- Horario cortado
- Horario corrido
- Múltiples tramos en un mismo día

---

## Identificación de empleados

El identificador único del empleado es:

- ID de persona

El campo Nombre es descriptivo y se utiliza solo para mostrar en reportes.

---

## Reglas de cálculo

No usar el campo `Trabajo` del Excel como fuente principal.

La app debe calcular los minutos trabajados como:

Registro de salida - Registro de entrada

Luego debe redondear el resultado a bloques de 30 minutos (0.5 horas).

Ejemplo:

Entrada: 07:30  
Salida: 12:02  
Tiempo real: 4h 32m  
Tiempo redondeado: 4h 30m  

Si hay más de una fila para el mismo empleado y la misma fecha, se deben sumar todos los tramos.

Ejemplo:

Mañana: 270 min  
Tarde: 270 min  
Total diario: 540 min / 9:00 hs  

---

## Ordenamiento

El Excel suele venir ordenado por `ID de persona`, pero la app no debe depender de ese orden.

Antes de procesar, ordenar por:

1. ID de persona  
2. Fecha  
3. Registro de entrada  
4. Horario  

---

## Inconsistencias

Ante errores o datos incompletos, la app debe marcar inconsistencia y continuar procesando.

No debe detener toda la ejecución por una fila problemática.

Detectar como mínimo:

- Falta Registro de entrada  
- Falta Registro de salida  
- Ambos vacíos  
- Salida menor que entrada  
- Fecha inválida  
- ID de persona vacío  
- Nombre vacío  

Las inconsistencias deben exportarse en una hoja separada.

---

## Excepciones

El sistema debe contemplar:

- Feriados  
- Ausencias justificadas  
- Enfermedad  
- Vacaciones  
- Permisos laborales  

Se deben manejar de dos formas:

1. Archivo externo de configuración  
2. Carga manual desde la UI  

Las excepciones deben impactar en los reportes finales y/o inconsistencias.

---

## Formato de salida

El sistema debe exportar un archivo Excel `.xlsx`.

Debe contener como mínimo estas hojas:

### Diario

- ID de persona  
- Nombre  
- Fecha  
- Departamento  
- Tramos trabajados  
- Minutos reales  
- Minutos redondeados  
- Horas totales  

### Mensual

- ID de persona  
- Nombre  
- Días trabajados  
- Minutos totales  
- Horas totales  

### Inconsistencias

- ID de persona  
- Nombre  
- Fecha  
- Tipo de inconsistencia  
- Detalle  

---

## Estructura del proyecto

- `main.py`: punto de entrada  
- `domain/`: modelos y constantes  
- `services/`: lógica de negocio  
- `tests/`: pruebas  
- `ui/`: interfaz gráfica  
- `samples/`: ejemplos  

---

## Entorno y comandos

Python 3.10+

Instalar dependencias:

pip install -r requirements.txt  

Ejecutar tests:

pytest -q  

Ejecutar app:

python main.py  

---

## Convenciones de desarrollo

- Mantener cambios simples y claros  
- No hacer refactors grandes sin necesidad  
- Separar lógica de negocio de UI  
- No agregar dependencias innecesarias  
- Agregar tests si se modifica lógica  

---

## Flujo de procesamiento

1. Leer archivo Excel  
2. Validar columnas  
3. Normalizar datos  
4. Ordenar  
5. Agrupar por empleado y fecha  
6. Calcular tramos  
7. Redondear a 30 minutos  
8. Aplicar excepciones  
9. Detectar inconsistencias  
10. Exportar Excel final  

---

## Prioridades

1. Correctitud del cálculo  
2. Compatibilidad con Hikvision  
3. Detección de errores  
4. Simplicidad  
5. Mantenibilidad  

---

## Validación antes de cambios

- Ejecutar tests: pytest -q  
- Ejecutar app: python main.py  
- Verificar Excel generado  

---

## Instrucciones para agentes IA

- No inventar formatos  
- Respetar columnas del Excel  
- No usar campo Trabajo  
- Calcular con entrada/salida  
- Redondear cada 30 minutos  
- Marcar inconsistencias y seguir  
- Mantener lógica separada de UI  
- Agregar tests si se cambia lógica  
- Evitar cambios innecesarios  
- Priorizar funcionamiento real sobre features nuevas  

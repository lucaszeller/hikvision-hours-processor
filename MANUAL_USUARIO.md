# Manual de Usuario - Hikvision Hours Processor

## 1. Objetivo

Este sistema procesa fichadas de Hikvision/iVMS, calcula horas por empleado y genera un Excel final listo para control interno y liquidacion.

Esta guia esta escrita para uso cotidiano. No hace falta conocimiento tecnico.

## 2. Que necesitas para usarlo

1. Archivo de fichadas exportado desde Hikvision (`.xls` o `.xlsx`).
2. Archivo `date.xlsx` en la carpeta del sistema (empleados, horarios y ausencias).
3. Ausencias y feriados configurados en `date.xlsx` (hoja `Ausencias`).

## 3. Abrir la aplicacion

1. Abri la app desde `abrir_app.bat` o ejecutando `python main.py`.
2. La app tiene 2 ventanas superiores:
   1. `Reporte`
   2. `Excepciones`

## 4. Flujo diario recomendado (paso a paso)

1. Ir a la ventana `Reporte`.
2. Presionar `Seleccionar archivo`.
3. Elegir el reporte de Hikvision del dia/periodo.
4. Verificar que el estado muestre archivo cargado.
5. Presionar `Procesar reporte`.
6. Esperar a `Proceso completado`.
7. Presionar `Abrir reporte generado`.

## 5. Resultado generado (hojas del Excel)

El archivo final se llama `reporte_horas_YYYYMMDD_HHMMSS.xlsx` y contiene:

1. `Diario`
   1. Detalle por empleado y fecha.
   2. Estado del dia.
   3. Minutos reales, redondeados, extras, horas totales.
2. `Mensual`
   1. Total de dias, minutos y horas por empleado.
3. `resumen-estudio`
   1. Dia de semana.
   2. Nombre del empleado.
   3. Horas normales.
   4. Horas extra.
4. `Liquidar`
   1. Matriz por fecha y empleado.
   2. Solo dias habiles (lunes a viernes).
   3. Totales por quincena separados en:
      1. Horas Comunes.
      2. Horas Extras.
   4. Totales de mes separados en:
      1. Horas Comunes.
      2. Horas Extras.

## 6. Como usar Excepciones

### 6.1 Archivo de excepciones

1. Ir a la ventana `Excepciones`.
2. Si queres usar otro archivo, presionar `Seleccionar archivo`.
3. Si queres volver al default, presionar `Restaurar archivo por defecto`.
4. Guardar con `Guardar configuracion`.

### 6.2 Carga manual rapida

En `Carga Rapida`:

1. Elegir empleado (`Todos` aplica a todos).
2. Elegir fecha inicio y fin (opcional) con `Calendario`.
3. Elegir tipo de excepcion.
4. Escribir detalle (opcional).
5. Presionar `Agregar a carga manual`.
6. Al terminar, presionar `Guardar configuracion`.

### 6.3 Formato por linea (manual)

Tambien podes pegar lineas en la caja manual:

`ID|YYYY-MM-DD|TIPO|DETALLE`

Ejemplo:

`20|2026-05-01|Feriado|Dia del trabajador`

Notas:

1. Si el ID esta vacio, aplica a todos.
2. Comentarios comienzan con `#`.

## 7. Reglas operativas importantes

1. El sistema calcula horas por `entrada/salida`. No usa el campo `Trabajo` como fuente principal.
2. Redondeo: a bloques de 30 minutos.
3. Si hay multiples tramos en un dia, se suman.
4. Tardanza:
   1. Si entra exactamente en su horario, queda `Normal`.
   2. Si entra 1 minuto despues, queda `Tarde`.
5. Sabados:
   1. Solo cuentan si hay fichada.
   2. Se consideran horas extra.
   3. La tardanza se evalua con inicio fijo 07:30.
6. Domingos:
   1. Solo aparecen si hay fichada.
   2. Se consideran horas extra.
7. Dias no laborales del empleado (segun `date.xlsx`):
   1. Si no hay fichada, no se marca ausente.
8. Si Hikvision marca `Ausente` pero hay fichada valida, prevalece la fichada.
9. Si el reporte trae varios meses, se procesa el mes predominante.

## 8. Colores de estados (visual en reportes)

1. Domingo: gris.
2. Ausencia/Ausente: rojo.
3. Vacaciones: amarillo.
4. Estudiar: celeste.
5. Capacitacion/Capacitaciones: rosado.
6. Suspencion/Suspension: azul.
7. No trabajado: marron claro.
8. Licencia/Enfermedad: naranja.
9. Feriado: verde manzana.
10. Accidente de trabajo/ART: violeta.
11. Tardanza/Tarde: verde oscuro.

## 9. Checklist rapido antes de procesar

1. El reporte de Hikvision corresponde al periodo correcto.
2. `date.xlsx` esta actualizado (empleados, horarios, dias, ausencias).
3. Si cargaste manuales en Excepciones, guardaste configuracion.
4. Ningun Excel de salida esta abierto (para evitar bloqueo al guardar).

## 10. Problemas frecuentes y solucion

### Problema: no me deja abrir calendario

Causa:

1. Falta dependencia `tkcalendar`.

Solucion:

1. Instalar dependencias del proyecto (`pip install -r requirements.txt`).

### Problema: un empleado aparece ausente y no deberia

Revisar:

1. Que tenga fichada completa o al menos fichada detectable en el archivo.
2. Que su horario/dias en `date.xlsx` sean correctos.
3. Que la fecha este dentro del mes procesado.

### Problema: no encuentro un empleado en Liquidar

Revisar:

1. Que exista en `date.xlsx` hoja `Empleados`.
2. Que el legajo este correcto.
3. Que no sea sabado/domingo (Liquidar solo muestra lunes a viernes).

### Problema: no guarda el reporte

Revisar:

1. Si el archivo destino esta abierto en Excel.
2. Permisos de carpeta.
3. Nombre de archivo con caracteres raros.

## 11. Buenas practicas de uso diario

1. Cerrar Excel antes de procesar.
2. Mantener `date.xlsx` actualizado semanalmente.
3. Cargar excepciones apenas ocurren, no al final del mes.
4. Verificar quincena en `Liquidar` antes de enviar al estudio contable.
5. Guardar una copia historica de cada reporte generado.

## 12. Nombre y ubicacion de archivos clave

1. Entrada fichadas: `report.xls` o export equivalente.
2. Empleados y horarios: `date.xlsx`.
3. Excepciones default: `feriados_nacionales_argentina_2026.xlsx`.
4. Reportes de salida: `reporte_horas_YYYYMMDD_HHMMSS.xlsx`.

---

Si queres, se puede preparar una version 1 pagina tipo "instructivo rapido" para imprimir y dejar al lado de la PC.


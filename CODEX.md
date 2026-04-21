# Proyecto: Procesador de fichadas Hikvision

## Objetivo
Construir una app de escritorio en Python que procese archivos Excel exportados desde relojes Hikvision y genere reportes diarios y mensuales de horas trabajadas.

## Contexto de negocio
El proceso hoy es manual y tarda aproximadamente 5 horas.
La meta es bajarlo a menos de 30 minutos.
El usuario final es administrativo, no técnico.

## Reglas de negocio
- No hay horarios fijos.
- Las horas se calculan a partir de las fichadas reales del día.
- Se deben sumar todos los tramos trabajados del día.
- Puede haber salidas e ingresos intermedios.
- El resumen mensual surge de sumar los totales diarios.
- Las inconsistencias deben ir a una hoja separada.

## Reglas técnicas
- Priorizar simplicidad y robustez
- Separar UI y lógica
- El motor de cálculo debe ser reutilizable para futura app web
- Usar pandas para transformación de datos
- Exportar a Excel con formato claro
- Preparar el proyecto para futura compilación a .exe en Windows

## UX mínima
- Selección de archivo
- Proceso con un clic
- Mensajes claros
- Salida Excel fácil de encontrar

## Entregables
- app funcional
- README
- requirements.txt
- ejemplo de entrada y salida
- tests
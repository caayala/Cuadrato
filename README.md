Cuadrato
===============

Funciones de Excel en VBA para traspaso y manejo de datos desde una tabla personalizada de SPSS a tablas específicas en una planilla Excel. El objetivo es ayudar a generar gráficos y tablas para informes.

Archivo: **`cuadrato.bas`** contiene las siguientes funciones:

- `cuadrato`: Calcula el error muestral de para una muestra aleatoria simple obtenida de una población infinita (mayor a 100.000 elementos)
    
	Sus variables son:

    * `fil1`, String: Nombre de pregunta en fila 1 de tabla personalizada
    * `fil2`, Variant: Nombre de alternativa de interés en la pregunta de interés, pasada en `fil1`.
    * `col1`, String: Nombre del segmento de interés.
    * `rgMat`, string: Nombre de la tabla desde donde se buscará la información.
    * `posicion`, Integer: Posición de la columna del dato de interés dentro del segmento pasado en `col1`. Por defecto `posicion = 0`.
    * `error`, Boolean: Si `error = TRUE` el resultado de la función será 0. Por defecto `error = FALSE`.

- `cuadratoRango`
- `cuadrato2`
- `cuadrato2Rango`

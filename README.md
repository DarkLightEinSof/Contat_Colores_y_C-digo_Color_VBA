Macros VBA para Contar y Obtener Color de Celdas
Este proyecto contiene dos funciones VBA útiles para trabajar con colores de celda en Microsoft Excel: una para contar celdas con un color específico y otra para obtener el código de color de una celda.
Funcionalidad
 * ContarVerdes(rango As Range): Esta función cuenta el número de celdas dentro de un rango dado que tienen un color de fondo específico (por defecto, verde estándar #00B050).
 * ColorCeldas(celdas As Range): Esta función auxiliar devuelve el código numérico de color RGB del color de fondo de la primera celda en un rango especificado. Esto es útil para identificar el código exacto de un color que deseas contar.
Cómo Usar
 * Abre tu Libro de Excel: Abre el archivo .xlsm (o .xlsb) donde quieres usar estas funciones.
 * Accede al Editor VBA: Presiona ALT + F11.
 * Inserta un Nuevo Módulo: En el explorador de proyectos (panel izquierdo), haz clic derecho en "VBAProject (TuLibro.xlsm)", ve a Insertar > Módulo.
 * Copia y Pega el Código: Copia todo el código de las funciones (ContarVerdes y ColorCeldas) y pégalo en el nuevo módulo.
Usando ContarVerdes en tu Hoja de Cálculo
Puedes usar ContarVerdes directamente en cualquier celda de Excel, como si fuera una fórmula estándar:
=ContarVerdes(A1:C10)

Esto contará las celdas de color verde estándar (#00B050) en el rango A1:C10.
Para contar un color diferente:
Primero, usa la función ColorCeldas para obtener el código del color que te interesa:
 * En una celda vacía (ej. E1), escribe =ColorCeldas(A1) (donde A1 es una celda con el color deseado).
 * La celda E1 mostrará un número (ej. 12611584). Este es el código de color.
 * Luego, modifica la función ContarVerdes en el Editor VBA, cambiando la línea:
   colorVerde = 5287936 ' Verde estándar (#00B050)

   por el nuevo código:
   colorVerde = 12611584 ' (Tu nuevo color)

   Guarda los cambios en VBA (Ctrl+S).
Usando ColorCeldas para Identificar Colores
Simplemente escribe la función en cualquier celda de tu hoja de Excel, refiriéndote a una celda con el color que deseas identificar:
=ColorCeldas(A1)

Esto te dará el código numérico del color de fondo de la celda A1.
Consideraciones
 * Estas funciones son UDFs (User-Defined Functions), lo que significa que recalculan cuando cambian los valores de las celdas a las que hacen referencia, pero no cuando cambia solo el color de una celda. Si cambias el color de una celda, es posible que necesites forzar un recálculo (presionando F9 o editando la fórmula) para que el conteo se actualice.
 * Los códigos de color son valores numéricos largos. Asegúrate de obtener el código exacto para el color que deseas contar.

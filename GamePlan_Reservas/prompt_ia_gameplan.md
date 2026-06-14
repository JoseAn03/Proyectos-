Actúa como un asistente avanzado en automatización de Excel. Debes procesar el archivo Excel adjunto y generar uno nuevo con las siguientes especificaciones exactas.

1. DIVISIÓN EN PESTAÑAS POR MARCA
- Lee la hoja activa del libro original.
- Crea tres nuevas pestañas con los nombres: ALAMO, ENTERPRISE, NATIONAL.
- Distribuye cada fila según el contenido de la columna Ubic. rec.:
  * ALAMO: si la celda contiene "E71" o "T71" (ej. SJOE71, SJOT71).
  * ENTERPRISE: si contiene "C61" o "T61".
  * NATIONAL: si contiene "E01" o "T01".
- Las filas que no coincidan con ninguna condición se ignoran.

2. ORDEN Y COLUMNAS DE SALIDA EN CADA PESTAÑA
Dentro de cada pestaña, ordena las filas alfabéticamente por la columna Nombre (A→Z). Las columnas deben aparecer en este orden exacto, con los siguientes encabezados:

A: Letra
B: N.º reserva
C: Marca
D: **
E: Nombre
F: Clase
G: **MAIN VIP**
H: Fecha recogida
I: Fecha entrega
J: Ubic. rec.
K: Ubic. ent.
L: Tarifa Diaria
M: Agencia de recom.

3. COLUMNA "Letra" (A)
- Agrupa las filas por la primera letra del Nombre (orden alfabético).
- En la primera fila de cada grupo, escribe en la columna A la letra mayúscula correspondiente (A, B, C...). Las demás filas del mismo grupo dejan la celda A vacía.
- La letra va en la misma fila del primer nombre del grupo (no insertar filas separadas).

4. COLUMNA "**" (D) – Combinación de indicadores
Evalúa dos condiciones sobre las columnas originales A y C (primera y tercera columna del archivo original):
   - Condición FL (col A original): si la celda contiene "~" o "*", asigna "⭐FL⭐".
   - Condición Marca (col C original): si la celda contiene "+", asigna la letra según la marca de la fila: Alamo → I, Enterprise → P, National → E.
- Resultado en columna D:
   * Solo FL: "⭐FL⭐"
   * Solo Marca: "I", "P" o "E"
   * Ambos: "⭐I-FL⭐", "⭐P-FL⭐" o "⭐E-FL⭐"
   * Ninguno: celda vacía.

5. COLUMNA "Marca" (C)
- Escribe el nombre de la marca con barra diagonal:
   * ALAMO/
   * ENTERPRISE/
   * NATIONAL/

6. COLUMNA "**MAIN VIP**" (G)
- Si Ubic. rec. termina exactamente en "E71", "E01" o "C61", asigna "* MAIN VIP*".
- Si Agencia de recom. contiene "11617270", asigna "* EXPEDIA*".
- Si se cumplen ambas: "* MAIN VIP* * EXPEDIA*".
- Si ninguna: celda vacía.

7. FORMATO EXIGIDO EN TODO EL DOCUMENTO
- Negrita: aplicar a TODAS las celdas de las tres pestañas.
-Fuente: Calibri
- Colores de fuente específicos:
   * Columna A (Letra): la letra en color ROJO y negrita.
   * Columna C (Marca): el texto con peso normal (anular la negrita general) y color:
        - ALAMO/ → AZUL (ej. #0070c0, similar a anciano oscuro 2)
        - ENTERPRISE/ → NEGRO (#000000)
        - NATIONAL/ → VERDE (ej. #00b050, similar a verde 11)
        - Incluir el formato (Negrita) para las tres marcas.
   * Columna D (**): todos los valores (I, P, E, ⭐FL⭐, combinaciones) en color ROJO.
   * Columna E (Nombre) colocar los siguientes colores para cada marca
        - ALAMO/ → AZUL (ej. #0070c0, similar a anciano oscuro 2)
        - ENTERPRISE/ → NEGRO (#000000)
        - NATIONAL/ → VERDE (ej. #00b050, similar a verde 11)
        - Incluir el formato (Negrita) para las tres marcas.
   * Resto de celdas: color negro automático, en negrita.

8. ENTREGA
Genera un nuevo archivo Excel (.xlsx) con las tres pestañas, manteniendo los valores calculados y todo el formato descrito. No alteres los datos originales más allá de lo especificado.
# Formateador TXT a CSV y Excel

Aplicación local en español. Requiere Python 3.10 o superior, sin instalar bibliotecas adicionales.

## Abrir

Haz doble clic en `iniciar.bat`. Se abrirá tu navegador. Mantén la ventana del programa abierta mientras lo usas; ciérrala para detenerlo.

1. Carga un TXT o pulsa **Probar con un ejemplo**.
2. Revisa los registros y las observaciones de lectura.
3. Marca o desmarca las columnas. Todas aparecen seleccionadas al cargar un archivo.
4. Usa **← →** junto a cada columna para moverla. El orden de la lista se refleja de izquierda a derecha en la vista previa y en las descargas. Puedes mover también las columnas desmarcadas y seleccionarlas después; conservan su posición. Al cargar otro archivo, se restablece el orden detectado.
5. Pulsa **Descargar CSV** para Google Sheets o **Descargar Excel**. La vista previa muestra hasta 100 registros; ambos formatos contienen todos, con solo las columnas seleccionadas y en el orden elegido.

También puedes ejecutar `python app.py` desde esta carpeta.

## Formatos de entrada

```text
name : Angel Reyes
email : AAA@gmail.com
telefono : 001234567

name : Maria Lopez email : maria@gmail.com
```

- Admite `:` y `=` entre etiqueta y valor.
- Una línea en blanco o una línea de tres o más guiones, asteriscos o guiones bajos separa registros.
- Una etiqueta repetida inicia otro registro y genera una observación.
- Las etiquetas con espacios se admiten al inicio de una línea. En medio de una línea admite etiquetas de una palabra y las expresiones `nombre completo`, `full name` y `correo electrónico`.
- Unifica name/nombre, email/correo y phone/teléfono. Los campos nuevos crean columnas.
- Conserva los campos ausentes como celdas vacías.
- El texto sin etiqueta se conserva y genera una observación para revisarlo.
- Codificaciones: UTF-8, UTF-16 con BOM y Windows-1252. Hasta 20 MB por solicitud.

## Límites de esta primera versión

La detección usa reglas. Los bloques separados por líneas en blanco se consideran registros distintos, incluso si la línea en blanco era decorativa. Un valor que contenga algo con forma de etiqueta (por ejemplo `nota: texto`) puede dividirse en otra columna. Revisa la vista previa y usa un formato consistente. Esta versión procesa un TXT a la vez y no guarda configuraciones entre sesiones.

El Excel almacena los valores como texto para conservar ceros iniciales y evitar que los ejecute como fórmulas. Incluye encabezados, filtros y la primera fila fija.

El CSV usa UTF-8 con BOM, comas como separador y comillas cuando un dato contiene comas, comillas o saltos de línea. Conserva los valores originales, incluidos ceros iniciales. CSV no almacena tipos de celda: al importarlo en Google Sheets, usa la opción de no convertir texto a números, fechas y fórmulas para mantener esos datos como texto. No incluye colores ni filtros guardados.

### Conservar los datos como texto en Google Sheets

Usa **Archivo → Importar → Subir**, selecciona el CSV, elige **Coma** como separador y desmarca **Convertir texto en números, fechas y fórmulas** antes de importar. Así se conservan valores como `00123`, `12345678901234567890`, `1E10`, `01/02`, `+59170000000` y `1.234,56`.

CSV no tiene una opción interna que obligue al programa que lo abre a tratarlo como texto. Poner todos los campos entre comillas tampoco fija su tipo. El exportador no añade apóstrofos ni fórmulas a los valores. Si prefieres un archivo con el tipo texto definido, usa **Excel · todo como texto**.

Si una importación ya cambió los datos, vuelve a importar el CSV original con la conversión desactivada. Cambiar el formato de las celdas después no recupera ceros o dígitos perdidos. El lector del TXT conserva los signos `;` y `|` que formen parte de los valores; solo recorta espacios exteriores de cada valor.

El servicio escucha únicamente en 127.0.0.1. No usa servicios externos ni guarda los TXT cargados en el servidor. Los archivos se guardan mediante las descargas del navegador.

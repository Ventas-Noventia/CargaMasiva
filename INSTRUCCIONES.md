# Noventia · Editor de catálogo

Proyecto en HTML, CSS y JavaScript, con archivos separados y código editable.

## Archivos

- `index.html`: estructura de la página, formularios, tabla y ventanas de edición.
- `styles.css`: diseño, colores y ajustes para computadora y celular.
- `app.js`: carga de archivos, eventos, edición, filtros y guardado local.
- `core.js`: lectura de Excel/CSV, conversión, IVA, validaciones y exportación.
- `plantilla-base.xlsx`: plantilla de carga incluida.
- `vendor/jszip.min.js`: biblioteca incluida para abrir y generar archivos XLSX.
- `vendor/JSZIP-LICENSE.md`: licencia de la biblioteca.

## Abrir en tu computadora

1. Descomprime el ZIP completo.
2. Abre la carpeta `Noventia_catalogo` en Visual Studio Code.
3. Usa la extensión **Live Server**: haz clic derecho sobre `index.html` y selecciona **Open with Live Server**.
4. En la página, carga tu inventario o presiona **Probar ejemplo**.

También puedes servir la carpeta con cualquier servidor estático. Si ya tienes Python 3, abre una terminal dentro de la carpeta y ejecuta:

```bash
python3 -m http.server 8080 --bind 127.0.0.1
```

Después abre `http://127.0.0.1:8080` en tu navegador. Detén el servidor con Ctrl+C cuando termines.

Usa un servidor local para que funcionen los módulos JavaScript y la lectura de la plantilla. Abrir el HTML mediante doble clic (`file://`) puede bloquear esas funciones.

## Subir a tu propio alojamiento

Sube todos los archivos y conserva la carpeta `vendor/` y la plantilla junto al HTML. La página es estática: no necesita instalación de paquetes, compilación ni una base de datos.

Esta copia no incorpora un inicio de sesión. Si la alojas en internet, el control de acceso depende del alojamiento que elijas.

## Uso

1. Carga un Excel `.xlsx` o un CSV de hasta 30 MB.
2. Elige la hoja y verifica la fila de encabezados.
3. Usa la plantilla incluida o carga otra compatible con sus columnas.
4. Relaciona las columnas de origen con SKU, nombre, descripción, precios, categoría y ubicación.
5. Configura el IVA y los valores comunes. Para conservar existencias, asigna la columna **Existencia** a **Cantidad de piezas**; si no se asigna, la plantilla base utiliza 1.
6. Aplica las reglas y revisa los productos. Puedes editar celdas, abrir cualquier producto o usar **Editar lote** sobre la selección, el filtro o toda la carga.
7. Resuelve las categorías y los almacenes que no coincidan con los catálogos de la plantilla.
8. Descarga todo el catálogo, las filas sin errores o los productos seleccionados.

Los SKU se conservan como texto para mantener los ceros iniciales. Los duplicados se señalan y se mantienen hasta que decidas excluir filas. Los precios faltantes quedan vacíos para su revisión. Las hojas auxiliares de la plantilla se conservan al exportar.

## Dónde cambiar las reglas iniciales

En `core.js`, el objeto `DEFAULT_RULES` establece IVA, condición, estado, moneda, almacén y las reglas de categorías y ubicaciones. Puedes cambiar estos valores en el código o desde la propia página.

El avance se guarda en el navegador del mismo dispositivo. Los archivos de inventario no se envían a un servidor. Para trasladar el trabajo a otra computadora, descarga el Excel y vuelve a cargarlo allí; si ya tiene IVA incluido, elige **Conservar el precio original** para evitar sumarlo de nuevo.

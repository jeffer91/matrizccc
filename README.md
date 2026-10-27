# Matriz CCC

Aplicación estática para convertir **tres archivos Excel institucionales** en tres copias nuevas **sin formato de celdas**.

## Cómo usarla

1. Abre la web publicada en GitHub Pages.
2. Selecciona en cualquier orden los tres archivos `.xls` o `.xlsx`: Redes base, PEA unidades y PEA actividades.
3. Comprueba las tarjetas y la vista previa de las primeras filas.
4. Usa **Descargar los 3 Excel** para obtener un ZIP con tres `.xlsx` limpios o descarga un archivo individual en su tarjeta.

La aplicación detecta cada archivo a partir de los encabezados, no a partir del nombre del archivo.

## Qué cambia y qué conserva

- Elimina colores, bordes, tipografías, celdas combinadas, formatos numéricos, anchos/altos personalizados, comentarios y características de presentación al reconstruir la hoja.
- Conserva los valores, textos, tildes y saltos de línea. Mantiene el orden de los registros y las columnas con datos.
- Elimina los espacios invisibles / finales **solo de los encabezados** y escribe los nombres de encabezado esperados.
- No elimina filas académicas repetidas ni modifica descripciones.
- Convierte fórmulas a valores almacenados si hubiera fórmulas. No recalcula fórmulas.
- Produce archivos separados sin hojas adicionales de informe.

## Esquemas esperados

**Redes base:** `codigoComponente`, `ordenComponente`, `descripcionComponente`, `descripcionComponente2`, `descripcionComponente3`.

**PEA unidades:** `codigoComponente`, `ordenComponente`, `descripcionComponente`.

**PEA actividades:** `nivel`, `Mecanismo`, `tema`, `descripcion`.

## Privacidad

Los archivos se abren y procesan en el navegador. No se suben al servidor ni se guardan en el repositorio.

## Tecnología y despliegue

HTML, CSS y JavaScript. SheetJS 0.18.5 para leer/exportar Excel; JSZip 3.10.1 para empaquetar. Se cargan desde CDNs públicos, por lo que el dispositivo necesita internet al abrir la aplicación. No requiere Node, backend ni base de datos.

El workflow `.github/workflows/pages.yml` publica el sitio en GitHub Pages desde la rama `main` mediante GitHub Actions.

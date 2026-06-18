# Facturación – Validación de pagos (PDF vs Excel)

Compara los **conteos de un PDF de facturación** contra la **planilla histórica de
Excel**, cruzando por fecha. Sirve para validar:

- **perfiles** – niveles/perfiles facturados (columna `Nivel/Perfil`).
- **equipos** – tipos de equipo.
- **servicios** – tipos de servicio.
- **equipos y servicios** – ambos a la vez, incluso si vienen en un mismo PDF que
  mezcla páginas de los dos formatos.

Para cada fecha y concepto se muestra el valor del PDF, el del Excel y un estado
(`OK` o `Valores diferentes`).

## Componente principal

`gui_validation_app.py` define la clase **`ServicesValidationApp`**. Métodos clave
que reutiliza la web:

- `_extraer_conteo_pdf(path, tipo)` – conteos por fecha desde el PDF. `tipo` puede
  ser `perfiles`, `equipos`, `servicios` o `equipos_servicios` (alias:
  `"equipos y servicios"`, `"todos"`).
- `_extraer_conteo_excel(path, fecha)` – conteos de la planilla para una fecha.
- `_clave_equipo` – clave de emparejamiento robusta: pliega tildes/mayúsculas,
  descarta signos y conjunciones (y/o/e/u) y **elimina espacios**, de modo que
  `(10H)` y `(10 H)` cruzan igual.
- `_extraer_valor_etiqueta` / `_limpiar_nombre_equipo` – leen el nombre del
  equipo/servicio de la etiqueta `EQUIPO:` / `SERVICIO:`. Exigen que la celda
  **sea** la etiqueta (no que la contenga, para no confundir la palabra
  “EQUIPOS” de un texto largo) y descartan texto duplicado que algunos PDFs
  arrastran tras el `)` (p. ej. `… (24 H) Motoso`).
- `_normalizar_perfil`, `_normalizar_fecha`, `_normalizar_busqueda` – normalización
  de texto para que el cruce sea robusto a tildes, mayúsculas y formato.

Equipos y servicios (formato vigente) comparten estructura (etiqueta + detalle
con `FECHA`/`CANTIDAD` por fila), por eso un único extractor por etiqueta sirve a
ambos; para servicios hay además un *fallback* al formato antiguo (fecha de
reporte + columna de servicio).

La clase también incluye una GUI de escritorio (tkinter). En la web se instancia
con `__new__` para reutilizar los métodos **sin** abrir ventanas.

## Uso

Se usa a través de la interfaz Streamlit del repositorio (ver el
[README raíz](../README.md)): herramienta **“Validación PDF + Excel”**, eligiendo
el tipo (perfiles / equipos / servicios / equipos y servicios) y subiendo el PDF
y el Excel. Los resultados se pueden **filtrar por fecha y por tipo**
(servicio/equipo/perfil), de forma independiente o combinada.

El Excel debe contener la columna `DESCRIPCION TARIFA` y columnas de fecha
(tipo fecha) con las cantidades por concepto.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `gui_validation_app.py` | Lógica de extracción/comparación + GUI de escritorio. |
| `revisar.ipynb` | Notebook de exploración (referencia, no producción). |
| `requirements.txt` | Dependencias del módulo. |

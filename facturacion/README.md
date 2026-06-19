# Facturación – Validación de pagos (PDF vs Excel)

Compara los **conteos de uno o varios PDFs de facturación** contra la **planilla
histórica de Excel**, cruzando por fecha. Sirve para validar:

- **perfiles** – niveles/perfiles facturados (columna `Nivel/Perfil`).
- **equipos y servicios** – equipos y/o servicios a la vez, incluso si vienen en
  un mismo PDF que mezcla páginas de los dos formatos.

Se pueden subir **varios PDF** en una misma validación; los conteos de todos se
acumulan antes de cruzarlos contra el Excel.

Para cada fecha y concepto se muestra el valor del PDF, el del Excel y un estado
(`OK` o `Valores diferentes`).

La comparación es **bidireccional**: además de verificar que lo del PDF esté en el
Excel, comprueba que lo del Excel (en la sección correspondiente) esté en el PDF.
Para no comparar contra secciones ajenas (perfiles, etc.), se detecta a qué
sección del histograma corresponde cada PDF por el **título de la página**
(p. ej. “…ELEMENTOS, HERRAMIENTAS Y EQUIPOS TRANSVERSALES” → `5.5`; “…OBRAS O
SERVICIOS TÍPICOS” → `5.6`) y solo se cruza esa sección. Una tarifa que esté en el
Excel con **valor 0** y sin registro en el PDF se considera válida (no se muestra);
una con valor > 0 ausente del PDF se marca como diferencia.

## Componente principal

`gui_validation_app.py` define la clase **`ServicesValidationApp`**. Métodos clave
que reutiliza la web:

- `_extraer_conteo_pdf(path, tipo)` – conteos por fecha desde el PDF. `tipo` puede
  ser `perfiles`, `equipos`, `servicios` o `equipos_servicios` (alias:
  `"equipos y servicios"`, `"todos"`).
- `_extraer_conteo_excel(path, fecha)` – conteos de la planilla para una fecha
  (valor **tal cual** del Excel, sin conversión de unidades).
- `_leer_histograma_largo(path, prefijos_cod=None)` – histograma en formato largo
  (`FECHA`, `DESCRIPCION TARIFA`, `CLAVE`, `VALOR`), **conservando ceros** y
  filtrable por sección (`COD. TAR.`). Es la base del cruce bidireccional.
- `_prefijos_seccion_pdf(path_hist, paths_pdf)` – casa el título de cada página de
  los PDF con la descripción del encabezado de sección del histograma y devuelve
  su `COD. TAR.` (p. ej. `5.5`, `5.6`).
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
el tipo (perfiles / equipos y servicios) y subiendo uno o varios PDF y uno o
varios Excel (la planilla histórica se arma sumando todos los Excel).
Los resultados se pueden **filtrar por fecha y por tipo**
(servicio/equipo/perfil), de forma independiente o combinada.

El Excel debe contener la columna `DESCRIPCION TARIFA` y columnas de fecha
(tipo fecha) con las cantidades por concepto.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `gui_validation_app.py` | Lógica de extracción/comparación + GUI de escritorio. |
| `revisar.ipynb` | Notebook de exploración (referencia, no producción). |
| `requirements.txt` | Dependencias del módulo. |

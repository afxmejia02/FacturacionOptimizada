# Facturación – Validación de pagos (PDF vs Excel)

Compara los **conteos de un PDF de facturación** contra la **planilla histórica de
Excel**, cruzando por fecha. Sirve para validar tres tipos de información:

- **perfiles** – niveles/perfiles facturados (columna `Nivel/Perfil`).
- **equipos** – tipos de equipo.
- **servicios** – tipos de servicio.

Para cada fecha y concepto se muestra el valor del PDF, el del Excel y un estado
(`OK` o `Valores diferentes`).

## Componente principal

`gui_validation_app.py` define la clase **`ServicesValidationApp`**. Métodos clave
que reutiliza la web:

- `_extraer_conteo_pdf(path, tipo)` – conteos por fecha desde el PDF.
- `_extraer_conteo_excel(path, fecha)` – conteos de la planilla para una fecha.
- `_normalizar_perfil`, `_normalizar_fecha`, `_normalizar_busqueda` – normalización
  de texto para que el cruce sea robusto a tildes, mayúsculas y formato.

La clase también incluye una GUI de escritorio (tkinter). En la web se instancia
con `__new__` para reutilizar los métodos **sin** abrir ventanas.

## Uso

Se usa a través de la interfaz Streamlit del repositorio (ver el
[README raíz](../README.md)): herramienta **“Validación PDF + Excel”**, eligiendo
el tipo (perfiles / equipos / servicios) y subiendo el PDF y el Excel.

El Excel debe contener la columna `DESCRIPCION TARIFA` y columnas de fecha
(tipo fecha) con las cantidades por concepto.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `gui_validation_app.py` | Lógica de extracción/comparación + GUI de escritorio. |
| `revisar.ipynb` | Notebook de exploración (referencia, no producción). |
| `requirements.txt` | Dependencias del módulo. |

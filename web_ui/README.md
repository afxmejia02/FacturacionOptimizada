# Web UI (Streamlit)

Interfaz **Streamlit** que reúne las revisiones del proyecto y reutiliza la
lógica de `facturacion/gui_validation_app.py`, `mapa-de-cargos/gui_app.py` y
`mapa-de-cargos/mano_obra.py` sin abrir ninguna ventana de escritorio.

Herramientas disponibles: validación de pagos, mapa de cargos (transferencias /
seguridad social) y **mapa de cargos – mano de obra** (Informe de Costo vs ODS).

## Estructura

| Archivo | Responsabilidad |
|---------|-----------------|
| `app.py` | Punto de entrada. Sólo UI: formularios, estado de sesión y render. |
| `processing.py` | Orquestación: guarda los archivos subidos, llama a los módulos de extracción y arma los resultados. |
| `pdf_export.py` | Construye el `.pdf` descargable (reportlab): multipágina, con título, nombres de los archivos ingresados y tablas coloreadas. Es el **único** formato de exportación. |
| `rendering.py` | Tablas HTML coloreadas y formato de valores (sin dependencia de Streamlit). En mano de obra colorea **solo la celda** inconsistente. |
| `codigos.py` | Conjunto `excluded_codes` usado al filtrar perfiles. |

El flujo es: `app.py` (UI) → `processing.py` (orquestación) → módulos de
`facturacion/` y `mapa-de-cargos/` (extracción) → `rendering.py` / `excel_export.py`
(presentación).

## Ejecutar

```bash
python -m pip install -r ../requirements.txt
streamlit run app.py
```

Abre la URL que muestra Streamlit, elige la herramienta y sube los archivos
(tipo equipos/servicios/perfiles, o el modo de conciliación con su formato
TABARCA/ITALCO).

## Notas

- La app importa los módulos de las carpetas hermanas; **no muevas**
  `facturacion/` ni `mapa-de-cargos/`.
- `processing.py` agrega la raíz del repo a `sys.path` para poder importarlos.
- Variable de entorno opcional `VALIDATION_DEBUG=0` para silenciar los logs de
  depuración.

# Facturación Optimizada

Herramientas para **optimizar la revisión de documentos de facturación y nómina**.
El proyecto compara automáticamente la información de archivos PDF contra las
planillas de Excel (o contra soportes de pago) y resalta las diferencias, de modo
que una revisión que antes era manual se hace en segundos.

La interfaz se publica con **Streamlit** para poder compartirla sin instalar nada.

## Revisiones disponibles

| Revisión | Carpeta | Qué compara |
|----------|---------|-------------|
| **Validación de pagos** (perfiles / equipos / servicios) | [`facturacion/`](facturacion/) | Conteos extraídos del PDF contra la planilla de Excel, cruzados por fecha. |
| **Mapa de cargos – transferencias** | [`mapa-de-cargos/`](mapa-de-cargos/) | Neto de los desprendibles contra los valores transferidos. |
| **Mapa de cargos – seguridad social** | [`mapa-de-cargos/`](mapa-de-cargos/) | Devengado de los desprendibles contra el IBC reportado. |

### Formatos: TABARCA e ITALCO

Cada revisión puede recibir documentos en dos formatos:

- **TABARCA** – formato usado por defecto (soportado en todas las revisiones).
- **ITALCO** – soportado en **transferencias** y **seguridad social**. Sus
  desprendibles son comprobantes de pago donde la cédula viene tras `CC:`, el
  neto tras `Total Neto:` y el devengado tras `TOTAL INGRESOS` (sin símbolo `$`);
  los soportes de transferencia son la consulta de pagos a terceros del banco y
  la seguridad social es la “Planilla Resumen” de aportes en línea. El selector
  de formato en la UI activa los parsers ITALCO.

## Estructura del repositorio

```
.
├── facturacion/        # Validación de pagos (PDF vs Excel)
│   └── gui_validation_app.py
├── mapa-de-cargos/     # Conciliación de nómina (desprendibles/transferencias/IBC)
│   └── gui_app.py
├── web_ui/             # Interfaz Streamlit (punto de entrega)
│   ├── app.py          # entry point delgado
│   ├── processing.py   # orquestación (reúsa los módulos de arriba)
│   ├── excel_export.py # generación del Excel descargable
│   └── rendering.py    # tablas HTML y formato de valores
└── requirements.txt
```

Las carpetas `facturacion/` y `mapa-de-cargos/` contienen además la lógica original
en clases (`ServicesValidationApp`, `PayrollReconciliationApp`). La web reutiliza
esos métodos sin abrir las ventanas de escritorio (tkinter).

## Ejecutar localmente

```bash
python -m venv .venv && .venv\Scripts\activate   # Windows
pip install -r requirements.txt
streamlit run web_ui/app.py
```

Abre la URL que muestra Streamlit, elige la herramienta y sube los archivos.

## Desplegar en Streamlit Community Cloud

1. Conecta este repositorio en <https://share.streamlit.io>.
2. **Main file path:** `web_ui/app.py`.
3. **Requirements:** `requirements.txt` (raíz).

## Notas para desarrollo

- Los notebooks (`*.ipynb`) son material de exploración/referencia; la lógica de
  producción vive en los `.py`.
- El entorno virtual y las salidas generadas (`Lib/`, `Scripts/`, `build/`,
  `dist/`, `etc/`, `docs/`) están excluidos por `.gitignore`. Las carpetas `docs/`
  contienen documentos reales de nómina y **no** deben subirse.

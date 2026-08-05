# facturacion

Compara los conteos de uno o varios **PDF de facturación** contra la **planilla
histórica de Excel**, cruzando por fecha. Sin interfaz: la UI vive en `web_ui/`.

| Módulo | Qué hace |
|---|---|
| `normalizacion.py` | Formas canónicas de textos, fechas y cantidades. |
| `pdf.py` | Extracción de conteos desde el informe PDF. |
| `histograma.py` | Lectura del Excel histórico. |

```python
from facturacion import extraer_conteo_pdf, leer_histograma_largo
```

Valida **perfiles** (columna `Nivel/Perfil`) y **equipos y servicios**, incluso
si vienen mezclados en un mismo PDF. Se pueden subir varios PDF: los conteos se
acumulan antes de cruzar. Para cada fecha y concepto se muestra el valor del PDF,
el del Excel y un estado (`OK` o `Valores diferentes`).

El Excel debe traer la columna `DESCRIPCION TARIFA` y columnas de fecha con las
cantidades por concepto.

## Por qué el código está escrito así

### El cruce es bidireccional, pero acotado por sección

No basta con verificar que lo del PDF esté en el Excel: también se comprueba lo
contrario. Para no comparar contra secciones ajenas se detecta a qué sección del
histograma corresponde cada PDF por el **título de la página**
(“…ELEMENTOS, HERRAMIENTAS Y EQUIPOS TRANSVERSALES” → `5.5`; “…OBRAS O SERVICIOS
TÍPICOS” → `5.6`) y solo se cruza esa sección.

Una tarifa que esté en el Excel con **valor 0** y sin registro en el PDF es
válida (no se muestra); con valor > 0 y ausente del PDF, es una diferencia.

En **perfiles** el lado Excel se toma como “todo lo que **no** es equipos ni
servicios”: se excluyen los `COD. TAR.` `5.5` y `5.6`, en vez de exigir que la
descripción diga “Nivel/Perfil”. Así entran tarifas de mano de obra que no usan
esa palabra (p. ej. `Inspector certificado: API/ASME NACIONAL`, cód. `5.4`).

### Por qué se normaliza tanto antes de comparar

Los mismos conceptos se escriben distinto en cada archivo. `clave_equipo` pliega
tildes y mayúsculas, descarta signos y conjunciones (y/o/e/u) y **elimina
espacios**, de modo que `(10H)` y `(10 H)` cruzan igual.

### Cantidades: convención colombiana

Tanto el PDF como el Excel usan **punto para miles y coma para decimales**:
`3.139` → 3139, `3.139,00` → 3139, `1.452,6` → 1452.6, `153,67` → 153.67. Los
resultados se muestran y exportan en ese mismo formato.

### Leer el nombre del equipo: dos trampas

`extraer_valor_etiqueta` exige que la celda **sea** la etiqueta (`EQUIPO:`), no
que la contenga — si no, la palabra “EQUIPOS” dentro de un texto largo hace tomar
la celda equivocada.

El texto sobrante tras el último `)` se descarta **solo si es un fragmento
inicial duplicado del propio nombre** (artefacto de superposición del PDF, p. ej.
`… (24 H) Motoso`). Una continuación legítima como `Torno … (Diurno / Nocturno)
para bridas >4 NPS <= 48 NPS` **se conserva**: sí está en el PDF y coincide con
el Excel.

### La columna Observaciones de perfiles

`parsear_observacion_perfil` devuelve `(recategorizado, es_ef, no_facturable,
es_24h)`. Los marcadores pueden coexistir y venir en cualquier orden:

| Marcador | Efecto |
|---|---|
| `RECATEGORIZADO … COMO <nivel>` | El turno se cuenta en el nivel indicado. |
| `E y F` | Aunque la jornada sea de 24 h, cuenta como **1 unidad**. |
| `NO FACTURABLE` | La fila **no cuenta**. |
| `24` / `24H` / `24HRS` / `24 HORAS` | Jornada de 24 h: el turno cuenta **1/3**. |

La jornada de 24 h se detecta tanto al inicio de la hoja (`tabla[4][2]`) como en
esta columna.

### Un solo extractor para equipos y servicios

Ambos comparten estructura (etiqueta + detalle con `FECHA`/`CANTIDAD` por fila),
así que un único extractor por etiqueta sirve a los dos. Para servicios queda un
*fallback* al formato antiguo (fecha de reporte + columna de servicio).

## Pruebas

```bash
python -m unittest discover -s facturacion -p "test_*.py" -t .
```

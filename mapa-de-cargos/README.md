# Mapa de cargos – Conciliación de nómina

Concilia los **desprendibles de nómina** contra dos fuentes:

- **Transferencias / soportes bancarios** – ¿el neto pagado coincide con lo
  transferido por persona?
- **Seguridad social (IBC)** – ¿el devengado coincide con el IBC reportado?

## Componente principal

`gui_app.py` define la clase **`PayrollReconciliationApp`**. Métodos clave que
reutiliza la web:

| Método | Rol |
|--------|-----|
| `_process_desprendibles(folder, formato)` | Extrae `Identificacion`, `Neto`, `Devengado`, `Cuenta`. Enruta a TABARCA o ITALCO. |
| `_process_transferencia(folder, formato)` | Extrae los valores transferidos. Enruta a TABARCA o ITALCO. |
| `procesar_seguridad_social(folder, formato)` | Extrae CC e IBC de las planillas. Enruta a TABARCA o ITALCO. |
| `_reconcile_data(despr, trans, seg)` | Agrupa por cédula y compara sumas; devuelve `(df_transfers, df_seguridad)`. |

### Formatos TABARCA vs ITALCO

`_process_desprendibles` y `_process_transferencia` reciben un parámetro
`formato` (`"tabarca"` por defecto, o `"italco"`):

- **TABARCA** – desprendibles tipo “Comprobante de Nómina” con `Neto a Pagar $…`.
- **ITALCO** – desprendibles tipo “Comprobante de pago de Nomina”: la cédula va
  tras `CC:`, el neto tras `Total Neto:` y el **devengado** tras `TOTAL INGRESOS`
  (sin `$`). Los soportes de transferencia son la “consulta de pagos a terceros”
  del banco (líneas tipo `<nombre> <doc> [cuenta] [fecha] [factura] PAGO NOMINA
  BCA <valor>`). La seguridad social es la “Planilla Resumen” de aportes en
  línea: el documento está en la columna 2 y el IBC de pensión en la columna 26
  (página 1) o 27 (páginas siguientes).

  El renglón de transferencia se extrae de forma **robusta y genérica**
  (`_match_linea_transferencia`), no por un layout fijo: detecta el documento
  (primer grupo de 5-15 dígitos), el valor (primer importe **después** de la
  etiqueta de destino, para no confundirlo con la columna `ods`/consecutivo que
  algunos soportes ponen al final), la cuenta (primer grupo de 9+ dígitos,
  opcional) y la fecha de la **factura** (`YYMMDD` antes de `PAGO`, que indica la
  quincena). Tolera la ausencia de columna de fecha de pago, distintos formatos
  monetarios, espacios y ruido de OCR, y variantes de la etiqueta de destino
  (basta que aparezca “NÓMINA”).

  Las líneas se reconstruyen a partir de **palabras y coordenadas**
  (`_lineas_desde_palabras`), no de `extract_text()`: algunos soportes traen una
  columna “Productos” que se entrelaza con el nombre y pega el documento al
  número de producto; reconstruir por `x` separa de nuevo los campos. Si una
  página no trae renglones, hay un *fallback* a soporte tipo desprendible
  (`CC:` / `Total Neto:`).

  El cruce es por documento o por cuenta (normalizando ceros de relleno) y,
  además, **se filtra por periodo**: de cada desprendible se extrae
  `Periodo: <inicio> al <fin>` y solo se conservan las transferencias cuya fecha
  de factura cae dentro de esa ventana. Así no se suman quincenas/meses ajenos
  (p. ej. una transferencia de marzo o mayo al conciliar abril). Hay logs que
  explican cada “Transferencia no encontrada” y cada transferencia descartada
  por periodo.

> Importante: el formato de los desprendibles debe coincidir con el de las
> transferencias / seguridad social. La web pasa el mismo `formato` a todos los
> parsers (desprendibles, transferencias y seguridad social).

### Conciliación de seguridad social (devengado vs IBC)

Para ambos formatos, `_reconcile_data` agrupa los desprendibles por cédula, suma
el **devengado** y lo compara contra los **IBC** reportados en la planilla:

- `OK` – el devengado total coincide con un IBC único.
- `Devengado no coincide` – no hay coincidencia (o falta el IBC).
- `IBC sin soporte` – coincide pero hay más de un IBC para la cédula.
- `Devengado no encontrado` – el desprendible no traía devengado.

### Mano de obra (Informe de Costo vs ODS)

`mano_obra.py` define **`comparar_mano_obra(informe, ods)`**, que cruza el
**Informe de Costo** (nómina) contra el registro de la **ODS** (empleados del
contrato) por número de documento. Ambos Excel traen la misma información bajo
nombres de columna distintos. El mapeo (`MAPEO_COLUMNAS`, con tipo de
comparación) es:

| Concepto | Informe | ODS | Tipo |
|----------|---------|-----|------|
| Documento (clave) | `Identificación` | `NumeroDocumento` | dígitos |
| OS | derivada de `Nombre Centro Costo` (`…Os050…`→50) | `No_de_orden_de_servicio_conocido_por_el_contratista` | número |
| Nombres / Apellidos | `Nombres` / `Apellidos` | `Nombres` / `Apellidos` | texto |
| Cargo | `Cargo` | `CargoContratoLaboral` | texto |
| Fecha Inicio | `Fecha Inicio` | `Fecha_de_inicio_de_actividades_…` | fecha |
| Fecha Vencimiento | `Fecha Vencimiento` | `Fecha_fin_de_actividades_…` | fecha |
| Días Trabajados | `Días Trabajados` | `DiasTrabajadosEnMes` | número |
| Salario | `Salario Diario Contratado` | `SalarioDiarioPesos` | moneda |

> **Fechas:** los campos de actividades de la ODS son la vigencia del
> **contrato**, por eso se comparan contra `Fecha Inicio` / `Fecha Vencimiento`
> del Informe, **no** contra `Fecha de Ingreso` / `Fecha de retiro` (vínculo
> laboral, un concepto distinto que generaba falsos positivos).

> **Salario:** se compara como `moneda` (valor numérico normalizado, tolerante a
> `$`, separadores de miles/decimales y espacios) y se muestra en pesos
> colombianos (`$120.000`). Cuando difiere, la celda lista ambos: **Informe** y
> **Lista ODS**. Las utilidades `normalizar_moneda` / `formatear_cop` son
> reutilizables.

El Informe se lee de la hoja `Informe` con el encabezado real en la fila 10;
como varias columnas con tildes llegan corruptas en el xlsx, las columnas usadas
se referencian **por posición** y se renombran. El resultado es un DataFrame
donde cada campo es una **lista**: `[valor]` si ambos coinciden,
`[valor_informe, valor_ods]` si difieren. La comparación normaliza por tipo
(texto sin acentos, fecha por día, número entero, moneda), de modo que
`2025-06-08 00:00:00` y `2025-06-08` se consideran iguales. **No hay columna de
estado/observaciones**: el resaltado por celda (listas de dos elementos) es el
único indicador de inconsistencia.

## Uso

A través de la interfaz Streamlit (ver el [README raíz](../README.md)):
herramientas **“Mapa de cargos - transferencias”**,
**“Mapa de cargos - seguridad social”** y **“Mapa de cargos - mano de obra”**.
En transferencias se elige el formato (TABARCA / ITALCO). En mano de obra se
suben los dos Excel (Informe y ODS) y se resalta **solo la celda** del campo
inconsistente, no toda la fila. Todos los resultados se descargan como Excel.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `gui_app.py` | Lógica de conciliación (transferencias / seguridad social) + GUI de escritorio. |
| `mano_obra.py` | Comparación Informe de Costo vs ODS (mano de obra). |
| `main.ipynb` | Notebook de referencia (incluye la exploración del formato ITALCO). |
| `ssocial.ipynb` | Notebook de exploración de seguridad social. |
| `mano-obra.ipynb` | Notebook de exploración de la comparación de mano de obra. |
| `requirements.txt` | Dependencias del módulo. |

> La carpeta `docs/` contiene documentos reales de nómina y está excluida por
> `.gitignore`; no se versiona.

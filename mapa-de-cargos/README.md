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
  `Periodo: <inicio> al <fin>` y solo se conservan las transferencias **confiables**
  (con etiqueta NÓMINA) cuya fecha de factura cae dentro de esa ventana. Así no se
  suman quincenas/meses ajenos (p. ej. una transferencia de marzo o mayo al
  conciliar abril).

  Algunos soportes vienen en otro layout (la “consulta de pagos a terceros” del
  banco) cuyos renglones **no traen etiqueta NÓMINA ni la fecha-factura de
  quincena** (solo la fecha de consignación, que puede ser de otro mes). Esos
  renglones se extraen como **candidatos** (`EsNomina=False`) y se aceptan **solo
  si su valor coincide con un neto del desprendible** aún no cubierto por una
  transferencia confiable; los importes que no estén en los netos se descartan
  (pueden ser de otra quincena, o conceptos como prima). Así el pago real se
  rescata aunque el renglón no traiga etiqueta/fecha, sin introducir falsos
  positivos. Hay logs que explican cada “Transferencia no encontrada”, cada
  transferencia descartada por periodo y cada candidata descartada por valor.

> Importante: el formato de los desprendibles debe coincidir con el de las
> transferencias / seguridad social. La web pasa el mismo `formato` a todos los
> parsers (desprendibles, transferencias y seguridad social).

### Conciliación de seguridad social (devengado vs IBC)

Para ambos formatos, `_reconcile_data` agrupa los desprendibles por cédula, suma
el **devengado** y lo compara contra la **suma de los IBC** reportados en la
planilla (misma lógica de suma que transferencias):

- `OK` – la suma de devengados coincide con la suma de IBC.
- `Devengado no coincide` – las sumas difieren (o falta el IBC).
- `Devengado no encontrado` – el desprendible no traía devengado.

> **Devengados repetidos:** en seguridad social los devengados **no se
> deduplican** (`_normalizar_lista_completa`): dos quincenas con el mismo
> devengado se suman ambas. (La deduplicación, vía `_normalizar_lista`, se aplica
> solo al cruce de **transferencias**.)

> **IBC sumados:** todos los IBC de la cédula se suman antes de comparar, no se
> exige un IBC único que coincida por sí solo (p. ej. `182.210 + 8.274.618 =
> 8.456.828` cruza con un devengado de `8.456.828`).

Tanto transferencias como seguridad social incluyen una columna **`Diferencia`**:
neto − transferencia en el primero, devengado − IBC en el segundo (lo ausente
cuenta como 0). Aplica a los formatos TABARCA e ITALCO.

### Mano de obra (Informe de Costo vs ODS)

`mano_obra.py` define **`comparar_mano_obra(informe, ods)`**, que cruza el
**Informe de Costo** (nómina) contra el registro de la **ODS** (empleados del
contrato) por número de documento. Cada parámetro acepta una ruta/buffer o una
**lista** de varios Excel por lado: cada archivo se lee por separado
(`_leer_y_concatenar`) y se concatena antes de cruzar, de modo que una persona
de cualquier Informe puede emparejarse con cualquier ODS. Ambos Excel traen la misma información bajo
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

El Informe se lee de la hoja `Informe` con el encabezado real en la fila 10; las
columnas usadas se localizan **por nombre** (normalizado: mayúsculas, sin acentos
y espacios colapsados, con alias) en vez de por posición, porque el layout varía
entre exportes mensuales (distinto número y orden de columnas). Si una columna
requerida no existe en ese archivo (p. ej. el exporte de abril no trae
`Días Trabajados` ni `Salario Diario Contratado`), el campo queda **vacío** en vez
de tomar por error otra columna en esa posición. El resultado es un DataFrame
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
suben uno o varios Excel por lado (Informe y ODS) y se resalta **solo la celda**
del campo inconsistente, no toda la fila. Todos los resultados se descargan como Excel.

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

# nomina

Conciliación de nómina y comparación de mano de obra. Sin interfaz: la UI vive
en `web_ui/`.

| Módulo | Qué hace |
|---|---|
| `formato.py` | Limpieza de números, documentos y líneas de texto de los PDF. |
| `desprendibles.py` | Lee los desprendibles (PDF), formatos TABARCA e ITALCO. |
| `transferencias.py` | Lee las transferencias bancarias y empareja sus líneas. |
| `seguridad_social.py` | Lee las planillas de seguridad social (IBC). |
| `conciliacion.py` | Cruza los tres orígenes (`conciliar`). |
| `mano_obra.py` | Informe de Costo contra el registro de la ODS. |

```python
from nomina import procesar_desprendibles, procesar_transferencias, conciliar
```

Concilia los desprendibles contra dos fuentes: **transferencias** (¿el neto
pagado coincide con lo transferido?) y **seguridad social** (¿el devengado
coincide con el IBC?).

## TABARCA vs ITALCO

Casi todas las funciones reciben `formato` (`"tabarca"` por defecto o
`"italco"`) porque los PDF de cada cliente tienen otro layout:

- **TABARCA** – desprendibles “Comprobante de Nómina”, con `Neto a Pagar $…`.
- **ITALCO** – “Comprobante de pago de Nomina”: la cédula va tras `CC:` (si ese
  campo viene vacío, se toma de `Documento <número>`), el neto tras
  `Total Neto:` y el devengado tras `TOTAL INGRESOS`. Los importes se leen con o
  sin `$`.

## Por qué el código está escrito así

Decisiones que no se deducen leyendo el código, y conviene entender antes de
cambiarlo.

### Las columnas se buscan por nombre, nunca por posición

El layout de los Excel **cambia entre exportes mensuales**: distinto número y
orden de columnas, e incluso distinta fila de encabezado (se ha visto la fila 10
un mes y la 7 al siguiente). Por eso la fila del encabezado se **detecta**
(primera fila que trae a la vez una columna de identificación y `Nombres`) y cada
columna se localiza por **nombre normalizado** (mayúsculas, sin acentos, espacios
colapsados) con una lista de alias.

Si una columna no aparece, el campo queda **vacío**. Es deliberado: preferible a
tomar por error la columna que esté en esa posición.

### Contra qué fechas se compara

Las fechas de actividades de la ODS son la vigencia del **contrato**. Se comparan
contra `Fecha Inicio` / `Fecha Vencimiento` del Informe — **no** contra
`Fecha de Ingreso` / `Fecha de retiro`, que son del vínculo laboral y son otra
cosa. Confundirlos genera diferencias falsas en masa.

### Documentos: cuidado con los números de Excel

El documento puede llegar como texto (`1.096.198.448`) o como número. Basta un
vacío en la columna para que pandas lea toda la columna como `float64`, y
entonces `str(1096198448.0)` deja un `.0`; si solo se quitan los no dígitos, el
documento **gana un `0` final** y no cruza con nadie. Por eso `solo_digitos` y
`_os_comparable` convierten a `int` los float enteros. Igual con la OS: `37.0`
daría `"0"`, porque se toma el último grupo de dígitos.

### El Informe ITALCO (la "progresión")

- Una sola hoja cuyo nombre incluye el mes (`PROGRESION JULIO 2025`), así que se
  lee la primera hoja, no una llamada `Informe`.
- La primera columna no tiene nombre y marca `ACTUAL` / `ANTERIOR` /
  `DIFERENCIA` por persona. Solo interesan las de **DIFERENCIA**.
- El salario diario **no** sale de la fila DIFERENCIA (ahí es un delta): se toma
  el `Sueldo Base` de `ACTUAL` dividido entre 30 (convención de nómina
  colombiana), que coincide con el `SalarioDiarioPesos` de la ODS.
- El nombre viene como un único `Nombre Completo`; en la ODS se compara contra
  `Nombres + Apellidos`.
- La OS puede venir en `Perfil Contable` (`BCA OS 37 CONVENCIONAL`) o en una
  columna `OS` con el número suelto. Se aceptan ambas y se comparan por su número
  contra el `0DS37` de la ODS.
- El cargo trae el marcador `(PROGRE)` que la ODS no tiene, y expresa
  alternativas separadas por `/` (`ANDAMIERO B / D8`). Se ignora el marcador y
  basta que **una** alternativa coincida.

### El resultado son listas, no textos

`comparar_mano_obra` devuelve una fila por persona y, en cada campo, una lista:
`[valor]` si Informe y ODS coinciden, `[informe, ods]` si difieren. Esa lista es
la única fuente de verdad: la web resalta la celda cuando tiene dos elementos. No
hay columna de "estado".

### Transferencias: por qué tanta tolerancia al ruido

Los soportes bancarios llegan escaneados y con OCR. La etiqueta de destino
(`PAGO NOMINA BCA`) solo sirve como señal de que la línea es un pago; el cruce
nunca depende de su forma exacta y tolera confusiones típicas (O/0, I/1, A/4).
El cruce se hace por documento **o** por cuenta, y solo se marca "no encontrada"
cuando de verdad no hay coincidencia.

## Pruebas

```bash
python -m unittest discover -s nomina -p "test_*.py" -t .
```

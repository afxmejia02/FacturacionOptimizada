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
  del banco (líneas `… PAGO NOMINA BCA <valor>`). La seguridad social es la
  “Planilla Resumen” de aportes en línea: el documento está en la columna 2 y el
  IBC de pensión en la columna 26 (página 1) o 27 (páginas siguientes).

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

## Uso

A través de la interfaz Streamlit (ver el [README raíz](../README.md)):
herramientas **“Mapa de cargos - transferencias”** y
**“Mapa de cargos - seguridad social”**. En transferencias se elige el formato
(TABARCA / ITALCO). El resultado se puede descargar como Excel con estilos.

## Archivos

| Archivo | Descripción |
|---------|-------------|
| `gui_app.py` | Lógica de conciliación + GUI de escritorio. |
| `main.ipynb` | Notebook de referencia (incluye la exploración del formato ITALCO). |
| `ssocial.ipynb` | Notebook de exploración de seguridad social. |
| `requirements.txt` | Dependencias del módulo. |

> La carpeta `docs/` contiene documentos reales de nómina y está excluida por
> `.gitignore`; no se versiona.

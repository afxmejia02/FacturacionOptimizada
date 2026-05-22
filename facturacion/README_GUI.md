## Sistema de Validación de Servicios y Equipos - Aplicación GUI

### Resumen
Esta aplicación GUI en Tkinter valida registros de servicios y equipos de archivos PDF contra datos históricos almacenados en Excel. Identifica diferencias entre ambas fuentes y muestra un reporte detallado.

### Funcionalidades

#### 1. **Selección de archivos**
   - Buscar y seleccionar archivos PDF
   - Buscar y seleccionar archivos Excel
   - Compatible con formatos .pdf, .xlsx y .xls

#### 2. **Extracción automática de datos**
   - Extrae tipos de servicio/equipo y cantidades desde tablas en PDF
   - Reconoce varios formatos de fecha (español e internacional)
   - Maneja múltiples tablas por página PDF
   - Detección flexible de columnas en ambas fuentes

#### 3. **Validación de datos**
   - Compara servicios entre PDF y Excel por nombre y fecha
   - Verifica cantidades para detectar diferencias
   - Genera reporte de discrepancias
   - Calcula diferencias absolutas

#### 4. **Visualización de resultados**
   - Tabla interactiva con todas las diferencias
   - 🔴 Filas en rojo para errores
   - Muestra fecha, descripción, cantidad PDF, cantidad Excel y diferencia
   - Columnas desplazables para ver toda la información

#### 5. **Notificación de éxito**
   - Muestra mensaje de "Todo correcto" cuando la validación pasa
   - No muestra errores si las cantidades coinciden
   - Estado claro en la interfaz

#### 6. **Exportación de datos**
   - Exporta discrepancias a CSV
   - Conserva toda la información de validación
   - Permite guardar en una ubicación personalizada

#### 7. **Retroalimentación al usuario**
   - Indicadores de estado durante el proceso
   - Mensajes de error para archivos inválidos o fallos de procesamiento
   - Conteo de discrepancias
   - Actualización en tiempo real del estado

### Instalación y requisitos

1. **Paquetes de Python**:
   ```bash
   pip install pandas pdfplumber openpyxl
   ```

2. **Librerías requeridas**:
   - `tkinter` (interfaz gráfica)
   - `pandas` (procesamiento de datos)
   - `pdfplumber` (extracción desde PDF)
   - `re` (expresiones regulares)
   - `datetime` (manejo de fechas)
   - `threading` (procesamiento en segundo plano)

### Cómo usarla

#### Paso 1: Iniciar la aplicación
```bash
python gui_validation_app.py
```

#### Paso 2: Seleccionar archivos
- Haz clic en "Buscar" junto a "Informe PDF"
- Selecciona el archivo PDF
- Haz clic en "Buscar" junto a "Datos históricos de Excel"
- Selecciona el archivo Excel

#### Paso 3: Validar archivos
- Haz clic en el botón "Validar archivos"
- La aplicación:
  - Extrae datos desde las tablas del PDF
  - Compara contra los datos históricos de Excel
  - Identifica discrepancias
  - Muestra el progreso en la barra de estado

#### Paso 4: Revisar resultados
- Si la validación pasa: aparece el mensaje "Todo correcto"
- Si hay diferencias: la tabla muestra todas las discrepancias
- Las filas en rojo indican diferencias de cantidad

#### Paso 5: Exportar resultados (opcional)
- Haz clic en "Exportar a CSV" para guardar las diferencias
- Elige ubicación y nombre del archivo
- El resultado se guarda con todos los detalles de validación

### Diseño de la aplicación

```
┌─────────────────────────────────────────────────────────────┐
│  Sistema de Validación de Servicios y Equipos              │
├─────────────────────────────────────────────────────────────┤
│  Informe PDF:          [________________] [Buscar]         │
│  Datos históricos:     [________________] [Buscar]         │
├─────────────────────────────────────────────────────────────┤
│  [Validar archivos] [Exportar a CSV] [Limpiar resultados]   │
├─────────────────────────────────────────────────────────────┤
│  Fecha   │ Descripción      │ PDF │ Excel │ Diferencia      │
├─────────────────────────────────────────────────────────────┤
│ 2026-04  │ Servicio A       │ 100 │ 105   │ 5               │
│ 2026-04  │ Equipo B         │ 50  │ 45    │ 5               │
│ 2026-05  │ Servicio C       │ 200 │ 195   │ 5               │
└─────────────────────────────────────────────────────────────┘
```

### Detalles del procesamiento

#### Extracción desde PDF
- **Origen**: Archivos PDF con tablas de servicios/equipos
- **Datos extraídos**:
  - Fecha (desde el encabezado de la tabla)
  - Tipo de servicio/equipo (detección flexible de columna)
  - Cantidad (compatible con formato numérico colombiano)
- **Proceso**:
  1. Abrir cada página del PDF
  2. Extraer todas las tablas
  3. Identificar la fila de encabezado
  4. Extraer nombres y cantidades
  5. Interpretar fechas en varios formatos
  6. Agrupar por fecha y tipo

#### Extracción desde Excel
- **Origen**: Archivo Excel con datos históricos
- **Datos extraídos**:
  - Descripciones de servicio (columna "DESCRIPCION TARIFA")
  - Cantidades por fecha (columnas de fecha)
- **Proceso**:
  1. Leer el archivo Excel
  2. Identificar columnas de fecha (tipo Timestamp)
  3. Convertir de formato ancho a formato largo
  4. Filtrar por la fecha del reporte
  5. Agrupar por descripción
  6. Eliminar valores en cero

#### Lógica de validación
```
PARA CADA servicio en el PDF (por fecha):
    Obtener cantidad del PDF
    Buscar el servicio equivalente en Excel para la misma fecha
    
    SI existe el servicio:
        SI PDF != Excel:
            Registrar discrepancia
    SI NO existe:
        SI la cantidad > 0:
            Registrar discrepancia
```

### Cómo interpretar los resultados

**Cuando ves "Todo correcto"**
- ✓ Todas las cantidades coinciden
- ✓ No hay diferencias entre PDF y Excel
- No se requiere ninguna acción adicional

**Cuando ves filas con discrepancias**
- 🔴 El color rojo indica diferencias
- Revisa el nombre del servicio en ambas fuentes
- Verifica si la diferencia es válida
- Investiga el origen de la discrepancia

**Columnas explicadas**
- **Fecha**: Fecha del reporte en el PDF
- **Descripción**: Tipo de equipo o servicio
- **PDF**: Cantidad encontrada en el PDF
- **Excel**: Cantidad encontrada en el Excel
- **Diferencia**: Diferencia absoluta entre ambas

### Formato del CSV exportado

El archivo CSV exportado contiene:
- **Fecha**: Fecha del reporte
- **Servicio**: Nombre del servicio/equipo
- **PDF**: Cantidad del PDF
- **Excel**: Cantidad del Excel
- **Diferencia**: Diferencia absoluta

Ejemplo:
```csv
Fecha,Servicio,PDF,Excel,Diferencia
2026-04-15,Service A,100,105,5
2026-04-15,Equipment B,50,45,5
2026-05-20,Service C,200,195,5
```

### Solución de problemas

#### "No se encontraron servicios válidos en el PDF"
- Verifica que el PDF tenga tablas correctamente estructuradas
- Revisa que las tablas incluyan columnas con "Tipo Equipo" y "Cantidad"
- Prueba con otro archivo PDF

#### "No se detectaron columnas de fecha en el Excel"
- Asegúrate de que el Excel tenga columnas de fecha en formato Timestamp
- Verifica la existencia de la columna "DESCRIPCION TARIFA"
- Confirma que la estructura del Excel coincida con el formato esperado

#### El proceso es lento
- Es normal en archivos grandes con muchas tablas
- Espera a que termine el proceso
- Considera dividir archivos grandes en partes más pequeñas

#### No aparecen diferencias, pero esperabas errores
- Verifica que las rutas de los archivos sean correctas
- Comprueba que las fechas coincidan en ambos archivos
- Asegúrate de que los nombres estén escritos de forma consistente

### Rendimiento

- **Tiempo típico de procesamiento**:
  - Archivos pequeños (5-10 páginas): 5-10 segundos
  - Archivos medianos (20-50 páginas): 15-30 segundos
  - Archivos grandes (más de 100 páginas): 1-3 minutos

- **Uso de memoria**:
  - Bajo (~50-100MB para uso típico)
  - Depende del número de registros

### Funciones avanzadas

**Soporte de formatos de fecha**
- Español: "6 de abril de 2026"
- Alternativo: "06 abril 2026"
- ISO: "2026-04-06"
- Detección y normalización automática

**Normalización de nombres**
- Elimina espacios extra
- Convierte saltos de línea en espacios
- Comparación sin distinguir mayúsculas/minúsculas
- Maneja caracteres especiales

**Soporte de formato numérico**
- Formato colombiano: "1.000,00"
- Standard format: "1,000.00"
- Automatic conversion to numbers
- Decimal precision preserved

### Code Structure

**Main Components**:
1. `ServicesValidationApp` class: GUI and data processing
2. `_extraer_conteo_pdf()`: PDF data extraction
3. `_extraer_conteo_excel()`: Excel data extraction
4. `_comparar_conteos()`: Comparison and validation
5. `_display_results()`: Results visualization

**Helper Methods**:
- `_normalizar_texto_equipo()`: Text normalization
- `_normalizar_fecha()`: Date parsing
- `_format_number()`: Number formatting

### Future Enhancements

Potential improvements:
- Batch validation of multiple file pairs
- Email notifications for validation failures
- Database storage of historical validation results
- Advanced filtering and sorting options
- Automated validation on schedule
- Support for additional Excel sheet formats
- Tolerance threshold configuration
- Detailed audit log generation

### Support

For issues:
1. Verify file formats are correct
2. Check that required columns exist
3. Review error messages for specifics
4. Ensure file paths are valid
5. Try with a smaller test file first

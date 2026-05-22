"""
GUÍA RÁPIDA - Validación de Servicios y Equipos

Este archivo muestra cómo ejecutar la aplicación gráfica de validación.
"""

# ============================================================================
# INSTALACIÓN
# ============================================================================
# 1. Asegúrate de estar dentro de tu entorno virtual
# 2. Instala los paquetes requeridos (si aún no están instalados):
#    pip install pandas pdfplumber openpyxl
#
# 3. Ejecuta este archivo o inicia directamente:
#    python gui_validation_app.py

# ============================================================================
# USO
# ============================================================================
# 1. Abre tu terminal o PowerShell
# 2. Ve a la carpeta facturacion:
#    cd c:\Users\andres.mejia\venv1.2\facturacion
#
# 3. Activa tu entorno virtual:
#    En Windows: ..\Scripts\activate
#    En macOS/Linux: source ../bin/activate
#
# 4. Ejecuta la aplicación:
#    python gui_validation_app.py
#
# 5. La ventana de la aplicación se abrirá
# 6. Sigue las instrucciones en pantalla

# ============================================================================
# ESTRUCTURA ESPERADA DE ARCHIVOS
# ============================================================================
# Tus archivos deberían estar organizados así:
#
# Carpeta del proyecto:
#   ├── Archivos PDF (informes de servicios/equipos)
#   │   ├── report_april_2026.pdf
#   │   ├── report_may_2026.pdf
#   │   └── ...
#   │
#   └── Datos históricos de Excel
#       ├── historical_data.xlsx
#       ├── services_inventory.xlsx
#       └── ...

# ============================================================================
# GUÍA PASO A PASO
# ============================================================================
#
# PASO 1: Iniciar la aplicación
#   $ python gui_validation_app.py
#   La ventana de la aplicación aparecerá
#
# PASO 2: Seleccionar el PDF
#   - Haz clic en "Buscar" junto a "Informe PDF"
#   - Ubica tu archivo PDF
#   - Haz clic en "Abrir" o doble clic sobre el archivo
#
# PASO 3: Seleccionar el Excel
#   - Haz clic en "Buscar" junto a "Datos históricos de Excel"
#   - Ubica tu archivo Excel
#   - Haz clic en "Abrir" o doble clic sobre el archivo
#
# PASO 4: Validar
#   - Haz clic en "Validar archivos"
#   - Espera a que termine el proceso (revisa la barra de estado)
#   - Los resultados aparecerán en la tabla
#
# PASO 5: Revisar resultados
#   - Si todo está correcto: aparecerá un mensaje de éxito
#   - Si hay errores: las filas rojas mostrarán las diferencias
#   - Revisa las columnas Fecha, Descripción, PDF, Excel y Diferencia
#
# PASO 6: Exportar (opcional)
#   - Haz clic en "Exportar a CSV"
#   - Elige ubicación y nombre del archivo
#   - El archivo se guardará con los resultados de la validación

# ============================================================================
# CÓMO ENTENDER EL RESULTADO
# ============================================================================
#
# MENSAJE DE ÉXITO (Todo correcto):
#   Aparece cuando:
#   - Todas las cantidades coinciden entre PDF y Excel
#   - No se encontraron diferencias
#   - La validación terminó correctamente
#
# FILAS EN ERROR (Rojo):
#   Muestra:
#   - Fecha de la diferencia
#   - Nombre del servicio/equipo
#   - Cantidad en PDF
#   - Cantidad en Excel
#   - Diferencia absoluta
#
# QUÉ HACER:
#   1. Revisa si la diferencia es esperada
#   2. Verifica que los nombres coincidan
#   3. Investiga el origen de la diferencia
#   4. Corrige los archivos si es necesario

# ============================================================================
# ATAJOS DE TECLADO
# ============================================================================
#
# Tab             - Mover entre campos
# Enter/Return    - Activar un botón
# Ctrl+A (campos) - Seleccionar todo el texto
# Ctrl+C (tabla)  - Copiar filas seleccionadas

# ============================================================================
# REQUISITOS DE LOS ARCHIVOS
# ============================================================================
#
# ARCHIVO PDF:
#   - Debe contener tablas con datos de servicios/equipos
#   - Requiere columnas con "Tipo Equipo" y "Cantidad"
#   - Debe incluir información de fecha (encabezado o tabla)
#   - Formatos de fecha soportados:
#     * Español: "6 de abril de 2026"
#     * Estándar: "2026-04-06"
#     * Europeo: "06/04/2026"
#
# ARCHIVO EXCEL:
#   - Debe tener la columna "DESCRIPCION TARIFA"
#   - Debe tener columnas de fecha (tipo Timestamp)
#   - Contiene cantidades históricas de servicios
#   - Las fechas deben coincidir con las del PDF

# ============================================================================
# SOLUCIÓN DE PROBLEMAS
# ============================================================================
#
# P: La aplicación no inicia
# R: Instala los paquetes requeridos:
#    pip install pandas pdfplumber openpyxl
#
# P: "No se encontraron elementos válidos en el PDF"
# R: Verifica que:
#    - El PDF tenga el formato correcto
#    - Las tablas tengan las columnas adecuadas
#    - La fecha se detecte correctamente
#
# P: "No se detectaron columnas de fecha en el Excel"
# R: Verifica que el archivo Excel tenga:
#    - Columnas de fecha en formato Timestamp
#    - La columna "DESCRIPCION TARIFA"
#    - Estructura correcta
#
# P: La aplicación está lenta
# R: Es normal con archivos grandes. Espera a que termine.
#
# P: No se ven todas las columnas
# R: Usa la barra horizontal en la parte inferior de la tabla
#
# P: Falló la exportación
# R: Asegúrate de tener permisos de escritura en la ubicación seleccionada

# ============================================================================
# FUNCIONALIDADES DE LA APLICACIÓN
# ============================================================================
#
# ✓ Detección automática de formatos
# ✓ Detección inteligente de columnas
# ✓ Soporte para varios formatos de fecha
# ✓ Procesamiento en segundo plano
# ✓ Errores con colores
# ✓ Exportación a CSV
# ✓ Notificaciones de éxito/error
# ✓ Reporte detallado de errores

# ============================================================================
# QUÉ HACE LA APLICACIÓN
# ============================================================================
#
# 1. LEE EL PDF:
#    - Recorre todas las páginas buscando tablas
#    - Encuentra columnas: tipo de equipo y cantidad
#    - Extrae la información de fecha
#    - Agrupa por fecha y servicio
#
# 2. LEE EL EXCEL:
#    - Carga los datos históricos
#    - Identifica descripciones de servicio
#    - Agrupa por fecha
#    - Normaliza nombres
#
# 3. COMPARA:
#    - Relaciona servicios por fecha
#    - Compara cantidades
#    - Calcula diferencias
#    - Reporta discrepancias
#
# 4. MUESTRA:
#    - Presenta resultados en tabla
#    - Resalta errores en rojo
#    - Muestra "Todo correcto" si no hay errores
#    - Permite exportar CSV

if __name__ == "__main__":
    # Importa y ejecuta la aplicación gráfica
    from gui_validation_app import main
    main()

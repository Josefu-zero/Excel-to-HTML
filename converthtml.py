import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
import os


def es_texto_aislado(ws, row_idx, merged_cells):
    """
    Determina si una fila contiene texto aislado (no parte de una tabla estructurada).
    """
    row = ws[row_idx]
    
    # Contar celdas con contenido
    celdas_con_contenido = sum(1 for cell in row if cell.value is not None)
    
    # Si hay menos de 2 celdas con contenido, es texto aislado
    if celdas_con_contenido < 2:
        return True
    
    # Verificar si es un título de sección
    first_cell_value = str(row[0].value).strip() if row[0].value else ""
    if first_cell_value and first_cell_value.isupper() and len(first_cell_value) > 10:
        return True
        
    return False

def procesar_texto_aislado(ws, row_idx, merged_cells,sheet_name):
    """
    Procesa una fila de texto aislado y devuelve el HTML correspondiente.
    """
    row = ws[row_idx]
    contenido = []
    
    for cell in row:
        cell_value = cell.value
        
        # Verificar si la celda es parte de un rango combinado
        for merged in merged_cells:
            if cell.coordinate in merged['range']:
                if (cell.row, cell.column) == merged['first_cell']:
                    cell_value = merged['value']
                else:
                    cell_value = None
                break
        
        if cell_value is not None:
            contenido.append(str(cell_value).strip())
    # No agregar nada si el texto es igual al nombre de la hoja (sheet_name)
    if contenido and " ".join(contenido).strip() == sheet_name:
        return ""
    texto = " ".join(contenido).strip()
    
    # Determinar si es un título
    if texto.isupper() and len(texto) > 7:
        return f'<h2 class="titulo-seccion">{texto}</h2>\n'
    else:
        return f'<div class="texto-contenido">{texto}</div>\n'

def procesar_tabla(ws, start_row, merged_cells):
    """
    Procesa una tabla desde la fila start_row y devuelve el HTML y número de filas procesadas.
    Combina celdas de encabezado vacías adyacentes usando colspan.
    """
    # Determinar el alcance de la tabla
    end_row = start_row
    while end_row <= ws.max_row:
        if es_texto_aislado(ws, end_row, merged_cells):
            break
        end_row += 1
    
    # Crear DataFrame con el rango de la tabla
    data = []
    for row in ws.iter_rows(min_row=start_row, max_row=end_row-1):
        row_data = []
        for cell in row:
            cell_value = cell.value
            
            # Manejar celdas combinadas
            for merged in merged_cells:
                if cell.coordinate in merged['range']:
                    if (cell.row, cell.column) == merged['first_cell']:
                        cell_value = merged['value']
                    else:
                        cell_value = None
                    break
            
            row_data.append(cell_value)
        data.append(row_data)
    
    df = pd.DataFrame(data)
    
    # Limpiar DataFrame (eliminar filas/columnas completamente vacías)
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # Generar HTML de la tabla
    html = '<div class="tabla-contenedor">\n'
    
    if len(df) > 0:
        # Procesar encabezados para combinar celdas vacías adyacentes
        headers = df.iloc[0].tolist()
        processed_headers = []
        i = 0
        while i < len(headers):
            if headers[i] is not None and str(headers[i]).strip() != '':
                # Contar celdas vacías siguientes
                colspan = 1
                j = i + 1
                while j < len(headers) and (headers[j] is None or str(headers[j]).strip() == ''):
                    colspan += 1
                    j += 1
                
                if colspan > 1:
                    processed_headers.append(f'<th colspan="{colspan}">{headers[i]}</th>')
                else:
                    processed_headers.append(f'<th>{headers[i]}</th>')
                i = j
            else:
                i += 1
        
        # Construir manualmente la tabla HTML
        html += '<table class="tabla-estructurada">\n'
        html += '  <thead>\n    <tr>\n'
        html += '      ' + '\n      '.join(processed_headers) + '\n'
        html += '    </tr>\n  </thead>\n'
        
        # Procesar el cuerpo de la tabla
        html += '  <tbody>\n'
        for _, row in df.iloc[1:].iterrows():
            html += '    <tr>\n'
            for cell in row:
                html += f'      <td>{cell if cell is not None else ""}</td>\n'
            html += '    </tr>\n'
        html += '  </tbody>\n</table>\n'
    else:
        html += '<p>Tabla vacía</p>\n'
    
    html += '</div>\n'
    
    return html, end_row - start_row



def excel_a_html_multiple(archivo_excel, carpeta_salida='html_output'):
    # Crear carpeta de salida si no existe
    os.makedirs(carpeta_salida, exist_ok=True)
    
    # Cargar el libro de Excel
    wb = load_workbook(archivo_excel, data_only=True)
    
    # Lista para almacenar información del índice
    indice = []
    
    # Obtener el nombre del archivo Excel sin la extensión
    nombre_archivo_excel = os.path.splitext(os.path.basename(archivo_excel))[0]
    print(f"Procesando archivo: {nombre_archivo_excel}")
    
    # Procesar cada hoja del libro
    for sheet_name in wb.sheetnames:
        if sheet_name.strip().lower() == "índice":
            continue
        ws = wb[sheet_name]
        nombre_archivo = f"{slugify(sheet_name)}.html"
        # Tomar el primer texto no vacío de la hoja como nombre de sección si existe
        newname = None
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                if cell.value is not None and str(cell.value).strip() != "":
                    newname = str(cell.value).strip()
                    sheet_name = str(cell.value).strip()
                    break
            if newname  is not None:
                break
            
        indice.append({'nombre': sheet_name, 'archivo': nombre_archivo})
        
        # Crear HTML para esta hoja
        html = generar_html_hoja(ws, sheet_name, nombre_archivo_excel)
        
        # Guardar archivo HTML
        with open(os.path.join(carpeta_salida, nombre_archivo), 'w', encoding='utf-8') as f:
            f.write(html)
    
    # Generar archivo índice
    generar_indice(indice, carpeta_salida, nombre_archivo_excel)
    
    return indice

def slugify(texto):
    """Convierte un texto a formato válido para nombre de archivo"""
    import re
    texto = re.sub(r'[^\w\s-]', '', texto.lower())
    return re.sub(r'[-\s]+', '-', texto).strip('-_')

def generar_html_hoja(ws, sheet_name, nombre_archivo_excel):
    

    """Genera el HTML para una hoja específica"""
    # Procesar celdas combinadas
    merged_cells = []
    for merged_range in ws.merged_cells.ranges if ws.merged_cells else []:
        min_row, min_col, max_row, max_col = range_boundaries(str(merged_range))
        merged_cells.append({
            'range': merged_range,
            'value': ws.cell(row=min_row, column=min_col).value,
            'first_cell': (min_row, min_col)
        })
    
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{sheet_name} - {nombre_archivo_excel}</title>
    <link rel="stylesheet" href="styles.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css">
</head>
<body>
    <header>
        <h1>{sheet_name}</h1>
        <a href="index.html" class="btn-volver"><i class="fas fa-arrow-left"></i> Volver al índice</a>
    </header>
    <div class="contenido-hoja">
"""
    
    # Procesar filas en orden
    current_row = 1
    while current_row <= ws.max_row:
        # Verificar si la fila está vacía
        row_empty = all(cell.value is None for cell in ws[current_row])
        
        if not row_empty:
            # Determinar si es texto aislado o parte de una tabla
            if es_texto_aislado(ws, current_row, merged_cells):
                html += procesar_texto_aislado(ws, current_row, merged_cells,sheet_name)
                current_row += 1
            else:
                # Procesar como tabla
                table_html, rows_processed = procesar_tabla(ws, current_row, merged_cells)
                html += table_html
                current_row += rows_processed
        else:
            current_row += 1
    
    html += f"""
    </div>
    <footer>
        <div class="footer-flex">
            <img src="azul_secretaria-de-gobierno-digital-y-tecnología-de-la-información-y-comunicaciones.png" alt="Logo secretaria de Gobierno Digital y Tecnología de la Información y Comunicaciones">
            <span>&copy; 2025</span>
        </div>
    </footer>
</body>
</html>
"""
    return html

def generar_indice(indice, carpeta_salida, nombre_archivo_excel):
    """Genera el archivo índice HTML"""
    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Índice - {nombre_archivo_excel}</title>
    <link rel="stylesheet" href="styles.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css">
</head>
<body>
    <header>
        <h1>{nombre_archivo_excel}</h1>
    </header>
    <div class="contenedor-indice">
        <h2>Índice de Contenidos</h2>
        <ul class="lista-indice">
"""
    
    for item in indice:
        html += f'            <li><a href="{item["archivo"]}"><i class="fas fa-file-alt"></i> {item["nombre"]}</a></li>\n'
    
    html += f"""        </ul>
    </div>
    <footer>
        <div class="footer-flex">
            <img src="azul_secretaria-de-gobierno-digital-y-tecnología-de-la-información-y-comunicaciones.png" alt="Logo secretaria de Gobierno Digital y Tecnología de la Información y Comunicaciones">
            <span>&copy; 2025</span>
        </div>
    </footer>
</body>
</html>
"""
    
    # Generar archivo CSS
    css = """/* Estilos generales */
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    line-height: 1.6;
    margin: 0;
    padding: 0;
    color: #333;
    background-color: #f5f7fa;
}

header {
    background: linear-gradient(135deg, #28367f, #28367f);
    color: white;
    padding: 25px 20px;
    text-align: center;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    position: relative;
}

h1 {
    font-size: 28px;
    margin-bottom: 15px;
    font-weight: 600;
    letter-spacing: 0.5px;
}

.contenido-hoja {
    max-width: 98%;
    margin: 30px auto;
    padding: 0 20px;
    overflow-x: auto;
}

/* Estilos para la tabla */
.tabla-contenedor {
    width: 100%;
    margin: 25px 0;
    box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    border-radius: 8px;
    overflow: hidden;
}

.tabla-estructurada {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.9em;
    min-width: 1200px;
}

.tabla-estructurada thead tr:first-child {
    background: linear-gradient(135deg, #1275bb, #2980b9);
}

.tabla-estructurada thead tr:first-child th {
    font-size: 16px;
    padding: 15px;
    text-align: center;
    border: none;
    position: relative;
}

.tabla-estructurada thead tr:first-child th:not(:last-child)::after {
    content: "";
    position: absolute;
    right: 0;
    top: 15%;
    height: 70%;
    width: 1px;
    background-color: rgba(255,255,255,0.3);
}

/* Subencabezado */
.tabla-estructurada thead tr.subheader {
    background: linear-gradient(135deg, #1275bb, #1275bb);
    color: white;
    font-weight: 600;
}

.tabla-estructurada thead tr.subheader th {
    padding: 12px 10px;
    font-weight: 500;
    text-align: left;
    border-right: 1px solid rgba(255,255,255,0.1);
}

.tabla-estructurada thead tr.subheader th:last-child {
    border-right: none;
}

/* Cuerpo de la tabla */
.tabla-estructurada tbody tr {
    transition: all 0.2s ease;
}

.tabla-estructurada tbody tr:hover {
    background-color: #e8f4fc !important;
    transform: translateY(-1px);
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.tabla-estructurada tbody td {
    padding: 12px 10px;
    border: 1px solid #e0e0e0;
    vertical-align: middle;
}

.tabla-estructurada tbody tr:nth-child(even) {
    background-color: #f8f9fa;
}

/* Estilos para celdas específicas */

.tabla-estructurada td[data-criticidad="Critico"] {
    background-color: #ffecec;
    color: #d32f2f;
    font-weight: 600;
}

.tabla-estructurada td[data-criticidad="No Critico"] {
    background-color: #f0fff0;
    color: #388e3c;
}

/* Botón de volver */
.btn-volver {
    display: inline-flex;
    align-items: center;
    padding: 10px 18px;
    background: rgba(204, 204, 204,0.15);
    color: white;
    text-decoration: none;
    border-radius: 50px;
    transition: all 0.3s ease;
    font-size: 14px;
    border: 1px solid rgba(204, 204, 204,0.2);
    backdrop-filter: blur(5px);
    position: absolute;
    left: 30px;
    top: 40px;
    z-index: 2;
}

.btn-volver:hover {
    background: rgba(204, 204, 204,0.25);
    transform: translateX(-3px);
}

.btn-volver i {
    margin-right: 8px;
    transition: all 0.3s ease;
}

.btn-volver:hover i {
    transform: translateX(-2px);
}

/* Footer */
footer {
    text-align: center;
    padding: 20px;
    background: linear-gradient(135deg, #28367f, #28367f);
    color: white;
    margin-top: 40px;
    font-size: 14px;

}

.footer-flex {
    display: flex;
    flex-direction: column; /* Cambia a columna */
    align-items: center;
    justify-content: center;
    gap: 0px;
}

.footer-flex img {
    max-height: 110px;
    display: block;
    margin: 0;
}

/* Efectos de scroll para tablas grandes */
@media (max-width: 1200px) {
    .contenido-hoja {
        overflow-x: auto;
        padding: 0 10px;
    }

    .tabla-contenedor {
        margin: 15px 0;
    }
}

/* Mejoras para la legibilidad */
td {
    line-height: 1.5;
}

/* Resaltar campos importantes */
td[data-tipo="Maestro"] {
    font-weight: 600;
    color: #1a73e8;
}

td[data-tipo="Referencia"] {
    color: #666;
}
/* Estilos para tablas - Versión mejorada */
.tabla-contenedor {
    margin: 30px 0;
    overflow-x: auto;
    background-color: white;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    padding: 0;
    width: 100%;
    min-width: 100%;
    table-layout: fixed
}

table {
    border-collapse: collapse;
    width: 100%;
    margin: 0;
    font-size: 0.95em;
    min-width: 600px; /* Ancho mínimo para tablas responsivas */
}

th {
    background-color: #3498db;
    border: 1px solid #1275bb;
    color: white;
    font-weight: 600;
    padding: 14px 16px;
    text-align: center;
    position: sticky;
    top: 0;
}

/* Estilo para el subencabezado */
.tabla-estructurada thead tr.subheader th {
    background-color: #1275bb;
    color: white;
    font-weight: 500;
    padding: 10px 16px;
    text-align: left;
    border-bottom: 1px solid rgba(255,255,255,0.1);
}

td {
    padding: 12px 16px;
    border: 1px solid #e0e0e0;
    vertical-align: middle;
    line-height: 1.5;
}

tr:nth-child(even) {
    background-color: #f8f9fa;
}

tr:hover {
    background-color: #f1f7fd;
}

/* Estilo para celdas especiales */
td.destacado {
    background-color: #e3f2fd;
    font-weight: 500;
    color: #1275bb;
}

td.negativo {
    color: #e53935;
    font-weight: 500;
}

td.positivo {
    color: #43a047;
    font-weight: 500;
}

/* Bordes redondeados para la tabla */
table {
    border-radius: 8px;
    overflow: hidden;
}

/* Cabecera con gradiente */
.tabla-estructurada thead tr:first-child th {
    background: linear-gradient(135deg, #28367f 0%, #28367f 100%);
}

/* Estilo para tablas compactas */
.tabla-compacta th,
.tabla-compacta td {
    padding: 8px 12px;
    font-size: 0.9em;
}

/* Estilo para tablas con bordes visibles */
.tabla-bordeada {
    border: 1px solid #e0e0e0;
}

.tabla-bordeada th,
.tabla-bordeada td {
    border: 1px solid #e0e0e0;
}

/* Estilo para celdas de encabezado izquierdo */
th.celda-izquierda {
    text-align: left;
    background-color: #2c3e50;
}

/* Efecto de transición para filas */
tr {
    transition: background-color 0.2s ease;
}

/* Scroll personalizado para tablas con overflow */
.tabla-contenedor::-webkit-scrollbar {
    height: 8px;
}

.tabla-contenedor::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 4px;
}

.tabla-contenedor::-webkit-scrollbar-thumb {
    background: #6ebce9;
    border-radius: 4px;
}

.tabla-contenedor::-webkit-scrollbar-thumb:hover {
    background: #1275bb;
}





"""
    
    # Guardar archivos
    with open(os.path.join(carpeta_salida, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(html)
    
    with open(os.path.join(carpeta_salida, 'styles.css'), 'w', encoding='utf-8') as f:
        f.write(css)
    
    return html

## Dar formato a la salida


if __name__ == "__main__":
    print("Procesando archivo Excel...")
    indice = excel_a_html_multiple('Libro Dominio Infraestructura Urbana.xlsx')
    print(f"Proceso completado. Se generaron {len(indice)} archivos HTML.")
    print(f"Archivo índice creado en: html_output/index.html")
    
from bs4 import BeautifulSoup

def convert_second_row_to_subheader(html_file):
    # Leer el archivo HTML
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    # Parsear el HTML
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Encontrar la tabla
    table = soup.find('table', class_='tabla-estructurada')
    if not table:
        print(f"No se encontró la tabla en {html_file}")
        return
    
    # Obtener el thead y tbody
    thead = table.find('thead')
    tbody = table.find('tbody')
    
    if not tbody or len(tbody.find_all('tr')) < 2:
        print(f"No hay suficientes filas en el tbody de {html_file}")
        return
    
    # Obtener la segunda fila del tbody (índice 1 porque la primera es 0)
    second_row = tbody.find_all('tr')[0]
    
    for td in second_row.find_all('td'):
        td.name = 'th'

    # Mover la segunda fila al thead
    second_row.extract()
    thead.append(second_row)
    
    # Guardar los cambios
    with open(html_file, 'w', encoding='utf-8') as file:
        file.write(str(soup))
    print(f"Archivo {html_file} modificado exitosamente")

# Procesar los archivos HTML
archivos_a_procesar = [
    'html_output/planes-de-remediación.html',
    'html_output/calidad.html',
    'html_output/diccionario.html'
]

for archivo in archivos_a_procesar:
    convert_second_row_to_subheader(archivo)
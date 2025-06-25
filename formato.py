from bs4 import BeautifulSoup

def convert_second_row_to_subheader(html_file):
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table', class_='tabla-estructurada')
    
    if not table:
        print(f"No se encontró la tabla en {html_file}")
        return
    
    thead = table.find('thead')
    tbody = table.find('tbody')
    
    if not tbody or len(tbody.find_all('tr')) < 2:
        print(f"No hay suficientes filas en el tbody de {html_file}")
        return
    
    # Obtener la segunda fila y agregar clase CSS
    second_row = tbody.find_all('tr')[0]
    
    
    second_row['class'] = 'subheader'  # Agregamos esta línea
    
    # Mover la fila al thead
    second_row.extract()
    thead.append(second_row)
    
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

from bs4 import BeautifulSoup

def limpiar_versionamiento(html_file):
    """
    Elimina la primera tabla y el primer texto del archivo versionamiento.html
    """
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Eliminar la primera tabla
    primera_tabla = soup.find('div', class_='tabla-contenedor')
    if primera_tabla:
        primera_tabla.decompose()
    
    # Eliminar el primer texto con clase 'texto-contenido'
    textos_contenido = soup.find_all('div', class_='texto-contenido')
    if len(textos_contenido) > 0:
        textos_contenido[0].decompose()
    
    # Guardar los cambios
    with open(html_file, 'w', encoding='utf-8') as file:
        file.write(str(soup))
    print(f"Archivo {html_file} modificado exitosamente")

# Uso de la función
limpiar_versionamiento('html_output/versionamiento.html')
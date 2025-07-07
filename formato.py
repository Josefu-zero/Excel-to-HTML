from bs4 import BeautifulSoup

def convert_second_row_to_subheader(html_file):
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Seleccionar la tabla según el archivo
    if 'calidad.html' in html_file:
        # Obtener todas las tablas y seleccionar la segunda
        tables = soup.find_all('table', class_='tabla-estructurada')
        if len(tables) < 2:
            print(f"No hay suficientes tablas en {html_file}")
            return
        table = tables[1]  # Segunda tabla (índice 1)
    else:
        # Para otros archivos, usar la primera tabla
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

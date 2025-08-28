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

def delete_fuenteoficial(html_file):
    with open(html_file, 'r', encoding='utf-8') as file:
        html_content = file.read()
    
    soup = BeautifulSoup(html_content, 'html.parser')
    table = soup.find('table', class_='tabla-estructurada')
    if not table:
        print(f"No se encontró la tabla en {html_file}")
        return
    
    thead = table.find('thead')
    tbody = table.find('tbody')

    # Buscar la fila subheader en thead
    subheader_row = thead.find('tr', class_='subheader')
    if not subheader_row:
        print(f"No se encontró la fila subheader en {html_file}")
        return

    # Identificar los índices de las columnas que contienen "(Fuente Oficial)"
    cols_to_delete = []
    for idx, th in enumerate(subheader_row.find_all(['th', 'td'])):
        if '(Fuente Oficial)' in th.get_text():
            cols_to_delete.append(idx)

    if not cols_to_delete:
        print(f"No se encontraron columnas '(Fuente Oficial)' en {html_file}")
        return

    # Eliminar las columnas en todas las filas de thead y tbody
    for section in [thead, tbody]:
        for row in section.find_all('tr'):
            cells = row.find_all(['th', 'td'])
            for i in sorted(cols_to_delete, reverse=True):
                if i < len(cells):
                    cells[i].extract()

    with open(html_file, 'w', encoding='utf-8') as file:
        file.write(str(soup))
    print(f"Columnas '(Fuente Oficial)' eliminadas en {html_file}")
    
    

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

/* Resalta los porcentajes (Usar si es necesario)
.percentage-cell {
    text-align: right;
    font-weight: bold;
    color: #0056b3; 
}*/

"""
def crear_css():
    with open('html_output/styles.css', 'w', encoding='utf-8') as file:
        file.write(css)
    print("Archivo CSS creado exitosamente en 'css/estilos.css'")


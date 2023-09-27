import pyodbc
from fpdf import FPDF
import html2text
import re

# Conexión a la base de datos
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=localhost\MSSQLSERVER01;DATABASE=ordicon;UID=erick;PWD=123')
cursor = conn.cursor()

# Tu consulta SQL ajustada
cursor.execute("""
    SELECT 
        MIN(t1.ID) AS IDdelCaso,
        MIN(t1.date) AS Fecha,
        CONCAT(MIN(u.FirstName), ' ', MIN(u.LastName)) AS Usuario,
        MIN(c.CompanyName) AS Cliente,
        t1.Name AS DescripcionProblema,
    MIN(t1.Description) AS DetalleProblema,
	MAX(CASE 
        WHEN t2.Description NOT LIKE '%UPDATE%'
        AND t2.Description NOT LIKE '%CREATE%'
        AND t2.Description NOT LIKE '%DELETE%'
        AND t2.Description NOT LIKE '%INSERT%'
        -- Puedes agregar más condiciones aquí si es necesario
        THEN t2.Description 
        ELSE NULL 
    END) AS SolucionalCliente,
    MAX(CASE 
        WHEN t2.Description LIKE '%UPDATE%'
        OR t2.Description LIKE '%CREATE%'
        OR t2.Description LIKE '%DELETE%'
        OR t2.Description LIKE '%INSERT%'
        -- Puedes agregar más condiciones aquí si es necesario
        THEN t2.Description 
        ELSE NULL 
    END) AS SolucionInterna
    FROM 
        RequestHistory t1
    JOIN 
        Comment t2 ON t1.GeneralID = t2.GeneralID
    JOIN 
        TPUser u ON t1.ownerId = u.UserID
    JOIN 
        Project p ON t1.projectId = p.ProjectID
    JOIN 
        Company c ON p.CompanyID = c.CompanyID
    GROUP BY 
        t1.Name;
""")

def clean_string(s):
    if s is None:
        return ""
    s = s.encode('latin-1', 'replace').decode('latin-1')
    s = re.sub(r'http\S+', '', s)  # Eliminar enlaces
    return s

def remove_html_format(text):
    if text is None:
        return ""
    return html2text.html2text(text).replace("\n", " ").strip()
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)


for row in cursor.fetchall():
    id_del_caso, fecha, nombre_completo, nombre_empresa, descripcion_problema, detalle_problema, solucion_cliente, solucion_interna = row
    fecha_format = fecha.strftime("%d/%m/%Y")
    descripcion_problema = remove_html_format(descripcion_problema)
    detalle_problema = clean_string(remove_html_format(detalle_problema))
    solucion_cliente = clean_string(remove_html_format(solucion_cliente))
    solucion_interna = clean_string(remove_html_format(solucion_interna))

    pdf.multi_cell(0, 5, f"ID del Caso: {id_del_caso}\nFecha: {fecha_format}\nUsuario: {nombre_completo}\nCliente: {nombre_empresa}\nDescripción Problema: {descripcion_problema}\nDetalle Problema: {detalle_problema}\nSolución al Cliente: {solucion_cliente}\nSolución Interna: {solucion_interna}\n", border=0)

pdf.output("output.pdf")

cursor.close()
conn.close()
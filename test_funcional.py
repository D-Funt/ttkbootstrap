import tkinter as tk
from tkinter import filedialog
from lxml import etree
import openpyxl

# Función para cargar y procesar el archivo XML
def cargar_y_procesar_xml():
    # Abrir un cuadro de diálogo para seleccionar el archivo XML
    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo XML",
        filetypes=(("Archivos XML", "*.xml"), ("Todos los archivos", "*.*"))
    )
    
    if archivo:
        # Procesar el archivo XML
        resultado = procesar_xml(archivo)
        # Guardar los resultados en un archivo Excel
        guardar_en_excel(resultado)
    else:
        print("No se seleccionó ningún archivo.")

# Función para obtener el texto de un nodo de manera segura
def obtener_texto(nodo, etiqueta, namespace):
    elemento = nodo.find(etiqueta, namespace)
    return elemento.text if elemento is not None else "N/A"

# Función para procesar el archivo XML
def procesar_xml(archivo):
    try:
        # Usar un parser de lxml con un enfoque de fragmentos
        parser = etree.XMLParser(recover=True, encoding='utf-8')
        with open(archivo, 'r', encoding='utf-8', errors='ignore') as file:
            tree = etree.parse(file, parser)

        root = tree.getroot()

        # Namespace a utilizar para acceder a las etiquetas
        ns = {'ns': 'urn:OECD:StandardAuditFile-Tax:PT_1.04_01'}  # Namespace del archivo XML

        # Lista para almacenar los datos de todas las facturas
        datos = []

        # Recorremos todas las facturas dentro de <SalesInvoices>
        for invoice in root.findall('.//ns:SalesInvoices/ns:Invoice', ns):
            invoice_no = obtener_texto(invoice, 'ns:InvoiceNo', ns)
            invoice_date = obtener_texto(invoice, 'ns:InvoiceDate', ns)
            invoice_type = obtener_texto(invoice, 'ns:InvoiceType', ns)
            customer_id = obtener_texto(invoice, 'ns:CustomerID', ns)
            gross_total = obtener_texto(invoice, './/ns:DocumentTotals/ns:GrossTotal', ns)
            net_total = obtener_texto(invoice, './/ns:DocumentTotals/ns:NetTotal', ns)
            tax_payable = obtener_texto(invoice, './/ns:DocumentTotals/ns:TaxPayable', ns)

            # Productos de la factura
            productos = []
            for line in invoice.findall('ns:Line', ns):
                description = obtener_texto(line, 'ns:Description', ns)
                quantity = obtener_texto(line, 'ns:Quantity', ns)
                unit_price = obtener_texto(line, 'ns:UnitPrice', ns)
                total = obtener_texto(line, 'ns:CreditAmount', ns)
                productos.append([description, quantity, unit_price, total])

            # Agregar una fila con los datos de la factura
            datos.append([invoice_no, invoice_date, invoice_type, customer_id, gross_total, net_total, tax_payable, productos])

        return datos

    except etree.XMLSyntaxError as e:
        return f"Error al procesar el archivo: {e}"
    except Exception as e:
        return f"Error desconocido: {e}"

# Función para guardar los datos en un archivo Excel
def guardar_en_excel(datos):
    # Crear un libro de trabajo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Facturas"

    # Escribir las cabeceras
    ws.append(['Factura No.', 'Fecha', 'Tipo', 'Cliente ID', 'Total Bruto', 'Total Neto', 'IVA', 'Productos'])

    # Escribir los datos de las facturas
    for factura in datos:
        productos_str = "\n".join([f"{producto[0]} - Cantidad: {producto[1]} - Precio Unitario: {producto[2]} - Total: {producto[3]}" for producto in factura[7]])
        ws.append([factura[0], factura[1], factura[2], factura[3], factura[4], factura[5], factura[6], productos_str])

    # Guardar el archivo Excel
    wb.save("facturas.xlsx")
    print("Los datos se han guardado en 'facturas.xlsx'")

# Ejecución de la función para cargar el XML
if __name__ == "__main__":
    cargar_y_procesar_xml()

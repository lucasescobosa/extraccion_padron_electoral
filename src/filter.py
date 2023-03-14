from pypdf import PdfReader
import xlsxwriter
import re

# Abro el archivo PDF
reader = PdfReader("padron.pdf")
number_of_pages = len(reader.pages)
text = ''
name = ''
second_line = ''
third_line = ''
buffer = 0

# Creo un archivo Excel para almacenar los resultados
workbook = xlsxwriter.Workbook('padron.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})
worksheet.write('A1', 'Nombre', bold)
worksheet.write('B1', 'Apellido', bold)
worksheet.write('C1', 'Dirección', bold)
worksheet.write('D1', 'Documento', bold)
worksheet.write('E1', 'Nacimiento', bold)
row = 0

# Función que filtra todo el texto de los encabezados
def fun(variable):
    # Textos variables
    words = ['DISTRITO: 08', 'SECCIÓN: 00', 'CIRCUITO:', 'MESA:']
    # Textos fijos
    lines = ['REPÚBLICA', 'REGISTRO NACIONAL DE ELECTORES', 'CÁMARA NACIONAL ELECTORAL',
             'ELECCIONES GENERALES - 14 DE NOVIEMBRE DE 2021', 'SECCIÓN ELECTORAL',
             'PADRÓN DEFINITIVO DE ELECTORES INSCRIPTOS AL 18 DE MAYO DE 2021ARGENTINA']
    for word in words:
        if word in variable:
            return False
    for line in lines:
        if line == variable:
            return False
    return True

# Bucle que recorre todas las paginas   
for page in range(number_of_pages):
    current_page = reader.pages[page]
    current_text = current_page.extract_text(0)
    # Filtro de portadas e índices
    if 'NRO. ORDEN' in current_text:
        text += current_text

lines = text.split('\n')

filtered = filter(fun, lines)

# Patrones Regex para extraer campos de las lineas
pattern_name = re.compile(r"[A-ZÀ-ÿ\u00f1\u00d1][A-ZÀ-ÿ\u00f1\u00d1\s]+,[A-ZÀ-ÿ\u00f1\u00d1\s]+")
pattern_doc = re.compile(r"(?<=DOC\.\s)[\S]+")
pattern_year = re.compile(r"\d{4}")

# Bucle para recorrer linea por linea
for line in filtered:

    # Detecta el comienzo de línea personal
    if 'NRO. ORDEN' in line:
        buffer = 1
        row += 1

    else:

        # Primera línea de datos, nombre y apellido
        if buffer == 1:
            full_name = pattern_name.findall(line)
            if (full_name != []):
                full_name = str(full_name[0]).split(',')
                name = full_name[1].strip()
                last_name = full_name[0].strip()
                worksheet.write(row, 0, name)
                worksheet.write(row, 1, last_name)
            else:
                print('error: ' + line)
            buffer += 1
        
        # Segunda línea puede ser dirección o continuación de nombre
        elif buffer == 2:
            second_line = line
            worksheet.write(row, 2, line)
            buffer += 1

        # Tercera línea puede ser documento y año o dirección
        elif buffer == 3:
            # Única excepción de dirección con doble línea
            if (line == 'CORONA)'):
                second_line = second_line + line
                worksheet.write(row, 2, second_line)
            else:
                third_line = line
                document = pattern_doc.findall(line)
                year = pattern_year.findall(line)
                if (document != [] and year !=[]):
                    worksheet.write(row, 3, document[0])
                    worksheet.write_number(row, 4, int(year[0]))
                buffer +=1
        
        # Cuarta línea si el buffer llega hasta aca significa que el nombre tiene doble línea
        # por lo tanto sobreescribo las líneas anteriores con los datos guardados en memoria
        elif buffer == 4:
            document = pattern_doc.findall(line)
            year = pattern_year.findall(line)
            if (document != [] and year !=[]):
                name = name + second_line
                worksheet.write(row, 0, name)
                worksheet.write(row, 2, third_line)
                worksheet.write(row, 3, document[0])
                worksheet.write_number(row, 4, int(year[0]))
            buffer +=1

# Cierro el libro Excel
workbook.close()

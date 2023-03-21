from pypdf import PdfReader
import xlsxwriter
import re
import os

# Abro el archivo PDF
dirname = os.path.dirname(__file__)
reader = PdfReader(os.path.join(dirname, "./files/padron.pdf"))
number_of_pages = len(reader.pages)

# Creo un archivo Excel para almacenar los resultados
workbook = xlsxwriter.Workbook(os.path.join(dirname, "./files/padron.xlsx"))
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})
worksheet.write('A1', 'distrito', bold) # Columna 0
worksheet.write('B1', 'matricula', bold) # Columna 1
worksheet.write('C1', 'clase', bold) # Columna 2
worksheet.write('D1', 'apellido', bold) # Columna 3
worksheet.write('E1', 'nombre', bold) # Columna 4
worksheet.write('F1', 'profesion', bold) # Columna 5
worksheet.write('G1', 'domicilio', bold) # Columna 6
worksheet.write('H1', 'tipdoc', bold) # Columna 7
worksheet.write('I1', 'secc', bold) # Columna 8
worksheet.write('J1', 'circu', bold) # Columna 9
worksheet.write('K1', 'mesa', bold) # Columna 10
worksheet.write('L1', 'sexo', bold) # Columna 11
worksheet.write('M1', 'padr_numreg', bold) # Columna 12
worksheet.write('N1', 'partido', bold) # Columna 13
worksheet.write('O1', 'padron', bold) # Columna 14
worksheet.write('P1', 'orden', bold) # Columna 15
row = 0

# Patrones Regex para extraer campos de las lineas
pattern_distrito = re.compile(r"(?<=DISTRITO:)[^\n]+")
pattern_seccion = re.compile(r"(?<=SECCIÓN ELECTORAL:)[^\n]+")
pattern_circuito = re.compile(r"(?<=CIRCUITO:)[^\n]+")
pattern_mesa = re.compile(r"(?<=MESA NRO.:)[^\n]+")
patern_orden = re.compile(r"\d+(?=[A-ZÀ-ÿ\u00f1\u00d1])")
pattern_nombre = re.compile(r"[A-ZÀ-ÿ\u00f1\u00d1][A-ZÀ-ÿ\u00f1\u00d1\s]+,[A-ZÀ-ÿ\u00f1\u00d1\s]+")
pattern_matricula = re.compile(r"(?<=DOC\.\s)[\S]+")
pattern_tipdoc = re.compile(r"[L](?=\s)|[L][.][A-Z][.]|[L][\d]|[D][N][I][-][^\s]+|[D][N][I][\s][E][^\s]+")
pattern_clase = re.compile(r"\d{4}")

# Función que filtra los encabezados de los padrones
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
padron = distrito = secc = circ = mesa = ''
for page in range(number_of_pages):
    current_page = reader.pages[page]
    current_text = current_page.extract_text(0)
    
    #Filtro para diferenciar portadas de padrones

    # Extraigo campos de la portada
    if 'NRO. ORDEN' not in current_text and 'REFERENCIAS' not in current_text and 'DISTRITO' in current_text:
        distrito = pattern_distrito.findall(current_text)
        secc = pattern_seccion.findall(current_text)
        circ = pattern_circuito.findall(current_text)
        mesa = pattern_mesa.findall(current_text)
        if (distrito == [] or secc == [] or circ == [] or mesa == []):
            print('Error buscando uno de los campos de la portada')
        else:
            distrito = str(distrito[0]).strip()
            secc = str(secc[0]).strip()
            circ = str(circ[0]).strip()
            mesa = str(mesa[0]).strip()
    
    # Extraigo el texto del padrón
    elif 'NRO. ORDEN' in current_text:
        lines = current_text.split('\n')
        padron = filter(fun, lines)
        second_line = third_line = ''

        for line in padron:
            # Detecta el comienzo de línea personal
            if 'NRO. ORDEN' in line:
                buffer = 1
                row += 1
                worksheet.write(row, 0, distrito)
                worksheet.write(row, 8, secc)
                worksheet.write(row, 9, circ)
                worksheet.write(row, 10, mesa)
            else:

                # Primera línea de datos, nombre y apellido
                if buffer == 1:
                    orden = patern_orden.findall(line)
                    nombre_completo = pattern_nombre.findall(line)
                    if (orden != [] and nombre_completo != []):
                        nombre_completo = str(nombre_completo[0]).split(',')
                        apellido = nombre_completo[0].strip()
                        nombre = nombre_completo[1].strip()
                        worksheet.write_number(row, 15, int(orden[0]))
                        worksheet.write(row, 3, apellido)
                        worksheet.write(row, 4, nombre)
                    else:
                        print('error buscando nro orden y nombre completo: ')
                    buffer += 1
                
                # Segunda línea puede ser domicilio o continuación de nombre
                elif buffer == 2:
                    second_line = line
                    worksheet.write(row, 6, line)
                    buffer += 1

                # Tercera línea puede ser documento y año o domicilio
                elif buffer == 3:
                    # Única excepción de domicilio con doble línea
                    if (line == 'CORONA)'):
                        second_line = second_line + line
                        worksheet.write(row, 6, second_line)
                    else:
                        third_line = line
                        matricula = pattern_matricula.findall(line)
                        tipdoc = pattern_tipdoc.findall(line)
                        clase = pattern_clase.findall(line)
                        if (matricula != [] and tipdoc != [] and clase !=[]):
                            matricula = str(matricula[0]).replace('.','')
                            worksheet.write_number(row, 1, int(matricula))
                            worksheet.write(row, 7, tipdoc[0])
                            worksheet.write_number(row, 2, int(clase[0]))
                        buffer +=1
                
                # Cuarta línea si el buffer llega hasta aca significa que el nombre tiene doble línea
                # por lo tanto sobreescribo las líneas anteriores con los datos guardados en memoria
                elif buffer == 4:
                    matricula = pattern_matricula.findall(line)
                    tipdoc = pattern_tipdoc.findall(line)
                    clase = pattern_clase.findall(line)
                    if (matricula != [] and tipdoc != [] and clase !=[]):
                        matricula = str(matricula[0]).replace('.','')
                        nombre = nombre + second_line
                        worksheet.write(row, 4, nombre)
                        worksheet.write(row, 6, third_line) # domicilio
                        worksheet.write_number(row, 1, int(matricula))
                        worksheet.write(row, 7, tipdoc[0])
                        worksheet.write_number(row, 2, int(clase[0]))
                    buffer +=1

# Cierro el libro Excel
workbook.close()




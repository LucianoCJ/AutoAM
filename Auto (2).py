
import pandas as pd 
#import openpyxl
from openpyxl import Workbook 
from openpyxl.styles import Font, Alignment 
from openpyxl.styles import PatternFill
from openpyxl.styles import Border, Side
#from openpyxl import load_workbook
#from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
#from reportlab.lib.pagesizes import letter
#from reportlab.lib.units import inch
#from reportlab.pdfgen import canvas
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.utils import range_boundaries
from openpyxl.utils.cell import column_index_from_string
#from openpyxl.worksheet.pagebreak import Break
import re
import datetime
from xlsx2html import xlsx2html
from weasyprint import HTML

##############################################################


# Obtener la fecha del próximo domingo en curso
today = datetime.date.today()
#Domingo proximo
next_sunday = today + datetime.timedelta(days=(6-today.weekday()+7)%7)
#domingo anterior
before_sunday = today - datetime.timedelta(days=today.weekday() + 1)
#before_sunday = today - datetime.timedelta(days=(6-today.weekday()-7)%7)

# Dar formato a la fecha como mes-día-año
formatted_date = next_sunday.strftime("%b %dth, %Y")
formatted_date2 = next_sunday.strftime("%b %dth")
formatted_date3 = before_sunday.strftime("%b %dth")
#Fecha del archivo
formatted_date4 = next_sunday.strftime("%b%d%Y")

############################################################

# Leer el archivo CSV
df = pd.read_csv('ReporteOAGUS_20230312.csv', header = None)

# Crear un nuevo libro de trabajo de Excel
wb = Workbook()

# Seleccionar la hoja activa
ws = wb.active

# Establecer el ancho de las columnas
ws.column_dimensions['A'].width = 3.9
ws.column_dimensions['B'].width = 6.75 
ws.column_dimensions['C'].width = 3.61     #COLUMNA 1
ws.column_dimensions['D'].width = 7.32     #COLUMNA 2
ws.column_dimensions['E'].width = 7.04     #COLUMNA 3
ws.column_dimensions['F'].width = 4.32 
ws.column_dimensions['G'].width = 4.32    #COLUMNA 4
ws.column_dimensions['H'].width = 4.04    #COLUMNA 5
ws.column_dimensions['I'].width = 3.89    #COLUMNA 6
ws.column_dimensions['J'].width = 4.32     #COLUMNA 7
ws.column_dimensions['K'].width = 4.18     #COLUMNA 8
ws.column_dimensions['L'].width = 5.32     #COLUMNA 9
ws.column_dimensions['M'].width = 6.04     #COLUMNA 10



# Establecer la altura deseada en la fila y columna especificada
ws.row_dimensions[1].height = 17.25

# Establecer el estilo de fuente y alineación
header_font = Font(name='Calibri', size=6, bold=True)
header_alignment = Alignment(horizontal='center', vertical='center')
cell_font = Font(name='Calibri', size=6)
cell_alignment = Alignment(horizontal='left', vertical='center')

# Crear una lista con los nombres de las cabeceras
cabeceras = ['CA', 'Market', 'Ind AM', 'Region', 'Month', 'Chg', 'Prev Ops', 'New Ops', 'Ops Chg', 'Prev Seat', 'New Seat', 'Sea Chg', '%_Seat Chg']

headers = cabeceras                             #Indico el número de columna a iniciar las cabeceras.
for col_num, header_title in enumerate(headers, 3):
    cell = ws.cell(row=1, column=col_num, value=header_title)
    cell.font = header_font
    cell.alignment = header_alignment

# Justifica el texto de los encabezados
for idx, cell in enumerate(ws[1],1):
    if idx<9:  
        cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    else:
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


# Escribir los datos                
for row_num, row_data in enumerate(df.values, 2):
    for col_num, cell_value in enumerate(row_data, 2):
        cell = ws.cell(row=row_num, column=col_num, value=cell_value)
        cell.font = cell_font
        if col_num < 9:
            cell.alignment = cell_alignment
        else:
            cell.alignment = Alignment(horizontal='right', vertical='center')
    if str(ws.cell(row=row_num, column=4).value) == 'nan':
        ws.delete_rows(row_num)
        

for idx, value in enumerate(ws.iter_rows(),2):
    ws.row_dimensions[idx].height = 7.5






# Cambiar el color de la celda A1 a rojo

#fill = PatternFill(start_color='808080', end_color='FFFFFF', fill_type='solid')

# Iterar sobre todas las filas y aplicar el formato de relleno al patrón especificado
#for row in ws.iter_rows():
 #   if row[0].value == 'TOTAL':
  #      cell_font = Font(bold = True)
   #     for cell in row:
    #        cell.fill = fill

#color1 = 'BFBFBF' COLOR DE LA CABECERA


# Seleccionar las celdas que se van a ajustar
#cell_range = ws['I1:M1']

# Ajustar el ancho de las columnas para que el texto quepa
#for row in cell_range:
 #   for cell in row:
  #      ws.column_dimensions[cell.column_letter].width = len(str(cell.value))


# definir los colores para los renglones
color1 = 'F2F2F2'
color2 = 'FFFFFF'
change = False


# cambiar el formato de color de los renglones 
for idx, row in enumerate(ws.iter_rows(),1):
    if ((str(ws.cell(row=idx, column=3).value) == 'None') or (str(ws.cell(row=idx, column=4).value) == 'TOTAL')):
        fill = PatternFill(start_color=color2, end_color=color2, fill_type='solid')
        change = True              
    elif change:
        fill = PatternFill(start_color=color1, end_color=color1, fill_type='solid')
        change = False                    
    else:
        fill = PatternFill(start_color=color2, end_color=color2, fill_type='solid')
        change = True
    for cell in row:
        cell.fill = fill
        
# Poner el renglon de total en negritas y porcentajes negativos en rojo
for idx, row in enumerate(ws.iter_rows(),1):
    if str(ws.cell(row=idx, column=4).value) == 'TOTAL':
        for cell in row:
            cell.font = Font(bold=True, name='Calibri', size=6)
        if re.search("-", str(ws.cell(row=idx, column=15).value)):
            ws.cell(row=idx, column=15).font = Font(color = "FF0000", name='Calibri', size=6, bold=True)
    elif re.search("-", str(ws.cell(row=idx, column=15).value)):
         ws.cell(row=idx, column=15).font = Font(color = "FF0000", name='Calibri', size=6)

columnaP = ws['C']
columnaU = ws['O']
mexterior = Side(style = 'thin', color = 'BFBFBF')

columnas = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
mnormal = Side(style = 'thin', color = 'F2F2F2')

# Pintar los bordes verticales interiores
for idx, value in enumerate(columnas,0):
    for idxcelda, celda in enumerate(ws[columnas[idx]],0):
        if idxcelda != 0:
            celda.border = Border(right = mnormal)

# Pintar los bordes verticales exteriores
for idx, celda in enumerate(columnaP,1):
    if idx == 1:
        celda.border = Border(left = mexterior)
    else:
        celda.border = Border(left = mexterior, right = mnormal)
for celda in columnaU:
    celda.border = Border(right = mexterior)

#Pintar los bordes totales

for idx, row in enumerate(ws.iter_rows(),1):
    if str(ws.cell(row=idx, column=4).value) == 'TOTAL':
        for idxcell, celda in enumerate(row,1):
            if re.search("Cell 'Sheet'.C", str(celda)):
                celda.border = Border(top = mnormal, bottom = mnormal, left = mexterior, right = mnormal)
            elif re.search("Cell 'Sheet'.O", str(celda)):
                celda.border = Border(top = mnormal, bottom = mnormal, left = mnormal, right = mexterior)
            else:
                celda.border = Border(top = mnormal, bottom = mnormal, left = mnormal, right = mnormal)


# Establecer color de relleno para los encabezados
fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')

# Justifica el texto de los encabezados
for cell in ws[1]:
    cell.fill = fill    #Agrego el color de relleno en el renglon 1


# Buscar la celda que contiene la palabra "TOTAL"
#for fila in ws.rows:
 #   for celda in fila:
  #      if celda.value == 'TOTAL':
            # Obtener la fila donde se encuentra la celda
   #         fila_total = celda.row
            # Poner el texto en negrita en toda la fila
    #        for celda_en_fila in ws[f'A{fila_total}:Z{fila_total}']:
     #           for celda_individual in celda_en_fila:
      #              celda_individual.font = openpyxl.styles.Font(bold=True)





#Convirtiendo a Decimales
# Selecciona la columna que deseas convertir (por ejemplo, columna A)
columna = ws['O']

# Itera sobre las celdas en la columna y convierte los valores de decimal a porcentaje
for celda in columna:
    if isinstance(celda.value, float):  # verifica si el valor de la celda es un decimal
        celda.value = celda.value * 1  # convierte el valor a porcentaje
        celda.number_format = '0%'  # establece el formato de número de la celda como porcentaje con dos decimales
    if (re.search("-", str(celda.value))):
        celda.value = str(int(celda.value * 100)).replace('-','(') + '%)'

#################################################################################
# Repitiendo cabeceras        
# Obtiene el número total de filas y columnas
num_rows = ws.max_row
num_cols = ws.max_column

# Establecer el estilo de fuente y alineación para las cabeceras
font = Font(name='Calibri', size=6, bold=True)
alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
# Define el número de filas entre cada repetición de la cabecera
#n = 48

# Itera a través de las filas y agrega las cabeceras cada n filas
#for i in range(1, num_rows + 1):
    #if (i - 1) % n == 0:
         #Copia las celdas de la primera fila a la fila actual
        #for j in range(1, num_cols + 1 ):
for max_row in range(2, num_rows + 1):
    if (max_row - 1) % 44 == 0: 
        # Inserta una nueva fila en la parte superior de la página
        ws.insert_rows(max_row)
        for col_num in range(1, num_cols + 1):
            cell = ws.cell(row=max_row, column=col_num)
            cell_name =get_column_letter(col_num) + '1'
            #header_cell = ws.cell(row=1, column=j)
            header_cell = ws[cell_name]
            cell.value = header_cell.value
            cell.font = font
            cell.alignment = alignment
            cell.fill = fill    #Agrego el color de relleno en el renglon 1


################################################################################

# Crear un objeto HeaderFooter y asignarle los textos
#hf = HeaderFooter()
#hf.center_header.text = "Encabezado"
#hf.center_footer.text = "Pie de página"
ws.oddHeader.center.text = "OAG Schedule Competitive Summary US " + '\n&"Calibri"&8Sunday ' +  formatted_date
ws.oddHeader.center.size = 14
ws.oddHeader.center.font = "Calibri,Bold"

ws.oddFooter.left.text = "Prev Snap:  " + formatted_date3 + " New Snap: " + formatted_date2 + '&R&"Calibri"&8&P - &N'
ws.oddFooter.left.size = 8
ws.oddHeader.left.font = "Calibri"


# Indicar el número de columna que deseas eliminar
num_columna = 1
num_columna2 = 1
num_columna17 = 14
num_columna18 = 15
num_columna19 = 14

# Eliminar la columnas

ws.delete_cols(num_columna)
ws.delete_cols(num_columna2)
ws.delete_cols(num_columna17)
ws.delete_cols(num_columna18)
ws.delete_cols(num_columna19)

# Guardar el archivo
wb.save('OAG Schedule Competitive Summary Sunday, ' + formatted_date + ' US.xlsx')

xlsx2html('OAG Schedule Competitive Summary Sunday, ' + formatted_date + ' US.xlsx', 'OAG Schedule Competitive Summary Sunday, ' + formatted_date + ' US.html')
HTML('OAG Schedule Competitive Summary Sunday, ' + formatted_date + ' US.html').write_pdf('OAG Schedule Competitive Summary Sunday, ' + formatted_date + ' US.pdf')

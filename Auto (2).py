import pandas as pd 
import openpyxl
from openpyxl import Workbook 
from openpyxl.styles import Font, Alignment 
from openpyxl.styles import PatternFill
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas



# Leer el archivo CSV
df = pd.read_csv('ReporteOAGUS_20230312.csv')

# Crear un nuevo libro de trabajo de Excel
wb = Workbook()

# Seleccionar la hoja activa
ws = wb.active


# Establecer el ancho de las columnas
ws.column_dimensions['A'].width = 3 # CA
ws.column_dimensions['B'].width = 8 # Market
ws.column_dimensions['C'].width = 3 # Ind AM
ws.column_dimensions['D'].width = 8 # Region
ws.column_dimensions['E'].width = 8 # Month
ws.column_dimensions['F'].width = 5 # Chg
ws.column_dimensions['G'].width = 5 # 
ws.column_dimensions['H'].width = 5    #COLUMNA 5
ws.column_dimensions['I'].width = 5    #COLUMNA 6
ws.column_dimensions['J'].width = 5     #COLUMNA 7
ws.column_dimensions['K'].width = 5     #COLUMNA 8
ws.column_dimensions['L'].width = 5     #COLUMNA 9
ws.column_dimensions['M'].width = 5     #COLUMNA 10


# Establecer la altura deseada en la fila y columna especificada
ws.row_dimensions[1].height = 18


# Establecer el estilo de fuente y alineación
header_font = Font(name='Calibri', size=6, bold=True)


# Crear una lista con los nombres de las cabeceras
cabeceras = ['CA', 'Market', 'Ind AM', 'Region', 'Month', 'Chg', 'Prev\nOps', 'New\nOps', 'Ops\nChg', 'Prev\nSeat', 'New\nSeat', 'Sea\nChg', '%_Seat\nChg']

#Indico el número de columna a iniciar las cabeceras
for col_num, header_title in enumerate(cabeceras, 3):
    cell = ws.cell(row=1, column=col_num, value=header_title)
    cell.font = header_font
    cell.alignment = Alignment(horizontal='left' if col_num < 9 else 'center', vertical='center', wrap_text=True)

# Escribir los datos                
for row_num, row_data in enumerate(df.values, 2):
    ws.row_dimensions[row_num].height = 10
    for col_num, cell_value in enumerate(row_data, 2):
        cell = ws.cell(row=row_num, column=col_num, value=cell_value)
        cell.alignment = Alignment(horizontal='left' if col_num < 9 else 'right', vertical='center')
    if str(ws.cell(row=row_num, column=4).value) == 'nan':
        ws.delete_rows(row_num)


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
bold = False


# cambiar el formato de color de los renglones 
for idx, row in enumerate(ws.iter_rows(),1):
    change = not change
    value_third_column = str(ws.cell(row=idx, column=3).value)
    value_fourth_column = str(ws.cell(row=idx, column=4).value)
    will_bold = (value_fourth_column == 'TOTAL')

    if ( value_third_column == 'None' or will_bold):
        fill = PatternFill(start_color=color2, end_color=color2, fill_type='solid')
        change = True              
    color_fill = color2 if change else color1
    for cell in row:
        cell.font = Font(name='Calibri', size=6, bold = will_bold)
        cell.fill = PatternFill(start_color=color_fill, end_color=color_fill, fill_type='solid')

# Establecer color de relleno para los encabezados
fill = PatternFill(start_color='BFBFBF', end_color='BFBFBF', fill_type='solid')

# Justifica el texto de los encabezados
for cell in ws[1]:
    cell.font = header_font
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
wb.save('archivo.xlsx')

from openpyxl import workbook
from openpyxl import load_workbook

#Crea la hoja 1
wb2 = load_workbook('5164_30_Limpieza.xlsx')
#ws1 = wb2.create_sheet('hoja1')


# Seleccionar la hoja de origen (Definitivo) y la hoja de destino (hoja1)
ws2_origen = wb2['Definitivo']
ws1_destino = wb2['hoja1']

#Columnas que deseas copiar 
columna_origen_1 = ws2_origen['H']  
columna_origen_2 = ws2_origen['M']  

# Iterar sobre las celdas en las columnas de origen y copiar los datos a la hoja de destino
for celda_origen_1, celda_origen_2 in zip(columna_origen_1, columna_origen_2):
    valor_1 = celda_origen_1.value
    valor_2 = celda_origen_2.value
    
    # Agregar los valores a las columnas correspondientes en la hoja de destino
    ws1_destino.append([valor_1, valor_2])


    # Seleccionar el rango de celdas a partir de B2 en hoja1 y reemplazar "(" por "_"
for row in ws1_destino.iter_rows(min_row=2, min_col=2, values_only=True):
    for cell in row:
        if cell is not None and '(' in cell:
            ws1_destino[cell.coordinate] = cell.replace('(', '_')

# Aplicar la función para dividir el texto en columnas en la columna B a partir de B2
# for row in ws1_destino.iter_rows(min_row=2, min_col=2, max_col=2):
#     for cell in row:
#         if cell.value is not None:
#             # Dividir el texto en columnas usando tabulaciones y ')' como delimitadores
#             split_values = [x.strip() for x in cell.value.split('\t') if x.strip() != '']
#             split_values = [x.replace(')', '_') for x in split_values]
#             # Actualizar la celda con los valores divididos
#             cell.value = split_values
  

# Guardar los cambios en el archivo Excel
wb2.save('5164_30_Limpieza.xlsx')

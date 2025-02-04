import openpyxl

# Ruta del archivo Excel
file_path = './excel/kuromi_hex.xlsx'  # Reemplaza con la ruta de tu archivo

# Cargar el archivo Excel
wb = openpyxl.load_workbook(file_path, data_only=True)

# Seleccionar la hoja (por ejemplo, "Hoja 4")
sheet = wb['Hoja 1']  # Asegúrate de que la hoja exista con este nombre

# Función para obtener el color hexadecimal de una celda
def get_hex_color(cell):
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.type == 'rgb':
        return f'#{fill.fgColor.rgb[:6]}'  # Extrae el color RGB en formato hexadecimal
    return '#FFFFFF'  # Si no hay color, asigna blanco por defecto

# Extraer los colores de cada celda y crear la matriz
colors = []
for row in sheet.iter_rows():
    row_colors = [get_hex_color(cell) for cell in row]
    colors.append(row_colors)

# Generar el código de la variable `colors`
colors_code = f"const colors = {colors};"

# Guardar el código en un archivo de texto
output_path = './colors/colors_variable.txt'
with open(output_path, 'w') as file:
    file.write(colors_code)

print(f'El archivo ha sido guardado exitosamente en {output_path}')

# ---------------------------- LIBRERIAS

import pandas as pd                          # Manejar datos en forma de tablas
import matplotlib.pyplot as plt              # Hacer gráficos
import seaborn as sns                        # Gráficos más facheros
from fpdf import FPDF
import xlsxwriter

# ---------------------------- AJUSTES INICIALES

# Cargar el dataset
archivo = "dataset.xlsx"
df = pd.read_excel(archivo, sheet_name="Planilla")  # Leer la hoja "Planilla"

# Configuración de estilo
sns.set(style="whitegrid")

# Por las dudas, convertir lo necesario a números
cols_numeric = ['Precio Original', 'Precio Usd', 'Monto Vendido USD', 'Unidades Vendidas']
df[cols_numeric] = df[cols_numeric].apply(pd.to_numeric, errors='coerce')

# Crear columna de fecha
if 'Año' in df.columns and 'Mes' in df.columns:
    df['Fecha'] = pd.to_datetime(df['Año'].astype(str) + '-' + df['Mes'].astype(str), errors='coerce')

# ---------------------------- INFORMES

# Información del DataFrame
print("\nInformación del DataFrame:")
print(df.info())                         # Info sobre columnas y tipos de datos

# Estadísticas descriptivas
print("\nEstadísticas Descriptivas:")
print(df.describe().round(2))            # Estadísticas básicas de columnas numéricas


# ---------------------------- ANALISIS

# ----- Marcas más populares

marcas_populares = df['Marca'].value_counts().head(10)     # 10 marcas más frecuentes en el dataset

plt.figure(figsize=(10,6))                      # Tamaño del Gráfico
sns.barplot(                                    # Gráfico de Barras Verticales
    x=marcas_populares.index, 
    y=marcas_populares.values,
    hue=marcas_populares.index, 
    palette='Blues_d',                          # Paleta de Colores
    legend=False)
plt.title(                                      # Título del Gráfico
    'Marcas Más Populares', 
    fontsize=16, 
    fontweight='bold', 
    color='darkblue', 
    family='sans-serif'
    )
plt.xlabel(                                     # Etiqueta del Eje X
    'Marca', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
    )
plt.ylabel(                                     # Etiqueta del Eje Y
    'Cantidad', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
    )
plt.xticks(                                     # Etiquetas del Eje X, rotadas a 45º, alineadas a derecha
    rotation=45, 
    ha='right', 
    fontsize=10, family='sans-serif'
    )
plt.grid(                                       # Grid ligero
    axis='y', 
    linestyle='--', 
    alpha=0.7
    )
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.gcf().set_facecolor('whitesmoke')           # Fondo del gráfico
plt.savefig(                                    # Guardar el gráfico en una imagen
    'marcas_populares.png', 
    dpi=300                                     # Aumentar resolución
    )
plt.close()

# ----- Productos más populares

productos_populares = df['Titulo Publicacion'].value_counts().head(10)   # 10 productos más frecuentes en el dataset

plt.figure(figsize=(10,6))                      # Tamaño del Gráfico
sns.barplot(                                    # Gráfico de Barras Horizontales
    x=productos_populares.values, 
    y=productos_populares.index, 
    hue=productos_populares.values, 
    palette='Reds_d',                          # Paleta de Colores
    legend=False
    )
plt.title(                                      # Título del gráfico
    'Productos Más Populares', 
    fontsize=16, 
    fontweight='bold', 
    color='darkred', 
    family='sans-serif'
)
plt.xlabel(                                     # Etiqueta del Eje X
    'Cantidad', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
)
plt.ylabel(                                     # Etiqueta del Eje Y
    'Producto', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
)

# Ajustar las etiquetas del eje X e Y
plt.xticks(fontsize=10, family='sans-serif')
plt.yticks(fontsize=10, family='sans-serif')

plt.grid(                                       # Grid ligero
    axis='x', 
    linestyle='--', 
    alpha=0.7)
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.gcf().set_facecolor('whitesmoke')           # Fondo del gráfico
plt.savefig(                                    # Guardar el gráfico en una imagen
    'productos_populares.png', 
    dpi=300                                     # Aumentar resolución
    )  
plt.close()


# ----- Promedio de productos comprados por marca
# Ordenar por promedio de unidades vendidas
productos_por_marca = df.groupby('Marca')['Unidades Vendidas'].mean().sort_values(ascending=False)   

# Limitar cantidad de marcas a mostrar
max_num_marcas = 20                                 
productos_top = productos_por_marca.head(max_num_marcas)

plt.figure(figsize=(12, 6))                     # Tamaño del Gráfico
sns.barplot(                                    # Gráfico de Barras Horizontales
    x=productos_top.values, 
    y=productos_top.index, 
    hue=productos_top.values, 
    palette='Blues_d',                          # Paleta de Colores
    legend=False
    )
plt.title(                                      # Título del Gráfico
    'Top 20 - Promedio de Productos Comprados por Marca', 
    fontsize=16, 
    fontweight='bold', 
    color='darkblue', 
    family='sans-serif'
)
plt.xlabel(                                     # Etiqueta del Eje X
    'Promedio de Unidades Vendidas', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
)
plt.ylabel(                                     # Etiqueta del Eje Y
    'Marca', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
)
plt.gca().invert_yaxis()                        # Invertir Eje Y para mostrar las marcas más populares arriba

# Ajustar las etiquetas del eje X e Y
plt.xticks(fontsize=10, family='sans-serif')
plt.yticks(fontsize=10, family='sans-serif')

plt.grid(                                       # Grid ligero
    axis='x', 
    linestyle='--', 
    alpha=0.7
    ) 
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.gcf().set_facecolor('whitesmoke')           # Fondo del gráfico
plt.savefig(                                    # Guardar el gráfico en una imagen
    'promedio_productos.png', 
    dpi=300                                     # Aumentar resolución
    )  
plt.close()


# ----- Gráfico de ventas mensuales

ventas_mensuales = df.groupby('Fecha')['Unidades Vendidas'].sum()

plt.figure(figsize=(12,6))                      # Tamaño del Gráfico
plt.plot(                                       # Gráfico de Lineas
    ventas_mensuales.index, 
    ventas_mensuales.values, 
    marker='o', 
    color='steelblue',                          # Color de la línea
    linewidth=2,                                # Grosor de la línea
    markersize=6,                               # Tamaño de los puntos
    markerfacecolor='darkblue',                 # Color de los puntos
    markeredgewidth=2,                          # Bordes de los puntos
    markeredgecolor='white'                     # Color de los bordes de los puntos
)
plt.title(                                      # Título del Gráfico
    'Evolución de Ventas Mensuales', 
    fontsize=16, 
    fontweight='bold', 
    color='darkblue', 
    family='sans-serif'
)
plt.xlabel(                                     # Etiqueta del Eje X
    'Fecha', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
)
plt.ylabel(                                     # Etiqueta del Eje Y
    'Unidades Vendidas', 
    fontsize=12, 
    fontstyle='italic', 
    color='gray', 
    family='sans-serif'
)
plt.xticks(                                     # Etiquetas del Eje X, rotadas a 45º, alineadas a derecha
    rotation=45, 
    ha='right', 
    fontsize=10, 
    family='sans-serif'
    )
plt.grid(                                       # Grid ligero
    axis='y', 
    linestyle='--', 
    alpha=0.7
    )
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.gcf().set_facecolor('whitesmoke')           # Fondo del gráfico
plt.savefig(                                    # Guardar el gráfico en una imagen
    'ventas_mensuales.png', 
    dpi=300                                     # Aumentar resolución
    )  
plt.close()

"""
# ----- Análisis de la inflación en los precios

# Calculamos el cambio porcentual de los precios de un mes a otro
df['Variación Precio'] = df['Precio Original'].pct_change(fill_method=None) * 100

tabla_data = df[['Fecha', 'Precio Original', 'Variación Precio']].dropna()    # Filtramos los valores NaN 

plt.figure(figsize=(10,6))                             # Tamaño del Gráfico

tabla = plt.table(                                     # Crear la tabla en el gráfico
    cellText=tabla_data.values,                        # Datos de la tabla
    colLabels=tabla_data.columns,                      # Nombres de las columnas
    loc='center',                                      # Ubicación de la tabla en el gráfico
    cellLoc='center',                                  # Alineación de los textos en las celdas
    colColours=['lightblue']*len(tabla_data.columns),  # Color de fondo de las columnas
    cellColours=[['whitesmoke']*len(tabla_data.columns) for _ in range(len(tabla_data))]  # Color de fondo de las celdas
    )

# Tamaño de la font y escalado de la tabla
tabla.auto_set_font_size(False)                        # Desactivamos tamaño automático 
tabla.set_fontsize(12)                                 # Tamaño de font = 12
tabla.scale(1.2, 1.2)                                  # Aumentamos el tamaño de la tabla un 20%

# Personalizamos las celdas de la tabla
for (i, j), cell in tabla.get_celld().items():  
    if i == 0:                                         # Si estamos en la primera fila
        cell.set_fontsize(14)                          # Letra más grande
        cell.get_text().set_fontweight('bold')         # Negrita
        cell.set_edgecolor('darkblue')                 # Color del borde de las celdas
        cell.set_facecolor('lightblue')                # Color de fondo de los encabezados
    else:                                              # Si estamos en las filas de datos
        cell.set_edgecolor('lightgray')                # Color del borde 
        cell.set_facecolor('whitesmoke')               # Color de fondo de las celdas
        cell.set_text_props(color='darkblue')          # Color del texto

plt.title(                                             # Título de la Tabla
    'Variación de Precios', 
    fontsize=16,  
    fontweight='bold',
    color='darkblue',  
    family='sans-serif'
)

plt.axis('off')                                        # Quitar ejes para que solo se vea la tabla

plt.tight_layout()                                     # Ajustar márgenes

plt.savefig(                                           # Guardar el gráfico en una imagen
    'variacion_precios_tabla.png', 
    dpi=300
    )

plt.close()

"""
# ----- Comparación de ventas entre productos nuevos y usados
ventas_estado = df.groupby('Estado')['Unidades Vendidas'].sum()

plt.figure(figsize=(6,6))                                            # Tamaño del gráfico

ventas_estado.plot(                                                  # Gráfico de Torta
    kind='pie', 
    autopct='%1.1f%%',                                               # Porcentaje con un decimal
    colors=['lightblue', 'lightseagreen'],  
    startangle=90,                                                   # Empezar la torta desde la parte superior
    wedgeprops={'edgecolor': 'black', 'linewidth': 1}                # Borde de las rebanadas
)

plt.title(                                                           # Título del gráfico
    'Comparación de Ventas: Productos Nuevos vs Usados', 
    fontsize=14, 
    fontweight='bold', 
    color='darkslategray'
    )

plt.ylabel('')                                                       # Oculta la etiqueta del eje Y

plt.legend(                                                          # Leyendas del Gráfico
    labels=ventas_estado.index,                                      # Nombres de las categorías
    loc='upper left',                                                # Ubicación
    fontsize=12,                                                     # Tamaño de font
    title='Estado del Producto',                                     # Título de la leyenda
    title_fontsize=13,                                               # Tamaño de font de leyenda
    frameon=False,                                                   # Leyenda sin borde
    handleheight=1.5  
)

plt.tight_layout()                                                   # Ajustar márgenes

plt.savefig(                                                         # Guardar el gráfico en una imagen
    'ventas_estado.png', 
    dpi=300)

plt.close()

# ----- Análisis del impacto de las ofertas
ventas_oferta = df.groupby('Esta en Oferta')['Unidades Vendidas'].mean()  

plt.figure(figsize=(6,6))                                            # Tamaño del gráfico

ventas_oferta.plot(                                                  # Gráfico de Torta
    kind='pie', 
    autopct='%1.1f%%',                                               # Porcentaje con un decimal
    colors=['lightblue', 'lightseagreen'], 
    startangle=90,                                                   # Empezar la torta desde la parte superior
    wedgeprops={'edgecolor': 'black', 'linewidth': 1}                # Borde de las rebanadas
)

plt.title(                                                           # Título del gráfico
    'Impacto de las Ofertas en las Ventas', 
    fontsize=14, 
    fontweight='bold', 
    color='darkslategray'
    )

plt.ylabel('')                                                       # Oculta la etiqueta del eje Y

plt.legend(                                                          # Leyendas del Gráfico
    labels=ventas_oferta.index,                                      # Nombres de las categorías
    loc='upper left',                                                # Ubicación
    fontsize=12,                                                     # Tamaño de font
    title='Estado de la Oferta',                                     # Título de la leyenda
    title_fontsize=13,                                               # Tamaño de font de leyenda
    frameon=False                                                    # Leyenda sin borde
)

plt.tight_layout()                                                   # Ajustar márgenes

plt.savefig(                                                         # Guardar el gráfico en una imagen
    'impacto_ofertas.png', 
    dpi=300)

plt.close()

# ----- Matriz de correlaciones para analizar relaciones entre variables
plt.figure(figsize=(10,6))                                           # Tamaño del gráfico

sns.heatmap(                                                         # Mapa de calor 
    df.corr(numeric_only=True),                                      # Correlación entre variables numéricas
    annot=True,                                                      # Valores numéricos dentro del gráfico
    cmap='coolwarm',                                                 # Esquema de colores (positivos, rojos) y (negativos, azules)
    fmt=".2f",                                                       # Números dentro del mapa con 2 decimales 
    cbar_kws={'label': 'Correlación'},                               # Etiqueta para la barra de color
    annot_kws={'size': 12, 'weight': 'bold', 'color': 'black'},      # Tamaño, peso y color del texto en las celdas
    linewidths=1,                                                    # Grosor de las líneas de las celdas
    linecolor='gray',                                                # Color de las líneas de separación
    vmin=-1, vmax=1                                                  # Rango de la barra de color
    )  
                                                     
plt.title(                                                           # Título del gráfico
    'Matriz de Correlaciones entre Variables', 
    fontsize=16, 
    fontweight='bold', 
    color='darkblue'
    )

plt.tight_layout()                                                   # Ajustar márgenes

plt.savefig(                                                         # Guardar el gráfico en una imagen
    'matriz_correlaciones.png', 
    dpi=300
    )

plt.close()

# ----- Comparación: ventas reales y predicción
# Se calcula un promedio de las ventas de los últimos 3 meses
df['Ventas Promedio Móvil'] = df['Unidades Vendidas'].rolling(window=3).mean()  

plt.figure(figsize=(12,6))                                            # Tamaño del gráfico

sns.lineplot(                                                         # Graficar las ventas reales
    data=df, x='Fecha', y='Unidades Vendidas',
    label="Ventas Reales", color='steelblue', linewidth=1.5, alpha=0.8
)

sns.lineplot(                                                         # Graficar la predicción
    data=df, x='Fecha', y='Ventas Promedio Móvil',
    label="Predicción (Prom. Móvil)", color='darkorange', linestyle='dashed', linewidth=2
)

plt.title(                                                            # Título del gráfico
    "Predicción Simple de Ventas",
    fontsize=16, 
    fontweight='bold', 
    color='darkblue'
)

plt.xlabel(                                                           # Etiqueta del Eje X
    "Fecha", 
    fontsize=12, 
    fontweight='bold', 
    color='black'
    )

plt.ylabel(                                                           # Etiqueta del Eje Y
    "Unidades Vendidas", 
    fontsize=12, 
    fontweight='bold', 
    color='black'
    )

plt.xticks(rotation=45)

plt.legend(
    frameon=True, 
    fontsize=12, 
    loc="upper left"
    )

plt.tight_layout()                                                   # Ajustar márgenes

plt.savefig(                                                         # Guardar el gráfico en una imagen
    'prediccion_ventas_mejorado.png', 
    dpi=300
    )

plt.close()

# ---------------------------- CREAR REPORTE EXCEL

with pd.ExcelWriter('reporte_ventas.xlsx', engine='xlsxwriter') as writer:
    # Exportar los DataFrames a hojas de Excel
    df.describe().to_excel(writer, sheet_name='Estadísticas')
    marcas_populares.to_frame().to_excel(writer, sheet_name='Marcas Populares')
    productos_populares.to_frame().to_excel(writer, sheet_name='Productos Populares')
    df[['Fecha', 'Precio Original']].dropna().to_excel(writer, sheet_name='Variación de Precios')

    # Obtener el objeto workbook y configurar formato
    workbook = writer.book
    bold_format = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#F4CCCC', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#FFDDC1', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    data_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    image_format = workbook.add_format({'border': 0, 'align': 'center'})

    # Agregar gráficos al Excel
    worksheet = writer.sheets['Estadísticas']
    worksheet.insert_image('E2', 'marcas_populares.png', {'x_scale': 0.5, 'y_scale': 0.5})  # Ajuste de escala para que el gráfico se vea más pequeño
    worksheet.insert_image('E20', 'productos_populares.png', {'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('E40', 'promedio_productos.png', {'x_scale': 0.5, 'y_scale': 0.5})
    # worksheet.insert_image('E60', 'variacion_precios.png', {'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('E80', 'prediccion_ventas_mejorado.png', {'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('E100', 'ventas_mensuales.png', {'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('E120', 'impacto_ofertas.png', {'x_scale': 0.5, 'y_scale': 0.5})
    worksheet.insert_image('E140', 'matriz_correlaciones.png', {'x_scale': 0.5, 'y_scale': 0.5})

    # Ajustes de formato en la hoja de "Estadísticas"
    worksheet.set_column('A:B', 20, data_format)
    worksheet.set_column('C:C', 15, data_format)
    worksheet.write(0, 0, 'Estadísticas Generales', bold_format)
    worksheet.write(1, 0, 'Descripción', header_format)
    worksheet.write(1, 1, 'Valor', header_format)

    # Añadir formato a la hoja "Marcas Populares"
    worksheet = writer.sheets['Marcas Populares']
    worksheet.set_column('A:B', 30, data_format)
    worksheet.write(0, 0, 'Marca', header_format)
    worksheet.write(0, 1, 'Número de Productos', header_format)

    # Añadir formato a la hoja "Productos Populares"
    worksheet = writer.sheets['Productos Populares']
    worksheet.set_column('A:B', 30, data_format)
    worksheet.write(0, 0, 'Producto', header_format)
    worksheet.write(0, 1, 'Número de Ventas', header_format)

    # Ajustar los anchos de las columnas según el contenido
    for sheet in writer.sheets.values():
        for col_num, col in enumerate(sheet.columns):
            max_length = 0
            column = col
            for cell in column:
                try:
                    if len(str(cell)) > max_length:
                        max_length = len(cell)
                except:
                    pass
            adjusted_width = (max_length + 2)
            sheet.set_column(col_num, col_num, adjusted_width, data_format)

    # Personalización adicional de la apariencia
    worksheet = writer.sheets['Estadísticas']
    worksheet.set_row(0, 25, bold_format)  # Hacer la primera fila más alta para títulos
    worksheet.set_row(1, 20, data_format)  # Asegurarse que las filas tengan el mismo tamaño

    # Para las hojas de 'Marcas Populares', 'Productos Populares' y 'Variación de Precios'
    for sheet_name in ['Marcas Populares', 'Productos Populares']:
        worksheet = writer.sheets[sheet_name]
        worksheet.set_row(0, 20, bold_format)  # Ajuste de la altura de la primera fila

    # Añadir bordes a las celdas de datos
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        worksheet.conditional_format('A1:Z1000', {'type': 'no_blanks', 'format': workbook.add_format({'border': 1})})


# ---------------------------- CREAR REPORTE PDF
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Reporte de Ventas', ln=True, align='C')
        self.ln(10)
    
    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, ln=True, align='L')
        self.ln(5)
    
    def chapter_body(self, body):
        self.set_font('Arial', '', 10)
        self.multi_cell(0, 7, body)
        self.ln()

pdf = PDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()
pdf.chapter_title('Resumen Ejecutivo')
pdf.chapter_body('Este informe presenta un análisis detallado de las ventas, las marcas más populares y la evolución de las ventas en el tiempo.')

pdf.chapter_title('Marcas Más Populares')
pdf.image('marcas_populares.png', x=10, w=180)

pdf.chapter_title('Productos Más Populares')
pdf.image('productos_populares.png', x=10, w=180)

pdf.chapter_title('Evolución de Ventas Mensuales')
pdf.image('ventas_mensuales.png', x=10, w=180)

pdf.chapter_title('Promedio de Productos Comprados por Marca')
pdf.image('promedio_producto.png', x=10, w=180)

pdf.chapter_title('Predicción de Ventas')
pdf.image('prediccion_ventas_mejorado.png', x=10, w=180)

pdf.chapter_title('Variación de Precios')
pdf.image('variacion_precios.png', x=10, w=180)

pdf.chapter_title('Impacto de Ofertas')
pdf.image('impacto_ofertas.png', x=10, w=180)

pdf.chapter_title('Matriz de Correlaciones')
pdf.image('matriz_correlaciones.png', x=10, w=180)

pdf.output('reporte_ventas.pdf')
print("Reporte generado con éxito: Excel y PDF guardados.")
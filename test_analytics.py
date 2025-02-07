# ---------------------------- LIBRERIAS

import pandas as pd                          # Manejar datos en forma de tablas
import matplotlib.pyplot as plt              # Hacer gráficos
import seaborn as sns                        # Gráficos más facheros
from fpdf import FPDF
import os 
from datetime import datetime

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

# ---------------------------- CREAR REPORTE PDF
class PDF(FPDF):
    def header(self):
        # Logo
        if os.path.exists('logo.jpeg'):
            self.image('logo.jpeg', 10, 8, 25)
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'Reporte de Ventas', ln=True, align='C')
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()} - {datetime.today().strftime("%Y-%m-%d")}', align='C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, title, ln=True, align='L')
        self.ln(5)

    def chapter_body(self, body):
        self.set_font('Arial', '', 10)
        self.multi_cell(0, 7, body)
        self.ln()

    def add_image(self, image_path):
        if os.path.exists(image_path):
            self.image(image_path, x=10, w=180)
        else:
            self.set_font('Arial', 'I', 10)
            self.cell(0, 10, f'Imagen no encontrada: {image_path}', ln=True, align='L')
        self.ln(5)

# Crear documento
pdf = PDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()

# Secciones
data = [
    ("Resumen Ejecutivo", "Este informe presenta un análisis detallado de las ventas, las marcas más populares y la evolución de las ventas en el tiempo."),
    ("Marcas Más Populares", "marcas_populares.png"),
    ("Productos Más Populares", "productos_populares.png"),
    ("Evolución de Ventas Mensuales", "ventas_mensuales.png"),
    ("Promedio de Productos Comprados por Marca", "promedio_productos.png"),
    ("Predicción de Ventas", "prediccion_ventas_mejorado.png"),
    ("Impacto de Ofertas", "impacto_ofertas.png"),
    ("Matriz de Correlaciones", "matriz_correlaciones.png")
]

for title, content in data:
    pdf.chapter_title(title)
    if isinstance(content, str) and content.endswith('.png'):
        pdf.add_image(content)
    else:
        pdf.chapter_body(content)

# Guardar PDF
pdf.output('reporte_ventas.pdf')
print("Reporte generado con éxito: PDF guardado.")
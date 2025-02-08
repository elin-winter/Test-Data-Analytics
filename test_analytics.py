# ---------------------------- LIBRERIAS

import pandas as pd                          # Manejar datos en forma de tablas
import matplotlib.pyplot as plt              # Hacer gráficos
import seaborn as sns                        # Gráficos más facheros
from fpdf import FPDF
import os 
from datetime import datetime

custom_palette = ['#2D336B', '#7886C7', '#A9B5DF', '#B7B7B7']

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

# ---------------------------- ANALISIS

# ----- Marcas más populares

marcas_populares = df['Marca'].value_counts().head(10)     # 10 marcas más frecuentes en el dataset

plt.figure(figsize=(10,6))                      # Tamaño del Gráfico
sns.barplot(                                    # Gráfico de Barras Verticales
    x=marcas_populares.index, 
    y=marcas_populares.values,
    hue=marcas_populares.index, 
    palette= custom_palette,                          # Paleta de Colores
    legend=False)
plt.title(                                      # Título del Gráfico
    'Top 10 - Marcas Más Populares', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B', 
    family='sans-serif'
    )
plt.xlabel(                                     # Etiqueta del Eje X
    'Marca', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
    )
plt.ylabel(                                     # Etiqueta del Eje Y
    'Cantidad', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
    )
plt.xticks(                                     # Etiquetas del Eje X, rotadas a 45º, alineadas a derecha
    rotation=45, 
    ha='right', 
    fontsize=14, family='sans-serif'
    )
plt.gcf().set_facecolor('#EEEEEE')              # Fondo del gráfico
plt.grid(                                       # Grid ligero
    axis='y', 
    linestyle='--', 
    alpha=0.7
    )
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.savefig(                                    # Guardar el gráfico en una imagen
    'marcas_populares.png', 
    dpi=300                                     # Aumentar resolución
    )
plt.close()

# ----- Productos más populares

productos_populares = df['Titulo Publicacion'].value_counts().head(10)   # 10 productos más frecuentes en el dataset

plt.figure(figsize=(12,6))                      # Tamaño del Gráfico
sns.barplot(                                    # Gráfico de Barras Horizontales
    x=productos_populares.values, 
    y=productos_populares.index, 
    hue=productos_populares.values, 
    palette=custom_palette,                          # Paleta de Colores
    legend=False
    )
plt.title(                                      # Título del gráfico
    'Top 10 - Productos Más Populares', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B', 
    family='sans-serif'
)
plt.xlabel(                                     # Etiqueta del Eje X
    'Cantidad', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
)
plt.ylabel(                                     # Etiqueta del Eje Y
    'Producto', 
    fontsize=12, 
    fontweight='bold',
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
)

# Ajustar las etiquetas del eje X e Y
plt.xticks(fontsize=14, family='sans-serif')
plt.yticks(fontsize=14, family='sans-serif')

plt.gcf().set_facecolor('#EEEEEE')           # Fondo del gráfico
plt.grid(                                       # Grid ligero
    axis='x', 
    linestyle='--', 
    alpha=0.7)
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.savefig(                                    # Guardar el gráfico en una imagen
    'productos_populares.png', 
    dpi=300                                     # Aumentar resolución
    )  
plt.close()


# ----- Promedio de productos comprados por marca
# Ordenar por promedio de unidades vendidas
productos_por_marca = df.groupby('Marca')['Unidades Vendidas'].mean().sort_values(ascending=False)   

# Limitar cantidad de marcas a mostrar
max_num_marcas = 10                                 
productos_top = productos_por_marca.head(max_num_marcas)

plt.figure(figsize=(14, 6))                     # Tamaño del Gráfico
sns.barplot(                                    # Gráfico de Barras Horizontales
    x=productos_top.values, 
    y=productos_top.index, 
    hue=productos_top.values, 
    palette=custom_palette,                          # Paleta de Colores
    legend=False
    )
plt.title(                                      # Título del Gráfico
    'Promedio de Productos Comprados por Marca', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B', 
    family='sans-serif'
)
plt.xlabel(                                     # Etiqueta del Eje X
    'Promedio de Unidades Vendidas', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
)
plt.ylabel(                                     # Etiqueta del Eje Y
    'Marca', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
)
plt.gca().invert_yaxis()                        # Invertir Eje Y para mostrar las marcas más populares arriba

# Ajustar las etiquetas del eje X e Y
plt.xticks(fontsize=14, family='sans-serif')
plt.yticks(fontsize=14, family='sans-serif')

plt.gcf().set_facecolor('#EEEEEE')              # Fondo del gráfico

plt.grid(                                       # Grid ligero
    axis='x', 
    linestyle='--', 
    alpha=0.7
    ) 
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
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
    color='#A9B5DF',                          # Color de la línea
    linewidth=2,                                # Grosor de la línea
    markersize=6,                               # Tamaño de los puntos
    markerfacecolor='#2D336B',                 # Color de los puntos
    markeredgewidth=2,                          # Bordes de los puntos
    markeredgecolor='white'                     # Color de los bordes de los puntos
)
plt.title(                                      # Título del Gráfico
    'Evolución de Ventas Mensuales', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B', 
    family='sans-serif'
)
plt.xlabel(                                     # Etiqueta del Eje X
    'Fecha', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
)
plt.ylabel(                                     # Etiqueta del Eje Y
    'Unidades Vendidas', 
    fontsize=15,
    fontweight='bold', 
    fontstyle='italic', 
    color='#7886C7', 
    family='sans-serif'
)
plt.xticks(                                     # Etiquetas del Eje X, rotadas a 45º, alineadas a derecha
    rotation=45, 
    ha='right', 
    fontsize=14, 
    family='sans-serif'
    )
plt.gcf().set_facecolor('#EEEEEE')              # Fondo del gráfico
plt.grid(                                       # Grid ligero
    axis='y', 
    linestyle='--', 
    alpha=0.7
    )
plt.tight_layout(pad=2)                         # Ajustar márgenes y agregar espacio extra
plt.gcf().set_facecolor('#EEEEEE')           # Fondo del gráfico
plt.savefig(                                    # Guardar el gráfico en una imagen
    'ventas_mensuales.png', 
    dpi=300                                     # Aumentar resolución
    )  
plt.close()

# ----- Análisis de la inflación en los precios

# Calculamos la variación porcentual de los precios
df['Variación Precio'] = df['Precio Original'].pct_change() * 100

# Gráfico de líneas para la variación de precios
plt.figure(figsize=(12, 6))
sns.lineplot(
    x='Fecha',
    y='Variación Precio',
    data=df,
    marker='o',
    color='#2D336B'
)

# Personalizamos el gráfico
plt.title(
    'Variación de Precios a lo Largo del Tiempo',
    fontsize=20,
    fontweight='bold',
    color='#2D336B',
    family='sans-serif'
)
plt.xlabel(
    'Fecha',
    fontsize=15,
    fontweight='bold',
    fontstyle='italic',
    color='#7886C7',
    family='sans-serif'
)
plt.ylabel(
    'Variación Porcentual de Precio',
    fontsize=15,
    fontweight='bold',
    fontstyle='italic',
    color='#7886C7',
    family='sans-serif'
)
plt.xticks(rotation=45)
plt.gcf().set_facecolor('#EEEEEE')           # Fondo del gráfico
plt.grid(
    True, 
    linestyle='--', 
    alpha=0.7
    )
plt.tight_layout()
plt.savefig('variacion_precio.png', dpi=300)
plt.close()

# ----- Comparación de ventas entre productos nuevos y usados
ventas_estado = df.groupby('Estado')['Unidades Vendidas'].sum()

plt.figure(figsize=(10,6))                                            # Tamaño del gráfico

ventas_estado.plot(                                                  # Gráfico de Torta
    kind='pie', 
    autopct='%1.1f%%',                                               # Porcentaje con un decimal
    colors=['#2D336B', '#7886C7'],  
    startangle=90,                                                   # Empezar la torta desde la parte superior
    wedgeprops={'edgecolor': 'black', 'linewidth': 0.5}                # Borde de las rebanadas
)

plt.title(                                                           # Título del gráfico
    'Comparación de Ventas: Productos Nuevos vs Usados', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B'
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

plt.gcf().set_facecolor('#EEEEEE')                                   # Fondo del gráfico

plt.tight_layout()                                                   # Ajustar márgenes

plt.savefig(                                                         # Guardar el gráfico en una imagen
    'ventas_estado.png', 
    dpi=300)

plt.close()

# ----- Análisis del impacto de las ofertas
ventas_oferta = df.groupby('Esta en Oferta')['Unidades Vendidas'].mean()  

plt.figure(figsize=(10,6))                                            # Tamaño del gráfico

ventas_oferta.plot(                                                  # Gráfico de Torta
    kind='pie', 
    autopct='%1.1f%%',                                               # Porcentaje con un decimal
    colors=['#2D336B', '#7886C7'], 
    startangle=90,                                                   # Empezar la torta desde la parte superior
    wedgeprops={'edgecolor': 'black', 'linewidth': 1}                # Borde de las rebanadas
)

plt.title(                                                           # Título del gráfico
    'Comparación de Ventas: En Oferta vs Sin Oferta', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B'
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

plt.gcf().set_facecolor('#EEEEEE')                                   # Fondo del gráfico

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
    linecolor='#7886C7',                                                # Color de las líneas de separación
    vmin=-1, vmax=1                                                  # Rango de la barra de color
    )  
                                                     
plt.title(                                                           # Título del gráfico
    'Matriz de Correlaciones entre Variables', 
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B'
    )

plt.gcf().set_facecolor('#EEEEEE')                                   # Fondo del gráfico

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
    label="Ventas Reales", color='#2D336B', linewidth=1.5, alpha=0.3
)

sns.lineplot(                                                         # Graficar la predicción
    data=df, x='Fecha', y='Ventas Promedio Móvil',
    label="Predicción (Prom. Móvil)", color='#FBA518', linestyle='dashed', linewidth=2
)

plt.title(                                                            # Título del gráfico
    "Comparación de Ventas: Predicción vs Realidad",
    fontsize=20, 
    fontweight='bold', 
    color='#2D336B'
)

plt.xlabel(                                                           # Etiqueta del Eje X
    "Fecha", 
    fontsize=15, 
    fontweight='bold', 
    color='#7886C7'
    )

plt.ylabel(                                                           # Etiqueta del Eje Y
    "Unidades Vendidas", 
    fontsize=12, 
    fontweight='bold', 
    color='#7886C7'
    )

plt.xticks(rotation=45)

plt.legend(
    frameon=True, 
    fontsize=12, 
    loc="upper left"
    )

plt.gcf().set_facecolor('#EEEEEE')                                   # Fondo del gráfico

plt.tight_layout()                                                   # Ajustar márgenes

plt.savefig(                                                         # Guardar el gráfico en una imagen
    'prediccion_ventas_mejorado.png', 
    dpi=300
    )

plt.close()
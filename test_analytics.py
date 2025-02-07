# ---------------------------- LIBRERIAS

import pandas as pd                          # Manejar datos en forma de tablas
import matplotlib.pyplot as plt              # Hacer gráficos
import seaborn as sns                        # Gráficos más facheros

# ---------------------------- AJUSTES INICIALES

# Cargar el dataset
archivo = "dataset.xlsx"
df = pd.read_excel(archivo, sheet_name="Planilla")  # Leer la hoja "Planilla"

# Configuración de estilo
sns.set(style="whitegrid")

# Por las dudas, convertir lo necesario a números
df['Precio Original'] = pd.to_numeric(df['Precio Original'], errors='coerce')
df['Precio Usd'] = pd.to_numeric(df['Precio Usd'], errors='coerce')
df['Monto Vendido USD'] = pd.to_numeric(df['Monto Vendido USD'], errors='coerce')

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
print("Marcas más populares:\n", marcas_populares)

plt.figure(figsize=(10, 6))                                # Tamaño del gráfico
sns.barplot(                                               # Gráfico de Barras Vertical
    x=marcas_populares.index,                              
    y=marcas_populares.values,                             
    palette='Blues'                                        # Paleta de Colores
    )  
plt.title(                                                 # Título del gráfico
    'Marcas Más Populares', 
    fontsize=16, 
    weight='bold'
    )
plt.xlabel('Marca', fontsize=12)                           # Etiqueta del eje X
plt.ylabel('Cantidad de Publicaciones', fontsize=12)       # Etiqueta del eje Y
plt.xticks(rotation=45, ha='right')                        # Etiquetas del eje X, rotadas 45 grados, alineadas a la derecha
plt.tight_layout()                                         # Ajusta márgenes 
plt.show()

# ----- Productos más populares

productos_populares = df['Titulo Publicacion'].value_counts().head(10)   # 10 productos más frecuentes en el dataset
print("Productos más populares:\n", productos_populares)

plt.figure(figsize=(10, 6))                                              # Tamaño del gráfico
productos_populares.plot(kind='barh', color='lightcoral')                # Gráfico de Barras Horizontal
plt.title(                                                               # Título del gráfico
    'Productos Más Populares', 
    fontsize=16, 
    weight='bold'
    ) 
plt.xlabel('Cantidad de Ventas', fontsize=12)                            # Etiqueta del eje X
plt.ylabel('Producto', fontsize=12)                                      # Etiqueta del eje Y
plt.tight_layout()                                                       # Ajusta márgenes
plt.show()

# ----- Promedio de productos comprados por marca
productos_por_marca = df.groupby('Marca')['Unidades Vendidas'].mean().sort_values(ascending=False)
print("Promedio de productos comprados por marca:\n", productos_por_marca)

plt.figure(figsize=(10, 6))                                              # Tamaño del gráfico
productos_por_marca.plot(kind='bar', color='lightgreen')                 # Gráfico de Barras Vertical
plt.title('Promedio de Productos Comprados por Marca', fontsize=16)      # Título del gráfico
plt.xlabel('Marca', fontsize=12)                                         # Etiqueta del eje X
plt.ylabel('Promedio de Unidades Vendidas', fontsize=12)                 # Etiqueta del eje Y
plt.xticks(rotation=45, ha='right')                                      # Etiquetas del eje X, rotadas 45 grados, alineadas a la derecha
plt.tight_layout()                                                       # Ajusta márgenes 
plt.show()

# ----- Gráfico de ventas mensuales

# Crear una columna con la fecha completa para analizar tendencias
# Une el año y el mes en una fecha real
df['Fecha'] = pd.to_datetime(df['Año'].astype(str) + '-' + df['Mes'].astype(str))
ventas_mensuales = df.groupby('Fecha')['Unidades Vendidas'].sum()

# Generar gráfico
plt.figure(figsize=(12,6))                              # Tamaño del gráfico
plt.plot(
    ventas_mensuales.index, 
    ventas_mensuales.values, 
    marker='o', linestyle='-')                          # Evolución de ventas
plt.xlabel('Fecha')                                     # Etiqueta del eje X
plt.ylabel('Unidades Vendidas')                         # Etiqueta del eje Y
plt.title('Evolución de las Ventas Mensuales')          # Título del gráfico
plt.grid(True)                                          # Agrega una cuadrícula para facilitar la lectura 
plt.show()                                              # Muestra el gráfico

# ----- Análisis de la inflación en los precios

# Calculamos el cambio porcentual de los precios de un mes a otro
df['Variación Precio'] = df['Precio Original'].pct_change() * 100
print("Variación de Precios:\n", df[['Fecha', 'Precio Original', 'Variación Precio']].dropna())

# ----- Comparación de ventas entre productos nuevos y usados
ventas_estado = df.groupby('Estado')['Unidades Vendidas'].sum()
plt.figure(figsize=(6,6))                                            # Tamaño del gráfico
ventas_estado.plot(kind='pie', autopct='%1.1f%%')                    # Gráfico de torta con porcentajes
plt.title('Comparación de Ventas: Productos Nuevos vs Usados')       # Título del gráfico
plt.ylabel('')                                                       # Oculta la etiqueta del eje Y para mayor claridad
plt.show()

# ----- Análisis del impacto de las ofertas
ventas_oferta = df.groupby('Esta en Oferta')['Unidades Vendidas'].mean()  
print("Impacto de las ofertas:\n", ventas_oferta)

# ----- Matriz de correlaciones para analizar relaciones entre variables
plt.figure(figsize=(10,6))                                           # Tamaño del gráfico
sns.heatmap(                                                         # Mapa de calor 
    df.corr(numeric_only=True),                                      # Correlación entre variables numéricas
    annot=True,                                                      # Valores numéricos dentro del gráfico
    cmap='coolwarm',                                                 # Esquema de colores (positivos, rojos) y (negativos, azules)
    fmt=".2f")                                                       # Números dentro del mapa con 2 decimales
plt.title("Matriz de Correlaciones")                                 # Título del gráfico
plt.show()

# ----- Predicción simple de ventas usando un promedio móvil
# Se calcula un promedio de las ventas de los últimos 3 meses
df['Ventas Promedio Móvil'] = df['Unidades Vendidas'].rolling(window=3).mean()  

# Gráfico de Comparación: ventas reales y predicción
plt.figure(figsize=(12,6))                                            # Tamaño del gráfico
plt.plot(                                                             # Para las ventas reales
    df['Fecha'], 
    df['Unidades Vendidas'], 
    label="Ventas Reales")                                           
plt.plot(                                                             # Para la predicción
    df['Fecha'], 
    df['Ventas Promedio Móvil'], 
    label="Predicción (Prom. Móvil)", 
    linestyle='dashed')                                              
plt.legend()                                                          # Agrega una leyenda para diferenciar las líneas
plt.title("Predicción Simple de Ventas")                              # Título del gráfico
plt.xlabel("Fecha")                                                   # Etiqueta del eje X (Fechas)
plt.ylabel("Unidades Vendidas")                                       # Etiqueta del eje Y (Ventas)
plt.show()                                                            # Muestra el gráfico

with pd.ExcelWriter('reporte_ventas.xlsx') as writer:
    marcas_populares.to_frame().to_excel(writer, sheet_name='Marcas Populares')
    productos_populares.to_frame().to_excel(writer, sheet_name='Productos Populares')
    productos_por_marca.to_frame().to_excel(writer, sheet_name='Promedio por Marca')
    df.describe().round(2).to_excel(writer, sheet_name='Estadísticas Descriptivas')
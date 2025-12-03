import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.style as mpl_style
import numpy as np
import os

archivo = 'vendedores-1.xlsx'

if not os.path.exists(archivo):
    print(f"No se pudo encontrar '{archivo}'.")
    print("Asegurate de que el archivo este en la misma carpeta que este archivo de Python.")
    exit()

print(f"Archivo {archivo}' encontrado. Iniciando analisis...\n")

try:
    df = pd.read_excel(archivo)
except Exception as e:
    print(f"Error al leer el Excel: {e}")
    exit()

mpl_style.use('ggplot')

print("-- RESUMEN DE DATOS --")
print(f"Total de registros: {df.shape[0]}")
print(f"Total de columnas: {df.shape[1]}")
print(f"\nEstadisticas clave: ")
print(df[['SALARIO', 'VENTAS TOTALES']].describe())

nulos = df.isnull().sum().sum()
print(f"\nDatos faltantes encontrados: {nulos}")
if nulos > 0:
    print("Advertencia: Hay datos vacios, se recomienda revisarlos.")

ventas_region = df.groupby('REGION')['VENTAS TOTALES'].sum()

plt.figure(figsize=(10, 6))
barras = plt.bar(ventas_region.index, ventas_region.values,
                 color='cornflowerblue', edgecolor='black')
plt.title('Ventas Totales por Region', fontsize=14)
plt.ylabel('Monto ($)')
plt.grid(axis='y', linestyle='--', alpha=0.6)

for bar in barras:
    yval = bar.get_height()
    plt.text(bar.get_x() + bar.get_width()/2, yval,
             f'${yval:,.0f}', ha='center', va='bottom')

plt.tight_layout()
print("Generando grafico: Ventras por region...")
plt.show()

regiones = df['REGION'].unique()

for region in regiones:
    data_reg = df[df['REGION'] == region]
    nombre = data_reg['NOMBRE'] + " " + data_reg['APELLIDO']

    plt.figure(figsize=(10, 5))
    plt.bar(nombre, data_reg['VENTAS TOTALES'], color='mediumseagreen')
    plt.title(f"Desempeño Vendedores: Region {region}")
    plt.xticks(rotation=45, ha='right')
    plt.ylabel('Ventas')
    plt.tight_layout()
    print(f"Generando grafico: Vendedores Region {region}")
    plt.show()

idx_max = df.groupby('REGION')['VENTAS TOTALES'].idxmax()
mejores = df.loc[idx_max]
etiquetas_mejores = mejores['REGION'] + '\n' + mejores['NOMBRE']

plt.figure(figsize=(8, 8))
plt.pie(mejores['VENTAS TOTALES'], labels=etiquetas_mejores,
        autopct='%1.1f%%', startangle=90, colors=plt.cm.Pastel1.colors)
plt.title('Los Mejores Vendedores de cada Region')
print("Generando grafico: Mejores Vendedores (Grafica de Pastel)...")
plt.show()

idx_min = df.groupby('REGION')['VENTAS TOTALES'].idxmin()
peores = df.loc[idx_min]
etiquetas_peores = peores['REGION'] + '\n' + peores['NOMBRE']

plt.figure(figsize=(8, 8))
plt.pie(peores['VENTAS TOTALES'], labels=etiquetas_peores,
        autopct='%1.1f%%', startangle=90, colors=plt.cm.Set3.colors)
plt.title('Vendedores con Menores Ventas por Región')
print("📊 Generando gráfico: Peores Vendedores (Pie Chart)...")
plt.show()


top5 = df.nlargest(5, 'UNIDADES VENDIDAS').sort_values(
    'UNIDADES VENDIDAS', ascending=True)
nombres_top5 = top5['NOMBRE'] + " " + \
    top5['APELLIDO'] + " (" + top5['REGION'] + ")"

plt.figure(figsize=(10, 6))
plt.barh(nombres_top5, top5['UNIDADES VENDIDAS'],
         color='orange', edgecolor='black')
plt.title('Top 5: Mayores Unidades Vendidas (Global)')
plt.xlabel('Unidades')
print("Generando grafico: Top 5 Unidades...")
plt.tight_layout()
plt.show()

print("\n" + "="*40)
print("Datos y Conclusiones")
print("="*40)

mejor_region = ventas_region.idxmax()
peor_region = ventas_region.idxmin()
promedio_global = df['VENTAS TOTALES'].mean()

print(
    f"1. La región más productiva es **{mejor_region}** con ventas totales de ${ventas_region.max():,.2f}.")
print(
    f"2. La región con menores ingresos es **{peor_region}**, lo que sugiere enfocar estrategias de marketing allí.")
print(f"3. El promedio de ventas por vendedor es de: ${promedio_global:,.2f}.")
print(
    f"4. El vendedor 'Estrella' de toda la empresa es: {df.loc[df['VENTAS TOTALES'].idxmax(), 'NOMBRE']} {df.loc[df['VENTAS TOTALES'].idxmax(), 'APELLIDO']} ({df.loc[df['VENTAS TOTALES'].idxmax(), 'REGION']}).")

print("\nProceso finalizado.")

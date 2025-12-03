# %% [markdown]
# 
# ## Cargar librerias necesarias

# %%
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
# from pandas import DataFrame, Series

# %%
## Importar CSV de Estructura funcional Ubicacion
import pandas as pd
import re
import os

# This function will convert the url to a download link
def convert_gsheets_url(u):
    try:
        worksheet_id = u.split('#gid=')[1]
    except:
        # Couldn't get worksheet id. Ignore it
        worksheet_id = None
    u = re.findall('https://docs.google.com/spreadsheets/d/.*?/',u)[0]
    u += 'export'
    u += '?format=csv'
    if worksheet_id:
        u += '&gid={}'.format(worksheet_id)
    return u

URL = 'https://docs.google.com/spreadsheets/d/1OkECu7qNfGZxX_rc2RDbaz0A-oE_gUwJ0P2tjU_x-q0/edit#gid=1115106678'
csv_raw = convert_gsheets_url(URL)
print(csv_raw)


# Si existe el CSV importarlo, sino cargar de la URL
filename = "csv/MaestroActividades.csv"

if os.path.exists(filename):
    print("Leyendo desde Archivo")
    df_equipos = pd.read_csv(filename,header=3)
else:
    print("Leyendo desde URL")
    df_equipos = pd.read_csv(csv_raw,header=3)
print(df_equipos.shape)


# %%
## Importar CSV de Planes de Mantto
URL = 'https://docs.google.com/spreadsheets/d/1OkECu7qNfGZxX_rc2RDbaz0A-oE_gUwJ0P2tjU_x-q0/edit#gid=1199302294'
csv_raw = convert_gsheets_url(URL)
print(csv_raw)
# Si existe el CSV importarlo, sino cargar de la URL
filename = "csv/Planes.csv"
if os.path.exists(filename):
    print("Leyendo desde Archivo")
    df_planes = pd.read_csv(filename,header=3)
else:
    print("Leyendo desde URL")
    df_planes = pd.read_csv(csv_raw,header=3)
print(df_planes.shape)
# Eliminar si tiene valores nulos
df_planes = df_planes [df_planes .Plan.notnull()]
print(df_planes.shape)

# %%
# En el dataframe de planes, en la fila del plan se coloca si tiene alguna frecuencia al inicio del plan
df_update = df_planes[~df_planes.Accion.isna()]
gb = df_update.groupby('Plan',sort=False).any()
gb = gb.reset_index()
# Concatenamos los dos dataframes por filas
df_planes = pd.concat([gb,df_update], axis=0, ignore_index=True)

# %%
# En caso no este instalado, no se puede manipular excel
#!pip install openpyxl

# %%
# Hacemos un join de los dos DataFrames por la izquierda
df_merged = pd.merge(df_equipos, df_planes, how='left', left_on= 'Plan', right_on= 'Plan',suffixes=('_equipo', '_plan'))
# aplicar una funcion en una columna de un dataframe dependiendo si la columna que actualizamos no estan dentro de ['Actividad','Tarea'], si no copiar el valor de la columna 'tipo_equipo', si tiene valor, mantener el valor de la columna 'tipo_plan'
df_merged['Tipo_plan']=df_merged.apply(lambda row: row['Tipo_equipo'] if  row['Tipo_plan']  not in ['Actividad','Tarea']  else  row ['Tipo_plan'] , axis='columns')

# realizar una funcion apply en el dataframe df_merge que evalue el contenido de la columna 'Tipo_Plan', si su valor es 'Componente' que vacie el contenido de la columna 'Accion' y la columna 'Actividad', de lo contrario que mantenga su valor
df_merged['Accion'] = df_merged.apply(lambda row : '' if row['Tipo_plan'] == 'Componente' else row['Accion'],axis='columns')
df_merged['Actividad'] = df_merged.apply(lambda row : '' if row['Tipo_plan'] == 'Componente' else row['Actividad'],axis='columns')

#Exportar dataframe a Excel para visualizacion 
# df_merged.to_excel('xls/merged.xlsx') # 12Mb el archivo, dificil de analizar

# %%
# Leer dataframe desde archivo excel
#df = df_merged.copy(deep=True) # pd.read_excel('xls/test_group.xlsx', sheet_name='Sheet1', index_col='id')

# EQUIPOS
index = df_merged[df_merged['Tipo_plan']=='Equipo'].index.to_list()
#print(len(index))
# obtener las columnas que se desean actualizar
cols = ['Sist','Subs','Equ','Comp','D', 'S', 'M', '2M', 'T', '4M', 'SE', '8M', 'A', '1.5A', '2A', '3A', '4A', '5A', '6A', '9A', '10A', '1000', '6000', '22500', '40000', '55000']
m_cols = df_merged.columns.to_list()
update_cols = [i for i in m_cols if i not in cols]
copy_cols = update_cols + ['Sist','Subs','Equ','Comp']
print (copy_cols)

df_g = df_merged[df_merged['Tipo_plan'].isin(['Equipo','Componente'])][cols]

# Agrupar por Equ
gb= df_g.groupby(['Equ'],sort=False).any().reset_index()
gb['id'] = index
gb.set_index('id',inplace=True)
#print (gb)

#TODO: Necesario mejorar el rendimiento del algoritmo aparecen mensajes de warning
#gb = pd.concat([gb,df[cols]],axis = 1)
gb[copy_cols] = df_merged.loc[gb.index,copy_cols]
df_merged.loc[gb.index,:] = gb

#df.head(302).to_excel('xls/df.xlsx')

# Subsistema
index = df_merged[df_merged['Tipo_plan']=='Subsistema'].index.to_list()
print((index))
df_g = df_merged[df_merged['Tipo_plan'].isin(['Equipo','Subsistema'])][cols]

# Agrupar por Subs
gb= df_g.groupby(['Subs'],sort=False).any().reset_index()
gb['id'] = index
gb.set_index('id',inplace=True)
# dar formato para copiar
gb[copy_cols] = df_merged.loc[gb.index,copy_cols]
df_merged.loc[gb.index,:] = gb

#df.head(302).to_excel('xls/df.xlsx')

############# Sistema
index = df_merged[df_merged['Tipo_plan']=='Sistema'].index.to_list()
#print(index)
df_g = df_merged[df_merged['Tipo_plan'].isin(['Sistema','Subsistema'])][cols]
# Agrupar por Sist
gb= df_g.groupby(['Sist'],sort=False).any().reset_index()
gb['id'] = index
gb.set_index('id',inplace=True)
# dar formato para copiar
gb[copy_cols] = df_merged.loc[gb.index,copy_cols]
df_merged.loc[gb.index,:] = gb
df_merged.head(302).to_excel('xls/df_merged.xlsx')
#print(df)

# %%
# Abrir archivo json
import json
with open('./csv/Lineas.json') as f:
    diccionario_lineas = json.load(f)

frecuencias=['D', 'S', 'M', '2M', 'T', '4M', 'SE', '8M', 'A', '1.5A', '2A', '3A', '4A', '5A', '6A', '9A', '10A', '1000', '6000', '22500', '40000', '55000']

# df_merged[ ['id', 'Cod', 'Accion', 'Actividad', 'Tipo_plan', 'Parada', 'Relevancia', 'Especialidad']+frecuencias]
head = ['id', 'Cod', 'Equipo','Componente', 'Accion', 'Actividad', 'Tipo_plan', 'Parada', 'Relevancia', 'Especialidad'] 

version = "_v1_1_2023"

for linea in diccionario_lineas['Lineas']:
    print (linea['Linea'])
    print ( head + frecuencias + linea['Columnas'])       
    # exportar plan maestro de linea
    df_linea = df_merged[head + frecuencias + linea['Columnas']]

    # A nivel de componente, solamente mantiene el primer valor, si encuentra mas adelante coloca ''
    # df_linea.loc[df_linea.duplicated(subset=['Componente']), 'Componente'] = '' 

    #Filtrar solamente los datos de la columna de la linea correspondiente, no todas las demas columnas
    df_linea = df_linea[df_linea[linea['Linea']] == True]
    ## Exportar Excel de Estructura - Planes    
    df_linea.to_excel(f"xls/{linea['Linea']}{version}.xlsx" , index= False)     
    print(df_linea.shape)        



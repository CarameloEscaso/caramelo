#!/usr/bin/env python
# coding: utf-8

# In[7]:


import pandas as pd
import glob


# In[8]:


def clean_str(string):
    """ 
    funcion que toma unstring y lo limpia de tildes y caracteres especiales, devuelve un string en minuscula
    """
    string_new = string.replace("¿","").replace("?","").lstrip().rstrip().replace(",","").\
                            replace("á","a").replace("é","e").replace("í","i").replace("ó","o").\
                            replace("ú","u").lower().replace("¡","").replace("!","").replace(".","")
    return string_new


# In[57]:


def tabla_insumo_agg():
    # Leer archivos de 
    file_cat = glob.glob('catalogos/*.xlsx')
    file_ins = glob.glob('insumo/*.xlsx')

    # Catalogo de respuestas

    inverso = file_cat[0] #r"C:\Users\yeimunoz\Documents\catalogos\catalogo_1.xlsx"
    puntaje = file_cat[1] #r"C:\Users\yeimunoz\Documents\catalogos\catalogo_2.xlsx"

    # Leer archivos
    catalogo_punt = pd.read_excel(puntaje)
    catalogo_inverso= pd.read_excel(inverso)
    insumo = pd.read_excel(file_ins[0], sheet_name="Empresa", skiprows=[0]).set_index("#")

    # Limpiar el valor de las respuestas de signos
    catalogo_punt["respuesta2"] = catalogo_punt.apply(lambda row: clean_str(row["Respuestas"]), axis=1) 
    catalogo_inverso["Inversa"] = catalogo_inverso.apply(lambda row: str(row["Inversa"]), axis=1) 


    # indentificar las preguntas inversas

    insumo_inv = insumo.T.reset_index().merge(catalogo_inverso[["Identificador_pregunta", "Inversa"]]
                                 , left_on="index", right_on="Identificador_pregunta", how="left").set_index("index").T

    columnas = insumo.loc[:, ~insumo.columns.str.contains('^Unnamed')].columns

    #limpiar respuestas del insumo

    df_cols = pd.DataFrame()
    for i in columnas:
        insumo_inv[i+"_lim"] = insumo_inv.apply(lambda row: clean_str(row[i]), axis=1) 

    # Pega los puntajes a las respuestas
    df_temp = []
    for i in columnas:
        temp = insumo_inv.reset_index()[["index", i+"_lim"]].merge(catalogo_punt[["respuesta2","puntaje"]]
                                                                , left_on=i+"_lim"
                                                                , right_on="respuesta2"
                                                                , how="left").set_index("index").drop(columns=["respuesta2"])
        df_temp.append(temp)

    # Limpia la basura que queda despues de pegar los puntajes
    df_total = pd.DataFrame()
    j=0
    cols = []
    for i in df_temp:
        if j==0:
            col = "pnt"+"_"+columnas[j]
            df_total = i.rename(columns={"puntaje":col})
        else:
            col = "pnt"+"_"+columnas[j]
            df_total = df_total.merge(i, left_index=True, right_index=True, how ="left").\
                                        rename(columns={"puntaje":col})
        cols.append(col)
        j+=1
    df = df_total.T.fillna(method="ffill")
    df_total_pnt = df.T

    # base calificada con los puntajes invertidos
    df_t = df_total_pnt[cols].T
    invers = {1:5,2:4,3:3,4:2,5:1}
    for c in df_t.columns[:-2]:
        df_t[c] = df_t.apply(lambda row:  row[c] if  row["Inversa"] == "00"
                           else  invers[row[c]] , axis=1)

    df_tabla_final = df_t.T

#     resultado con las agregaciones y conteos
    df_agg = pd.DataFrame([1,2,3,4,5],columns={"index"}).set_index("index")
    df_mean = pd.DataFrame()
    df_sum = pd.DataFrame()
    for c in df_tabla_final.columns:
        temp1 = df_tabla_final[~df_tabla_final.index.isin(["Identificador_pregunta","Inversa"])]
        temp_cn = temp1[[c]].groupby(c)[[c]].count().rename(columns={c:"cnt_"+c})
        df_agg = df_agg.merge(temp_cn, left_index=True, right_index=True, how="left")
        temp_mean = pd.DataFrame(temp1[[c]].mean()).rename(columns={0:"cnt_"+c}).\
                                                    reset_index().drop(columns={"index"}).T.\
                                                    rename(columns={0:"prom"})
        df_mean = df_mean.append(temp_mean)
        temp_sum = pd.DataFrame(temp_cn.sum()).rename(columns={0:"cnt_"+c}).\
                                                   reset_index().drop(columns={"index"}).T.\
                                                   rename(columns={0:"sum"})
        df_sum = df_sum.append(temp_sum)

    df_agg = df_agg.T.merge(df_mean, left_index=True, right_index=True, how="left").T
    df_agg = df_agg.T.merge(df_sum, left_index=True, right_index=True, how="left").T.fillna(0)
    df_agg.to_csv("resultado/resultado.csv")


# In[58]:


## Ejecución general
if __name__ == "__main__":
    tabla_insumo_agg()
    


# In[ ]:





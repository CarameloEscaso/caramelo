#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
import scipy.stats as ss
import glob


# In[2]:


def clean_str(string):
    """ 
    funcion que toma unstring y lo limpia de tildes y caracteres especiales, devuelve un string en minuscula
    """
    string_new = string.replace("¿","").replace("?","").lstrip().rstrip().replace(",","").                replace("á","a").replace("é","e").replace("í","i").replace("ó","o").replace("ú","u").lower().                replace("¡","").replace("!","").replace(".","")
    return string_new

def cramers_corrected_stat(x,y):

        """ calculate Cramers V statistic for categorial-categorial association.
            uses correction from Bergsma and Wicher, 
            Journal of the Korean Statistical Society 42 (2013): 323-328
        """
        result=-1
        if len(x.value_counts())==1 :
            print("First variable is constant")
        elif len(y.value_counts())==1:
            print("Second variable is constant")
        else:   
            conf_matrix=pd.crosstab(x, y)

            if conf_matrix.shape[0]==2:
                correct=False
            else:
                correct=True

            chi2 = ss.chi2_contingency(conf_matrix, correction=correct)[0]

            n = sum(conf_matrix.sum())
            phi2 = chi2/n
            r,k = conf_matrix.shape
            phi2corr = max(0, phi2 - ((k-1)*(r-1))/(n-1))    
            rcorr = r - ((r-1)**2)/(n-1)
            kcorr = k - ((k-1)**2)/(n-1)
            result=np.sqrt(phi2corr / min( (kcorr-1), (rcorr-1)))
        return round(result,6)


# In[3]:


def tabla_insumo_agg():
    # Leer archivos de insumos y catalogo
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
            df_total = df_total.merge(i, left_index=True, right_index=True, how ="left").                            rename(columns={"puntaje":col})
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
    lista_cols = df_tabla_final.columns
    cols_final = [cols.split('_')[1] for cols in lista_cols ]
    dct = {}
    for i in range(len(lista_cols)):
        dct[lista_cols[i]] = cols_final[i]

    df_tabla_final.T.drop(columns=["Identificador_pregunta","Inversa"]).T.rename(columns=dct).to_csv("resultado/id_puntajes.csv")

#     resultado con las agregaciones y conteos
    df_agg = pd.DataFrame([1,2,3,4,5],columns={"index"}).set_index("index")
    df_mean = pd.DataFrame()
    df_sum = pd.DataFrame()
    for c in df_tabla_final.columns:
        temp1 = df_tabla_final[~df_tabla_final.index.isin(["Identificador_pregunta","Inversa"])]
        temp_cn = temp1[[c]].groupby(c)[[c]].count().rename(columns={c:"cnt_"+c})
        df_agg = df_agg.merge(temp_cn, left_index=True, right_index=True, how="left")
        temp_mean = pd.DataFrame(temp1[[c]].mean()).rename(columns={0:"cnt_"+c}).                                            reset_index().drop(columns={"index"}).T.                                                rename(columns={0:"prom"})
        df_mean = df_mean.append(temp_mean)
        temp_sum = pd.DataFrame(temp_cn.sum()).rename(columns={0:"cnt_"+c}).                                            reset_index().drop(columns={"index"}).T.                                                rename(columns={0:"sum"})
        df_sum = df_sum.append(temp_sum)

    df_agg = df_agg.T.merge(df_mean, left_index=True, right_index=True, how="left").T
    df_agg = df_agg.T.merge(df_sum, left_index=True, right_index=True, how="left").T.fillna(0)
    lista_cols = df_agg.columns
    cols_final = [cols.split('_')[2] for cols in lista_cols ]
    dct = {}
    for i in range(len(lista_cols)):
        dct[lista_cols[i]] = cols_final[i]

    df_agg.rename(columns=dct, inplace=True)
    df_agg = df_agg.T.merge(catalogo_inverso[["Identificador_pregunta","actuacion","nivel"]]
                   , left_index=True
                   , right_on=["Identificador_pregunta"]).set_index("Identificador_pregunta").T
    df_agg = df_agg.T.reset_index().set_index("actuacion")
    for i in [1,2,3,4,5,"prom","sum"]:
        df_agg[i] = df_agg.apply(lambda row: str(row[i]).replace(".",","), axis=1)
    df_agg.to_csv("resultado/resultado.csv",sep=";")
#     return df_agg


# In[10]:



def demografico():
    # Leer resultado de id_puntajes 
    file_ins = glob.glob('insumo/*.xlsx')
    file_res = glob.glob('resultado/id_puntajes.csv')
    demo = pd.read_excel(file_ins[1]).set_index("index")
    
    # categorizacion de las columnas demograficas
    for i in demo.columns:
        demo[i] = demo[i].astype("category").cat.codes

    insumo_punt = pd.read_csv(file_res[0]).set_index("index")
    # unir las tablas que contiene la relacion entre variables demograficas con las calificaciones
    df_corr = insumo_punt.merge(demo, left_index=True, right_index=True)
    # Generar el coeficiente de correlacion entre variables categoricas 
    cols_demo = demo.columns
    cols_punt = insumo_punt.columns

    df_corr_cramer = pd.DataFrame()
    for col_demo in cols_demo:
        for col_punt in cols_punt:
            temp = pd.DataFrame({'vars1':[col_demo]
                                ,"vars2":[col_punt]
                                ,"cramer_corr":[cramers_corrected_stat(df_corr[col_demo],df_corr[col_punt])]})
            df_corr_cramer = df_corr_cramer.append(temp)
    #         cramers_corrected_stat(df_corr[col_demo],df_corr[col_punt])

    df_corr_vars = df_corr_cramer.sort_values("cramer_corr",ascending=False)

    file_cat = glob.glob('catalogos/*.xlsx')
    inverso = file_cat[0]
    catalogo_inverso= pd.read_excel(inverso)
    df_corr = df_corr_vars.merge(catalogo_inverso[["Identificador_pregunta","actuacion","nivel"]]
                                 , left_on="vars2"
                                 ,right_on="Identificador_pregunta")[["vars1","vars2","cramer_corr","actuacion","nivel","Identificador_pregunta"]]
    df_corr["cramer_corr"] = df_corr.apply(lambda row: str(row["cramer_corr"]).replace(".",","), axis=1)
    df_corr.to_csv("resultado/correlacion_variables.csv",sep=";")
#     return df_corr_vars


# In[11]:


## Ejecución general
if __name__ == "__main__":
    tabla_insumo_agg()
    demografico()


# In[ ]:





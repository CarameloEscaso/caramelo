#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
import scipy.stats as ss
import glob
import os


# In[ ]:


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
def genera_case():
    file_res = glob.glob('resultado/correlacion_variables.csv')
    punt = pd.read_csv(file_res[0], sep=";")["vars1"].unique()
    not_in = ['estado civil','educacion','antigüedad en el cargo','nivel del cargo', 'hijos'
              ,'genero','area','antigüedad en la organización',"edad"]

    listadomin = [x for x in punt if x not in not_in]
    listadomay = [x.replace("t","T").replace("n","N").replace("y","Y") for x in punt if x not in not_in]
    string1 = "CASE [Parámetros].[Seleccionar pregunta] \n" # para actuacion Seleccionar Actuacion

    stringn = "WHEN '{0}' THEN [{1}] \n"

    stringfinal1 = "ELSE '' \n"
    stringfinal2 = "END"
    with open("resultado/case_preguntas.txt", 'w') as x_file:
        x_file.write(string1)
        for i in range(len(listadomin)):
            n = stringn.format(listadomin[i],listadomin[i])
            x_file.write(n)
        x_file.write(stringfinal1)
        x_file.write(stringfinal2)
        
def genera_case_actuacion():
    file_res = glob.glob('resultado/correlacion_variables.csv')
    punt = pd.read_csv(file_res[0], sep=";")["vars1"].unique()
    actuaciones = ['Conexion',
                'Equilibrio',
                'Proposito',
                'Orgullo (eNps)',
                'Pedir ayuda',
                'Contribucion',
                'Retroalimentacion',
                'Reconocimiento',
                'Liderazgo y trabajo en equipo',
                'Rituales',
                'Comunicacion',
                'Conversaciones',
                'Entorno saludable',
                'Innovacion',
                'Adaptabilidad',
                'Toma de decisiones',
                'Empoderamiento',
                'Relaciones',
                'Superar metas']

    listadomin = [x for x in actuaciones]
    listadomay = [x for x in actuaciones]
    string1 = "CASE [Parametros].[Seleccionar actuacion] \n" # para actuacion Seleccionar Actuacion

    stringn = "WHEN '{0}'  THEN avg([{1}]) \n"

    stringfinal2 = "END"
    with open("resultado/case_actuaciones.txt", 'w') as x_file:
        x_file.write(string1)
        for i in range(len(actuaciones)):
            n = stringn.format(actuaciones[i],actuaciones[i])
            x_file.write(n)
        x_file.write(stringfinal2)


# In[ ]:


def tabla_insumo_agg():
    # Leer archivos de insumos y catalogo
    file_cat = glob.glob('catalogos/*.xlsx')
    file_ins = glob.glob('insumo/*.xlsx')

    # Catalogo de respuestas

    inverso = file_cat[0] #catalogo_1.xlsx"
    puntaje = file_cat[1] #catalogo_2.xlsx"

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
        df_t[c] = df_t.apply(lambda row: row[c] if row["Inversa"] == "0" 
                                    else invers[row[c]] , axis=1)
    df_tabla_final = df_t.T


    lista_cols = df_tabla_final.columns
    cols_final = [cols.split('_')[1] for cols in lista_cols ]
    dct = {}
    i = 0
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
    i = 0
    for i in range(len(lista_cols)):
        dct[lista_cols[i]] = cols_final[i]

    df_agg.rename(columns=dct, inplace=True)
    df_agg = df_agg.T.merge(catalogo_inverso[["Identificador_pregunta","actuacion","nivel"]]
                   , left_index=True
                   , right_on=["Identificador_pregunta"]).set_index("Identificador_pregunta").T
    df_agg = df_agg.T.reset_index().set_index("actuacion")
    # agregacion por actuación
    temp = df_agg.reset_index()
    temp["prom"] = temp["prom"].astype("float")
    act_glob = temp.groupby("actuacion")[["prom"]].mean().reset_index().rename(columns={"prom":"actuacion global"})
    act_glob2 = temp.merge(act_glob, on="actuacion")[["actuacion","Identificador_pregunta","actuacion global"]]
    act_glob2.to_csv("resultado/actuacion_global.csv")
    # convertir en proporcion la cantidad de registros
    i=0
    for i in [1,2,3,4,5]:
        df_agg[i] = df_agg[i]/df_agg["sum"]
    
    # formatear columnas que se exportaran
    for cols_to_form in [1,2,3,4,5,"prom","sum"]:
        df_agg[cols_to_form] = df_agg.apply(lambda row: str(row[cols_to_form]).replace(".",","), axis=1)
    df_agg.to_csv("resultado/resultado.csv",sep=";")
#     return df_agg


# In[ ]:



def demografico():
    # Leer resultado de id_puntajes 
    file_ins = glob.glob('insumo/*.xlsx')
    file_res = glob.glob('resultado/id_puntajes.csv')
    demo = pd.read_excel(file_ins[1]).set_index("index")
    insumo_punt = pd.read_csv(file_res[0]).set_index("index")
    df_corr = insumo_punt.merge(demo, left_index=True, right_index=True)
    df_corr.to_csv('resultado/id_puntajes.csv')

    # categorizacion de las columnas demograficas
    for i in demo.columns:
        demo[i] = demo[i].astype("category").cat.codes

    # unir las tablas que contiene la relacion entre variables demograficas con las calificaciones
    df_corr = insumo_punt.merge(demo, left_index=True, right_index=True)

    # Generar el coeficiente de correlacion entre variables categoricas 
    cols_demo = demo.columns
    cols_punt = insumo_punt.columns
    cols_tot = cols_demo.append(cols_punt)
    df_corr_cramer = pd.DataFrame()
    combinaciones = []
    for col_demo in cols_tot:
        for col_punt in cols_punt:
            if col_demo != col_punt:
                if (col_punt,col_demo) not in combinaciones:
                    combinaciones.append((col_demo, col_punt))
                    temp = pd.DataFrame({'vars1':[col_demo]
                                        ,"vars2":[col_punt]
                                        ,"cramer_corr":[cramers_corrected_stat(df_corr[col_demo],df_corr[col_punt])]})
                    df_corr_cramer = df_corr_cramer.append(temp)

    df_corr_vars = df_corr_cramer.sort_values("cramer_corr",ascending=False)

    file_cat = glob.glob('catalogos/*.xlsx')
    inverso = file_cat[0]
    catalogo_inverso= pd.read_excel(inverso)
    df_corr = df_corr_vars.merge(catalogo_inverso[["Identificador_pregunta","actuacion","nivel"]]
                                 , left_on="vars2"
                                 ,right_on="Identificador_pregunta")[["vars1","vars2","cramer_corr","actuacion","nivel","Identificador_pregunta"]]
    df_corr["cramer_corr"] = df_corr.apply(lambda row: str(row["cramer_corr"]).replace(".",","), axis=1)
    df_corr.to_csv("resultado/correlacion_variables.csv",sep=";")
    # #     return df_corr_vars


# In[ ]:


def aggregado_id():
    """
    Funcion que ejecuta la agregación por actuacion  global
    """
    file_res = glob.glob('resultado/id_puntajes.csv')  
    file_act_glo = glob.glob("resultado/actuacion_global.csv")
    punt = pd.read_csv(file_res[0])#.rename(columns={"Unnamed: 0":"id"}).set_index("id")
    act_glo = pd.read_csv(file_act_glo[0])
    temp1 = punt.T.merge(act_glo, left_index=True,right_on="Identificador_pregunta", how="left")
    temp1 = temp1.set_index("Identificador_pregunta").drop(columns="Unnamed: 0").T.to_csv("resultado/id_puntajes.csv")
    os.remove('resultado/actuacion_global.csv')
    punt = pd.read_csv(file_res[0]).rename(columns={"Unnamed: 0":"id"}).set_index("id")
    temp = punt.T
    temp.drop(columns="actuacion global",inplace=True)

    actuaciones = [act for act in temp["actuacion"].dropna().unique() if act is not None]

    df_final = pd.DataFrame(index=temp.T.index)
    for actua in actuaciones:
        temp2 = pd.DataFrame(temp[temp["actuacion"]==actua].iloc[:,:-1].astype(float).mean()).rename(columns={0:actua})
        df_final = df_final.merge(temp2, left_index=True, right_index=True, how="left")
    df_final.dropna(inplace=True)
    df_final_act = punt.merge(df_final, left_index=True,right_index=True).dropna()
    to_num = ['antigüedad en el cargo',
    'antigüedad en la organización',
    "edad",
    'Conexion',
    'Equilibrio',
    'Proposito',
    'Orgullo (eNps)',
    'Pedir ayuda',
    'Contribucion',
    'Retroalimentacion',
    'Reconocimiento',
    'Liderazgo y trabajo en equipo',
    'Rituales',
    'Comunicacion',
    'Conversaciones',
    'Entorno saludable',
    'Innovacion',
    'Adaptabilidad',
    'Toma de decisiones',
    'Empoderamiento',
    'Relaciones',
    'Superar metas'
         ]

    df_final_act["edad_bkt"] = df_final_act.apply(lambda row: '18-25' if row['edad'] < 26 
                                                              else '25-35' if row['edad'] < 36 
                                                              else '35-45' if row['edad'] < 46
                                                              else '45-55' if row['edad'] < 56
                                                              else '+55' if row['edad'] >= 56
                                                              else ''
                                                  , axis=1) 
    
    df_final_act["antigüedad en el cargo bkt"] = df_final_act.apply(lambda row: '0-1' if row['antigüedad en el cargo'] < 2 
                                                                          else '1-3' if row['antigüedad en el cargo'] < 4 
                                                                          else '3-5' if row['antigüedad en el cargo'] < 6
                                                                          else '5-8' if row['antigüedad en el cargo'] < 9
                                                                          else '8-10' if row['antigüedad en el cargo'] < 11
                                                                          else '+10' if row['antigüedad en el cargo'] >= 11
                                                                          else ''
                                                                    , axis=1) 
    df_final_act["antigüedad en la organización bkt"] = df_final_act.apply(lambda row: '0-1' if row['antigüedad en la organización'] < 2 
                                                                          else '1-3' if row['antigüedad en la organización'] < 4 
                                                                          else '3-5' if row['antigüedad en la organización'] < 6
                                                                          else '5-8' if row['antigüedad en la organización'] < 9
                                                                          else '8-10' if row['antigüedad en la organización'] < 11
                                                                          else '+10' if row['antigüedad en la organización'] >= 11
                                                                          else ''
                                                                    , axis=1)
    for i in to_num:
        df_final_act[i] = df_final_act[[i]].apply(lambda row: row[i].astype("str").replace(".",","), axis=1)
    df_final_act.to_csv("resultado/id_puntajes.csv")


# In[ ]:


## Ejecución general
if __name__ == "__main__":
#     print("inicio ejecución")
    df = tabla_insumo_agg()
#     print("demograficos")
    demografico()
#     print("agregado global")
    aggregado_id()
#     print("generar case preguntas y actuaciones")
    genera_case_actuacion()
    genera_case()
#     print("finalizado exitosamente")


import pandas as pd
import numpy as np
import shutil

#Para calcular velocidad angular

Dframe_f_output=pd.read_excel("data_io.xlsx","f_and_ouput", header=None)

Frecuencia=Dframe_f_output.iloc[0, 1]
Velocidad_Angular=round(Frecuencia*2*np.pi)

#Se leen datos del excel

#Fuente voltaje

Dframe_V_fuente = pd.read_excel("data_io.xlsx", "V_fuente")

Valores_V_fuente = Dframe_V_fuente.astype(float, errors="ignore")
Dframe_V_fuente.fillna(0, inplace=True)    # Rellenar vacíos con 0.
DfWarnings_V = pd.DataFrame(Valores_V_fuente)

Desfase_V_fuente = np.array(Dframe_V_fuente.iloc[:, 3], dtype="float_")         # Angulo de desfase fuente voltaje.
V_pico_V_fuente = np.array(Dframe_V_fuente.iloc[:, 2] / np.sqrt(2))             # Voltaje pico de la fuente voltaje.
Resistencia_V_fuente = np.array(Dframe_V_fuente.iloc[:, 4])                     # Resistencia de fuente voltaje.
Inductancia_V_fuente = np.array(Dframe_V_fuente.iloc[:, 5]) * (10 ** -3)        # Inductancia de fuente voltaje.
Capacitancia_V_fuente = np.array(Dframe_V_fuente.iloc[:, 6]) * (10 ** -6)        # Capacitancia de fuente voltaje.


Nodo_V_fuente_i = np.array(Dframe_V_fuente.iloc[:, 0])                  # Nodo i fuente voltaje.
Nodo_V_fuente_j = np.full((len(Dframe_V_fuente.iloc[:, 0])), 0)         # Nodo j fuente voltaje.

#Fuente de corriente

Dframe_I_fuente = pd.read_excel("data_io.xlsx","I_fuente")
Valores_I_fuente = Dframe_I_fuente.astype(float, errors="ignore")
Dframe_I_fuente.fillna(0, inplace=True)    # Rellenar vacíos con 0.
DfWarnings_I = pd.DataFrame(Valores_I_fuente)

I_pico_I_fuente = np.array(Dframe_I_fuente.iloc[:, 2] / np.sqrt(2))     # Corriente pico de la fuente de corriente.
Desfase_I_fuente = np.array(Dframe_I_fuente.iloc[:, 3], dtype="float_") # Angulo de desfase fuente de corriente.
Resistencia_I_fuente = np.array(Dframe_I_fuente.iloc[:, 4])                     # Resistencia de la fuente de corriente.
Inductancia_I_fuente = np.array(Dframe_I_fuente.iloc[:, 5]) * (10 ** -3)        # Inductancia de la fuente de corriente.
Capacitancia_I_fuente = np.array(Dframe_I_fuente.iloc[:, 6]) * (10 ** -6)        # Capacitancia de la fuente de corriente.

Nodo_I_fuente_i = np.array(Dframe_I_fuente.iloc[:, 0])                  # Nodo i fuente de corriente.
Nodo_I_fuente_j = np.full((len(Dframe_I_fuente.iloc[:, 0])), 0)         # Nodo j fuente de corriente.

index_carga = np.concatenate(([Nodo_I_fuente_i], [Nodo_I_fuente_j]))    # Matriz de conexion de las cargas.
index_carga = np.transpose(index_carga)


#Valores de resistencia, inductancia, capacitancia

Dframe_Z = pd.read_excel("data_io.xlsx","Z")
Valores_Z = Dframe_Z.astype(float, errors="ignore")
Dframe_Z.fillna(0, inplace=True)    # Rellenar vacíos con 0.
DfWarnings_Z = pd.DataFrame(Valores_Z)


Res_Z = np.array(Dframe_Z.iloc[:, 3])                                   # Resistores.
Ind_Z = np.array(Dframe_Z.iloc[:, 4]) * (10 ** -6)                      # Inductores.
Cap_Z = np.array(Dframe_Z.iloc[:, 5]) * (10 ** -6)                      # Capacitores.

Nodo_Z_i = np.array(Dframe_Z.iloc[:, 0])                                # Bus i.
Nodo_Z_j = np.array(Dframe_Z.iloc[:, 1])                                # Bus j.

#Dataframes para guardado

Dframe_VZth = pd.read_excel("data_io.xlsx","VTH_AND_ZTH")
Dframe_Sfuente = pd.read_excel("data_io.xlsx","Sfuente")
Dframe_SZ = pd.read_excel("data_io.xlsx","S_Z")
Dframe_BalanceS = pd.read_excel("data_io.xlsx","Balance_S")

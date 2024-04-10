import numpy as np
import math


                                # -Impedancias de las fuentes de voltaje- #

def V_fuente(res_v_fuente, indc_v_fuente, cap_v_fuente, v_ang, bus_V):

    # Impedancia de V_fuente

    Impd_resist_v = res_v_fuente
    Impd_induct_v = 1j * v_ang * indc_v_fuente
    Impd_capact_v = np.zeros((len(bus_V)), dtype="complex_")
    
    for a in range(len(Impd_capact_v)):
        if cap_v_fuente[a] == 0:

            Impd_capact_v[a] = 0

        else:

            Impd_capact_v[a] = (-1j) / (v_ang * cap_v_fuente[a])

    Impedancia_V_fuente = Impd_resist_v + Impd_induct_v + Impd_capact_v
    
    # En caso de no haber una impedancia se añade una inductancia de 10^-6

    for i in range(len(Impedancia_V_fuente)):

        if Impedancia_V_fuente[i] == 0:

            Impedancia_V_fuente[i] = 0 + (0.000001)*1j

    
    
    return Impedancia_V_fuente, Impd_resist_v, Impd_induct_v, Impd_capact_v

                                # -Impedancias de las fuentes de corriente- #

def I_fuente(res_i_fuente, indc_i_fuente, cap_i_fuente, v_ang, bus_I):

    Impd_resist_i = res_i_fuente
    Impd_induct_i = 1j * v_ang * indc_i_fuente
    Impd_capact_i = np.zeros((len(bus_I)), dtype="complex_")
    
    for b in range(len(Impd_capact_i)):
                
        if cap_i_fuente[b] == 0:

            Impd_capact_i[b] = 0
            
        else:

            Impd_capact_i[b] = (-1j) / (v_ang * cap_i_fuente[b])

    Impedancia_I_fuente = Impd_resist_i + Impd_induct_i + Impd_capact_i
    
    # En caso de no haber una impedancia se añade una resistencia de 10^-6

    for i in range(len(Impedancia_I_fuente)):

        if Impedancia_I_fuente[i] == 0:

            Impedancia_I_fuente[i] = 0 + (0.000001)*1j
    
    return Impedancia_I_fuente, Impd_resist_i, Impd_induct_i, Impd_capact_i

                    # -Impedancias de los elementos resistivos, capacitivos e inductivos- #

def Z(Resis_Z, Indc_Z, Cap_Z, V_ang, Bus_Z):

    Imp_Resis_Z = Resis_Z
    Imp_Ind_Z = 1j * V_ang * Indc_Z
    Imp_Cap_Z = np.zeros((len(Bus_Z)), dtype="complex_")

    for b in range(len(Bus_Z)):
        
        if Cap_Z[b] == 0:

            Imp_Cap_Z[b] = 0

        else:

            Imp_Cap_Z[b] = (-1j) / (V_ang * Cap_Z[b])

    Impedancia_Z = Imp_Resis_Z + Imp_Ind_Z + Imp_Cap_Z
    
    return Impedancia_Z, Imp_Resis_Z, Imp_Ind_Z, Imp_Cap_Z

                                            # -Vector de corrientes- #
                                            
def Matriz_Corrientes(Voltaje, Corriente, Desface_v, Desface_I, Impedancia_v, Nro_nodos, Nodo_v_i, Nodo_I_i):

    Vec_Corrientes = np.zeros((Nro_nodos,1), dtype="complex_")

    for i in range(len(Nodo_v_i)):

        indicefuentevoltaje = Nodo_v_i[i]-1
        Desface_v[i]= (Desface_v[i] * math.pi) / 180
        Vec_Corrientes[indicefuentevoltaje] = (Voltaje[i] * (math.cos(Desface_v[i]) + 1j*math.sin(Desface_v[i])) / Impedancia_v[i]) 

    for i in range(len(Nodo_I_i)):

        indicefuentecorriente = Nodo_I_i[i]-1
        Desface_I[i]= (Desface_I[i] * math.pi) / 180
        Vec_Corrientes[indicefuentecorriente] += Corriente[i] * (math.cos(Desface_I[i]) + 1j*math.sin(Desface_I[i]))
        
    #print(Vec_Corrientes)       
    return Vec_Corrientes

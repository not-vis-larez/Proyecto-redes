import numpy as np
import cmath
import math

                                                # -Matríz Ybus- #
def Matriz_Y_Bus(V_fuentes, I_fuentes, Zs, Nro_Nodos, Nro_N_i, Nro_N_j):

    Datos_Entrada = np.concatenate([V_fuentes, Zs], axis=0)
    
    Ybus_Salida = np.zeros((Nro_Nodos, Nro_Nodos), dtype="complex_")

    filas, columnas = Datos_Entrada.shape

    # Admitancias fuera de la diagonal.

    for k in range(filas):

        i = int(Datos_Entrada[k,0].real-1)
        j = int(Datos_Entrada[k,1].real-1)
        
        if i == -1 or j == -1:

            continue

        else:
                        
            Ybus_Salida[i,j] = Ybus_Salida[i,j] + (-1) / (Datos_Entrada[k,2])
            Ybus_Salida[j,i] = Ybus_Salida[i,j]

    # Admitancias de la diagonal.

    Matriz_Auxiliar = Ybus_Salida.sum(axis=1)
    Matriz_Auxiliar = np.diag(Matriz_Auxiliar)
    
    Ybus_Salida = Ybus_Salida - Matriz_Auxiliar
    
    for k in range(filas):

        i = int(Datos_Entrada[k, 0].real - 1)
        j = int(Datos_Entrada[k, 1].real - 1)
        
        if i == -1 or j == -1:

            if i == -1:
                    
                Ybus_Salida[j,j] = Ybus_Salida[j, j] + (1 / Datos_Entrada[k, 2])

            elif j == -1:
                
                Ybus_Salida[i,i] = Ybus_Salida[i, i] + (1 / Datos_Entrada[k, 2])
                
    #Ybus_Salida = np.round(Ybus_Salida,4)
    #print(Ybus_Salida)
    print(f"Matriz de admitancias:\n\n{Ybus_Salida}")
    return Ybus_Salida

                                                        # -Zth- #
def Zth(Ybus):

    Zbus = np.linalg.inv(Ybus)

    Zth = np.diag(Zbus)

    return Zth, Zbus

                                                        # -Vth- #
def Vth(Zbus, corrientes, num_barra):

    V_thevenin = np.dot(Zbus, corrientes)

    Matriz_Vth_Polar = np.zeros((num_barra, 2))
    
    # Forma polar.
        
    for i in range(num_barra):

        Matriz_Vth_Polar[i,0],Matriz_Vth_Polar[i,1] = cmath.polar(V_thevenin[i])
        
    # Radianes a grados.

    for i in range(num_barra):

        Matriz_Vth_Polar[i,1] = Matriz_Vth_Polar[i,1]*180/math.pi
    print(f"Voltajes en nodos (Polar)\n\n {Matriz_Vth_Polar}")
    return Matriz_Vth_Polar, V_thevenin

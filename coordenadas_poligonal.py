from inicializador import Inicializador
from grupo_seleccion import Seleccion
import numpy as np
import pandas as pd

class Coordenadas(Inicializador, Seleccion ):
    def __init__(self):
        Inicializador.__init__(self)
        Seleccion.__init__(self)
        self.coordenadas_list = []
        self.columnas = {"OBJETO" : [],
                         "ESTE" : [],
                         "NORTE" : [],
                         "ALTITUD" : []}

    def coordenadas(self)-> list:
        self.lineas = ["AcDbPolyline",
                       "AcDb3dPolyline"]
        for elemento in self.grupo:
            if elemento.EntityName in self.lineas:
                self.coordenadas_list.append(elemento.Coordinates)
        return self.coordenadas_list
    
    def formato(self)-> pd.DataFrame:
        j = 1
        for coordenadas, name in zip(self.coordenadas_list, self.grupo):
            ent = name.EntityName
            i = 0
            if ent in self.lineas:
                puntos =  int((len(coordenadas)-2)/2) if ent == self.lineas[0] else int((len(coordenadas)-3)/3)
                for _ in range(puntos):
                    altitud = 0 if ent == self.lineas[0] else coordenadas[i+2]
                    self.columnas["OBJETO"].append(j)
                    self.columnas["ESTE"].append(coordenadas[i])
                    self.columnas["NORTE"].append(coordenadas[i+1])
                    self.columnas["ALTITUD"].append(altitud)
                    i+= 2 if ent == self.lineas[0] else 3
                j+=1
        self.tabla = pd.DataFrame(self.columnas)    
        return self.tabla

    def exportar_coordenadas(self):
        self.tabla.to_excel("Z:/Coordenadas/coordenadas.xlsx", index=False)
    
if __name__ == '__main__':
    valor = Coordenadas()
    valor.coordenadas()
    valor.formato()
    valor.exportar_coordenadas()


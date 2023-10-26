import pandas as pd
import numpy as np

class Planilha:
    def __init__(self, path:str, require_cols:list):        
        self.data = pd.read_excel(path, usecols=require_cols, 
                             header=None, skiprows=24,
                             names=["Item","Nº Tombo","Descrição do Material"]).dropna()
        self.data=self.data.astype({'Nº Tombo':'int'}, errors='ignore')

    def getTombo(self):                        
        return self.data['Nº Tombo'].values.tolist()
    
    def getSaida(self, other:object):        
        return [s for s in self.getTombo() if s not in other.getTombo()]
    
    def getEntrada(self, other:object):        
        return [o for o in other.getTombo() if o not in self.getTombo()]
    
    def compara(self, other:object):
        entrada = self.getEntrada(other)
        saida = self.getSaida(other)
        le = len(entrada)
        ls = len(saida)
        if(le>ls):
            for i in range(le-ls):
                saida.append('-')
        elif(le<ls):
            for i in range(ls-le):
                entrada.append('-')
        return np.column_stack((saida, entrada)).tolist()

import pandas as pd
import os


class DadosPlanilha:
    def __init__(self,numero_linha, dados_linha): 
        self.numero_linha = numero_linha #Número de linhas no Excel
        self.dados_linha = dados_linha  #Dicionario com {coluna: valor}
        
        
        
linha_exemplo = DadosPlanilha(
    numero_linha=2,
    dados_linha={"Nome": "João", "Idade": 25, "Cidade": "São Paulo"}
    )
print(f"Linha: {linha_exemplo.numero_linha}")
print(f"Dados: {linha_exemplo.dados_linha}")
print(f"Nome: {linha_exemplo.dados_linha['Nome']}")
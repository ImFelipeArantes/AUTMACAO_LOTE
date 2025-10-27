import pandas as pd
import numpy as np
from unidecode import unidecode

facilidades = pd.read_excel('./arquivos/facilidades.xlsx').FACILIDADE.tolist()

class tratamentoResumosoe():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.__resumo_soe = file.readlines()
        file.close()


    def trata_resumosoe(self):
        for i in range(len(self.__resumo_soe)):
            self.__resumo_soe[i] = self.__resumo_soe[i].removesuffix('\n')
            self.__resumo_soe[i] = self.__resumo_soe[i].split("\t")

        acum = 0
        while acum < len(self.__resumo_soe[0]):
            self.__resumo_soe[0][acum] = self.__resumo_soe[0][acum].replace(" ","_").upper()
            if self.__resumo_soe[0][acum] == '%DISPONIBILIDADE':
                self.__resumo_soe[0][acum] = self.__resumo_soe[0][acum].replace('%','')
            acum += 1
        acum_resumo = 0
        acum_facilidade = 0
        while acum_resumo < len(self.__resumo_soe[0]):
            if self.__resumo_soe[0][acum_resumo] == facilidades[acum_facilidade]:
                self.__resumo_soe[0][acum_resumo + 1] = f'{self.__resumo_soe[0][acum_resumo]}_{unidecode(self.__resumo_soe[0][acum_resumo + 1])}'
                self.__resumo_soe[0][acum_resumo + 2] = f'{self.__resumo_soe[0][acum_resumo]}_{unidecode(self.__resumo_soe[0][acum_resumo + 2])}'
                self.__resumo_soe[0][acum_resumo + 3] = f'{self.__resumo_soe[0][acum_resumo]}_{unidecode(self.__resumo_soe[0][acum_resumo + 3])}'
                if acum_facilidade < len(facilidades)-1:
                    acum_facilidade += 1
            acum_resumo += 1

        df = pd.DataFrame(np.array(self.__resumo_soe[1:]), columns=self.__resumo_soe[0])
        df[['ID', 'SEV']] = df[['ID', 'SEV']].apply(pd.to_numeric)

        return df
    
class tratamentoResultado():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.__resultado = file.readlines()
        file.close()
    
    def trata_resultado(self):
        
        for i in range(len(self.__resultado)):
            self.__resultado[i] = self.__resultado[i].removesuffix('\n')
            self.__resultado[i] = self.__resultado[i].split("\t")

        df = pd.DataFrame(np.array(self.__resultado[1:]),columns=self.__resultado[0])
        df[['ID', 'SEV']] = df[['ID', 'SEV']].apply(pd.to_numeric)

        return df
    
class tratamentoRestricao():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.restricao = file.readlines()
        file.close()

    def trata_restricao(self):
        self.restricao.pop(0)

        for i in range(len(self.restricao)):
            self.restricao[i] = self.restricao[i].removesuffix('\n')
            self.restricao[i] = self.restricao[i].split("\t")

        self.restricao[0].remove('Camada')

        cols_restricao = self.restricao[0]
        self.restricao.pop(0)

        for i in cols_restricao:
            ind = cols_restricao.index(i)
            cols_restricao[ind] = unidecode(cols_restricao[ind]).replace(" ","_").upper()
        
        i = 0
        tam = len(self.restricao)
        while i < tam:
            if i != (len(self.restricao) - 1):
                sev_atual = self.restricao[i][1]
                tam_restricao_atual = len(self.restricao[i])
                if self.restricao[i+1][1] == sev_atual:
                    if len(self.restricao[i+1]) > tam_restricao_atual:
                        self.restricao.pop(i)
                        tam -= 1
                    else:
                        i += 1
                else:
                    i += 1
            else: i += 1


        i = 0
        while i < len(self.restricao):
            x = 1
            while x < len(self.restricao[i]):
                if x == 1:
                    sev = self.restricao[i][x]
                    id = self.restricao[i][x-1]
                    x += 15
                else:
                    self.restricao[i].insert(x, sev)
                    self.restricao[i].insert(x, id)
                    x += 16
            i += 1
            
        i = 0
        new_data = []
        while i < len(self.restricao):
            x = 0
            while x < len(self.restricao[i]):
                new_data.append(self.restricao[i][x:x+16])
                x += 16
            i += 1
        self.restricao = new_data

        i = 0
        while i < len(self.restricao):
            if len(self.restricao[i]) != 16:
                self.restricao.pop(i)
            else:
                i += 1


        df = pd.DataFrame(np.array(self.restricao[0:]),columns=cols_restricao)
        df[['ID', 'SEV']] = df[['ID', 'SEV']].apply(pd.to_numeric)

        return df
    

class tratamentoNuvens():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.nuvens = file.readlines()
        file.close()

    def trata_nuvens(self):
        self.nuvens.pop(0)
        for i in range(len(self.nuvens)):
            self.nuvens[i] = self.nuvens[i].removesuffix('\n')
            self.nuvens[i] = self.nuvens[i].split("\t")
            

        self.nuvens[0].remove('Camada')

        data = self.nuvens[1:]


        i = 0
        tam = len(data)
        while i < tam:
            if i != (len(data) - 1):
                sev_atual = data[i][1]
                tam_nuvem_atual = len(data[i])
                if data[i+1][1] == sev_atual:
                    if len(data[i+1]) > tam_nuvem_atual:
                        data.pop(i)
                        tam -= 1
                    else:
                        i += 1
                else:
                    i += 1
            else: i += 1

        i = 0
        while i < len(data):
            tecnologia_pos = 0
            while tecnologia_pos < len(data[i]):
                if ('DOCSIS 3.0' in data[i][tecnologia_pos]) | ('DOCSIS 3.1' in data[i][tecnologia_pos]) | ('XDSL' in data[i][tecnologia_pos]):
                    for x in range(11):
                        data[i].insert(tecnologia_pos + 9,' ')
                    tecnologia_pos += 1
                tecnologia_pos += 1


            i += 1

        i = 0
        while i < len(data):
            x = 1
            while x < len(data[i]):
                if x == 1:
                    sev = data[i][x]
                    id = data[i][x-1]
                    x += 35
                else:
                    data[i].insert(x, sev)
                    data[i].insert(x, id)
                    x += 36
            i += 1
            
        i = 0
        new_data = []
        while i < len(data):
            x = 0
            while x < len(data[i]):
                new_data.append(data[i][x:x+36])
                x += 36
            i += 1
        data = new_data

        colunas = ['ID','SEV','OBJECTID','ESTACAO_ENTREGA','UF','SIGLA_LOC','NOME_NUVEM','REDE','TECNOLOGIA','MEIO_TRANSMISSAO','SITUACAO','ALT_NUVEM','PADRAO_PON','PROPRIETARIO','OBSERVACAO','ORIGEM','DATA_ATUALIZACAO','DATA_PREVISAO','VEL_MAX_VIABILIDADE','VEL_MAX_SEV_AUTOMATICA','MOTIVO','OPERADORA','NUMERO_OPERACIONAL','SIGLA_ESTACAO_CLARO','CODIGO_DESCARGA_CSL','SIGLA_ESTACAO_RESID','TIPO_INFRA','TIPO','STATUS','DONO','ROTEADORES_QTD','TX','FABRICANTE_OLT','ABRANGENCIA','CONCENTRADOR_OLT','POSICAO']

        nuvens_df = pd.DataFrame(data, columns=colunas)

        nuvens_df['SEV'] = pd.to_numeric(nuvens_df['SEV'],downcast='signed', errors='coerce')
        nuvens_df = nuvens_df.drop(columns=['ID','POSICAO'])
        nuvens_df = nuvens_df.drop_duplicates()

        return nuvens_df
    
class tratamentoNuvensTerceiros():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.nuvens_terceiros = file.readlines()
        file.close()

    def trata_nuvens(self):
        self.nuvens_terceiros.pop(0)
        for i in range(len(self.nuvens_terceiros)):
            self.nuvens_terceiros[i] = self.nuvens_terceiros[i].removesuffix('\n')
            self.nuvens_terceiros[i] = self.nuvens_terceiros[i].split("\t")

        self.nuvens_terceiros[0].remove('Camada')

        data = self.nuvens_terceiros[1:]


        i = 0
        tam = len(data)
        while i < tam:
            if i != (len(data) - 1):
                sev_atual = data[i][1]
                tam_nuvem_atual = len(data[i])
                if data[i+1][1] == sev_atual:
                    if len(data[i+1]) > tam_nuvem_atual:
                        data.pop(i)
                        tam -= 1
                    else:
                        i += 1
                else:
                    i += 1
            else: i += 1

        i = 0
        while i < len(data):
            x = 1
            while x < len(data[i]):
                if x == 1:
                    sev = data[i][x]
                    id = data[i][x-1]
                    x += 20
                else:
                    data[i].insert(x, sev)
                    data[i].insert(x, id)
                    x += 21
            i += 1
            
        i = 0
        new_data = []
        while i < len(data):
            x = 0
            while x < len(data[i]):
                new_data.append(data[i][x:x+21])
                x += 21
            i += 1
        data = new_data

        nuvens_terceiros_df = pd.DataFrame(data, columns=self.nuvens_terceiros[0])

        nuvens_terceiros_df['SEV'] = pd.to_numeric(nuvens_terceiros_df['SEV'],downcast='signed', errors='coerce')
        nuvens_terceiros_df = nuvens_terceiros_df.drop(columns=['ID','POSICAO'])
        nuvens_terceiros_df = nuvens_terceiros_df.drop_duplicates()

        return nuvens_terceiros_df
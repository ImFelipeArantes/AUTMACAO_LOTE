import pandas as pd
from unidecode import unidecode
import warnings
warnings.filterwarnings('ignore')

class extracaoTeia:

    def __init__(self,df):
        self.__df = df


    def __tratar_end(self,tipo):

        match tipo:

            case "A":
                tipo = "ACESSO"
                
            case "AC":
                tipo = "ACESSO"
            
            case "ACA":
                tipo = "ACAMPAMENTO"
            
            case "ACL":
                tipo = "ACESSO LOCAL"
            
            case "AD":
                tipo = "ADRO"
            
            case "AE":
                tipo = "AREA ESPECIAL"
            
            case "AER":
                tipo = "AEROPORTO"
            
            case "AL":
                tipo = "ALAMEDA"
                
            case "AMD":
                tipo = "AVENIDA MARGINAL DIREITA"
                
            case "AME":
                tipo = "AVENIDA MARGINAL ESQUERDA"
                
            case "AN":
                tipo = "ANEL VIARIO"
                
            case "ANT":
                tipo = "ANTIGA ESTRADA"
                
            case "ART":
                tipo = "ARTERIA"
                
            case "ATL":
                tipo = "ATALHO"
                
            case "A V":
                tipo = "AREA VERDE"
            
            case "AV":
                tipo = "AVENIDA"
                
            case "AVC":
                tipo = "AVENIDA CONTORNO"
            
            case "AVM":
                tipo = "AVENIDA MARGINAL"
            
            case "AVV":
                tipo = "AVENIDA VELHA"
            
            case "BAL":
                tipo = "BALNEARIO"
            
            case "BC":
                tipo = "BECO"
            
            case "BCO":
                tipo = "BURACO"
            
            case "BEL":
                tipo = "BELVEDERE"
            
            case "BL":
                tipo = "BLOCO"
            
            case "BLO":
                tipo = "BALAO"
            
            case "BLS":
                tipo = "BLOCOS"
            
            case "BLV":
                tipo = "BULEVAR"
            
            case "BSQ":
                tipo = "BOSQUE"
            
            case "BVD":
                tipo = "BOULEVARD"
            
            case "BX":
                tipo = "BAIXA"
            
            case "C":
                tipo = "CAIS"
            
            case "CAL":
                tipo = "CALCADA"
            
            case "CAM":
                tipo = "CAMINHO"
            
            case "CAN":
                tipo = "CANAL"
            
            case "CH":
                tipo = "CHACARA"
            
            case "CHA":
                tipo = "CHAPADAO"
            
            case "CIC":
                tipo = "CICLOVIA"
            
            case "CIR":
                tipo = "CIRCULAR"
            
            case "CJ":
                tipo = "CONJUNTO"
            
            case "CJM":
                tipo = "CONJUNTO MUTIRAO"
            
            case "CMP":
                tipo = "COMPLEXO VIARIO"
            
            case "COL":
                tipo = "COLONIA"
            
            case "COM":
                tipo = "COMUNIDADE"
            
            case "CON":
                tipo = "CONDOMINIO"
            
            case "COR":
                tipo = "CORREDOR"
            
            case "CPO":
                tipo = "CAMPO"
            
            case "CRG":
                tipo = "CORREGO"
            
            case "CTN":
                tipo = "CONTORNO"
            
            case "DSC":
                tipo = "DESCIDA"
            
            case "DSV":
                tipo = "DESVIO"
            
            case "DT":
                tipo = "DISTRITO"
            
            case "EB":
                tipo = "ENTRO BLOCO"
            
            case "EIM":
                tipo = "ESTRADA INTERMUNICIPAL"
            
            case "ENS":
                tipo = "ENSEADA"
            
            case "ENT":
                tipo = "ENTRADA PARTICULAR"
            
            case "EQ":
                tipo = "ENTRE QUADRA"
            
            case "ESC":
                tipo = "ESCADA"
            
            case "ESD":
                tipo = "ESCADARIA"
            
            case "ESE":
                tipo = "ESTRADA ESTADUAL"
            
            case "ESI":
                tipo = "ESTRADA VICINAL"
            
            case "ESL":
                tipo = "ESTRADA DE LIGACAO"
            
            case "ESM":
                tipo = "ESTRADA MUNICIPAL"
            
            case "ESP":
                tipo = "ESPLANADA"
            
            case "ESS":
                tipo = "ESTRADA DA SERVIDAO"
            
            case "EST":
                tipo = "ESTRADA"
            
            case "ESV":
                tipo = "ESTRADA VELHA"
            
            case "ETA":
                tipo = "ESTRADA ANTIGA"
            
            case "ETC":
                tipo = "ESTACAO"
            
            case "ETD":
                tipo = "ESTADIO"
            
            case "ETN":
                tipo = "ESTANCIA"
            
            case "ETP":
                tipo = "ESTRADA PARTICULAR"
            
            case "ETT":
                tipo = "ESTACIONAMENTO"
                
            case "EVA":
                tipo = "EVANGELICA"
            
            case "EVD":
                tipo = "ELEVADA"
            
            case "EX":
                tipo = "EIXO INDUSTRIAL"
            
            case "FAV":
                tipo = "FAVELA"
            
            case "FAZ":
                tipo = "FAZENDA"
            
            case "FER":
                tipo = "FERROVIA"
            
            case "FNT":
                tipo = "FONTE"
            
            case "FRA":
                tipo = "FEIRA"
            
            case "FTE":
                tipo = "FORTE"
            
            case "GAL":
                tipo = "GALERIA"
            
            case "GJA":
                tipo = "GRANJA"
            
            case "HAB":
                tipo = "NUCLEO HABITACIONAL"
            
            case "IA":
                tipo = "ILHA"
            
            case "IND":
                tipo = "INDETERMINADO"
            
            case "IOA":
                tipo = "ILHOTA"
            
            case "JD":
                tipo = "JARDIM"
            
            case "JDE":
                tipo = "JARDINETE"
            
            case "LD":
                tipo = "LADEIRA"
            
            case "LGA":
                tipo = "LAGOA"
            
            case "LGO":
                tipo = "LAGO"
            
            case "LOT":
                tipo = "LOTEAMENTO"
            
            case "LRG":
                tipo = "LARGO"
            
            case "LT":
                tipo = "LOTE"
            
            case "MER":
                tipo = "MERCADO"
            
            case "MNA":
                tipo = "MARINA"
            
            case "MOD":
                tipo = "MODULO"
                
            case "MRG":
                tipo = "PROJECAO"
                
            case "MRO":
                tipo = "MORRO"
            
            case "MTE":
                tipo = "MONTE"
                
            case "NUC":
                tipo = "NUCLEO"
            
            case "NUR":
                tipo = "NUCLEO RURAL"
                
            case "OUT":
                tipo = "OUTEIRO"
                
            case "PAR":
                tipo = "PARALELA"
                
            case "PAS":
                tipo = "PASSEIO"
                
            case "PAT":
                tipo = "PATIO"
                
            case "PC":
                tipo = "PRACA"
            
            case "PCE":
                tipo = "PRACA DE ESPORTES"
            
            case "PDA":
                tipo = "PARADA"
            
            case "PDO":
                tipo = "PARADOURO"
            
            case "PNT":
                tipo = "PONTA"
            
            case "PR":
                tipo = "PRAIA"
            
            case "PRL":
                tipo = "PROLONGAMENTO"
            
            case "PRM":
                tipo = "PARQUE MUNICIPAL"
            
            case "PRQ":
                tipo = "PARQUE"
            
            case "PRR":
                tipo = "PARQUE RESIDENCIAL"
            
            case "PSA":
                tipo = "PASSARELA"
            
            case "PSG":
                tipo = "PASSAGEM"
            
            case "PSP":
                tipo = "PASSAGEM DE PEDESTRE"
            
            case "PSS":
                tipo = "PASSAGEM SUBTERRANEA"
            
            case "PTE":
                tipo = "PONTE"
            
            case "PTO":
                tipo = "PORTO"
            
            case "Q":
                tipo = "QUADRA"
            
            case "QTA":
                tipo = "QUINTA"
            
            case "QTS":
                tipo = "QUINTAS"
            
            case "R":
                tipo = "RUA"
            
            case "R I":
                tipo = "RUA INTEGRACAO"
            
            case "R L":
                tipo = "RUA DE LIGACAO"
            
            case "R P":
                tipo = "RUA PARTICULAR"
            
            case "R V":
                tipo = "RUA VELHA"
            
            case "RAM":
                tipo = "RAMAL"
                
            case "RCR":
                tipo = "RECREIO"
            
            case "REC":
                tipo = "RECANTO"
            
            case "RER":
                tipo = "RETIRO"
            
            case "RES":
                tipo = "RESIDENCIAL"
            
            case "RET":
                tipo = "RETA"
            
            case "RLA":
                tipo = "RUELA"
            
            case "RMP":
                tipo = "RAMPA"
            
            case "ROA":
                tipo = "RODO ANEL"
            
            case "ROD":
                tipo = "RODOVIA"
            
            case "ROT":
                tipo = "ROTULA"
            
            case "RPE":
                tipo = "RUA DE PEDESTRE"
            
            case "RPR":
                tipo = "MARGEM"
            
            case "RTN":
                tipo = "RETORNO"
            
            case "RTT":
                tipo = "ROTATORIA"
            
            case "SEG":
                tipo = "SEGUNDA AVENIDA"
            
            case "SIT":
                tipo = "SITIO"
            
            case "SRV":
                tipo = "SERVIDAO"
                
            case "ST":
                tipo = "SETOR"
            
            case "SUB":
                tipo = "SUBIDA"
            
            case "TCH":
                tipo = "TRINCHEIRA"
            
            case "TER":
                tipo = "TERMINAL"
            
            case "TR":
                tipo = "TRECHO"
            
            case "TRV":
                tipo = "TREVO"
            
            case "TUN":
                tipo = "TUNEL"
            
            case "TV":
                tipo = "TRAVESSA"
            
            case "TVP":
                tipo = "TRAVESSA PARTICULAR"
            
            case "TVV":
                tipo = "TRAVESSA VELHA"
            
            case "UNI":
                tipo = "UNIDADE"
            
            case "V":
                tipo = "VIA"
            
            case "V C":
                tipo = "VIA COLETORA"
            
            case "V L":
                tipo = "VIA LOCAL"
            
            case "VAC":
                tipo = "VIA DE ACESSO"
            
            case "VAL":
                tipo = "VALA"
            
            case "VCO":
                tipo = "VIA COSTEIRA"
            
            case "VD":
                tipo = "VIADUTO"
            
            case "V-E":
                tipo = "VIA EXPRESSA"
            
            case "VER":
                tipo = "VEREDA"
            
            case "VEV":
                tipo = "VIA ELEVADO"
            
            case "VL":
                tipo = "VILA"
            
            case "VLA":
                tipo = "VIELA"
            
            case "VLE":
                tipo = "VALE"
            
            case "VLT":
                tipo = "VIA LITORANEA"
            
            case "VPE":
                tipo = "VIA DE PEDESTRE"
            
            case "VRT":
                tipo = "VARIANTE"
            
            case "ZIG":
                tipo = "ZIGUE-ZAGUE"
            
            case _:
                tipo = "-"
                
        return tipo

    def __remover_sevs(self,df_modelo):
        df_sevs_retiradas = pd.DataFrame(columns=['SEV','MOTIVO'])

        for i, v in df_modelo.iterrows():
            if v['Serviço/Produto'] == 'LAN - LAN TO LAN':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'LAN TO LAN'
                df_modelo = df_modelo.drop(index=i)
            elif v['Serviço/Produto'] == 'EIN - E-ACCESS':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'E-ACCESS'
                df_modelo = df_modelo.drop(index=i)
            elif v['Serviço/Produto'] == 'LAN - LAN EPL':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'LAN - LAN EPL'
                df_modelo = df_modelo.drop(index=i)
            elif v['Serviço/Produto'] == 'LAN - LAN EPL MEF':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'LAN - LAN EPL MEF'
                df_modelo = df_modelo.drop(index=i)
            elif v['Serviço/Produto'] == 'DTN - PRIMELINK(EX.MEGADATA)':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'PRIMELINK'
                df_modelo = df_modelo.drop(index=i)

            elif (v['CAIXA'] == 'REANÁLISE DE SEV CONTESTAÇÃO') | (v['CAIXA'] == 'ANALISE_RADIO'):
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'REANALISE DE CONTESTACAO'
                df_modelo = df_modelo.drop(index=i)

            elif (v.Velocidade[-4:] == 'Gbps') & (int(v.Velocidade[:-4]) > 1):
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'ALTAS VELOCIDADES'
                df_modelo = df_modelo.drop(index=i)

            elif v.ACAO == 'Upgrade':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = 'UPGRADE'
                df_modelo = df_modelo.drop(index=i)
            ## TODO
            # CRIAR TRATAMENTO PARA SEVS DE ACESSO DISTINTO, PEGANDO A SEV INFORMADA NO CAMPO CODGIO DO PROJETO E VERIFICAR NO BANCO SE TEM ALGUMA SEV COM ESSE CODIGO,
            # E REALIZAR A VIABILIDADE COM OUTRO ACESSO
            elif v.PROJETO != 'PORTFOLIO':
                df_sevs_retiradas.at[i,'SEV'] = v.Sequencial
                df_sevs_retiradas.at[i,'MOTIVO'] = v.PROJETO
                df_modelo = df_modelo.drop(index=i)
        
        df_modelo = df_modelo.drop(columns=['CAIXA','ACAO','PROJETO'])

        return df_modelo.reset_index().drop(columns='index'), df_sevs_retiradas.reset_index().drop(columns='index')
    
        
        

    def tratar_modelo_gaia(self,removed_sevs='N'):
        cols_modelo = ['Sequencial', 'Cliente', 'Tipo', 'Logradouro', 'Numero', 'Complemento',
            'Bairro', 'Cidade', 'UF', 'CEP', 'Serviço/Produto', 'Velocidade',
            'Qtd.Circuitos', 'Necessário Contingência', 'Latitude', 'Longitude',
            'Observação', 'Distância Abordado', 'Distância Cabo',
            'Distância Infraestrutura', 'Cliente Primesys',
            'Tipo Calculo Distância', 'Backup3G', 'Conta Corrente',
            'Designação do Serviço', 'Migração PABX', 'Segmento Mercado',
            '%Disponibilidade Rede Desejável']
        
        cols_format = ['Cliente', 'Tipo', 'Logradouro', 'Complemento',
        'Bairro', 'Cidade', 'UF', 'CEP','Necessário Contingência',
        'Latitude', 'Longitude','Observação', 
        'Distância Infraestrutura', 'Cliente Primesys',
        'Tipo Calculo Distância', 'Backup3G', 'Conta Corrente',
        'Designação do Serviço', 'Migração PABX', 'Segmento Mercado',
        '%Disponibilidade Rede Desejável']
            
        df_modelo = pd.DataFrame(columns=cols_modelo)
        

        df_modelo[['CAIXA','ACAO','PROJETO','Sequencial','Cliente','Tipo','Logradouro','Numero','Complemento','Bairro','Cidade','UF','CEP',
            'Serviço/Produto','Velocidade','Qtd.Circuitos','Latitude','Longitude']] = self.__df[['CAIXA','ACAO','PROJETO','SEV',
            'CLIENTE','TIPO_LOGRADOURO','NOME_DO_LOGRADOURO','NUMERO','COMPLEMENTO','BAIRRO','CIDADE','UF','CEP',
            'SERVICO','VELOCIDADE_SERV','QTDE_CIRCUITOS','LATITUDE','LONGITUDE']]
            
        df_modelo.Cliente = 'CLIENTE'
        df_modelo['Necessário Contingência'] = 'N'
        df_modelo['Distância Abordado'] = 50
        df_modelo['Distância Cabo'] = 50
        df_modelo['Migração PABX'] = 'N'
        df_modelo['Segmento Mercado'] = 'CORPORATIVO'
        df_modelo['%Disponibilidade Rede Desejável'] = '99,650%'
        df_modelo['Numero'] = pd.to_numeric(df_modelo['Numero'],downcast='signed', errors='coerce')
        df_modelo.Numero = df_modelo.Numero.fillna(0)
        df_modelo['Numero'] = df_modelo.Numero.astype(int)
        df_modelo = df_modelo.fillna('')
        
        for i, v in df_modelo.iterrows():
            for c in cols_format:
                df_modelo.at[i,c] = unidecode(str(v[c])).upper()

        if removed_sevs == 'S':
            df_modelo, df_sevs_retirdas = self.__remover_sevs(df_modelo)

        for i, v in df_modelo.iterrows():
            df_modelo.at[i,'Tipo'] = self.__tratar_end(v.Tipo)
            df_modelo.at[i,'Velocidade'] = f'{v.Velocidade[:-4]} {v.Velocidade[-4:]}'
            
        
        if removed_sevs == 'N':
            df_modelo.to_excel('atendimento_gaia.xlsx',index=False)
            return 0
            
        elif removed_sevs == 'S':
            df_modelo.to_excel('atendimento_gaia.xlsx',index=False)
            df_sevs_retirdas.to_excel('sevs_removidas.xlsx',index=False)
            return df_sevs_retirdas.SEV.values
        
       


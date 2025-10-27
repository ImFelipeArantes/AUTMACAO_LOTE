from extracao_teia import extracaoTeia
from tratamento_gaia import tratamentoNuvens, tratamentoRestricao, tratamentoResultado, tratamentoResumosoe
import pandas as pd
import warnings
from PIL import Image
import customtkinter as ctk
warnings.filterwarnings('ignore')
import time
from unidecode import unidecode
from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.support.ui import Select
import chromedriver_autoinstaller
from sqlalchemy import create_engine
service = Service(chromedriver_autoinstaller.install())

pd.set_option('display.max_columns', None)
pd.set_option('future.no_silent_downcasting', True)

capacity = pd.read_excel('./arquivos/capacity.xlsx')
tecnologia_capacity = pd.read_excel('./arquivos/tecnologia_capacity.xlsx')
capacity_fixa = pd.read_excel('./arquivos/capacity_fixa.xlsx')
capacity_funil = pd.read_excel('./arquivos/capacity_funil.xlsx')
facilidades = pd.read_excel('./arquivos/facilidades_tecnologia_prioridade.xlsx').sort_values('PRIORIDADE',ascending=True).fillna('')
custos_proprios = pd.read_excel('./arquivos/custos_nfv.xlsx')
tecnologia_bbip_epl_eaccess = pd.read_excel('./arquivos/tecnologia_bbip_eaccess_epl.xlsx')
bb_municipio_estacao = pd.read_excel('./arquivos/bb_municipio_estacao.xlsx')
estacoes_entregas = pd.read_excel('./arquivos/estacoes_entregas.xlsx')
municipios = pd.read_excel('./arquivos/cidades_bbip.xlsx')
id_tecnologia_facilidade = pd.read_excel('./arquivos/id_tecnologia_facilidades.xlsx').fillna('')
id_provedores = pd.read_excel('./arquivos/id_provedores.xlsx').fillna('')
estacoes_newteia = pd.read_excel('./arquivos/lista_estacoes_newteia.xlsx').fillna('')
municipio_localidade = pd.read_excel('./arquivos/municipio_localidade.xlsx')

engine = create_engine('mysql+pymysql://viabilidade:senha_segura123#@10.0.15.243:3306/desenvolvimento_viabilidade')


# teia = pd.read_csv(r"C:\Users\F257064\Documents\Codes\AUTOMACOES_LOTE\arquivos_teste\ext_20250102_020045-512941a0b32d0e350ccae43a67681edf.csv",sep=';').fillna('')
arquivo_modelo = pd.read_excel('arquivo_padrao.xlsx')

try:
    valores_ethernet = pd.read_sql('SELECT * FROM valores_terceiros_eth_filtered', engine)
    valores_banda_larga = pd.read_sql('SELECT * FROM valores_terceiros_internet', engine)
    valores_internet_id = pd.read_sql('SELECT * FROM valores_terceiros_internet_id', engine)
    status = pd.read_sql('SELECT * FROM status', engine)
    status = status.fillna('OK')
except:
    valores_ethernet = pd.read_csv("./arquivos/valores_ethernet.csv")
    valores_banda_larga = pd.read_csv("./arquivos/valores_banda_larga.csv")
    valores_internet_id = pd.read_csv("./arquivos/valores_internet_id.csv")
    status = pd.read_csv("./arquivos/status.csv")
    pass


def converter_velocidade(valor):
    if not isinstance(valor, str) or len(valor) < 2:
        return None
    numero, unidade = valor[:-1], valor[-1]
    try:
        numero = float(numero.replace(',', '.'))
        if unidade == 'M':
            return numero
        elif unidade == 'G':
            return numero * 1000
        elif unidade == 'K':
            return numero / 1000
    except ValueError:
        return None
    return None

def padronizar_obs(valor):
    if pd.isna(valor):
        return None
    return unidecode(str(valor)).upper()

def processar_dataframe(df, padronizar_obs_flag=False):
    if padronizar_obs_flag and "OBS" in df.columns:
        df["OBS"] = df["OBS"].apply(padronizar_obs)
    if "VELOCIDADE" in df.columns:
        df["VEL"] = df["VELOCIDADE"].apply(converter_velocidade)
    return df

# Aplicando nos 3 dataframes
valores_ethernet = processar_dataframe(valores_ethernet)
valores_banda_larga = processar_dataframe(valores_banda_larga, padronizar_obs_flag=True)
valores_internet_id = processar_dataframe(valores_internet_id, padronizar_obs_flag=True)



def arquivo_teia():
    global teia, arquivo_modelo,nome_arquivo_padrao
    arquivo_entrada_sevs_teia = ctk.filedialog.askopenfilename(title='Abrir arquivo de extração do TEIA')
    teia = pd.read_csv(arquivo_entrada_sevs_teia, sep=';')
    sevs_removidas = extracaoTeia(teia).tratar_modelo_gaia(removed_sevs=check_remover.get())

    
    if type(sevs_removidas) != int:
        teia = teia[~teia.SEV.isin(sevs_removidas)]
        
    for index, value in teia.iterrows():
        arquivo_modelo.at[index,'SEV_PONTA_A'] = value.PONTA_A
        arquivo_modelo.at[index,'SEV'] = value.SEV
        arquivo_modelo.at[index,'CLIENTE'] = value.CLIENTE
        arquivo_modelo.at[index,'VELOCIDADE'] = value.VELOCIDADE
        arquivo_modelo.at[index,'CNL'] = value.CNL
        arquivo_modelo.at[index,'TIPO_LOGRADOURO'] = value.TIPO_LOGRADOURO
        arquivo_modelo.at[index,'NOME_DO_LOGRADOURO'] = value.NOME_DO_LOGRADOURO
        arquivo_modelo.at[index,'NUMERO'] = str(value.NUMERO).split('.')[0]
        arquivo_modelo.at[index,'COMPLEMENTO'] = value.COMPLEMENTO
        arquivo_modelo.at[index,'BAIRRO'] = value.BAIRRO	
        arquivo_modelo.at[index,'CIDADE'] = value.CIDADE
        arquivo_modelo.at[index,'UF'] = value.UF
        arquivo_modelo.at[index,'CEP'] = value.CEP
        arquivo_modelo.at[index,'SERVICO'] = value.SERVICO
        arquivo_modelo.at[index,'QTDE_CIRCUITOS'] = value.QTDE_CIRCUITOS
        arquivo_modelo.at[index,'ID_TEIA'] = value.ID_ANALISE
        arquivo_modelo.at[index,'LATITUDE'] = value.LATITUDE
        arquivo_modelo.at[index,'LONGITUDE'] = value.LONGITUDE

    arquivo_modelo = arquivo_modelo.reset_index(drop=True)

    for index, value in arquivo_modelo.iterrows():
        aux = municipio_localidade[municipio_localidade.SIGLA_LOC == value.CNL]
        if len(aux) > 0:
            arquivo_modelo.at[index,'CNL'] = aux.CNL.values[0]
        else:
            print(f'{value.SEV} nao encontrou CNL')

    arquivo_modelo.to_excel('04_PADRAO.xlsx',index=False)
    nome_arquivo_padrao = '04_PADRAO.xlsx'

    button_browse = ctk.CTkButton(janela,text="Arquivo TEIA",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=arquivo_teia)
    button_browse.place(x=150,y=100)

def arquivo_padrao():
    global nome_arquivo_padrao
    nome_arquivo_padrao = ctk.filedialog.askopenfilename(title='Abrir arquivo de extração do TEIA')
    button_browse_arq_padrao = ctk.CTkButton(janela,text="Arquivo PADRAO",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=arquivo_padrao)
    button_browse_arq_padrao.place(x=300,y=100)

def selecionar_resumosoe():
    global resumosoe
    arquivo_resumosoe = ctk.filedialog.askopenfilename(title='Abrir arquivo resumoSoE')
    resumosoe = tratamentoResumosoe(f"{arquivo_resumosoe}").trata_resumosoe()
    
    for index, value in resumosoe.iterrows():
        if value.TERCEIROS_ETH == 'Viável':
            if 'MOBWIRE' in value.TERCEIROS_ETH_INFORMACAO:
                aux_res = value.TERCEIROS_ETH_INFORMACAO
                aux_res = aux_res.replace('/ MOBWIRE /','').replace(' MOBWIRE /','').replace(' MOBWIRE','')
                if aux_res.split(':')[-1] == '':
                    resumosoe.at[index,'TERCEIROS_ETH'] = 'Inviável'
                    resumosoe.at[index,'TERCEIROS_ETH_INFORMACAO'] = ''
                else:
                    resumosoe.at[index,'TERCEIROS_ETH_INFORMACAO'] = aux_res

    resumosoe.to_excel('arquivo_resumosoe.xlsx',index=False)
    button_resumosoe = ctk.CTkButton(janela,text="ResumoSoE",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_resumosoe)
    button_resumosoe.place(x=10,y=200)

def selecionar_nuvens():
    global nuvens
    arquivo_nuvens = ctk.filedialog.askopenfilename(title='Abrir arquivo nuvens')
    nuvens = tratamentoNuvens(f"{arquivo_nuvens}").trata_nuvens()
    nuvens.to_excel('arquivo_nuvens.xlsx',index=False)
    button_nuvens = ctk.CTkButton(janela,text="Nuvens",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_nuvens)
    button_nuvens.place(x=100,y=200)

def selecionar_resultado():
    global resultado
    arquivo_resultado = ctk.filedialog.askopenfilename(title='Abrir arquivo resultado')
    resultado = tratamentoResultado(f"{arquivo_resultado}").trata_resultado()
    resultado.to_excel('arquivo_resultado.xlsx',index=False)
    button_resultado = ctk.CTkButton(janela,text="Resultado",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_resultado)
    button_resultado.place(x=165,y=200)


def selecionar_restricao():
    global restricao
    arquivo_restricao = ctk.filedialog.askopenfilename(title='Abrir arquivo restricao')
    restricao = tratamentoRestricao(f"{arquivo_restricao}").trata_restricao()
    restricao.to_excel('arquivo_restricao.xlsx',index=False)
    button_restricao = ctk.CTkButton(janela,text="Restrição",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_restricao)
    button_restricao.place(x=245, y=200)


def inclui_restricao():

    if check_restricao.get() == 'S':

        button_restricao.place(x=245, y=200)

    else:
        button_restricao.place(x=1000, y=1000)

def tratativa_inicial():
    global nuvens, restricao, resultado, resumosoe

    sevs_tratar = pd.read_excel(nome_arquivo_padrao)


    acum_nuvens = pd.DataFrame(columns=nuvens.columns)

    for i, v in nuvens.iterrows():

        try:
            tec = v.TECNOLOGIA.split(' / ')
            
            for value in tec:
                aux = nuvens.loc[i].to_dict()
                aux.update(TECNOLOGIA=value)
                acum_nuvens.loc[len(acum_nuvens)] = aux
        except:
            pass
    nuvens = acum_nuvens.drop_duplicates()

    for i, v in nuvens.iterrows():
        aux_capcity = tecnologia_capacity[tecnologia_capacity.TECNOLOGIA == v.TECNOLOGIA]
        if len(aux_capcity) > 0:
            if aux_capcity.CAPACITY.values[0] not in ['SIGLA_ESTACAO_CLARO','ESTACAO_ENTREGA']:
                if v.TECNOLOGIA == 'VIRTUA':
                    if sevs_tratar[sevs_tratar.SEV == v.SEV].SERVICO.values[0] == 'VPE - VIP BSOD LIGHT':
                        nuvens.at[i,'CAPACITY_NUVEM'] = 1000
                    else:
                        nuvens.at[i,'CAPACITY_NUVEM'] = 0
                else:
                    nuvens.at[i,'CAPACITY_NUVEM'] = int(aux_capcity.CAPACITY.values[0])
            else:
                if aux_capcity.CAPACITY.values[0] == 'ESTACAO_ENTREGA':
                # CONSULTAR UTILIZANDO AS COLUNAS DA TABELA TECNOLOGIA_CAPACITY, DO FUNIL E TABELA DE CAPACITYS.
                    if v.TECNOLOGIA == 'FO EDD NET':
                        if len(capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)]) > 0:
                            nuvens.at[i,'CAPACITY_NUVEM'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0]
                            nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CENTRO_ROTEAMENTO.values[0]
                    elif v.TECNOLOGIA == 'GPON NET':
                        if len(capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)]) > 0:
                            if capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0] > 200:
                                nuvens.at[i,'CAPACITY_NUVEM'] = 200
                                nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CENTRO_ROTEAMENTO.values[0]
                            else:
                                nuvens.at[i,'CAPACITY_NUVEM'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0]
                    elif v.TECNOLOGIA == 'SDH':
                        if len(capacity[(capacity.NUVEM == 'FO SDH') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)]) > 0:
                            nuvens.at[i,'CAPACITY_NUVEM'] = capacity[(capacity.NUVEM == 'FO SDH') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0]
                            nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity[(capacity.NUVEM == 'FO SDH') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CENTRO_ROTEAMENTO.values[0]
                    else:
                        if len(capacity_fixa[(capacity_fixa.TECNOLOGIA == v.TECNOLOGIA) & (capacity_fixa.SIGLA_EMBRATEL == v.ESTACAO_ENTREGA)]) > 0:
                            nuvens.at[i,'CAPACITY_NUVEM'] = capacity_fixa[(capacity_fixa.TECNOLOGIA == v.TECNOLOGIA) & (capacity_fixa.SIGLA_EMBRATEL == v.ESTACAO_ENTREGA)].TOTAL.values[0]
                            nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity_fixa[(capacity_fixa.TECNOLOGIA == v.TECNOLOGIA) & (capacity_fixa.SIGLA_EMBRATEL == v.ESTACAO_ENTREGA)].ESTACAO_BB.values[0]
                elif aux_capcity.CAPACITY.values[0] == 'SIGLA_ESTACAO_CLARO':
                    aux_funil = capacity_funil[capacity_funil.SITE == v.SIGLA_ESTACAO_CLARO]
                    if len(aux_funil) > 0:
                        nuvens.at[i,'CENTRO_ROTEAMENTO'] = v.ESTACAO_ENTREGA
                        if nuvens.at[i,'REDE'] == 'CORTE CAPACIDADE-BANDA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        elif 'CORTE PLANEJAMENTO REGIONAL' in nuvens.at[i,'REDE']:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'ESGOTADA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'CONCLUIDA':
                            if nuvens.at[i,'MEIO_TRANSMISSAO'] == 'REDE OPTICA':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 100
                            elif nuvens.at[i,'MEIO_TRANSMISSAO'] == 'ENLACE DE RADIO':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        nuvens.at[i,'CAPACITY_NUVEM'] = aux_funil.BANDA.values[0]
                        if (v.TECNOLOGIA == 'GPON MOVEL') & (nuvens.at[i,'CAPACITY_NUVEM'] > 200):
                            nuvens.at[i,'CAPACITY_NUVEM'] = 200

                    else:
                        nuvens.at[i,'CENTRO_ROTEAMENTO'] = v.ESTACAO_ENTREGA
                        if nuvens.at[i,'REDE'] == 'CORTE CAPACIDADE-BANDA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        elif 'CORTE PLANEJAMENTO REGIONAL' in nuvens.at[i,'REDE']:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'ESGOTADA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'CONCLUIDA':
                            if nuvens.at[i,'MEIO_TRANSMISSAO'] == 'REDE OPTICA':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 100
                            elif nuvens.at[i,'MEIO_TRANSMISSAO'] == 'ENLACE DE RADIO':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        else:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        
                    if nuvens.at[i,'CAPACITY_NUVEM'] == 0:
                        nuvens.at[i,'CENTRO_ROTEAMENTO'] = v.ESTACAO_ENTREGA
                        if nuvens.at[i,'REDE'] == 'CORTE CAPACIDADE-BANDA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        elif 'CORTE PLANEJAMENTO REGIONAL' in nuvens.at[i,'REDE']:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'ESGOTADA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'CONCLUIDA':
                            if nuvens.at[i,'MEIO_TRANSMISSAO'] == 'REDE OPTICA':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 100
                            elif nuvens.at[i,'MEIO_TRANSMISSAO'] == 'ENLACE DE RADIO':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 10
    try:
        nuvens['CAPACITY_NUVEM'] = nuvens['CAPACITY_NUVEM'].fillna(0)
    except:
        pass
    resumosoe['BANDA_ABORDADO'] = 0
    for index, value in resumosoe.iterrows():
        aux_capacity = capacity[(capacity.NUVEM_ABORDADO == value.FACILIDADE_ABORDADO) & (capacity.ESTACAO_ENTREGA == value.ESTACAO_ENTREGA_ABORDADO)]
        aux_capacity_fixa = capacity_fixa[(capacity_fixa.FACILIDADE == value.FACILIDADE_ABORDADO) & (capacity_fixa.SIGLA_EMBRATEL == value.ESTACAO_ENTREGA_ABORDADO)]
        if value.FACILIDADE_ABORDADO == 'FOetherNET':
            resumosoe.at[index,'BANDA_ABORDADO'] = 1000
        elif value.FACILIDADE_ABORDADO == 'FO SDH':
            resumosoe.at[index,'BANDA_ABORDADO'] = 0
        elif len(aux_capacity_fixa) > 0:
            resumosoe.at[index,'BANDA_ABORDADO'] = aux_capacity_fixa.TOTAL.values[0]
        elif len(aux_capacity) > 0:
            resumosoe.at[index,'BANDA_ABORDADO'] = aux_capacity.CAPACITY_MB.values[0]


    for index, value in sevs_tratar.iterrows():
        if value.VELOCIDADE[-4] == 'M':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4])
        elif value.VELOCIDADE[-4] == 'G':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) * 1000
        elif value.VELOCIDADE[-4] == 'K':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) / 1000


    nuvens.NOME_NUVEM = nuvens.NOME_NUVEM.replace(' ','')

    for i, v in nuvens.iterrows():
        if v.NOME_NUVEM == '':
            nuvens.at[i,'NOME_NUVEM'] = v.ESTACAO_ENTREGA

    for index,value in resumosoe.iterrows():
        for i,v in facilidades.iterrows():
            if 'NUVEM: /' in value[f'{v.FACILIDADE}_INFORMACAO']:
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = value[f'{v.FACILIDADE}_INFORMACAO'].replace('NUVEM: /','')
            elif 'NUVEM:  / ' in value[f'{v.FACILIDADE}_INFORMACAO']:
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = value[f'{v.FACILIDADE}_INFORMACAO'].replace('NUVEM:  / ','')
            elif 'NUVEM: ' in value[f'{v.FACILIDADE}_INFORMACAO']:
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = value[f'{v.FACILIDADE}_INFORMACAO'].replace('NUVEM: ','')
            
            if resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] == '':
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = resumosoe.at[index,f'{v.FACILIDADE}_ESTACAO_ENTREGA']

    sevs_tratar = sevs_tratar.fillna('')

    for index, value in sevs_tratar.iterrows():
        try:
            sevs_tratar.at[index,'PROTOCOLO_GAIA'] = resultado[resultado.SEV == value.SEV].PROTOCOLO.values[0]
        except:
            sevs_tratar.at[index,'PROTOCOLO_GAIA'] = 0
        if value.TRAVA_ACESSO != 'X':
            aux_resumosoe = resumosoe[resumosoe.SEV == value.SEV]
            if len(aux_resumosoe) > 0:
                
                if aux_resumosoe.FACILIDADE_ABORDADO.values[0] != '':
                    if value.SERVICO != 'VPE - VIP BSOD LIGHT':
                        if value.FACILIDADE_ACESSO_DISTINTO.upper() not in aux_resumosoe.FACILIDADE_ABORDADO.values[0].upper():
                            if aux_resumosoe.BANDA_ABORDADO.values[0] >= value.VEL:
                                sevs_tratar.at[index, 'RESPOSTA_FACILIDADE'] = aux_resumosoe.FACILIDADE_ABORDADO.values[0].upper()
                                sevs_tratar.at[index, 'ESTACAO_DE_ENTREGA'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                                match aux_resumosoe.FACILIDADE_ABORDADO.values[0]:
                                    case 'FO EDD ETH':
                                        sevs_tratar.at[index, 'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'FO EDD FIXA'

                                    case 'FOetherNET':
                                        sevs_tratar.at[index, 'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'FO EDD NET'
                                    
                                    case 'FO SDH':
                                        sevs_tratar.at[index, 'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'SDH'
                                    
                                    case 'FO GPON ETH':
                                        sevs_tratar.at[index, 'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'GPON FIXA'
                                
                                sevs_tratar.at[index, 'DESIGNACAO'] = aux_resumosoe.ID_ACESSO_ABORDADO.values[0]
                
                if value.RESPOSTA_FACILIDADE == '':
                    
                    for i,v in facilidades.iterrows():
                        if v.FACILIDADE != value.FACILIDADE_ACESSO_DISTINTO.upper().replace(' ','_'):
                            if ('Viável' in aux_resumosoe[v.FACILIDADE].values[0]) | ('Nuvem Avaliar Capacidade' in aux_resumosoe[v.FACILIDADE].values[0]):
                                if v.VERIFICA_CAPACITY == 'N':
                                    if v.FACILIDADE == 'HFC_BSOD':
                                        
                                        if 'HP GED' in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]:
                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                            sevs_tratar.at[index,'HP_GED'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split('HP GED ')[-1]
                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                            if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA HFC'
                                            else:
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'HFC BSOD'
                                            break
                                        else:
                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = f"ESTACAO ENTREGA {aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0]}"
                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].replace('ESTAÇÃO ENTRONCAMENTO:','')
                                            if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA HFC'
                                            else:
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'HFC BSOD'
                                            break
                                    elif v.FACILIDADE == '4G':
                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]
                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'LTE (4G)'
                                        break
                                    elif v.FACILIDADE == 'FO_GPON_RESID_ETH_PRE_VIAVEL':
                                        for tec in v.TECNOLOGIA.split('/'):
                                            if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == '':
                                                for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                    aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                    if len(aux_nuvem) > 0:
                                                        if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                            break
                                    elif v.FACILIDADE == 'FO_XGSPON_RESID_ETH':
                                        for tec in v.TECNOLOGIA.split('/'):
                                            for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                if len(aux_nuvem) > 0:
                                                    if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                        break
                                    elif v.FACILIDADE == 'FO_GPON_RESID_ETH':
                                        
                                        if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                            if 'HP GED' in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]:
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'HP_GED'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split('HP GED ')[-1]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA GPON'
                                                break
                                            else:
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == '':
                                                        
                                                        
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                
                                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA GPON'
                                                                    else:
                                                                        if tec != 'VIRTUA GPON':
                                                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                    break
                                            
                                        else:
                                            for tec in v.TECNOLOGIA.split('/'):
                                                if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == '':
                                                    
                                                    
                                                    for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                        aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                        if len(aux_nuvem) > 0:
                                                            
                                                            if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA GPON'
                                                                else:
                                                                    if tec != 'VIRTUA GPON':
                                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                break

                                    elif (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = 'RJO AM'
                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                        break
                                else:
                                    
                                    if v.FACILIDADE == 'TERCEIROS_ETH':
                                        if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == '':
                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                            sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0].split('PROPRIETÁRIO ')[-1]
                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe.TERCEIROS_ETH_ESTACAO_ENTREGA.values[0].split(' / ')[-1]
                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'TERCEIROS ETH'
                                            break
                                        
                                    else:
                                        for tec in v.TECNOLOGIA.split('/'):

                                            if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == '':
                                                for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                    
                                                    aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem.replace(':','')) & (nuvens.TECNOLOGIA == tec)]
                                                    if len(aux_nuvem) > 0:
                                                        if (value.SERVICO == 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                            if 'EPL MEF - NOK' not in aux_nuvem.OBSERVACAO.values[0]:
                                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                if tec in v.CONSULTA_FUNIL.split('/'):
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                                else:
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'{aux_nuvem.NOME_NUVEM.values[0]} {aux_nuvem.ESTACAO_ENTREGA.values[0]}'
                                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec

                                                                break

                                                        if (value.SERVICO != 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                            if tec in v.CONSULTA_FUNIL.split('/'):
                                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                            else:
                                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'{aux_nuvem.NOME_NUVEM.values[0]} {aux_nuvem.ESTACAO_ENTREGA.values[0]}'
                                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec

                                                            break

                                    
                            if aux_resumosoe[f'{v.FACILIDADE}'].values[0] == '%Disponibilidade não atende ao desejado':
                                if (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = 'RJO AM'
                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA


    sevs_tratar = sevs_tratar.drop(columns=['VEL'])

    if check_restricao.get() == 'S':
        for index, value in sevs_tratar.iterrows():
            aux_restricao = restricao[restricao.SEV == value.SEV]
            if len(aux_restricao) > 0:

                if len(aux_restricao) > 1:
                    if 'TOTAL' in aux_restricao.TIPO_DE_IMPACTO.tolist():
                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''

                else:
                    if aux_restricao.TIPO_DE_IMPACTO.values[0] == 'TOTAL':
                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
            

    for index, value in sevs_tratar.iterrows():
        if value.RESPOSTA_FACILIDADE == '':
            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'

    sevs_tratar.to_excel(nome_arquivo_padrao,index=False)
    button_fase_1 = ctk.CTkButton(janela,text="Tratativa inicial",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=tratativa_inicial)
    button_fase_1.place(x=10,y=250)

def prox_acesso():
    global nuvens, restricao, resultado, resumosoe

    sevs_tratar = pd.read_excel(nome_arquivo_padrao).fillna('')

    for index, value in sevs_tratar.iterrows():
        if value.VELOCIDADE[-4] == 'M':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4])
        elif value.VELOCIDADE[-4] == 'G':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) * 1000
        elif value.VELOCIDADE[-4] == 'K':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) / 1000

    for index, value in sevs_tratar.iterrows():
        sevs_tratar.at[index,'HP_GED'] = ''
        if value.RESPOSTA_FACILIDADE != 'INVIAVEL':
            # print(1)
            if value.TRAVA_ACESSO == '':
                # print(2)
            
                if value.DESIGNACAO != '':
                    # print(3)
                    sevs_tratar.at[index,'DESIGNACAO'] = ''


                aux_resumosoe = resumosoe[resumosoe.SEV == value.SEV]
                if len(aux_resumosoe) > 0:
                    # print(4)
                    
                    prioridade_tecnologia = int(facilidades[facilidades.TECNOLOGIA == value.TECNOLOGIA_ACESSO_PRINCIPAL].PRIORIDADE.values[0])
                    for i,v in facilidades[facilidades.PRIORIDADE > prioridade_tecnologia].iterrows():
                        # print(v.TECNOLOGIA)
                        if v.FACILIDADE != value.FACILIDADE_ACESSO_DISTINTO.upper().replace(' ','_'):
                            # print(5)
                            if v.FACILIDADE != value.FACILIDADE_ACESSO_DISTINTO.upper():
                                # print(6)
                                if sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL:
                                    # print(7)

                                    
                                    if ('Viável' in aux_resumosoe[v.FACILIDADE].values[0]) | ('Nuvem Avaliar Capacidade' in aux_resumosoe[v.FACILIDADE].values[0]):
                                        # print('a')

                                        if v.VERIFICA_CAPACITY == 'N':
                                            # print('b')
                                            if v.FACILIDADE == 'HFC_BSOD':
                                                # print('c')
                                                if (value.SERVICO == 'VPE - VIP BSOD LIGHT') & (v.TECNOLOGIA == 'VIRTUA HFC'):
                                                    if 'HP GED' in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]:
                                                        # print('d')
                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                        sevs_tratar.at[index,'HP_GED'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split('HP GED ')[-1]
                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                        break
                                                    else:
                                                        # print('e')
                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f"ESTACAO ENTREGA {aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0]}"
                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].replace('ESTAÇÃO ENTRONCAMENTO:','')
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                        break
                                                elif (value.SERVICO != 'VPE - VIP BSOD LIGHT') & (v.TECNOLOGIA != 'VIRTUA HFC'):
                                                    if 'HP GED' in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]:
                                                        # print('f')
                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                        sevs_tratar.at[index,'HP_GED'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split('HP GED ')[-1]
                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                        break
                                                    else:
                                                        # print('g')
                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f"ESTACAO ENTREGA {aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0]}"
                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].replace('ESTAÇÃO ENTRONCAMENTO:','')
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                        break
                                            elif v.FACILIDADE == '4G':
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'LTE (4G)'
                                                break
                                            elif v.FACILIDADE == 'FO_GPON_RESID_ETH_PRE_VIAVEL':
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    if sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL:
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                    break
                                            elif v.FACILIDADE == 'FO_XGSPON_RESID_ETH':
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                        aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                        if len(aux_nuvem) > 0:
                                                            if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                break
                                            elif v.FACILIDADE == 'FO_GPON_RESID_ETH':
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    if sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL:
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA GPON'
                                                                    else:
                                                                        if tec != 'VIRTUA GPON':
                                                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                    break

                                            elif (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = 'RJO AM'
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                break
                                        else:
                                            
                                            if v.FACILIDADE == 'TERCEIROS_ETH':
                                                
                                                if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == 'TERCEIROS ETH':
                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                                                    break
                                                if sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL:
                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0].split('PROPRIETÁRIO ')[-1]
                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe.TERCEIROS_ETH_ESTACAO_ENTREGA.values[0].split(' / ')[-1]
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'TERCEIROS ETH'

                                                    break
                                                
                                                
                                            else:
                                                
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    

                                                    if sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL:
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem.replace(':','')) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                if (value.SERVICO == 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                                    if 'EPL MEF - NOK' not in aux_nuvem.OBSERVACAO.values[0]:
                                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                        if tec in v.CONSULTA_FUNIL.split('/'):
                                                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                                        else:
                                                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'{aux_nuvem.NOME_NUVEM.values[0]} {aux_nuvem.ESTACAO_ENTREGA.values[0]}'
                                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                        

                                                                        break

                                                                elif (value.SERVICO != 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                                    
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    if tec in v.CONSULTA_FUNIL.split('/'):
                                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                                    else:
                                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'{aux_nuvem.NOME_NUVEM.values[0]} {aux_nuvem.ESTACAO_ENTREGA.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec

                                                                    break
                                                                elif value.VEL >= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                    pass                                                            
                                                                else:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''

                                            
                                    elif aux_resumosoe[f'{v.FACILIDADE}'].values[0] == '%Disponibilidade não atende ao desejado':
                                        
                                        if sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL:
                                            if (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = 'RJO AM'
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                break
                                    elif prioridade_tecnologia >= 21:
                                        
                                        if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                            sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                                            break
                                    # else:
                                    #     if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                    #         sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                    #         sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                    #         sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                    #         sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                    #         sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                                    #         break
                if (sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] == value.TECNOLOGIA_ACESSO_PRINCIPAL) & (value.DESIGNACAO == ''):
                    # print(value.TECNOLOGIA_ACESSO_PRINCIPAL)
                    # print(sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'])
                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                    sevs_tratar.at[index,'HP_GED'] = ''
                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''

    sevs_tratar = sevs_tratar.drop(columns=['VEL'])

    sevs_tratar.to_excel(nome_arquivo_padrao,index=False)
    
    print('RODOU PROXIMO ACESSO!')


def acesso_anterior():
    global nuvens, restricao, resultado, resumosoe

    sevs_tratar = pd.read_excel(nome_arquivo_padrao).fillna('')

    for index, value in sevs_tratar.iterrows():
        if value.VELOCIDADE[-4] == 'M':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4])
        elif value.VELOCIDADE[-4] == 'G':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) * 1000
        elif value.VELOCIDADE[-4] == 'K':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) / 1000

    for index, value in sevs_tratar.iterrows():


        if value.TRAVA_ACESSO == '':
            if value.DESIGNACAO != '':
                sevs_tratar.at[index,'DESIGNACAO'] = ''

            aux_resumosoe = resumosoe[resumosoe.SEV == value.SEV]
            if len(aux_resumosoe) > 0:
                                
                    if value.RESPOSTA_FACILIDADE != 'INVIAVEL':
                        prioridade_tecnologia = int(facilidades[facilidades.TECNOLOGIA == value.TECNOLOGIA_ACESSO_PRINCIPAL].PRIORIDADE.values[0])
                    else:
                        prioridade_tecnologia = len(facilidades)
                    for i,v in facilidades[facilidades.PRIORIDADE < prioridade_tecnologia].sort_values(by='PRIORIDADE',ascending=False).iterrows():
                        if v.FACILIDADE != value.FACILIDADE_ACESSO_DISTINTO.upper().replace(' ','_'):
                            if v.FACILIDADE != value.FACILIDADE_ACESSO_DISTINTO.upper():
                                if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                    if ('Viável' in aux_resumosoe[v.FACILIDADE].values[0]) | ('Nuvem Avaliar Capacidade' in aux_resumosoe[v.FACILIDADE].values[0]):
                                        
                                        if v.VERIFICA_CAPACITY == 'N':
                                            if v.FACILIDADE == 'HFC_BSOD':
                                                if 'HP GED' in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]:
                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                    sevs_tratar.at[index,'HP_GED'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split('HP GED ')[-1]
                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA HFC'
                                                    else:
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'HFC BSOD'
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                    break
                                                else:
                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f"ESTACAO ENTREGA {aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0]}"
                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].replace('ESTAÇÃO ENTRONCAMENTO:','')
                                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA HFC'
                                                    else:
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'HFC BSOD'
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                    break
                                            elif v.FACILIDADE == '4G':
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'LTE (4G)'
                                                sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                break
                                            elif v.FACILIDADE == 'FO_GPON_RESID_ETH_PRE_VIAVEL':
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                    break
                                            elif v.FACILIDADE == 'FO_XGSPON_RESID_ETH':
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                        aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                        if len(aux_nuvem) > 0:
                                                            if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                break
                                            elif v.FACILIDADE == 'FO_GPON_RESID_ETH':
                                                for tec in v.TECNOLOGIA.split('/'):
                                                    if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0]                                                                     
                                                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'VIRTUA GPON'
                                                                    else:
                                                                        if tec != 'VIRTUA GPON':
                                                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                    break

                                            elif (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = 'RJO AM'
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                break
                                        else:
                                            
                                            if v.FACILIDADE == 'TERCEIROS_ETH':
                                                if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == 'TERCEIROS ETH':
                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                    break
                                                if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0].split('PROPRIETÁRIO ')[-1]
                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_resumosoe.TERCEIROS_ETH_ESTACAO_ENTREGA.values[0].split(' / ')[-1]
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = 'TERCEIROS ETH'

                                                    break
                                                
                                                
                                            else:
                                                for tec in v.TECNOLOGIA.split('/'):

                                                    if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                                        
                                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                                            
                                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem.replace(':','')) & (nuvens.TECNOLOGIA == tec)]
                                                            if len(aux_nuvem) > 0:
                                                                if (value.SERVICO == 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                                    if 'EPL MEF - NOK' not in aux_nuvem.OBSERVACAO.values[0]:
                                                                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                        if tec in v.CONSULTA_FUNIL.split('/'):
                                                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                                        else:
                                                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'{aux_nuvem.NOME_NUVEM.values[0]} {aux_nuvem.ESTACAO_ENTREGA.values[0]}'
                                                                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                        sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''

                                                                        break

                                                                elif (value.SERVICO != 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):

                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                                    if tec in v.CONSULTA_FUNIL.split('/'):
                                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                                    else:
                                                                        sevs_tratar.at[index,'OBS_FECHAMENTO'] = f'{aux_nuvem.NOME_NUVEM.values[0]} {aux_nuvem.ESTACAO_ENTREGA.values[0]}'
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = tec
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''

                                                                    break
                                                                
                                                                else:
                                                                    sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                                                    sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                                                    sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                                                                    sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''

                                            
                                    elif aux_resumosoe[f'{v.FACILIDADE}'].values[0] == '%Disponibilidade não atende ao desejado':
                                        if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] == value.RESPOSTA_FACILIDADE:
                                            if (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                                sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = v.FACILIDADE.replace('_',' ')
                                                sevs_tratar.at[index,'OBS_FECHAMENTO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                                sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = 'RJO AM'
                                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = v.TECNOLOGIA
                                                sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                                break
                                    else:
                                        if sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] != value.RESPOSTA_FACILIDADE:

                                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                                            sevs_tratar.at[index,'OBS_FECHAMENTO'] = ''
                                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                                            sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                                            sevs_tratar.at[index,'CONCATENADO_PROVEDOR'] = ''
                                


    sevs_tratar = sevs_tratar.drop(columns=['VEL'])

    sevs_tratar.to_excel(nome_arquivo_padrao,index=False)

    print('RODOU ACESSO ANTERIOR!')

def precifica_sevs():

    sevs_tratar = pd.read_excel(nome_arquivo_padrao).fillna('')

    print('PRECIFICANDO SEVS...')

    for index, value in sevs_tratar.iterrows():
        if value.VELOCIDADE[-4] == 'M':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4])
        elif value.VELOCIDADE[-4] == 'G':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) * 1000
        elif value.VELOCIDADE[-4] == 'K':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) / 1000

    # for i,v in valores_ethernet.iterrows():
    #     try:
    #         if v.velocidade[-1] == 'M':
    #             valores_ethernet.at[i,'VEL'] = int(v.velocidade[:-1])
    #         elif v.velocidade[-1] == 'G':
    #             valores_ethernet.at[i,'VEL'] = float(v.velocidade[:-1].replace(',','.')) * 1000
    #         elif v.velocidade[-1] == 'K':
    #             valores_ethernet.at[i,'VEL'] = float(v.velocidade[:-1].replace(',','.')) / 1000
    #     except:
    #         pass

    # for i,v in valores_banda_larga.iterrows():
    #     try:
    #         valores_banda_larga.at[i,'obs'] = unidecode(v.obs).upper()
    #         if v.velocidade[-1] == 'M':
    #             valores_banda_larga.at[i,'VEL'] = int(v.velocidade[:-1])
    #         elif v.velocidade[-1] == 'G':
    #             valores_banda_larga.at[i,'VEL'] = float(v.velocidade[:-1].replace(',','.')) * 1000
    #         elif v.velocidade[-1] == 'K':
    #             valores_banda_larga.at[i,'VEL'] = float(v.velocidade[:-1].replace(',','.')) / 1000
    #     except:
    #         pass


    for index,value in sevs_tratar.iterrows():
        print(round((index/len(sevs_tratar)*100),2),end="\r")
        valores_ethernet_ = valores_ethernet[valores_ethernet.PROVEDOR.isin(value.CONCATENADO_PROVEDOR.split(' / '))]
        valores_banda_larga_ = valores_banda_larga[valores_banda_larga.PROVEDOR.isin(value.CONCATENADO_PROVEDOR.split(' / '))]
        valores_internet_id_ = valores_internet_id[valores_internet_id.PROVEDOR.isin(value.CONCATENADO_PROVEDOR.split(' / '))]
        if value.TRAVA_CUSTO == '':
            
            aux_custo_proprio = custos_proprios[custos_proprios.FACILIDADE == value.RESPOSTA_FACILIDADE.replace(' ','_')]
            if len(aux_custo_proprio) > 0:
                
                if value.DESIGNACAO != '':
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo_proprio.ACESSO_ABORDADO.values[0]
                elif value.VEL <= 202:
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo_proprio.NOVO_ACESSO_200M.values[0]
                elif value.VEL <= 502:
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo_proprio.NOVO_ACESSO_500M.values[0]
                elif value.VEL <= 1000:
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo_proprio.NOVO_ACESSO_1G.values[0]
            elif value.RESPOSTA_FACILIDADE == 'TERCEIROS ETH':
                if value.SERVICO != 'VPE - VIP BSOD LIGHT':
                    if value.SINALIZADOR_SIMETRICO == '':
                        
                        ######## PRECIFICA LAST MILE ################
                        melhor_provedor = ''
                        melhor_instal = 0
                        melhor_mensal = 0
                        custo_mensalizado = 0
                        for p in value.CONCATENADO_PROVEDOR.split(' / '):
                            
                            aux_provedor = valores_ethernet_[(valores_ethernet_.PROVEDOR == p) & (valores_ethernet_.SIGLA_MUNICIPIO == value.CNL) & (valores_ethernet_.UF == value.UF)]
                            if len(aux_provedor) > 0:
                                
                                aux_status = status[(status.PROVEDOR == aux_provedor.PROVEDOR.values[0]) & (status.UF == aux_provedor.UF.values[0])]
                                if len(aux_status) > 0:
                                    if (aux_status.STATUS.values[0] != 'BLOQUEADO') & (aux_status.STATUS.values[0] != 'AEROPORTO'):
                                        
                                        aux_valores = aux_provedor[(aux_provedor.VEL == value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                                        if len(aux_valores) > 0:
                                            if melhor_provedor == '':
                                                melhor_provedor = p
                                                melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                            else:
                                                if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                                    melhor_provedor = p
                                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal
                                        else:
                                            aux_valores = aux_provedor[(aux_provedor.VEL >= value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                                            if len(aux_valores) > 0:
                                                if melhor_provedor == '':
                                                    melhor_provedor = p
                                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                                else:
                                                    if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                                        melhor_provedor = p
                                                        melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                        melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                        custo_mensalizado = (melhor_instal / 24) + melhor_mensal     
                        if melhor_provedor != '':
                            sevs_tratar.at[index,'PROVEDOR_FINAL_TER'] = melhor_provedor
                            sevs_tratar.at[index,'INSTALACAO_TER'] = melhor_instal
                            sevs_tratar.at[index,'MENSAL_TER'] = melhor_mensal
                        else:
                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                    else:
                        ###### PRECIFICA PROVEDOR SIMETRICO
                        melhor_provedor = ''
                        melhor_instal = 0
                        melhor_mensal = 0
                        custo_mensalizado = 0
                        for p in value.CONCATENADO_PROVEDOR.split(' / '):
                            aux_provedor = valores_internet_id_[(valores_internet_id_.PROVEDOR == p) & (valores_internet_id_.SIGLA_MUNICIPIO == value.CNL) & (valores_internet_id_.UF == value.UF)]
                            if len(aux_provedor) > 0:
                                aux_status = status[(status.PROVEDOR == aux_provedor.PROVEDOR.values[0]) & (status.UF == aux_provedor.UF.values[0])]
                                if len(aux_status) > 0:
                                    if (aux_status.STATUS.values[0] != 'BLOQUEADO') & (aux_status.STATUS.values[0] != 'AEROPORTO'):
                                        aux_valores = aux_provedor[(aux_provedor.VEL == value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                                        if len(aux_valores) > 0:
                                            if melhor_provedor == '':
                                                melhor_provedor = p
                                                melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                            else:
                                                if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                                    melhor_provedor = p
                                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal
                                        else:
                                            aux_valores = aux_provedor[(aux_provedor.VEL >= value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                                            if len(aux_valores) > 0:
                                                if melhor_provedor == '':
                                                    melhor_provedor = p
                                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                                else:
                                                    if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                                        melhor_provedor = p
                                                        melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                        melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                        custo_mensalizado = (melhor_instal / 24) + melhor_mensal     
                        if melhor_provedor != '':
                            sevs_tratar.at[index,'PROVEDOR_FINAL_TER'] = melhor_provedor
                            sevs_tratar.at[index,'INSTALACAO_TER'] = melhor_instal
                            sevs_tratar.at[index,'MENSAL_TER'] = melhor_mensal
                        else:
                            sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                            sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''
                else:
                    ############# PRECIFICA BLC/ASSIMETRICO
                    melhor_provedor = ''
                    melhor_instal = 0
                    melhor_mensal = 0
                    custo_mensalizado = 0
                    for p in value.CONCATENADO_PROVEDOR.split(' / '):
                        aux_provedor = valores_banda_larga_[(valores_banda_larga_.PROVEDOR == p) & (valores_banda_larga_.SIGLA_MUNICIPIO == value.CNL) & (valores_banda_larga_.UF == value.UF)]
                        if len(aux_provedor) > 0:
                            aux_status = status[(status.PROVEDOR == aux_provedor.PROVEDOR.values[0]) & (status.UF == aux_provedor.UF.values[0])]
                            if len(aux_status) > 0:
                                if (aux_status.STATUS.values[0] != 'BLOQUEADO') & (aux_status.STATUS.values[0] != 'AEROPORTO'):
                                    aux_valores = aux_provedor[(aux_provedor.VEL == value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                                    if len(aux_valores) > 0:
                                        if melhor_provedor == '':
                                            melhor_provedor = p
                                            melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                            melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                            custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                        else:
                                            if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                                melhor_provedor = p
                                                melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                custo_mensalizado = (melhor_instal / 24) + melhor_mensal
                                    else:
                                        aux_valores = aux_provedor[(aux_provedor.VEL >= value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                                        if len(aux_valores) > 0:
                                            if melhor_provedor == '':
                                                melhor_provedor = p
                                                melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                            else:
                                                if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                                    melhor_provedor = p
                                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal     
                    if melhor_provedor != '':
                        sevs_tratar.at[index,'PROVEDOR_FINAL_TER'] = melhor_provedor
                        sevs_tratar.at[index,'INSTALACAO_TER'] = melhor_instal
                        sevs_tratar.at[index,'MENSAL_TER'] = melhor_mensal
                    else:
                        sevs_tratar.at[index,'RESPOSTA_FACILIDADE'] = 'INVIAVEL'
                        sevs_tratar.at[index,'ESTACAO_DE_ENTREGA'] = ''
                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO_PRINCIPAL'] = ''

    sevs_tratar = sevs_tratar.drop(columns=['VEL'])

    sevs_tratar.to_excel(nome_arquivo_padrao,index=False)

    print('SEVS PRECIFICADAS!')


def roda_bbip():

    sevs_tratar = pd.read_excel(nome_arquivo_padrao).fillna('')

    for index, value in sevs_tratar.iterrows():
        if value.VELOCIDADE[-4] == 'M':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4])
        elif value.VELOCIDADE[-4] == 'G':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) * 1000
        elif value.VELOCIDADE[-4] == 'K':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) / 1000


    for index, value in sevs_tratar.iterrows():
        if value.SERVICO != 'VPE - VIP BSOD LIGHT':
            if value.SINALIZADOR_SIMETRICO == '':
                if (value.RESPOSTA_FACILIDADE != 'INVIAVEL') &  (value.RESPOSTA_FACILIDADE != ''):
                    if value.BBIP == '':
                        if (value.SERVICO == 'EIN - E-ACCESS') | (value.SERVICO == 'LAN - LAN EPL MEF') | (value.SERVICO == 'LAN - LAN EPL'):
                            if value.VEL > 102:
                                aux_tec = tecnologia_bbip_epl_eaccess[tecnologia_bbip_epl_eaccess.TECNOLOGIA == value.TECNOLOGIA_ACESSO_PRINCIPAL]
                                if value.SERVICO != 'EIN - E-ACCESS':
                                    if aux_tec.BBIP_EPL.values[0] == 'S':
                                        aux_municipio = bb_municipio_estacao[bb_municipio_estacao.ESTACAO == value.ESTACAO_DE_ENTREGA]
                                        if len(aux_municipio) > 0:
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                                        else:
                                            aux_municipio = estacoes_entregas[estacoes_entregas.UF == value.UF]
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                                else:
                                    if aux_tec.BBIP_EACCESS.values[0] == 'S':
                                        aux_municipio = bb_municipio_estacao[bb_municipio_estacao.ESTACAO == value.ESTACAO_DE_ENTREGA]
                                        if len(aux_municipio) > 0:
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                                        else:
                                            aux_municipio = estacoes_entregas[estacoes_entregas.UF == value.UF]
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                        else:
                            aux_facilidade = facilidades[facilidades.FACILIDADE == value.RESPOSTA_FACILIDADE.replace(' ','_')]
                            
                            if (aux_facilidade.VERIFICA_CAPACITY.values[0] == 'S') & (value.TECNOLOGIA_ACESSO_PRINCIPAL !='GPON FIXA'):
                                if (value.RESPOSTA_FACILIDADE =='FO GPON ETH') & (value.TECNOLOGIA_ACESSO_PRINCIPAL ==''):
                                    continue
                                elif (value.RESPOSTA_FACILIDADE =='FO GPON ETH') & (value.TECNOLOGIA_ACESSO_PRINCIPAL !=''):
                                    if value.VEL > 102:
                                        aux_municipio = bb_municipio_estacao[bb_municipio_estacao.ESTACAO == value.ESTACAO_DE_ENTREGA]
                                        if len(aux_municipio) > 0:
                                            
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                                        else:
                                            aux_municipio = estacoes_entregas[estacoes_entregas.UF == value.UF]
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                                elif (value.RESPOSTA_FACILIDADE !='FO GPON ETH'):
                                    if value.VEL > 102:
                                        aux_municipio = bb_municipio_estacao[bb_municipio_estacao.ESTACAO == value.ESTACAO_DE_ENTREGA]
                                        if len(aux_municipio) > 0:
                                            
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'
                                        else:
                                            aux_municipio = estacoes_entregas[estacoes_entregas.UF == value.UF]
                                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio.MUNICIPIO.values[0]}|{aux_municipio.ESTACAO.values[0]}'




    navegador = webdriver.Chrome(service=service)
    navegador.implicitly_wait(5)
    navegador.get('http://10.100.1.30/admredes/admredes/RPVB_BLD_Cadastrar_2.asp')
    time.sleep(25)


    for index, value in sevs_tratar[(sevs_tratar.BBIP != '')].iterrows():
        if value.SERVICO != 'VPE - VIP BSOD LIGHT':
            try:
                navegador.get('http://10.100.1.30/admredes/admredes/RPVB_BLD_Cadastrar_2.asp')
                time.sleep(1)
                if (value.BBIP != None) & (value.BBIP != 'BBIP') & ("ID" not in str(value.BBIP)):
                    navegador.find_element('name','cliente').send_keys(value.CLIENTE)
                    navegador.find_element('name','sev').send_keys(value.SEV)
                    municipio, estacao = value.BBIP.split('|')
                    municipio = unidecode(municipio).upper()
                    id_municipio = str(municipios[municipios.CIDADE == municipio].ID.values[0])
                    Select(navegador.find_element('id', 'combo1')).select_by_value(id_municipio)
                    Select(navegador.find_element('id', 'combo2')).select_by_visible_text(estacao)
                    if 'E-ACCESS' in value.SERVICO or 'EPL' in value.SERVICO:
                        Select(navegador.find_element('name', 'servico')).select_by_value('EPL')
                    else:
                        Select(navegador.find_element('name', 'servico')).select_by_value('Internet')
                    navegador.find_element('name', 'velocidade').send_keys(int(value.VEL))
                    navegador.find_element('xpath', '/html/body/div[6]/font/form/div/table[1]/tbody/tr[2]/td/table[5]/tbody/tr/td/input').click()
                    time.sleep(4)
                    id_bbip = navegador.find_element('xpath','/html/body/div[6]/font/form/div/table[1]/tbody/tr[2]/td/table[1]/tbody/tr[1]/td').text[:-4]
                    status_bbip = navegador.find_element('xpath','/html/body/div[6]/font/form/div/table[1]/tbody/tr[2]/td/table[4]/tbody/tr/td[2]/font/b').text
                    sevs_tratar.at[index,'BBIP'] = f'{id_bbip} / {status_bbip}'
            except:
                continue

    navegador.close()

    sevs_tratar.to_excel(nome_arquivo_padrao,index=False)

    print('RODOU BACKBONE!')

def finaliza_sevs():

    fechamento_teia = pd.DataFrame(columns=['sequencial','latitude','longitude','uf','cnl','facilidade','id_facilidade','provedor','id_provedor','entrega',
                                        'abordado','custo_de_acesso_proprio','instalacao_terceiros','mensalidade_terceiros','tipo_terceiros','id_da_sev','prazo',
                                        'bb_ip','hp_bsod','codigo_spe','sinalizacao_sip','protocolo_gaia','obs','justificativa','ID Justificativa','status','tecnologia'])

    sevs_tratar = pd.read_excel(nome_arquivo_padrao).fillna('')

    for i,v in id_tecnologia_facilidade.iterrows():
        try:
            id_tecnologia_facilidade.at[i,'ID_EMPRESA'] = str(v.ID_EMPRESA).split('.')[0]
        except:
            pass

    tipo_fechamento = 'NFV FASE 2'

    for index, value in sevs_tratar.iterrows():

        if value.RESPOSTA_FACILIDADE != 'INVIAVEL':
            if 'Indeferido' not in value.BBIP:
                if len(fechamento_teia) == 0:
                    fechamento_teia.at[0,'sequencial'] = value.SEV
                    fechamento_teia.at[0,'latitude'] = value.LATITUDE
                    fechamento_teia.at[0,'longitude'] = value.LONGITUDE
                    fechamento_teia.at[0,'uf'] = value.UF
                    fechamento_teia.at[0,'cnl'] = value.CNL
                    fechamento_teia.at[0,'facilidade'] = value.RESPOSTA_FACILIDADE

                    aux_id_facilidade = id_tecnologia_facilidade[id_tecnologia_facilidade.FACILIDADE_FECHAMENTO == value.RESPOSTA_FACILIDADE]

                    fechamento_teia.at[0,'id_facilidade'] = aux_id_facilidade.ID.values[0]
                    if value.RESPOSTA_FACILIDADE != 'TERCEIROS ETH':
                        fechamento_teia.at[0,'provedor'] = aux_id_facilidade.EMPRESA.values[0]
                        fechamento_teia.at[0,'id_provedor'] = aux_id_facilidade.ID_EMPRESA.values[0]
                    else:
                        if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                            aux_provedores = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains('BANDA LARGA')]
                            aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                            aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.UF)]
                            try:
                                fechamento_teia.at[0,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                fechamento_teia.at[0,'id_provedor'] = aux_provedores.ID.values[0]
                            except:
                                print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                fechamento_teia.at[0,'provedor'] = 'NAO ENCONTRADO'
                                fechamento_teia.at[0,'id_provedor'] = 'NAO ENCONTRADO'
                        else:
                            if value.SINALIZADOR_SIMETRICO == '':
                                aux_provedores = id_provedores[~id_provedores.PROVEDOR_TEIA.str.contains('BANDA LARGA')]
                                aux_provedores = aux_provedores[~aux_provedores.PROVEDOR_TEIA.str.contains('SIMETRICO')]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.UF)]
                                if len(aux_provedores) == 0:
                                    aux_provedores = id_provedores[~id_provedores.PROVEDOR_TEIA.str.contains('BANDA LARGA')]
                                    aux_provedores = aux_provedores[~aux_provedores.PROVEDOR_TEIA.str.contains('SIMETRICO')]
                                    aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                                try:
                                    fechamento_teia.at[0,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                    fechamento_teia.at[0,'id_provedor'] = aux_provedores.ID.values[0]
                                except:
                                    print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                    fechamento_teia.at[0,'provedor'] = 'NAO ENCONTRADO'
                                    fechamento_teia.at[0,'id_provedor'] = 'NAO ENCONTRADO'
                            else:
                                aux_provedores = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains('SIMETRICO')]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.UF)]
                                try:
                                    fechamento_teia.at[0,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                    fechamento_teia.at[0,'id_provedor'] = aux_provedores.ID.values[0]
                                except:
                                    print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                    fechamento_teia.at[0,'provedor'] = 'NAO ENCONTRADO'
                                    fechamento_teia.at[0,'id_provedor'] = 'NAO ENCONTRADO'
                    if value.ESTACAO_DE_ENTREGA == '':
                        aux_estacao = estacoes_entregas[estacoes_entregas.UF == value.UF]
                        fechamento_teia.at[0,'entrega'] = aux_estacao.ESTACAO.values[0]
                    else:
                        aux_estacao = estacoes_newteia[estacoes_newteia.estacao == value.ESTACAO_DE_ENTREGA]
                        if len(aux_estacao) > 0:
                            fechamento_teia.at[0,'entrega'] = value.ESTACAO_DE_ENTREGA
                        else:
                            aux_estacao = estacoes_entregas[estacoes_entregas.UF == value.UF]
                            fechamento_teia.at[0,'entrega'] = aux_estacao.ESTACAO.values[0]

                    fechamento_teia.at[0,'id_da_sev'] = value.ID_TEIA
                    fechamento_teia.at[0,'prazo'] = ''
                    fechamento_teia.at[0,'codigo_spe'] = value.COD_SPE
                    fechamento_teia.at[0,'sinalizacao_sip'] = ''
                    fechamento_teia.at[0,'protocolo_gaia'] = value.PROTOCOLO_GAIA
                    fechamento_teia.at[0,'status'] = '1'

                    try:
                        if value.DESIGNACAO != '':
                            fechamento_teia.at[0,'abordado'] = 'SIM'
                            fechamento_teia.at[0,'obs'] = f'FECHAMENTO {tipo_fechamento}/{value.OBS_FECHAMENTO}/ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}/DESIG ABORDADO {value.DESIGNACAO}'
                        else:
                            fechamento_teia.at[0,'abordado'] = 'NAO'
                            fechamento_teia.at[0,'obs'] = f'FECHAMENTO {tipo_fechamento}/{value.OBS_FECHAMENTO}/ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}'
                    except:
                        fechamento_teia.at[0,'abordado'] = 'NAO'
                        fechamento_teia.at[0,'obs'] = f'FECHAMENTO {tipo_fechamento}/{value.OBS_FECHAMENTO}/ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}'

                    if value.RESPOSTA_FACILIDADE != 'TERCEIROS ETH':
                        fechamento_teia.at[0,'custo_de_acesso_proprio'] = str(value.CUSTO_ACESSO_PROPRIO).replace('.',',')
                        fechamento_teia.at[0,'obs'] = f'FECHAMENTO {tipo_fechamento}/ ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}/{value.OBS_FECHAMENTO}'
                    else:
                        fechamento_teia.at[0,'instalacao_terceiros'] = str(value.INSTALACAO_TER).replace('.',',')
                        fechamento_teia.at[0,'mensalidade_terceiros'] = str(value.MENSAL_TER).replace('.',',')
                        fechamento_teia.at[0,'tipo_terceiros'] = '3'
                        fechamento_teia.at[0,'justificativa'] = 'FORA DE REDE'
                        fechamento_teia.at[0,'ID Justificativa'] = '1'
                        if value.TERCEIRO_COTACAO == '':
                            fechamento_teia.at[len(fechamento_teia) - 1,'obs'] += '/PRECO PADRAO'
                        else:
                            fechamento_teia.at[len(fechamento_teia) - 1,'obs'] += '/PRECO COTACAO'

                    if value.BBIP != '':
                        fechamento_teia.at[0,'bb_ip'] = value.BBIP.split(' / ')[0].split(': ')[-1]

                    if value.HP_GED != '':
                        fechamento_teia.at[0,'hp_bsod'] = str(value.HP_GED).replace('.0','')
                    fechamento_teia.at[0,'tecnologia'] = value.TECNOLOGIA_ACESSO_PRINCIPAL
                else:
                    fechamento_teia.at[len(fechamento_teia),'sequencial'] = value.SEV
                    fechamento_teia.at[len(fechamento_teia) - 1,'latitude'] = value.LATITUDE
                    fechamento_teia.at[len(fechamento_teia) - 1,'longitude'] = value.LONGITUDE
                    fechamento_teia.at[len(fechamento_teia) - 1,'uf'] = value.UF
                    fechamento_teia.at[len(fechamento_teia) - 1,'cnl'] = value.CNL
                    fechamento_teia.at[len(fechamento_teia) - 1,'facilidade'] = value.RESPOSTA_FACILIDADE
                    
                    aux_id_facilidade = id_tecnologia_facilidade[id_tecnologia_facilidade.FACILIDADE_FECHAMENTO == value.RESPOSTA_FACILIDADE]
                    fechamento_teia.at[len(fechamento_teia) - 1,'id_facilidade'] = aux_id_facilidade.ID.values[0]
                    if value.RESPOSTA_FACILIDADE != 'TERCEIROS ETH':
                        fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = aux_id_facilidade.EMPRESA.values[0]
                        fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = aux_id_facilidade.ID_EMPRESA.values[0]
                    else:
                        if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                            aux_provedores = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains('BANDA LARGA')]
                            aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                            aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.UF)]
                            try:
                                fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = aux_provedores.ID.values[0]
                            except:
                                print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = 'NAO ENCONTRADO'
                                fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = 'NAO ENCONTRADO'
                        else:
                            if value.SINALIZADOR_SIMETRICO == '':
                                aux_provedores = id_provedores[~id_provedores.PROVEDOR_TEIA.str.contains('BANDA LARGA')]
                                aux_provedores = aux_provedores[~aux_provedores.PROVEDOR_TEIA.str.contains('SIMETRICO')]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.UF)]
                                if len(aux_provedores) == 0:
                                    aux_provedores = id_provedores[~id_provedores.PROVEDOR_TEIA.str.contains('BANDA LARGA')]
                                    aux_provedores = aux_provedores[~aux_provedores.PROVEDOR_TEIA.str.contains('SIMETRICO')]
                                    aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA == value.PROVEDOR_FINAL_TER]
                                try:
                                    fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                    fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = aux_provedores.ID.values[0]
                                except:
                                    print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                    fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = 'NAO ENCONTRADO'
                                    fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = 'NAO ENCONTRADO'
                            else:
                                aux_provedores = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains('SIMETRICO')]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_FINAL_TER)]
                                aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA.str.contains(value.UF)]
                                try:
                                    fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                    fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = aux_provedores.ID.values[0]
                                except:
                                    # print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                    # fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = 'NAO ENCONTRADO'
                                    # fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = 'NAO ENCONTRADO'
                                    aux_provedores = aux_provedores[aux_provedores.PROVEDOR_TEIA == value.PROVEDOR_FINAL_TER]
                                    try:
                                        fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = aux_provedores.PROVEDOR_TEIA.values[0]
                                        fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = aux_provedores.ID.values[0]
                                    except:
                                        print(f'SEV {value.SEV} PROVEDOR NAO ENCONTRADO NA BASE')
                                        fechamento_teia.at[len(fechamento_teia) - 1,'provedor'] = 'NAO ENCONTRADO'
                                        fechamento_teia.at[len(fechamento_teia) - 1,'id_provedor'] = 'NAO ENCONTRADO'
                    if value.ESTACAO_DE_ENTREGA == '':
                        aux_estacao = estacoes_entregas[estacoes_entregas.UF == value.UF]
                        fechamento_teia.at[len(fechamento_teia) - 1,'entrega'] = aux_estacao.ESTACAO.values[0]
                    else:
                        aux_estacao = estacoes_newteia[estacoes_newteia.estacao == value.ESTACAO_DE_ENTREGA]
                        if len(aux_estacao) > 0:
                            fechamento_teia.at[len(fechamento_teia) - 1,'entrega'] = value.ESTACAO_DE_ENTREGA
                        else:
                            aux_estacao = estacoes_entregas[estacoes_entregas.UF == value.UF]
                            fechamento_teia.at[len(fechamento_teia) - 1,'entrega'] = aux_estacao.ESTACAO.values[0]

                    fechamento_teia.at[len(fechamento_teia) - 1,'id_da_sev'] = value.ID_TEIA
                    fechamento_teia.at[len(fechamento_teia) - 1,'prazo'] = ''
                    fechamento_teia.at[len(fechamento_teia) - 1,'codigo_spe'] = value.COD_SPE
                    fechamento_teia.at[len(fechamento_teia) - 1,'sinalizacao_sip'] = ''
                    fechamento_teia.at[len(fechamento_teia) - 1,'protocolo_gaia'] = value.PROTOCOLO_GAIA
                    fechamento_teia.at[len(fechamento_teia) - 1,'status'] = '1'

                    try:
                        if value.DESIGNACAO != '':
                            
                            fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'SIM'
                            fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO {tipo_fechamento}/{value.OBS_FECHAMENTO}/ ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}/DESIG ABORDADO {value.DESIGNACAO}'
                        else:
                            fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'NAO'
                            fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO {tipo_fechamento}/{value.OBS_FECHAMENTO}/ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}'
                    except:
                        fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'NAO'
                        fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO {tipo_fechamento}/{value.OBS_FECHAMENTO}/ESTACAO ENTREGA {value.ESTACAO_DE_ENTREGA}'

                    if value.RESPOSTA_FACILIDADE != 'TERCEIROS ETH':
                        fechamento_teia.at[len(fechamento_teia) - 1,'custo_de_acesso_proprio'] = str(value.CUSTO_ACESSO_PROPRIO).replace('.',',')

                    else:
                        fechamento_teia.at[len(fechamento_teia) - 1,'instalacao_terceiros'] = str(value.INSTALACAO_TER).replace('.',',')
                        fechamento_teia.at[len(fechamento_teia) - 1,'mensalidade_terceiros'] = str(value.MENSAL_TER).replace('.',',')
                        fechamento_teia.at[len(fechamento_teia) - 1,'tipo_terceiros'] = '3'
                        fechamento_teia.at[len(fechamento_teia) - 1,'justificativa'] = 'FORA DE REDE'
                        fechamento_teia.at[len(fechamento_teia) - 1,'ID Justificativa'] = '1'
                        if value.TERCEIRO_COTACAO == '':
                            fechamento_teia.at[len(fechamento_teia) - 1,'obs'] += '/PRECO PADRAO'
                        else:
                            fechamento_teia.at[len(fechamento_teia) - 1,'obs'] += '/PRECO COTACAO'
                    if value.BBIP != '':
                        fechamento_teia.at[len(fechamento_teia) - 1,'bb_ip'] = value.BBIP.split(' / ')[0].split(': ')[-1]

                    if value.HP_GED != '':
                        fechamento_teia.at[len(fechamento_teia) - 1,'hp_bsod'] = str(value.HP_GED).replace('.0','')
                    fechamento_teia.at[len(fechamento_teia) - 1,'tecnologia'] = value.TECNOLOGIA_ACESSO_PRINCIPAL

    fechamento_teia.fillna('').to_csv('fechamento_lote_semiauto.csv',sep=';',index=False)

    print('ARQUIVO FECHAMENTO GERADO!')

janela = ctk.CTk()

janela.title("NFV")
janela.geometry("700x400")
janela.resizable(width=False, height=False)

ctk.CTkLabel(janela,text='NFV',font=('Arial',20)).place(x=220,y=20)

img = ctk.CTkImage(dark_image=Image.open('./claro.png'),light_image=Image.open('./claro.png'), size=(75,75))
ctk.CTkLabel(janela,text='',image=img).place(x=5,y=0)

titulo_teia= ctk.CTkLabel(janela,text='Selecione o arquivo de entrada (TEIA)')
titulo_teia.place(x=10,y=70)

check_remover = ctk.StringVar(value='N')

checkbox_remover = ctk.CTkCheckBox(janela, text="Remover SEVs?",
                                    variable=check_remover, onvalue="S", offvalue="N",checkbox_height=20,checkbox_width=20)
checkbox_remover.place(x=10, y=100)

button_browse = ctk.CTkButton(janela,text="Arquivo TEIA",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=arquivo_teia)
button_browse.place(x=150,y=100)

titulo_arq_padrao = ctk.CTkLabel(janela,text='Selecione o arquivo de tratamento')
titulo_arq_padrao.place(x=300,y=70)

button_browse_arq_padrao = ctk.CTkButton(janela,text="Arquivo PADRAO",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=arquivo_padrao)
button_browse_arq_padrao.place(x=300,y=100)


# button_gera_mapinfo = ctk.CTkButton(janela,text="Gerar Arquivo para mapinfo",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue')
# button_gera_mapinfo.place(x=10,y=200)

titulo_arquivo_gaia= ctk.CTkLabel(janela,text='Selecione os arquivo de retorno do GAIA')
titulo_arquivo_gaia.place(x=10,y=130)

button_resumosoe = ctk.CTkButton(janela,text="ResumoSoE",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_resumosoe)
button_resumosoe.place(x=10,y=200)

button_nuvens = ctk.CTkButton(janela,text="Nuvens",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_nuvens)
button_nuvens.place(x=100,y=200)

button_resultado = ctk.CTkButton(janela,text="Resultado",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_resultado)
button_resultado.place(x=165,y=200)

button_restricao = ctk.CTkButton(janela,text="Restrição",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_restricao)
button_restricao.place(x=1000, y=1000)

check_restricao = ctk.StringVar(value='N')

checkbox = ctk.CTkCheckBox(janela, text="Possuí arquivo de restrição?",command=inclui_restricao,
                                    variable=check_restricao, onvalue="S", offvalue="N",checkbox_height=20,checkbox_width=20)
checkbox.place(x=10, y=160)

titulo_tratar_sev= ctk.CTkLabel(janela,text='Tratar SEVs')
titulo_tratar_sev.place(x=10,y=230)

button_fase_1 = ctk.CTkButton(janela,text="Tratativa inicial",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=tratativa_inicial)
button_fase_1.place(x=10,y=250)

button_prox_acesso = ctk.CTkButton(janela,text="Proximo acesso",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=prox_acesso)
button_prox_acesso.place(x=120,y=250)

button_acesso_anterior = ctk.CTkButton(janela,text="Acesso anterior",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=acesso_anterior)
button_acesso_anterior.place(x=10,y=280)

button_precifica_sevs = ctk.CTkButton(janela,text="Precifica SEVs",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=precifica_sevs)
button_precifica_sevs.place(x=120,y=280)

button_backbone = ctk.CTkButton(janela,text="Roda Backbone",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=roda_bbip)
button_backbone.place(x=10,y=310)

button_finaliza_sevs = ctk.CTkButton(janela,text="Finalizar SEVs",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=finaliza_sevs)
button_finaliza_sevs.place(x=120,y=310)

# button_bbip = ctk.CTkButton(janela,text="Rodar BBIP",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=rodar_bbip)
# button_bbip.place(x=70,y=250)

# button_finalizar = ctk.CTkButton(janela,text="Gerar arquivos finais",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=gerar_arquivos_finais)
# button_finalizar.place(x=160,y=250)

janela.mainloop()
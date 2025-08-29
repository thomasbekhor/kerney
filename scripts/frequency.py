# ======================================================================================
# Cálculo de frequência de abastecimento (visitas/semana) por patrimônio (máquina)
# --------------------------------------------------------------------------------------
# Fluxo em alto nível:
# 1) Utilidades: fechar Excel aberto; padronizar códigos; ler configs (txt) -> dict
# 2) Função de "repasses": quando a frequência baseada em consumo excede muito a agenda,
#    divide patrimônio/parceiro em _A/_B para permitir janelas distintas no mesmo dia
# 3) main():
#    3.1) Confirmação do usuário (sobrescreve aba "Frequências" no Excel)
#    3.2) Resolução de diretórios/arquivos e abertura do workbook .xlsm
#    3.3) Leitura dos parâmetros na planilha e do arquivo de formato do consumo (txt)
#    3.4) Leitura e padronização de dados (parceiros, patrimônios, insumos, consumo)
#    3.5) Cálculo do consumo médio semanal por (FILIAL, PARCEIRO, PATRIMONIO, INSUMO)
#    3.6) Cálculo de FREQ_BASEADO_EM_CONSUMO (nível de reposição global ou por SKU)
#    3.7) (Opcional) Repasses e ajustes de janelas e depara_point_id
#    3.8) Definição da FREQUENCIA_REPOSICAO respeitando limites (dias/semana, atual)
#    3.9) Aplicação de FREQUENCIA_SEMANAL_MINIMA e flexibilidade (freq_flex)
#    3.10) (Opcional) Padronização intra-parceiro: mesma frequência para todos os ativos
#    3.11) Escrita de saídas (parceiros/patrimônios atualizados, parquet e Excel)
#    3.12) Estilização na aba "Frequências" e finalização
# ======================================================================================

import pandas as pd
import os
import time
from tqdm import tqdm
import xlwings as xw
import numpy as np
import win32com.client

def close_excel_file_if_open(filename):
    # Fecha um arquivo Excel específico se estiver aberto (usa automação COM via pywin32)
    # Evita erro de gravação ao sobrescrever planilhas/arquivos usados pelo Excel.
    """Check if an Excel file is open and close it using pywin32."""
    filename = os.path.basename(filename)  # Extract just the filename
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")  # Get running Excel
        for wb in excel.Workbooks:
            if wb.Name == filename:
                print(f"Closing {filename}...")
                wb.Close(SaveChanges=False)  # Close without saving
                time.sleep(1)  # Give Excel a moment
                return True
    except Exception:
        pass  # No running Excel instance found
    return False


def std_codes(code):
    # Padroniza códigos vindos de Excel/CSV: remove quebras de linha e _x000D_, cast numérico
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')


def read_txt_as_dict(file_path):
    # Lê um arquivo .txt no formato "chave: valor" e retorna dict com pares limpos.
    """
    Lê um arquivo .txt com linhas no formato `chave: valor` e retorna um dicionário.

    Args:
        file_path (str): Caminho para o arquivo .txt.

    Returns:
        dict: Dicionário com as chaves e valores do arquivo.
    """
    config_dict = {}
    with open(file_path, 'r') as file:
        for line in file:
            line = line.strip()
            if line:  # Ignorar linhas vazias
                key, value = line.split(":", 1)  # Separar chave e valor
                config_dict[key.strip()] = value.strip().strip('"')  # Remover espaços e aspas
                
    return config_dict


def repasses(visitas_df, patrimonios_df, parceiros_df, depara_point_id):
    # ----------------------------------------------------------------------------------
    # Divide patrimônios/parceiros em _A/_B quando a FREQ_BASEADO_EM_CONSUMO excede
    # significativamente o número de DIAS_POR_SEMANA (aqui usa fator 1.5 como exemplo).
    # Efeito:
    #   - Duplica patrimônio e sócios (_A e _B)
    #   - Ajusta FREQUENCIA_SEMANAL_MINIMA do patrimônio original (metade, arred. p/ cima)
    #   - Divide janelas de funcionamento do parceiro em dois blocos com sobreposição
    #   - Atualiza depara_point_id para refletir _A/_B
    # ----------------------------------------------------------------------------------
    ## Caso frequência baseada em consumo maior que # dias por semana, aplica repasse dividindo o patrimônio em dois
    index = (visitas_df['FREQ_BASEADO_EM_CONSUMO'] > visitas_df['DIAS_POR_SEMANA']*1.5) ##### 1.5 deve ser input do usuário
    patrimonios_repasse = list(visitas_df.loc[index, 'PATRIMONIO'].unique())
    parceiros_repasse = list(visitas_df.loc[index, 'PARCEIRO'].unique())
    horario_inicio_default = ''
    horario_fim_default = ''

    ## Visitas
    # Duplica patrimônios para repasse
    visitas_df_aux = visitas_df.loc[index].copy()
    # Adiciona sufixo _A e define o consumo igual aos dias por semana
    visitas_df.loc[index, 'PARCEIRO'] = visitas_df.loc[index, 'PARCEIRO'] + '_A'
    visitas_df.loc[index, 'PATRIMONIO'] = visitas_df.loc[index, 'PATRIMONIO'] + '_A'
    visitas_df.loc[index, 'FREQ_BASEADO_EM_CONSUMO'] = visitas_df.loc[index, 'DIAS_POR_SEMANA']
    # Adiciona sufixo _B e define consumo igual ao restante
    visitas_df_aux['PARCEIRO'] = visitas_df_aux['PARCEIRO'] + '_B'
    visitas_df_aux['PATRIMONIO'] = visitas_df_aux['PATRIMONIO'] + '_B'
    visitas_df_aux['FREQ_BASEADO_EM_CONSUMO'] = visitas_df_aux['FREQ_BASEADO_EM_CONSUMO'] - visitas_df_aux['DIAS_POR_SEMANA']
    visitas_df = pd.concat([visitas_df, visitas_df_aux]).reset_index(drop=True)

    # Patrimônios
    index_patrimonios = (patrimonios_df['PATRIMONIO'].isin(patrimonios_repasse))
    patrimonios_df.loc[index_patrimonios, 'FREQUENCIA_SEMANAL_MINIMA'] = np.ceil(patrimonios_df.loc[index_patrimonios, 'FREQUENCIA_SEMANAL_MINIMA']/2)

    patrimonios_df_aux = patrimonios_df.loc[index_patrimonios].copy()

    patrimonios_df.loc[index_patrimonios, 'PATRIMONIO'] = patrimonios_df.loc[index_patrimonios, 'PATRIMONIO'] + '_A'
    patrimonios_df.loc[index_patrimonios, 'PARCEIRO'] = patrimonios_df.loc[index_patrimonios, 'PARCEIRO'] + '_A'

    patrimonios_df_aux['PATRIMONIO'] = patrimonios_df_aux['PATRIMONIO'] + '_B'
    patrimonios_df_aux['PARCEIRO'] = patrimonios_df_aux['PARCEIRO'] + '_B'

    patrimonios_df = pd.concat([patrimonios_df, patrimonios_df_aux]).reset_index(drop=True)
    parceiros_final = list(patrimonios_df['PARCEIRO'].unique()) 

    # Parceiros
    # Ajusta janelas de funcionamento com uma sobreposição (overlap_fraction) para transição
    index = parceiros_df['PARCEIRO'].isin(parceiros_repasse)
    parceiros_df.loc[index, 'INICIO_FUNCIONAMENTO'] = np.datetime64('1900-01-01T06:00:00.000000000')
    parceiros_df.loc[index, 'FIM_FUNCIONAMENTO'] = np.datetime64('1900-01-01T16:45:00.000000000')

    parceiros_df['HORAS_FUNCIONAMENTO'] = (parceiros_df['FIM_FUNCIONAMENTO']-parceiros_df['INICIO_FUNCIONAMENTO']).dt.total_seconds()/3600
    parceiros_df_aux = parceiros_df[index].copy()
    parceiros_df_aux_og = parceiros_df[index].copy()

    overlap_fraction = 0.05
    horas_total = parceiros_df['HORAS_FUNCIONAMENTO']
    overlap_timedelta = pd.to_timedelta((horas_total * overlap_fraction).round(1), unit='h')
    split_point = pd.to_timedelta((horas_total * 0.5).round(0), unit='h') # Divide no meio para o exemplo

    # Ajuste de horário
    parceiros_df.loc[index, 'PARCEIRO'] = parceiros_df.loc[index, 'PARCEIRO'] + '_A'
    parceiros_df.loc[index, 'FIM_FUNCIONAMENTO'] = parceiros_df['INICIO_FUNCIONAMENTO'] + split_point + overlap_timedelta

    parceiros_df_aux_og.loc[index, 'FIM_FUNCIONAMENTO'] = parceiros_df_aux_og['INICIO_FUNCIONAMENTO'] + split_point + overlap_timedelta

    parceiros_df_aux['PARCEIRO'] = parceiros_df_aux['PARCEIRO'] + '_B'
    parceiros_df_aux['FIM_FUNCIONAMENTO'] = parceiros_df_aux['FIM_FUNCIONAMENTO']
    parceiros_df_aux['INICIO_FUNCIONAMENTO'] = parceiros_df_aux['FIM_FUNCIONAMENTO'] - split_point - overlap_timedelta
    
    parceiros_df = pd.concat([parceiros_df_aux_og, parceiros_df, parceiros_df_aux])
    parceiros_df = parceiros_df.loc[parceiros_df['PARCEIRO'].isin(parceiros_final)].reset_index(drop=True)
    parceiros_df.drop(columns='HORAS_FUNCIONAMENTO', inplace=True)

    # Depara point id (duplica _A/_B para manter coerência espacial do parceiro)
    index = depara_point_id['PARCEIRO'].isin(parceiros_repasse)
    depara_point_id_aux = depara_point_id[index].copy()
    depara_point_id_aux_2 = depara_point_id[index].copy()

    depara_point_id_aux['PARCEIRO'] = depara_point_id_aux['PARCEIRO'] + '_A'
    depara_point_id_aux_2['PARCEIRO'] = depara_point_id_aux_2['PARCEIRO'] + '_B'

    depara_point_id = pd.concat([depara_point_id, depara_point_id_aux, depara_point_id_aux_2])
    depara_point_id = depara_point_id.loc[depara_point_id['PARCEIRO'].isin(parceiros_final)].reset_index(drop=True)

    return visitas_df, patrimonios_df, parceiros_df, depara_point_id

def main(developer=False, visitas_duplas=False, nivel_reposicao=None,
         padronizacao_intra_patrimonio=False, freq_flex=None):
    # ----------------------------------------------------------------------------------
    # Parâmetros:
    #   developer: pula interação de confirmação
    #   visitas_duplas: ativa lógica de repasses (divisão _A/_B)
    #   nivel_reposicao: se fornecido, usa NÍVEL DE REPOSIÇÃO global (0..1) p/ todos SKUs
    #   padronizacao_intra_patrimonio: força mesma frequência entre patrimônios do mesmo parceiro
    #   freq_flex: flexibilidade (+/-) em relação à frequência atual mínima por patrimônio
    # Saídas: escreve resultados na aba "Frequências" e arquivos auxiliares em "Dados Intermediários"
    # ----------------------------------------------------------------------------------
    
    if developer:
        ans='S'
    else:
        print('Essa Operação Sobrescreverá os dados da aba "Frequências", deseja continuar? (s/n)')
        ans = str(input())
    
    if ans.upper() != 'S':
        # Cancelamento controlado pelo usuário
        print('Calculo de Frequência Cancelado')
        time.sleep(0.5)
        return 0

    start = time.time()
    print("Iniciando Cálculo de Frequência")

    section_start = time.time()
    print("Lendo dados de Input e Configurações...")

    # 1) Resolução de diretórios e localização do workbook .xlsm
    current_script_directory = os.path.dirname(os.path.abspath(__file__))
    main_dir = '/'.join(current_script_directory.split('\\')[:-1]) + '/'
    model_data_folder = main_dir + 'Dados Intermediários/'
    input_folder = main_dir + 'Dados de Input/'

    # Descobrir arquivo .xlsm na pasta raiz do projeto (main_dir)
    for file in os.listdir(main_dir):
        if len(file) > 4:
            if file[-5:] == '.xlsm':
                workbook = file
                
    # Garante que o Excel não está com o arquivo aberto (evita lock ao gravar)
    close_excel_file_if_open(main_dir + workbook)
    wb = xw.Book(main_dir + workbook)
    config_sheet = wb.sheets['Configurações']

    # 2) Ler parâmetros de configuração da planilha
    dias_semana = int(config_sheet['D7'].value)
    input_mapping = {"Sim": True, "Não": False}
    visitas_duplas = input_mapping[config_sheet['J10'].value]

    # 3) Ler arquivo de configuração de formato do CSV de consumo
    read_config = read_txt_as_dict(input_folder + 'formato_consumo.txt')

    # 4) Importar dados de entrada (com padronização de códigos)
    parceiros_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='parceiros')
    parceiros_df['PARCEIRO'] = parceiros_df['PARCEIRO'].apply(std_codes)
    parceiros_df['INICIO_FUNCIONAMENTO'] = pd.to_datetime(parceiros_df['INICIO_FUNCIONAMENTO'], format='%H:%M:%S')
    parceiros_df['FIM_FUNCIONAMENTO'] = pd.to_datetime(parceiros_df['FIM_FUNCIONAMENTO'], format='%H:%M:%S')

    patrimonios_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='patrimonios')
    patrimonios_df['PATRIMONIO'] = patrimonios_df['PATRIMONIO'].apply(std_codes)
    patrimonios_df['PARCEIRO'] = patrimonios_df['PARCEIRO'].apply(std_codes)

    insumos_df = pd.read_excel(input_folder + 'Dados.xlsx', sheet_name='insumos')
    insumos_df['PATRIMONIO'] = insumos_df['PATRIMONIO'].apply(std_codes)
    insumos_df['PARCEIRO'] = insumos_df['PARCEIRO'].apply(std_codes)
    insumos_df['INSUMO'] = insumos_df['INSUMO'].apply(std_codes)

    consumo_df = pd.read_csv(input_folder + 'Dados de Consumo.csv', encoding=read_config['encoding'], sep=read_config['sep'], decimal=read_config['decimal']).dropna()
    consumo_df['PATRIMONIO'] = consumo_df['PATRIMONIO'].apply(std_codes)
    consumo_df['PARCEIRO'] = consumo_df['PARCEIRO'].apply(std_codes)
    consumo_df['INSUMO'] = consumo_df['INSUMO'].apply(std_codes)

    # De-para geográfico gerado na etapa de matriz de distâncias
    depara_point_id = pd.read_parquet(model_data_folder + 'depara_point_id.parquet')[['FILIAL', 'PARCEIRO', 'POINT_ID', 'LAT', 'LON']]
    depara_point_id['PARCEIRO'] = depara_point_id['PARCEIRO'].apply(std_codes)

    print('Leitura de dados encerrada, tempo: {:.1f}s'.format(time.time() - section_start))
    section_start = time.time()
    print('Iniciando cálculo de frequência...')

    # 5) Preparação do consumo médio semanal por patrimônio e insumo
    consumo_med_df = consumo_df.copy()
    for col in ['INICIO', 'FIM']:
        # Se datas vierem como string, converter dd/mm/YYYY
        if str(consumo_med_df[col].dtypes) == 'object':
            consumo_med_df[col] = pd.to_datetime(consumo_med_df[col], format='%d/%m/%Y')
    consumo_med_df['INICIO'] = consumo_med_df['INICIO'].astype('datetime64[ns]')
    consumo_med_df['FIM'] = consumo_med_df['FIM'].astype('datetime64[ns]')
    consumo_med_df['DIAS_TOTAL'] = np.maximum(1, (consumo_med_df['FIM'] - consumo_med_df['INICIO']).dt.days)
    consumo_med_df['SEMANAS'] = consumo_med_df['DIAS_TOTAL']/7
    consumo_med_df = consumo_med_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'INSUMO']).agg({'CONSUMO':'sum', 'SEMANAS':'sum'}).reset_index()
    consumo_med_df['CONSUMO_SEMANAL'] = consumo_med_df['CONSUMO']/consumo_med_df['SEMANAS']
    consumo_med_df = (consumo_med_df
                    .merge(insumos_df,
                            on=['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'INSUMO'],
                            how='left')
                    .merge(patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'DIAS_POR_SEMANA', 'FREQUENCIA_ATUAL']],
                            on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                            how='outer')
                    .loc[lambda x: x['DIAS_POR_SEMANA'].notna()]
                    .reset_index(drop=True))
    
    # 6) Cálculo base por item (cada linha é um SKU/insumo do patrimônio)
    visitas_df = consumo_med_df.copy()

    print(consumo_med_df.loc[consumo_med_df["PATRIMONIO"] == "15134"])

    # Valores default para preenchimento de faltas (observa-se 'INSUMO ' com espaço à direita)
    fill_values = {'INSUMO': "COPO", 'CONSUMO': 0, 'SEMANAS':1 ,'CAPACIDADE':1, 'MAQUINAS':"-", 'NIVEL_REPOSICAO':0 , 'CONSUMO_SEMANAL': 0} # Aplicando fillna com dicionário
    visitas_df = visitas_df[visitas_df["CAPACIDADE"].notna()]
    visitas_df = visitas_df.fillna(fill_values)
 
    # 7) Dois modos de calcular FREQ_BASEADO_EM_CONSUMO:
    #    a) Se nivel_reposicao (global) vier: usa esse valor para todos os itens
    #    b) Caso contrário: usa NIVEL_REPOSICAO por linha (insumo)
    if nivel_reposicao is not None:
        # a) Usando nível de reposição global (0..1)
        visitas_df['FREQ_BASEADO_EM_CONSUMO'] = np.ceil(visitas_df['CONSUMO_SEMANAL']/(visitas_df['CAPACIDADE']*(1 - nivel_reposicao)))
        # Consolidar por patrimônio: usar a MAIOR frequência entre insumos (max)
        visitas_df = visitas_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO']).agg({'FREQ_BASEADO_EM_CONSUMO':'max'}).reset_index()
        # Trazer FREQUENCIA_SEMANAL_MINIMA para aplicar piso
        visitas_df = (patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'FREQUENCIA_SEMANAL_MINIMA']]
            .merge(
                visitas_df,
                on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                how='left'))
        # Preencher faltas com o mínimo semanal
        visitas_df['FREQ_BASEADO_EM_CONSUMO'] = (
            visitas_df['FREQ_BASEADO_EM_CONSUMO']
            .fillna(visitas_df['FREQUENCIA_SEMANAL_MINIMA']))
        wb.close()
        return visitas_df
    else:
        # b) Usando NIVEL_REPOSICAO específico de cada linha
        visitas_df['FREQ_BASEADO_EM_CONSUMO'] = np.ceil(visitas_df['CONSUMO_SEMANAL']/(visitas_df['CAPACIDADE']*(1 - visitas_df['NIVEL_REPOSICAO'])))
    

    # 8) (Opcional) Repasses: divide patrimônio e janelas do parceiro (_A/_B) quando habilitado
    if visitas_duplas:
        visitas_df, patrimonios_df, parceiros_df, depara_point_id = repasses(visitas_df, patrimonios_df, parceiros_df, depara_point_id)

    # 9) FREQUENCIA_REPOSICAO é limitada por DIAS_POR_SEMANA e pela FREQUENCIA_ATUAL
    #    (evita propor mais visitas que a janela semanal e respeita estado atual)
    visitas_df['FREQUENCIA_REPOSICAO'] = np.minimum(visitas_df['FREQ_BASEADO_EM_CONSUMO'],
                                                    np.minimum(visitas_df['DIAS_POR_SEMANA'],
                                                               visitas_df['FREQUENCIA_ATUAL'])).astype(int)

    # Consolidar por patrimônio pegando o máximo entre insumos (linha -> patrimônio)
    visitas_df = visitas_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO']).agg({'FREQUENCIA_REPOSICAO':'max'}).reset_index()

    # 10) Aplicar FREQUENCIA_SEMANAL_MINIMA (piso regulatório/operacional por patrimônio)
    visitas_df = (patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'FREQUENCIA_SEMANAL_MINIMA', 'FREQUENCIA_ATUAL']]
                .merge(
                    visitas_df,
                    on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                    how='left')
                .fillna(0)
                )
    
    # 11) Se houver flexibilidade (freq_flex), ajusta o piso mínimo com base na atual
    if freq_flex is not None:
        visitas_df['FREQUENCIA_SEMANAL_MINIMA'] = np.maximum(visitas_df['FREQUENCIA_SEMANAL_MINIMA'], (visitas_df['FREQUENCIA_ATUAL']-freq_flex))
    
    # 12) FREQUENCIA final = max(piso mínimo, frequência de reposição calculada)
    visitas_df['FREQUENCIA'] = np.maximum(visitas_df['FREQUENCIA_SEMANAL_MINIMA'], visitas_df['FREQUENCIA_REPOSICAO'])
    
    # 13) (Opcional) Padronização intra-parceiro: força mesma frequência para todos patrimônios
    if padronizacao_intra_patrimonio:
        visitas_df_aux = (
            visitas_df.groupby('PARCEIRO')['FREQUENCIA'].max()
            .reset_index().rename(columns={'FREQUENCIA':'FREQUENCIA_PARC'}))
        visitas_df = (visitas_df.merge(visitas_df_aux, on='PARCEIRO', how='left')
                      .assign(FREQUENCIA= lambda x: x['FREQUENCIA_PARC'])
                      .drop(columns='FREQUENCIA_PARC'))
        visitas_df["FREQUENCIA"]=visitas_df["FREQUENCIA"].astype(int)
    
    # 14) Renomear colunas para nomenclatura de apresentação (mantém subset de interesse)
    renames = {'FILIAL':'FILIAL', 'PARCEIRO': 'PARCEIRO', 'PATRIMONIO': 'PATRIMONIO', 
               'FREQUENCIA_ATUAL': 'Frequência Atual (visitas/semana)', 'FREQUENCIA_SEMANAL_MINIMA': 'Frequência Mínima (visitas/semana)',
               'FREQUENCIA_REPOSICAO': 'Frequência de Reposição (visitas/semana)', 'FREQUENCIA': 'Frequência (visitas/semana)'}
    visitas_df = visitas_df[renames.keys()].rename(columns=renames)

    # 15) Converter colunas de horário para time puro (remove data 1900-01-01)
    parceiros_df['INICIO_FUNCIONAMENTO'] = (parceiros_df['INICIO_FUNCIONAMENTO'].dt.time) #.astype(str)
    parceiros_df['FIM_FUNCIONAMENTO'] = (parceiros_df['FIM_FUNCIONAMENTO'].dt.time) #.astype(str)

    # 16) Persistir ajustes intermediários:
    #     - Grava parceiros/patrimônios atualizados em "Dados Intermediários/Dados.xlsx"
    #     - Atualiza depara_point_id (caso repasses)
    with pd.ExcelWriter(model_data_folder+'Dados.xlsx', engine="xlsxwriter") as writer:
        parceiros_df.to_excel(writer, sheet_name="parceiros", index=False)
        patrimonios_df.to_excel(writer, sheet_name="patrimonios", index=False)

    depara_point_id.to_parquet(model_data_folder+'depara_point_id_atualizado.parquet', index=False)

    print('Cálculo de frequência encerrado, tempo: {:.1f}s'.format(time.time() - section_start))
    section_start = time.time()
    print('Salvando dados e escrevendo resultados na aba "Frequências"...')

    # 17) Escrever resultado na aba "Frequências" do workbook e aplicar formatação
    sheet = wb.sheets['Frequências']  # Seleciona a aba específica

    # Transferir o DataFrame para o Excel
    sheet.range("7:1048576").clear_contents()
    sheet['B6'].options(index=False).value = visitas_df  # Transfere o DataFrame com o cabeçalho

    # Selecionar o intervalo de células a ser formatado (borda e estilo)
    cells = sheet.range('B7:H{}'.format(len(visitas_df) + 6))

    # Bordas (usa API COM do Excel)
    borders = cells.api.Borders

    # Bordas laterais e horizontais (cores em BGR)
    borders(11).Weight = 4  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco

    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xD2D2D2  # Preto
    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xD2D2D2  # Preto
    borders(9).Weight = 4  # xlEdgeBottom, fina
    borders(9).Color = 0xD2D2D2  # Preto

    # Destaque visual na coluna H (Frequência final)
    cells = sheet.range('H7:H{}'.format(len(visitas_df) + 6))
    cells.api.Font.Bold = True                # Negrito
    cells.api.Font.Italic = True
    cells.api.Font.Size = 12                  # Tamanho da fonte
    cells.api.Interior.Color = 0xF4EDFC

    # 18) Persistir série temporal agregada (consumo médio) para reuso posterior
    consumo_med_df.to_parquet(model_data_folder + 'consumo_medio.parquet', index=False)

    print('Dados salvos com sucesso, tempo: {:.1f}s'.format(time.time() - section_start))

    # 19) Encerramento
    print('Cálculo de Frequências Encerrado, tempo total: {:.1f}s\nPrecione Enter para continuar...'.format(time.time() - start))
    
    # Fecha workbook ou aguarda ENTER conforme modo
    if developer:
        wb.close()
    else:
        input()

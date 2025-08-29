# ======================================================================================
# Cronograma de visitas (distribuição semanal)
# --------------------------------------------------------------------------------------
# Objetivo:
#   Dado o número de visitas/semana por patrimônio, escolher em quais dias da semana
#   cada patrimônio será visitado, de forma a:
#     • Distribuir visitas de modo homogêneo ao longo da semana
#     • Minimizar o número MÁXIMO de visitas em um mesmo dia (minimiza "picos")
#
# Abordagem:
#   1) Ler frequências (visitas/semana) e parâmetros (número de dias/semana)
#   2) Gerar padrões de visita possíveis para cada frequência (ex.: 2x/sem ->
#      padrões espaçados em [0..n_p-1], com rotações para cobrir simetrias)
#   3) Modelar um problema de otimização (MIP) com OR-Tools:
#        - Variável x[i,p] = 1 se o patrimônio i adota o padrão p
#        - Variável max_w   = volume máximo de visitas em um dia
#        - Restrição: cada patrimônio escolhe exatamente 1 padrão
#        - Restrição: para todo dia t, soma de visitas do dia ≤ max_w
#        - Objetivo: minimizar max_w (uniformidade)
#   4) Salvar o cronograma (marcação por dia) na aba "Cronograma" (Excel)
#
# Observações:
#   • Sábados são bloqueados para certos casos via 'ALLOW_SATURDAY'
#   • A função generate_visit_patterns cria padrões equiespaçados e suas rotações
# ======================================================================================

import pandas as pd
import os
import math
import xlwings as xw
import time
from ortools.linear_solver import pywraplp
from tqdm import tqdm
import win32com.client

def close_excel_file_if_open(filename):
    # Fecha o arquivo do Excel se estiver aberto (evita lock ao sobrescrever). Usa pywin32/COM.
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
    # Padroniza códigos vindos de planilhas/CSV: remove quebras e _x000D_; se numérico, cast p/ int string.
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')
    
def generate_visit_patterns(dias_semana): #Dos 5 dias, cria padroes de frequencia para cada quantidade de dias
    # ----------------------------------------------------------------------------------
    # Gera padrões equiespaçados de visitas para cada frequência f ∈ {1..n_p},
    # e depois gera todas as rotações (shifts) possíveis desses padrões.
    # Ex.: n_p=5, f=2 -> base ~ [0, 2] e rotações: [0,2], [1,3], [2,4], [3,0], [4,1]
    # Retorna dict: {f: [tuplas_de_dias_ordenadas]}
    # ----------------------------------------------------------------------------------
    n_p = dias_semana
    visit_pattern = {}

    for f in range(1, n_p + 1):
        # Espaçamento fracionário para distribuir f visitas em n_p dias
        step = n_p / f
        pattern = []
        current = 0
        
        for _ in range(f):
            pattern.append(round(current) % n_p)
            current += step
        
        visit_pattern[f] = sorted(pattern)

    # Constrói todas as rotações únicas (usa set para evitar duplicatas)
    visit_patterns = {f: [] for f in range(1, n_p + 1)}

    for f in range(1, n_p + 1):
        unique_patterns = set()

        for s in range(n_p):
            shifted_pattern = tuple(sorted((t + s) % n_p for t in visit_pattern[f]))
            unique_patterns.add(shifted_pattern)

        visit_patterns[f] = sorted(unique_patterns)

    return visit_patterns

def main(developer=False):
    # ----------------------------------------------------------------------------------
    # Parâmetro:
    #   developer: pula confirmação interativa (mantém comportamento automatizado)
    # Efeitos:
    #   • Lê dados de frequências e configurações do workbook .xlsm
    #   • Constrói e resolve MIP (OR-Tools SCIP) para uniformizar as visitas
    #   • Escreve cronograma formatado na aba "Cronograma"
    # ----------------------------------------------------------------------------------
        
    if developer:
        ans='S'
    else:
        print('Essa operação sobrescreverá os dados da aba "Cronograma", deseja continuar? (s/n)')
        ans = str(input())
    
    if ans.upper() != 'S':
        # Saída amigável se o usuário cancelar
        print('Construção do cronograma de visitas cancelada')
        time.sleep(0.5)
        return 0
    
    start = time.time()
    print("Iniciando construção do cronograma de visitas")
    section_start = time.time()
    print("Lendo dados de frequências de visitas...")

    # 1) Diretórios e caminhos base do projeto
    current_script_directory = os.path.dirname(os.path.abspath(__file__)) # os.getcwd() # 
    main_dir = '/'.join(current_script_directory.split('\\')[:-1]) + '/'
    input_folder = main_dir + 'Dados de Input/'
    model_data_folder = main_dir + 'Dados Intermediários/'

    # 2) Localiza o workbook .xlsm na raiz do projeto e garante que não está aberto
    for file in os.listdir(main_dir):
        if len(file) > 4:
            if file[-5:] == '.xlsm':
                workbook = file
                
    close_excel_file_if_open(main_dir + workbook)
    wb = xw.Book(main_dir + workbook)

    # 3) Abre planilhas de configuração e de frequências
    config_sheet = wb.sheets['Configurações']
    frequency_sheet = wb.sheets['Frequências']

    # 4) Carrega dados auxiliares de patrimônios (dias por semana)
    patrimonios_df = pd.read_excel(model_data_folder + 'Dados.xlsx', sheet_name='patrimonios')
    patrimonios_df['PATRIMONIO'] = patrimonios_df['PATRIMONIO'].apply(std_codes)

    # 5) Parâmetro: número de dias de operação da semana (ex.: 5 ou 6)
    dias_semana = int(config_sheet['D7'].value)
    n_p = dias_semana

    # 6) Ler frequências calculadas (aba "Frequências")
    freq_df = frequency_sheet.range("B6").expand().value
    freq_df = pd.DataFrame(freq_df[1:], columns=freq_df[0])[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'Frequência (visitas/semana)']].rename(columns={'Frequência (visitas/semana)':'FREQUENCIA'})

    # Padronizações e tipos
    freq_df['FILIAL'] = freq_df['FILIAL'].astype(str)
    freq_df['PARCEIRO'] = freq_df['PARCEIRO'].apply(std_codes)
    freq_df['PATRIMONIO'] = freq_df['PATRIMONIO'].apply(std_codes)
    #freq_df['FREQUENCIA'] = freq_df['FREQUENCIA'].astype(int)

    # Trazer DIAS_POR_SEMANA e flag de permissão de sábado (ALLOW_SATURDAY)
    freq_df = freq_df.merge(patrimonios_df[['PATRIMONIO', 'DIAS_POR_SEMANA']],
                            on='PATRIMONIO',
                            how='left')

    # Permite sábado somente se (6x/semana E DIAS_POR_SEMANA=6) E não for suf. "_B"
    freq_df['ALLOW_SATURDAY'] = (((freq_df['DIAS_POR_SEMANA'] == 6) &
                                  (freq_df['FREQUENCIA'] == 6)) &
                                  (~freq_df['PATRIMONIO'].str.contains('_B')))

    print('Leitura de dados concluída, tempo: {:.1f}s'.format(time.time() - section_start))
    section_start = time.time()
    print('Iniciando otimizador para distribuição de visitas nos períodos de planejamento...')

    # 7) Gerar padrões de visita para todos os valores de DIAS_POR_SEMANA existentes
    n_dias = list(freq_df['DIAS_POR_SEMANA'].unique())
    merged_visit_patterns = {}

    for n_dia in n_dias:
        visit_patterns = generate_visit_patterns(n_dia)
        
        for f, patterns in visit_patterns.items():
            if f not in merged_visit_patterns:
                merged_visit_patterns[f] = set()  # Use set to avoid duplicates

            merged_visit_patterns[f].update(patterns)

    # Converter sets em listas ordenadas
    visit_patterns = {f: sorted(merged_visit_patterns[f]) for f in merged_visit_patterns}

    # (Mantidos comentários/originais abaixo como referência histórica do cálculo de padrões)
    # visit_pattern = {}
    # for f in range(1, n_p + 1):
    #     min_interval = n_p//f
    #     max_interval = math.ceil(n_p/f)
    #     max_interval_jumps = n_p%f
    #     pattern = []
    #     for p in range(0, max_interval_jumps*max_interval, max_interval):
    #         pattern.append(p)
    #     for p in range(max_interval_jumps*max_interval, n_p, min_interval):
    #         pattern.append(p)
    #     visit_pattern[f] = pattern

    # visit_patterns = {f:[] for f in range(1, n_p + 1)}
    # for f in range(1, n_p + 1):
    #     for s in range(0, n_p):
    #         pattern = [(t+s)%n_p for t in visit_pattern[f]]
    #         pattern.sort()
    #         pattern = tuple(pattern)
    #         visit_patterns[f].append(pattern)
    #     visit_patterns[f] = list(set(visit_patterns[f]))

    # 8) Resolver por filial para permitir paralelismo futuro e logs claros
    cronograma_df = pd.DataFrame()
    filiais = freq_df['FILIAL'].unique().tolist()

    for filial in filiais:
        filial_start = time.time()
        print(f'Obtendo Cronograma para a Filial {filial}')

        # Filtra dados da filial
        index = freq_df['FILIAL'] == filial

        patrimonios = freq_df[index]['PATRIMONIO'].unique().tolist()
        patr_freq = freq_df[index].set_index('PATRIMONIO')['FREQUENCIA'].to_dict()
        patr_parc = freq_df[index].set_index('PATRIMONIO')['PARCEIRO'].to_dict()

        # 8.1) Opções de padrão por patrimônio (todos os padrões de sua frequência)
        group_pattern = {
            i: {
                p:dias
                for p, dias in enumerate(visit_patterns[patr_freq[i]])}
            for i in patrimonios}
        
        # 8.2) Remover padrões com sábado quando não permitido (ALLOW_SATURDAY == False)
        group_pattern = {
            i: {
                p: dias 
                for p, dias in group_pattern[i].items()
                if (freq_df.loc[freq_df['PATRIMONIO'] == i, 'ALLOW_SATURDAY'].values[0] | (5 not in dias))
            }
            for i in patrimonios} 
        
        # 8.3) Reenumerar padrões (assegura chaves 0..k-1 por patrimônio após filtragem)
        group_pattern = {i:{p:dias for p, dias in enumerate(group_pattern[i].values())} for i in patrimonios}

        # Conjunto de dias [0..n_p-1]
        dias = [t for t in range(0, n_p)]

        # 9) Construção do solver MIP (SCIP)
        solver = pywraplp.Solver.CreateSolver('SCIP')

        # 9.1) Variáveis de decisão
        # x[i][p] = 1 se patrimônio i escolhe o padrão p; 0 caso contrário
        x = {i:{} for i in patrimonios}
        for i in patrimonios:
            for p in group_pattern[i].keys():
                x[i][p] = solver.IntVar(0, 1, f'x[{i},{p}]')

        # max_w: variável real que representa a carga máxima (visitas) em um dia
        max_w = solver.NumVar(0, solver.infinity(), f'max_w')

        # 9.2) Função objetivo: minimizar max_w (uniformizar o cronograma)
        obj = solver.Objective()
        obj.SetCoefficient(max_w, 1)
        obj.SetMinimization()

        # 9.3) Restrições
        # (a) Cada patrimônio deve adotar exatamente UM padrão
        for i in patrimonios:
            solver.Add(sum(x[i][p] for p in group_pattern[i].keys()) == 1)

        # (b) Definição de max_w: para cada dia t, a soma de visitas do dia ≤ max_w
        for t in dias: # deixar uniforme a distirbuição de dias
            solver.Add(max_w >= sum(sum(x[i][p] for p, days in group_pattern[i].items() if t in days) for i in patrimonios))

        # 9.4) Resolver MIP (tempo limite e gap)
        solver.SetTimeLimit(180*1000)
        solver.SetSolverSpecificParametersAsString("mip_gap = 0.01")  # Gap de 1%
        #solver.EnableOutput()
        status = solver.Solve()
        
        # 9.5) Extrair solução escolhida e montar linhas do cronograma
        result = {'FILIAL':[], 'PARCEIRO':[],'PATRIMONIO':[], 'FREQUENCIA':[], 'DIA':[]}
        for i in patrimonios:
            for p, days in group_pattern[i].items():
                if x[i][p].solution_value() == 1:
                    for day in days:
                        result['FILIAL'].append(filial)
                        result['PARCEIRO'].append(patr_parc[i])
                        result['PATRIMONIO'].append(i)
                        result['FREQUENCIA'].append(patr_freq[i])
                        result['DIA'].append(day)

        cronograma_df = pd.concat([cronograma_df, pd.DataFrame(result)])

        print(f'Cronograma obtido para a filial {filial}, tempo: {time.time() - filial_start:.1f}s')
    
    print('Cronogramas obtidos para todas as filiais, tempo: {:.1f}s'.format(time.time() - section_start))
    section_start = time.time()
    print('Salvando cronogramas obtidos na aba "Cronogramas"...')

    # 10) Pivotar em colunas 1..n_p marcando 'X' nos dias escolhidos
    for p in range(1, n_p+1):
        cronograma_df[f'{p}'] = (cronograma_df['DIA'] == (p - 1))
    cronograma_df = cronograma_df.groupby(['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'FREQUENCIA']).agg({f'{p}':'sum' for p in range(1, n_p+1)}).reset_index()
    for p in range(1, n_p+1):
        cronograma_df[f'{p}'] = cronograma_df[f'{p}'].astype(str).replace('0', '').replace('1', 'X')
    cronograma_df = cronograma_df.rename(columns={'FREQUENCIA':'Frequência (visitas/semana)'})

    # 11) Escrita e formatação na aba "Cronograma"
    sheet = wb.sheets['Cronograma']  # Seleciona a aba específica
    # Limpa conteúdo e formatos
    sheet.range("6:1048576").clear_contents()
    sheet.range("F7:XFD1048576").api.ClearFormats()
    # Escreve a tabela
    sheet['B7'].options(index=False).value = cronograma_df 

    # Cabeçalho "Dia" (celular F6)
    sheet.cells(6, 6).value = 'Dia'

    rows = len(cronograma_df)

    # Formatação: intervalo principal
    cells = sheet.range((7, 2), (7 + rows, 5 + n_p))
    cells.api.Font.Name = "Arial Narrow"

    # Bordas (API COM) — atenção: cores em BGR
    borders = cells.api.Borders

    # Bordas verticais (grossas e brancas)
    borders(11).Weight = 4  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco

    # Bordas horizontais (finas e cinza)
    borders(12).Weight = 2  # xlEdgeTop, fina
    borders(12).Color = 0xD2D2D2  # Preto
    borders(8).Weight = 2  # xlEdgeTop, fina
    borders(8).Color = 0xD2D2D2  # Preto
    borders(9).Weight = 4  # xlEdgeBottom, fina
    borders(9).Color = 0xD2D2D2  # Preto

    # Formatar cabeçalhos dos dias
    cells = sheet.range((7, 6), (8, 5 + n_p))
    borders = cells.api.Borders
    borders(11).Weight = 4  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco

    # Formatar corpo do cronograma (dias com 'X')
    cells = sheet.range((8, 6), (7 + rows, 5 + n_p))
    cells.api.Font.Bold = True
    cells.api.Font.Name = "Arial Narrow"
    cells.api.Interior.Color = 0xF4EDFC
    cells.api.HorizontalAlignment = -4108     # Alinhamento horizontal (centro)
    cells.api.VerticalAlignment = -4108       # Alinhamento vertical (centro)
        
    
    print('Cronogramas salvos com sucesso, tempo: {:.1f}s'.format(time.time() - section_start))

    print('Obtenção de Cronogramas Encerrada, tempo total: {:.1f}s\nPrecione Enter para continuar...'.format(time.time() - start))
   
    # 12) Encerramento (fecha workbook automaticamente no modo developer)
    if developer:
        wb.close()
    else:
        input()

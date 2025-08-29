# ======================================================================================
# Geração de rotas diárias e alocação de abastecedores (PARTE 1/2)
# --------------------------------------------------------------------------------------
# Objetivo geral:
#   • Montar “livros” (rotas) de visitas por dia/filial a partir de um cronograma
#     semanal e de uma matriz de distâncias/tempos, respeitando janelas, limites
#     diários de tempo/distância e capacidade semanal.
#   • Alocar abastecedores às rotas (a "BASE" é um nó fantasma equidistante usado
#     como origem/destino lógico; no modelo, custo/tempo a partir da BASE é controlado).
#
# Abordagem (alta visão):
#   1) Leitura de insumos (cronograma, frequências, matriz de distâncias por modal,
#      parceiros, patrimônios, parâmetros de configuração)
#   2) Pré-processamento: janelas de tempo, tempos de serviço/entrada, agrupamento de
#      patrimônios em “grupos” (GROUP) dentro de um mesmo parceiro/periodo respeitando
#      orçamento de tempo
#   3) Para cada (filial, supervisor, dia): montar um VRP com OR-Tools
#      (tempo com janelas + distância com limite + demanda/capacidade), resolvendo a
#      roteirização de grupos (cada GROUP é um “nó cliente”)
#   4) (PARTE 2) Pós-processar solução em visitas ordenadas por livro, ajustar tempos
#      de entrada/deslocamento, decidir modal, consolidar e escrever saídas
# ======================================================================================

import pandas as pd
import numpy as np
import os
import xlwings as xw
import time
from ortools.constraint_solver import pywrapcp
from ortools.constraint_solver import routing_enums_pb2
from tqdm import tqdm
import sys
from collections import defaultdict
import win32com.client

def close_excel_file_if_open(filename):
    # Fecha um arquivo do Excel se estiver aberto (via COM/pywin32), evitando lock na escrita
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



class Lote(): # lote = um problema VRP para (filial, dia, supervisor)
    def __init__(self):
        # -------------------------
        # Conjuntos
        self.clients = []     # nós clientes (grupos “GROUP”)
        self.vertices = []    # nós do grafo (BASE + clients)

        # -------------------------
        # Parâmetros globais
        self.cap = 108000     # capacidade (segundos/semana) inicial padrão
        self.route_cost = 10000  # penalidade fixa ao sair da BASE (estimula menos veículos/rotas)
        self.base = 'BASE'    # nó fantasma de origem/destino
        self.max_time = 7     # tempo máximo por rota (segundos) — ajustado depois por filial/dia
        self.max_dist = 15000 # distância máxima por rota (mesma lógica do tempo)

        # -------------------------
        # Parâmetros por nó
        self.demand = {}       # demanda do nó (aqui pode representar 1 por grupo)
        self.tw_start = {}     # janela de tempo: início (segundos relativos)
        self.tw_end = {}       # janela de tempo: fim (segundos relativos)
        self.service_time = {} # tempo de serviço no nó (segundos)

        # -------------------------
        # Parâmetros por arco
        self.distance = {}     # distância entre nós (m)
        self.time = {}         # tempo de deslocamento entre nós (s)

        # -------------------------
        # Controles auxiliares
        self.infeasible_clients = []  # clientes inviáveis por tempo mínimo base->i->base
        self.same_vehicle_groups = [] # grupos que devem estar no mesmo veículo (se usado)
    

    def solve(self, time_limit=600, verbose=False): # resolve o VRP com os parâmetros acima

        if verbose:
            print('Iniciando Parametrização')

        # =====================
        # 0. Parametrização
        # =====================

        # 0.1. Conjuntos-base
        clients = self.clients
        nodes = [self.base] + clients  # manager/roteamento usam índices desses nodes

        # 0.2. Subconjuntos de distância/tempo só com nós ativos (inclui BASE)
        dist = {i:{j:int(v) for j,v in v_i.items() if j in nodes} for i,v_i in self.distance.items() if i in nodes}
        t = {i:{j:int(v) for j,v in v_i.items() if j in nodes} for i,v_i in self.time.items() if i in nodes}

        # 0.3. Janelas de tempo relativas ao início da BASE
        #     e: earliest; l: latest — ambos normalizados subtraindo tw_start[BASE]
        e = {i: int(max(v - self.tw_start[self.base], 0))  for i,v in self.tw_start.items() if i in nodes}
        l = {i: int((v - self.tw_start[self.base])) for i,v in self.tw_end.items() if i in nodes}

        # 0.4. Demandas (BASE com 0)
        d = {i:v for i,v in self.demand.items() if i in clients} | {self.base:0}

        # 0.5. Serviço: tempo alocado no nó i; é somado ao “leg” de saída (i->j)
        s = {i:int(v) for i,v in self.service_time.items() if i in clients} | {self.base:0}

        # =====================
        # 1. Estrutura de dados no formato do OR-Tools
        # =====================
        data = {}

        # 1.1. Matriz de custos por distância
        #      ATENÇÃO: adiciona route_cost ao sair da BASE para punir “abrir” uma rota
        data['distance_matrix'] = [
            [
                dist[i][j] + (i == self.base)*self.route_cost
                for j in nodes
            ]
            for i in nodes
        ]

        # 1.2. Matriz de tempos (inclui tempo de serviço no nó de origem i)
        data['time_matrix'] = [
            [
                t[i][j] + s[i]
                for j in nodes
            ]
            for i in nodes
        ]

        # 1.3. Demandas por nó
        data['demands'] = [d[i] for i in nodes]

        # 1.4. Capacidade de veículo (homogênea)
        data['vehicle_capacity'] = self.cap

        # 1.5. Janelas de tempo (par, par)
        data['time_windows'] = [(e[i], l[i]) for i in nodes]

        # 1.6. Número de veículos: = #nodes (permite múltiplas rotas se necessário)
        data['num_vehicles'] = len(nodes)

        # 1.7. Depósito (índice do nó "BASE")
        data['depot'] = 0

        # =====================
        # 2. Criação do modelo
        # =====================

        # 2.1. Manager de índices
        manager = pywrapcp.RoutingIndexManager(len(data['distance_matrix']),
                                            data['num_vehicles'], data['depot'])

        # 2.2. Modelo de roteamento
        routing = pywrapcp.RoutingModel(manager)

        # 2.3. (Opcional) Restrições de "mesmo veículo" para grupos específicos
        if self.same_vehicle_groups != []:   
            group_representatives = []
            for group in self.same_vehicle_groups:
                indices = [manager.NodeToIndex(nodes.index(node)) for node in group if node in nodes]
                group_representatives.append(indices[0])
                print(indices)
                for i in range(len(indices) - 1):
                    routing.solver().Add(routing.VehicleVar(indices[i]) == routing.VehicleVar(indices[i+1]))
                    print(f'Restrição ativada para: {nodes[manager.IndexToNode(indices[i])]} e {nodes[manager.IndexToNode(indices[i+1])]}')

            # Assegura que representantes de grupos distintos não compartilham o mesmo veículo
            for i in range(len(group_representatives)):
                for j in range(i+1, len(group_representatives)):
                    routing.solver().Add(routing.VehicleVar(group_representatives[i]) != routing.VehicleVar(group_representatives[j]))
                    rep1 = nodes[manager.IndexToNode(group_representatives[i])]
                    rep2 = nodes[manager.IndexToNode(group_representatives[j])]
                    print(f'Different vehicle constraint activated between group representatives: {rep1} and {rep2}')

        # 2.4. Callback de distância e custo do arco
        def distance_callback(from_index, to_index):
            from_node = manager.IndexToNode(from_index)
            to_node = manager.IndexToNode(to_index)
            return data['distance_matrix'][from_node][to_node]

        transit_distance_callback_index = routing.RegisterTransitCallback(distance_callback)
        routing.SetArcCostEvaluatorOfAllVehicles(transit_distance_callback_index)

        # 2.5. Capacidade (via callback de demanda)
        def demand_callback(from_index):
            from_node = manager.IndexToNode(from_index)
            return data['demands'][from_node]

        demand_callback_index = routing.RegisterUnaryTransitCallback(demand_callback)
        routing.AddDimensionWithVehicleCapacity(
            demand_callback_index, 0, [data['vehicle_capacity']] * data['num_vehicles'], True, 'Capacity')

        # 2.6. Tempo de viagem (separado de distância)
        def time_callback(from_index, to_index):
            from_node = manager.IndexToNode(from_index)
            to_node = manager.IndexToNode(to_index)
            return data['time_matrix'][from_node][to_node]

        transit_time_callback_index = routing.RegisterTransitCallback(time_callback)

        # 2.7. Dimensão de TEMPO com janelas
        routing.AddDimension(
            transit_time_callback_index,
            self.max_time,  # folga/espera máxima permitida
            self.max_time,  # tempo máximo permitido por rota
            True,           # start cumul at zero (no waiting before start)
            'Time')

        time_dimension = routing.GetDimensionOrDie('Time')

        # 2.8. Dimensão de DISTÂNCIA com limite máximo
        routing.AddDimension(
            transit_distance_callback_index,  # usa a mesma callback da distância/custo
            0,                                # sem folga
            self.max_dist,                    # distância máxima por rota
            True,                             # start cumul at zero
            'Distance'
        )
        distance_dimension = routing.GetDimensionOrDie('Distance')

        # 2.9. Limite máximo de distância por veículo
        for vehicle_id in range(data['num_vehicles']):
            distance_dimension.CumulVar(routing.End(vehicle_id)).SetMax(self.max_dist)

        # 2.10. Janelas de tempo por nó (exceto depósito)
        for location_idx, time_window in enumerate(data['time_windows']):
            if location_idx == 0:
                continue
            index = manager.NodeToIndex(location_idx)
            time_dimension.CumulVar(index).SetRange(int(time_window[0]), int(time_window[1]))

        # 2.11. Janela do depósito (BASE)
        depot_idx = manager.NodeToIndex(data['depot'])
        time_dimension.CumulVar(depot_idx).SetRange(e[self.base], l[self.base])

        # 2.12. Estratégia de busca e limites
        search_parameters = pywrapcp.DefaultRoutingSearchParameters()
        # Estratégia de solução inicial escolhida pela maior restrição de arcos
        search_parameters.first_solution_strategy = (
            routing_enums_pb2.FirstSolutionStrategy.PATH_MOST_CONSTRAINED_ARC) # PATH_CHEAPEST_ARC

        search_parameters.time_limit.FromSeconds(time_limit)  # limite de tempo de busca
        search_parameters.log_search = verbose                # habilita logs

        # =====================
        # 3. Resolução
        # =====================
        if verbose:
            print('Resolvendo Problema')

        solution = routing.SolveWithParameters(search_parameters)
        
        if solution:
            a = 1
            # print("Solução encontrada!")
        else:
            print("Nenhuma solução encontrada! Verifique as restrições.")

        # =====================
        # 4. Extração da solução (para cada veículo/rota)
        # =====================
        solution_dict = {}
        # if solution:
        for vehicle_id in range(data['num_vehicles']):
            route_dict = {}
            index = routing.Start(vehicle_id)
            arc_count = 0

            # Se a rota termina imediatamente após o início, veículo não foi usado
            if routing.IsEnd(solution.Value(routing.NextVar(index))):
                continue

            # Percorre a rota adicionando arcos (i->j)
            while not routing.IsEnd(index):
                from_node = manager.IndexToNode(index)
                next_index = solution.Value(routing.NextVar(index))
                to_node = manager.IndexToNode(next_index)
                i = nodes[from_node]
                j = nodes[to_node]

                arc_data = {
                    'arc': (i, j),
                    'dist': self.distance[i][j],
                    'time': self.time[i][j],
                    'service_time': self.service_time.get(i, 0),  # 0 se for depósito
                    'demand': self.demand.get(i, 0)  # 0 se for depósito
                }
                route_dict[arc_count] = arc_data

                arc_count += 1
                index = next_index

            solution_dict[vehicle_id] = route_dict

        return solution_dict

    
    def remove_infesible_points(self):
        # Remove clientes inviáveis cujo ciclo BASE->i->BASE + serviço excede max_time
        self.infeasible_clients = [client for client in self.clients 
                        if self.time[self.base][client] + self.time[client][self.base] + self.service_time[client] > self.max_time]
        
        self.clients = [client for client in self.clients 
                        if self.time[self.base][client] + self.time[client][self.base] + self.service_time[client] <= self.max_time]


def std_codes(code):
    # Padroniza IDs/códigos vindos de planilhas/CSV: remove quebras e _x000D_;
    # se for numérico, converte para inteiro em string.
    if str(code).replace('.','').isdigit():
        return(str(int(code))).replace('_x000D_\n', '').replace('\n', '')
    else:
        return str(code).replace('_x000D_\n', '').replace('\n', '')


def main(developer=False, tempo_abastecimento=None, rota_1_pra_1=False,
         tempo_total_semana=44, tempo_total_dia=10, visitar_toda_planta=False,
         tempo_de_visita_min=5, output_rodada=False, tempo_sabado=4,
         modal_distance_matrix='carro',
         percentil_deslocamento=50):
    # ----------------------------------------------------------------------------------
    # Parâmetros (principais):
    #   • developer: pula confirmação interativa
    #   • tempo_abastecimento: sobrescreve tempo de serviço (min) por patrimônio (opcional)
    #   • rota_1_pra_1: modo em que a “rota principal” é quebrada em subrotas depois
    #   • tempo_total_semana: orçamento (horas) semanal por abastecedor
    #   • tempo_total_dia: orçamento (horas) diário (usado em rota_1_pra_1)
    #   • visitar_toda_planta: se True, visita ativos fora do cronograma com tempo mínimo
    #   • tempo_de_visita_min: tempo mínimo (min) para visita de “varredura”
    #   • tempo_sabado: limite de horas no sábado (periodo=6)
    #   • modal_distance_matrix: 'carro' ou 'a_pe' (define qual parquet usar)
    #   • percentil_deslocamento: percentil do deslocamento diário para compor demanda
    # ----------------------------------------------------------------------------------
    
    if developer:
        ans='S'
    else:
        print('Essa operação sobrescreverá os dados da aba "Livros", deseja continuar? (s/n)')
        ans = str(input())
    
    if ans.upper() != 'S':
        print('Otimização de livros cancelada')
        time.sleep(0.5)
        return 0

    start = time.time()
    print("Iniciando Otimização de Lotes")

    # =====================
    # 1. Importação e leitura de parâmetros/insumos
    # =====================
    current_script_directory = os.path.dirname(os.path.abspath(__file__))
    main_dir = '/'.join(current_script_directory.split('\\')[:-1]) + '/'
    model_data_folder = main_dir + 'Dados Intermediários/'
    input_folder = main_dir + 'Dados de Input/'

    section_start = time.time()
    print("Iniciando leitura e tratamentos iniciais de dados e parâmetros...")

    # 1.1. De-para de pontos (com POINT_ID/LAT/LON por parceiro/filial)
    depara_point_id = pd.read_parquet(model_data_folder + 'depara_point_id_atualizado.parquet')[['FILIAL', 'PARCEIRO', 'POINT_ID', 'LAT', 'LON']]
    depara_point_id['PARCEIRO'] = depara_point_id['PARCEIRO'].apply(std_codes)

    # 1.2. Configurações gerais (workbook/planilha)
    for file in os.listdir(main_dir):
        if len(file) > 4:
            if file[-5:] == '.xlsm':
                workbook = file
                
    close_excel_file_if_open(main_dir + workbook)
    wb = xw.Book(main_dir + workbook)

    config_sheet = wb.sheets['Configurações']

    # Número de dias operacionais por semana (ex. 5 ou 6)
    dias_semana = int(config_sheet['D7'].value)
    n_p = dias_semana
    periodos = [i for i in range(1, n_p+1)]

    # Map para inputs textualizados na planilha
    input_mapping = {"Sim": True, "Não": False, 'Carro/Moto':'carro', 'A pé':'a_pe'}
    time_limit = int(config_sheet['J7'].value)              # limite (s) do solver por lote
    modal_distance_matrix = input_mapping[config_sheet['J8'].value]
    error_margin = int(config_sheet['J9'].value)            # margem p/ decisão de modal a pé
    rota_1_pra_1 = input_mapping[config_sheet['J11'].value]
    visitar_toda_planta = input_mapping[config_sheet['J12'].value]
    
    # Tabela de config por filial: fator trânsito, tempo/distância máximos
    params = config_sheet.range("G14").expand().value
    config = {}
    if type(params[0]) == list:
        for row in params:
            config[str(row[0])] = {
                'Fator Transito':float(row[1]),
                'Tempo Max':int(round(row[2]*3600)),
                'Dist Max':int(row[3]*1000)
            }
    else:
        row = params
        config[str(row[0])] = {
            'Fator Transito':float(row[1]),
            'Tempo Max':int(round(row[2]*3600)),
            'Dist Max':int(row[3]*1000)
        }

    # Escalas de trabalho (nome -> horas/dia em segundos)
    nome_escalas = config_sheet.range("B11").expand().value
    horas_escalas = config_sheet.range("D11").expand().value
    escalas = {}
    for i, escala in enumerate(nome_escalas):
        escalas[escala] = int(horas_escalas[i]*3600)

    # 1.3. Cronograma semanal -> explode em linhas por (FILIAL, PARCEIRO, PATRIMONIO, PERIODO)
    cronograma_sheet = wb.sheets['Cronograma']
    cronograma_df = cronograma_sheet.range("B7").expand().value
    cronograma_df = pd.DataFrame(cronograma_df[1:], columns=['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'FREQUENCIA'] + [i for i in range(1, n_p+1)])
    cronograma_df['FILIAL'] = cronograma_df['FILIAL'].astype(str)
    cronograma_df['PARCEIRO'] = cronograma_df['PARCEIRO'].apply(std_codes)
    cronograma_df['PATRIMONIO'] = cronograma_df['PATRIMONIO'].apply(std_codes)
    cronograma_df = cronograma_df.melt(id_vars=['FILIAL', 'PARCEIRO', 'PATRIMONIO'], value_vars=[i for i in range(1, n_p+1)], var_name='PERIODO', value_name='VISITA')
    cronograma_df = cronograma_df[cronograma_df['VISITA'].notna()].drop(columns='VISITA')

    # 1.4. Frequências (visitas/semana por patrimônio)
    frequency_sheet = wb.sheets['Frequências']
    freq_df = frequency_sheet.range("B6").expand().value
    freq_df = pd.DataFrame(freq_df[1:], columns=freq_df[0])[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'Frequência (visitas/semana)']].rename(columns={'Frequência (visitas/semana)':'FREQUENCIA'})
    freq_df['FILIAL'] = freq_df['FILIAL'].astype(str)
    freq_df['PARCEIRO'] = freq_df['PARCEIRO'].apply(std_codes)
    freq_df['PATRIMONIO'] = freq_df['PATRIMONIO'].apply(std_codes)
    freq_df['FREQUENCIA'] = freq_df['FREQUENCIA'].astype(int)

    # 1.5. Matriz de distâncias/tempos ajustada pelo “Fator Trânsito”
    distance_matrix = pd.read_parquet(model_data_folder + f'{modal_distance_matrix}_distance_matrix.parquet')
    distance_matrix['DISTANCE'] = distance_matrix['DISTANCE'].round(0).astype(int)
    distance_matrix['DURATION'] = (distance_matrix['DURATION']*distance_matrix['FILIAL'].map({k:v['Fator Transito'] for k,v in config.items()})).round(0).astype(int)

    # 1.6. Patrimônios (tempo de serviço etc.)
    patrimonios_df = pd.read_excel(model_data_folder + 'Dados.xlsx', sheet_name='patrimonios')
    patrimonios_df['PATRIMONIO'] = patrimonios_df['PATRIMONIO'].apply(std_codes)
    patrimonios_df['PARCEIRO'] = patrimonios_df['PARCEIRO'].apply(std_codes)

    # 1.7. Parceiros (janelas de funcionamento, supervisor/abastecedor)
    parceiros_df = pd.read_excel(model_data_folder + 'Dados.xlsx', sheet_name='parceiros')
    parceiros_df['PARCEIRO'] = parceiros_df['PARCEIRO'].apply(std_codes)

    # =====================
    # 2. Tratamentos iniciais e montagem de "lotes" (grupos por parceiro/período)
    # =====================

    # 2.0. Parametrização por filial (Tempo Max Dia inicia como Tempo Max da filial)
    for filial in config.keys():
        config[filial]['Tempo Max Dia'] = config[filial]['Tempo Max']

    # Ajuste do Tempo Max no modo 1:1 para filial SP (mantido conforme original)
    if rota_1_pra_1:
        config['SP']['Tempo Max'] = int(tempo_total_dia*3600)
    
    # 2.1. Janelas de tempo (converte HH:MM -> segundos e normaliza para início mínimo)
    index = parceiros_df['INICIO_FUNCIONAMENTO'].notna()
    parceiros_df.loc[index, 'INICIO'] = parceiros_df.loc[index, 'INICIO_FUNCIONAMENTO'].astype(str).apply(lambda x: int(x.split(':')[0])*3600 + int(x.split(':')[1])*60).astype(int)
    index = parceiros_df['FIM_FUNCIONAMENTO'].notna()
    parceiros_df.loc[index, 'FIM'] = (parceiros_df.loc[index, 'FIM_FUNCIONAMENTO']
                                        .astype(str).apply(lambda x: int(x.split(':')[0])*3600 + 
                                                            int(x.split(':')[1])*60).astype(int))

    # Se a janela cruza meia-noite (INICIO > FIM), empurra FIM para o dia seguinte
    index = parceiros_df['INICIO'] > parceiros_df['FIM']
    parceiros_df.loc[index, 'FIM'] = parceiros_df.loc[index, 'FIM'] + 24*3600

    # Normalização: subtrai o menor INICIO para começar em 0
    min_inicio = parceiros_df['INICIO'].min()
    parceiros_df['INICIO'] = parceiros_df['INICIO'].fillna(min_inicio) - min_inicio
    parceiros_df['FIM'] = parceiros_df['FIM'] - min_inicio

    # Preenche FIM faltante com max(FIM observado, Tempo Max por filial)
    max_fim = parceiros_df['FIM'].max()
    index = parceiros_df['FIM'].isna()
    parceiros_df.loc[index, 'FIM'] = np.maximum(max_fim, parceiros_df.loc[index, 'FILIAL'].apply(lambda x: config[x]['Tempo Max'])).astype(float)

    # Casts e seleção de colunas relevantes; tempo de entrada em segundos
    parceiros_df['INICIO'] = parceiros_df['INICIO'].astype(int)
    parceiros_df['FIM'] = parceiros_df['FIM'].astype(int)
    parceiros_df = parceiros_df[['FILIAL', 'PARCEIRO', 'TEMPO_DE_ENTRADA_MIN', 'INICIO', 'FIM', 'SUPERVISOR','ABASTECEDOR']].rename(columns={'TEMPO_DE_ENTRADA_MIN':'TEMPO_DE_ENTRADA'})
    parceiros_df['TEMPO_DE_ENTRADA'] = (parceiros_df['TEMPO_DE_ENTRADA']*60).astype(int)
    parceiros_df['ABASTECEDOR'] = parceiros_df['ABASTECEDOR'].fillna('-')

    # 2.2. Unificação e preparação de dados por período
    if tempo_abastecimento is not None:
        # Sobrescreve tempo de serviço (min) se informado
        patrimonios_df['TEMPO_DE_SERVICO_MINUTOS'] = tempo_abastecimento 

    if rota_1_pra_1:
        # Modo 1:1 cria um único período (PERIODO=1) por parceiro/patrimônio
        patrimonios_df_aux = patrimonios_df.copy()
        parceiros_df_aux = parceiros_df.copy()
        cronograma_df_aux = freq_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO']].copy()
        cronograma_df_aux['PERIODO'] = 1
    else:
        # Modo normal: usa cronograma semanal (e opcionalmente visita toda a planta)
        patrimonios_df_aux = patrimonios_df.copy()
        parceiros_df_aux = parceiros_df.copy()
        if visitar_toda_planta:
            # Para visitar toda planta: combina parceiros x períodos e marca ativos fora do cronograma
            cronograma_df_aux = (cronograma_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO']]
                                 .drop_duplicates()
                                 .merge(cronograma_df[['FILIAL', 'PARCEIRO', 'PERIODO']]
                                        .drop_duplicates(),
                                        on=['FILIAL', 'PARCEIRO'],
                                        how='left')
                                 .merge(cronograma_df
                                        .assign(ABASTECIMENTO=1),
                                        on=['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'PERIODO'],
                                        how='left'))
        else:
            cronograma_df_aux = cronograma_df.copy()

    # 2.3. Monta base de “lotes”: cada linha é um patrimônio em um (FILIAL, PARCEIRO, PERIODO)
    lotes_df = patrimonios_df_aux[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'TEMPO_DE_SERVICO_MINUTOS']].copy().rename(columns={'TEMPO_DE_SERVICO_MINUTOS':'TEMPO_SERVICO'})
    lotes_df['TEMPO_SERVICO'] = (lotes_df['TEMPO_SERVICO']*60).astype(int)  # min -> s

    lotes_df = (
        lotes_df
        .merge(
            parceiros_df_aux,
            on=['PARCEIRO', 'FILIAL'],
            how='left'
        )
        .merge(
            cronograma_df_aux,
            on=['PARCEIRO', 'PATRIMONIO', 'FILIAL'],
            how='left'
        )
        .merge(
            depara_point_id[['PARCEIRO', 'FILIAL', 'POINT_ID']],
            on=['PARCEIRO', 'FILIAL'],
            how='left'
        )
    )

    # Se for visitar toda a planta, patrimônios fora do cronograma recebem tempo mínimo
    if ((not rota_1_pra_1) and visitar_toda_planta):        
        index = (lotes_df['ABASTECIMENTO'].isnull())
        lotes_df.loc[index, 'TEMPO_SERVICO'] = tempo_de_visita_min*60

    # 2.4. Agrupamento de patrimônios por parceiro/período respeitando tempo disponível
    lotes_df = lotes_df.sort_values(by=['PERIODO','FILIAL','PARCEIRO', 'TEMPO_SERVICO','PATRIMONIO'])

    lotes_df['MAX_TIME'] = lotes_df['FILIAL'].map({f:c['Tempo Max'] for f,c in config.items()})
    lotes_df['GROUP'] = lotes_df['PARCEIRO'] + 'P' + lotes_df['PERIODO'].astype(str) + 'G1'
    group_num = 1
    lotes_df['AC_TIME'] = lotes_df.groupby(['PERIODO','FILIAL','PARCEIRO', 'GROUP'])['TEMPO_SERVICO'].cumsum()

    # Enquanto a soma acumulada + tempo de entrada exceder MAX_TIME, empurra p/ próximo grupo
    index = ((lotes_df['AC_TIME'] + lotes_df['TEMPO_DE_ENTRADA']) > lotes_df['MAX_TIME'])
    while index.sum() > 0:
        group_num+=1
        lotes_df.loc[index, 'GROUP'] = lotes_df.loc[index, 'PARCEIRO'] + 'P' + lotes_df.loc[index, 'PERIODO'].astype(str) + f'G{group_num}'
        lotes_df['AC_TIME'] = lotes_df.groupby(['PERIODO','FILIAL','PARCEIRO', 'GROUP'])['TEMPO_SERVICO'].cumsum()
        index = ((lotes_df['AC_TIME'] + lotes_df['TEMPO_DE_ENTRADA']) > lotes_df['MAX_TIME'])

    # 2.5. Enriquecimento: frequências, demanda semanal e deslocamento típico por POINT_ID
    lotes_df = (
        lotes_df
        .merge(freq_df,
            on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
            how='left')
        .assign(WEEK_DEMAND=lambda x: x['FREQUENCIA']*x['TEMPO_SERVICO'])
        .rename(columns={'FREQUENCIA':'FREQUENCIA_PATRIMONIO'})
        .merge(freq_df
            .groupby(['FILIAL', 'PARCEIRO'])
            .agg({'FREQUENCIA':'max'})
            .reset_index(),
            on=['FILIAL', 'PARCEIRO'],
            how='left')
        .assign(TEMPO_DE_ENTRADA_SEMANAL=lambda x: x['FREQUENCIA']*x['TEMPO_DE_ENTRADA'])
        .merge(distance_matrix
            .groupby('POINT_ID_J')
            .agg({'DURATION':lambda v: np.percentile(v, percentil_deslocamento)})
            .reset_index()
            .rename(columns={'POINT_ID_J':'POINT_ID',
                             'DURATION':'TEMPO_DESLOCAMENTO_DIA'}),
            on='POINT_ID',
            how='left')
        # .assign(TEMPO_DESLOCAMENTO_SEMANAL=...)  # mantido comentado como original
    )

    # Se visitar toda a planta, adiciona varredura (tempo mínimo) além da frequência base
    if visitar_toda_planta:
        lotes_df['WEEK_DEMAND'] = (
            lotes_df['WEEK_DEMAND'] +
            ((lotes_df['FREQUENCIA'] - lotes_df['FREQUENCIA_PATRIMONIO']) * tempo_de_visita_min * 60))

    # 2.6. Controle de restrição semanal (modo 1:1 ajusta orçamento; modo normal relaxa)
    if rota_1_pra_1:
        tempo_total_semana_adj = tempo_total_semana
    else:
        tempo_total_semana_adj=tempo_total_semana*1000000  # “infinito” prático
        
    restricao_semanal = False

    # =====================
    # 3. Otimização de lotes (loop até respeitar restrição semanal no modo 1:1)
    # =====================
    while not restricao_semanal:
        # 3.1. Agrega por (filial, periodo, supervisor, abastecedor, parceiro, group...)
        grouped_df = (lotes_df
                        .groupby(['FILIAL', 'PERIODO', 'SUPERVISOR', 'ABASTECEDOR', 'PARCEIRO', 'PATRIMONIO', 'GROUP', 'POINT_ID', 'INICIO', 'FIM', 'TEMPO_DE_ENTRADA'])
                        .agg({'TEMPO_SERVICO':'sum',
                            'WEEK_DEMAND':'first',
                            'TEMPO_DE_ENTRADA_SEMANAL':'first'
                            #,'TEMPO_DESLOCAMENTO_SEMANAL': 'first'
                            })
                        .reset_index()                      
                        .groupby(['FILIAL', 'PERIODO', 'SUPERVISOR', 'ABASTECEDOR', 'PARCEIRO', 'GROUP', 'POINT_ID', 'INICIO', 'FIM', 'TEMPO_DE_ENTRADA'])
                        .agg({'TEMPO_SERVICO':'sum',
                            'WEEK_DEMAND':'sum',
                            'TEMPO_DE_ENTRADA_SEMANAL':'first'
                            #, 'TEMPO_DESLOCAMENTO_SEMANAL': 'first'
                            })
                        .reset_index())
        
        # grouped_df['WEEK_DEMAND'] = ...  # mantido como no original (comentado)

        print('Leitura e tratamento de dados concluídos, tempo: {:.1f}s'.format(time.time() - section_start))

        # 3.2. Laço de solução por supervisor e período
        section_start = time.time()
        print('Iniciando otimização de lotes...')

        n_p_aux = 1 if rota_1_pra_1 else n_p

        # Mapeia supervisor -> filial (para particionar o problema)
        supervisores_filial = (
            grouped_df
            [['SUPERVISOR', 'FILIAL']]
            .drop_duplicates()
            .sort_values(by='FILIAL')
            .set_index('SUPERVISOR')['FILIAL'].to_dict())
        
        result_df = pd.DataFrame()

        # Para cada supervisor: resolve período a período
        for supervisor in supervisores_filial.keys():    
            filial = supervisores_filial[supervisor]
            start_filial = time.time()
            for periodo in tqdm(range(1, n_p_aux + 1), file=sys.stdout):

                # Subconjunto do lote: (filial, período, supervisor)
                lote_df = grouped_df[(grouped_df['FILIAL'] == filial) & 
                                    (grouped_df['PERIODO'] == periodo) &
                                    (grouped_df['SUPERVISOR'] == supervisor)].copy()
                

                # Dicionários de distância/tempo por POINT_ID (i,j)
                dist_dict = distance_matrix[distance_matrix['FILIAL'] == filial].set_index(['POINT_ID_I','POINT_ID_J'])['DISTANCE'].to_dict()
                time_dict = distance_matrix[distance_matrix['FILIAL'] == filial].set_index(['POINT_ID_I','POINT_ID_J'])['DURATION'].to_dict()

                # Parâmetros por GROUP (nó cliente)
                # demanda_por_grupo = lote_df.set_index('GROUP')['WEEK_DEMAND'].to_dict() | {'BASE': 0}
                tempo_servico = lote_df.set_index('GROUP')['TEMPO_SERVICO'].to_dict() | {'BASE': 0}
                tempo_entrada = lote_df.set_index('GROUP')['TEMPO_DE_ENTRADA'].to_dict() | {'BASE': 0}
                parceiro = lote_df.set_index('GROUP')['PARCEIRO'].to_dict() | {'BASE': 'BASE'}
                point = lote_df.set_index('GROUP')['POINT_ID'].to_dict()
                inicio = lote_df.set_index('GROUP')['INICIO'].to_dict() | {'BASE': 0}
                fim = lote_df.set_index('GROUP')['FIM'].to_dict() | {'BASE': int(48*3600)}

                # =====================
                # 3.3. Montagem do Lote (VRP)
                # =====================
                lote = Lote()

                # Conjuntos
                lote.clients = lote_df['GROUP'].to_list()
                lote.vertices = ['BASE'] + lote.clients

                # Parâmetros globais do VRP
                # Capacidade semanal (s) com folga inicial de +9% (mantido)
                lote.cap = int(np.ceil(tempo_total_semana_adj*3600*1.09))
                lote.route_cost = 1000000     # custo alto para abrir rota (sair da BASE)
                lote.base = 'BASE'
                # Tempo diário: sábado tem limite próprio, demais usam config por filial
                if periodo == 6:
                    lote.max_time = int(tempo_sabado*3600)
                else:
                    lote.max_time = config[filial]['Tempo Max']
                # Distância máxima por rota (adiciona route_cost para “segurar” BASE)
                lote.max_dist = config[filial]['Dist Max'] + lote.route_cost

                # Parâmetros por nó
                lote.demand = {i:1 for i in lote.clients} | {'BASE':0} ########demanda_por_grupo
                lote.tw_start = inicio
                lote.tw_end = fim
                lote.service_time = tempo_servico

                # Parâmetros por arco (matrizes completas com BASE zerado)
                lote.distance = {i:{j:dist_dict[point[i], point[j]] for j in lote.clients} | {'BASE':0} for i in lote.clients} | {'BASE':{j:0 for j in lote.vertices}}
                lote.time = {i:{j:time_dict[point[i], point[j]] for j in lote.clients} | {'BASE':0} for i in lote.clients} | {'BASE':{j:0 for j in lote.vertices}}

                # Penaliza trocas de parceiro: ao mudar de parceiro, adiciona tempo de entrada do destino
                for i in lote.vertices:
                    for j in lote.vertices:
                        if (j != 'BASE') & (parceiro[i] != parceiro[j]):
                            lote.time[i][j] = lote.time[i][j] + tempo_entrada[j]

                # (Opcional) Força grupos no mesmo veículo (mantido desativado como original)
                same_vehicle_groups = []
                # ... blocos comentados mantidos ...

                if len(same_vehicle_groups) > 0:
                    lote.same_vehicle_groups = same_vehicle_groups

                # =====================
                # 3.4. Resolver VRP do lote
                # =====================
                result_dict = lote.solve(time_limit=time_limit, verbose=False)

                # =====================
                # 3.5. Linearizar solução em DataFrame de “livros/visitas”
                # =====================
                livros_df = {'FILIAL':[], 'PERIODO':[], 'LIVRO':[], 'VISITA':[], 'PARCEIRO':[], 'GROUP':[], 'DIST':[], 'TEMPO_DESLOCAMENTO':[], 'TEMPO_DE_SERVICO':[]}

                livro = 1
                for visitas in result_dict.values():
                    livro_name = f'{filial} - P{periodo}L{livro} - {supervisor}'
                    livro += 1
                    for visita, arco in visitas.items():
                        livros_df['FILIAL'].append(filial)
                        livros_df['PERIODO'].append(periodo)
                        livros_df['LIVRO'].append(livro_name)
                        livros_df['VISITA'].append(visita + 1)
                        livros_df['PARCEIRO'].append(parceiro[arco['arc'][1]])
                        livros_df['GROUP'].append(arco['arc'][1])
                        livros_df['DIST'].append(arco['dist'])
                        livros_df['TEMPO_DESLOCAMENTO'].append(arco['time'])
                        livros_df['TEMPO_DE_SERVICO'].append(tempo_servico[arco['arc'][1]])
                    
                livros_df = pd.DataFrame(livros_df)
                result_df = pd.concat([result_df, livros_df])

            print(f'Otimização de lotes para {supervisor} da filial {filial} finalizada, tempo de processamento: {time.time() - start_filial:.1f}s')

        print('Otimizações finalizadas para todos os lotes, tempo: {:.1f}s'.format(time.time() - section_start))

        # =====================
        # 4. Pré-processamento de resultados (início)
        # =====================
        section_start = time.time()
        print('Processando resultados...')

        # Remove arcos artificiais para BASE na listagem final de visitas
        result_df = result_df[result_df['PARCEIRO'] != 'BASE'].copy()



    # Quebrando rota principal do modelo 1:1 em subrotas
        if rota_1_pra_1:
            # ----------------------------------------------------------------------
            # Modo 1:1
            #   • Reexpande visitas por patrimônio usando cronograma (ou toda planta)
            #   • Reconstrói sequência de VISITAS por livro/dia
            #   • Recalcula deslocamentos (DIST/DURATION) com POINT_ID_I -> POINT_ID_J
            #   • Ajusta tempos de entrada (apenas ao trocar de parceiro)
            # ----------------------------------------------------------------------
            result_aux = result_df.drop(columns=['PERIODO', 'TEMPO_DE_SERVICO']).copy()

            if visitar_toda_planta:
                cronograma_aux = (cronograma_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO']]
                                .drop_duplicates()
                                .merge(cronograma_df[['FILIAL', 'PARCEIRO', 'PERIODO']]
                                        .drop_duplicates(),
                                        on=['FILIAL', 'PARCEIRO'],
                                        how='left')
                                .merge(cronograma_df
                                        .assign(ABASTECIMENTO=1),
                                        on=['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'PERIODO'],
                                        how='left'))
            else:
                cronograma_aux = cronograma_df.copy()

            result_aux = (
                result_aux
                .merge(lotes_df[['FILIAL', 'PARCEIRO', 'GROUP', 'PATRIMONIO', 'POINT_ID']],
                    on=['FILIAL', 'PARCEIRO', 'GROUP'],
                    how='left')
                .rename(columns={'POINT_ID':'POINT_ID_J'})            
                .merge(cronograma_aux,
                    on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                    how='left')
                .merge(patrimonios_df[['FILIAL', 'PARCEIRO', 'PATRIMONIO', 'TEMPO_DE_SERVICO_MINUTOS']],
                    on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                    how='left')
                .assign(TEMPO_SERVICO=lambda x: x['TEMPO_DE_SERVICO_MINUTOS']*60)
                .drop(columns='TEMPO_DE_SERVICO_MINUTOS')
                .merge(parceiros_df[['FILIAL', 'PARCEIRO', 'TEMPO_DE_ENTRADA']],
                    on=['FILIAL', 'PARCEIRO'],
                    how='left')
                )

            # Se visitando toda a planta, patrimônios “fora” do cronograma recebem tempo mínimo
            if visitar_toda_planta:
                index = (result_aux['ABASTECIMENTO'].isnull())
                result_aux.loc[index, 'TEMPO_SERVICO'] = tempo_de_visita_min*60

            # Normaliza numeração dos livros (N_LIVRO) e cria LIVRO_ADJ reindexado por período
            result_aux['N_LIVRO'] = result_aux['LIVRO'].str.extract(r'L(\d+)$').astype(int)
            result_aux = result_aux.sort_values(by=['FILIAL', 'PERIODO', 'N_LIVRO', 'LIVRO']).reset_index(drop=True)
            result_aux['N_LIVRO'] = result_aux.groupby('PERIODO')['N_LIVRO'].transform(lambda x: x.rank(method='dense').astype(int))
            result_aux['LIVRO_ADJ'] = result_aux.apply(lambda row: pd.Series(row['LIVRO']).str
                                                    .replace(r'P\d+(?=L)', f'P{row["PERIODO"]}', regex=True)
                                                    .replace(r'L(\d+)$', f'L{row["N_LIVRO"]}', regex=True)
                                                    .iloc[0], axis=1)

            # De-para entre LIVRO ajustado e original agregado
            depara_livro = (
                result_aux[['FILIAL', 'LIVRO_ADJ', 'LIVRO']]
                .drop_duplicates()
                .rename(columns={'LIVRO':'LIVRO_AGG', 'LIVRO_ADJ':'LIVRO'})
                .copy())

            # Substitui rótulo do livro e reindexa VISITA
            result_aux = (result_aux
                        .drop(columns='LIVRO')
                        .assign(LIVRO=result_aux['LIVRO_ADJ']))
            result_aux['VISITA'] = result_aux.groupby('LIVRO').cumcount() + 1
            result_aux = result_aux.sort_values(by=['FILIAL', 'PERIODO', 'N_LIVRO', 'LIVRO', 'VISITA']).reset_index(drop=True)

            # Cria pares (POINT_ID_I -> POINT_ID_J) por visita, para recompor DIST/DURATION
            result_aux['POINT_ID_I'] = result_aux['POINT_ID_J'].shift(1)
            result_aux.loc[result_aux['VISITA'] == 1, 'POINT_ID_I'] = None

            result_aux = (result_aux
                        .merge(distance_matrix,
                                on=['FILIAL', 'POINT_ID_I', 'POINT_ID_J'],
                                how='left')
                        .fillna({'DISTANCE': 0, 'DURATION':0})
                        .drop(columns=['DIST', 'TEMPO_DESLOCAMENTO', 'POINT_ID_J', 'POINT_ID_I'])
                        .rename(columns={'DISTANCE':'DIST', 'DURATION':'TEMPO_DESLOCAMENTO'}))

            # Ajusta TEMPO_DE_ENTRADA: só na primeira visita de cada parceiro dentro do livro
            result_aux['N_PATRIMONIO_POR_PARCERIO'] = result_aux.groupby(['LIVRO', 'PARCEIRO'])['PATRIMONIO'].transform(lambda x: x.rank(method='dense').astype(int))
            result_aux['REMOVER_TEMPO_DE_ENTRADA'] = (result_aux['N_PATRIMONIO_POR_PARCERIO'] == 1)
            result_aux.loc[(~result_aux['REMOVER_TEMPO_DE_ENTRADA']), 'TEMPO_DE_ENTRADA'] = 0

            # Limpeza e ajuste final de tempos: serviço vai a zero (somado antes em arcos)
            result_aux['N_PATRIMONIO_POR_PARCERIO'] = result_aux.groupby(['LIVRO', 'PARCEIRO'])['PATRIMONIO'].transform(lambda x: x.rank(method='dense').astype(int))
            result_aux['REMOVER_TEMPO_DE_ENTRADA'] = (result_aux['N_PATRIMONIO_POR_PARCERIO'] == 0)
            result_aux.loc[result_aux['REMOVER_TEMPO_DE_ENTRADA'], 'TEMPO_DE_ENTRADA'] = 0

            result_aux['TEMPO_DE_SERVICO'] = 0

            result_aux = result_aux.drop(columns=['N_PATRIMONIO_POR_PARCERIO', 'N_LIVRO'])
            result_df = result_aux[['FILIAL', 'PERIODO', 'LIVRO', 'VISITA', 'PARCEIRO', 'GROUP', 'DIST',
                                    'TEMPO_DESLOCAMENTO', 'TEMPO_DE_SERVICO', 'REMOVER_TEMPO_DE_ENTRADA',
                                    'PATRIMONIO', 'TEMPO_SERVICO', 'TEMPO_DE_ENTRADA']].copy()
            
        else:
            # ----------------------------------------------------------------------
            # Modo NORMAL
            #   • Reconstrói sequência VISITA por livro, marca trocas de parceiro
            #   • Ajusta TEMPO_DE_ENTRADA e deslocamento apenas quando muda de parceiro
            # ----------------------------------------------------------------------
            aux_df = result_df[['LIVRO', 'VISITA', 'PARCEIRO']].copy()
            aux_df['VISITA'] = aux_df['VISITA'] + 1
            result_df = result_df.merge(aux_df, on=['LIVRO', 'VISITA'], how='left', suffixes=('', '_ANT'))
            result_df['REMOVER_TEMPO_DE_ENTRADA'] = (result_df['PARCEIRO_ANT'] != result_df['PARCEIRO'])
            result_df = result_df.drop(columns='PARCEIRO_ANT')

            # Junta info de serviço/entrada por GROUP e recompõe visitabilidade
            result_df = result_df.merge(
                lotes_df[['FILIAL', 'PERIODO', 'SUPERVISOR', 'PARCEIRO', 'GROUP', 'PATRIMONIO', 'TEMPO_SERVICO', 'TEMPO_DE_ENTRADA']], 
                on=['FILIAL', 'PERIODO', 'PARCEIRO', 'GROUP'], how='left')

            result_df = result_df[result_df['PARCEIRO'] != 'BASE'].copy()
            result_df['VISITA'] = 1
            result_df['VISITA'] = result_df.groupby(['LIVRO'])['VISITA'].cumsum()

            # Ajusta deslocamento deduzindo TEMPO_DE_ENTRADA onde aplica (troca de parceiro)
            result_df['TEMPO_DESLOCAMENTO'] = result_df['TEMPO_DESLOCAMENTO'] - result_df['TEMPO_DE_ENTRADA']*result_df['REMOVER_TEMPO_DE_ENTRADA']

            # Reavalia trocas (olhando linha anterior) para setar TEMPO_DE_ENTRADA/DESLOCAMENTO/DIST
            aux_df = result_df[['LIVRO', 'VISITA', 'PARCEIRO']].copy()
            aux_df['VISITA'] = aux_df['VISITA'] + 1
            result_df = result_df.merge(aux_df, on=['LIVRO', 'VISITA'], how='left', suffixes=('', '_ANT'))
            result_df['TEMPO_DE_ENTRADA'] = (result_df['PARCEIRO_ANT'] != result_df['PARCEIRO'])*result_df['TEMPO_DE_ENTRADA']
            result_df['TEMPO_DESLOCAMENTO'] = (result_df['PARCEIRO_ANT'] != result_df['PARCEIRO'])*result_df['TEMPO_DESLOCAMENTO']
            result_df['DIST'] = (result_df['PARCEIRO_ANT'] != result_df['PARCEIRO'])*result_df['DIST']
            result_df = result_df.drop(columns='PARCEIRO_ANT')

        # --------------------------------------------------------------------------
        # 4.1. Alternativa a pé: calcula deslocamento equivalente (5 km/h) e decide modal
        # --------------------------------------------------------------------------
        result_df['TEMPO_DESLOCAMENTO_A_PE'] = (result_df['DIST']/(5/3.6)).astype(int)  # s = m / (m/s)

        # Agrega métricas por livro (resumo)
        result_livros_df = (
            result_df
            .groupby(['FILIAL', 'SUPERVISOR', 'PERIODO', 'LIVRO'])
            .agg({'PATRIMONIO':'nunique','DIST':'sum','TEMPO_SERVICO':'sum', 'TEMPO_DE_ENTRADA':'sum', 'TEMPO_DESLOCAMENTO':'sum', 'TEMPO_DESLOCAMENTO_A_PE':'sum', 'PARCEIRO':'nunique'})
            .reset_index()
        )

        # Tempo total por modal (s)
        result_livros_df['TEMPO_TOTAL'] = result_livros_df['TEMPO_SERVICO'] + result_livros_df['TEMPO_DE_ENTRADA'] + result_livros_df['TEMPO_DESLOCAMENTO']
        result_livros_df['TEMPO_TOTAL_A_PE'] = result_livros_df['TEMPO_SERVICO'] + result_livros_df['TEMPO_DE_ENTRADA'] + result_livros_df['TEMPO_DESLOCAMENTO_A_PE']
        
        # Limites por dia/filial e marca modal “A Pé” quando caber no limite (com margem)
        result_livros_df['MAX_TIME'] = result_livros_df['FILIAL'].map({f:c['Tempo Max Dia'] for f,c in config.items()})
        result_livros_df.loc[result_livros_df['PERIODO'] == 6, "MAX_TIME"] = int(tempo_sabado * 3600)

        result_livros_df['MODAL'] = 'Moto/Carro'

        index = (result_livros_df['TEMPO_TOTAL_A_PE']*100000 <= result_livros_df['MAX_TIME']*(1 + error_margin))
        result_livros_df.loc[index, 'MODAL'] = 'A Pé'
        result_livros_df.loc[index, 'TEMPO_TOTAL'] = result_livros_df.loc[index, 'TEMPO_TOTAL_A_PE']
        result_livros_df.loc[index, 'TEMPO_DESLOCAMENTO'] = result_livros_df.loc[index, 'TEMPO_DESLOCAMENTO_A_PE']
        result_livros_df = result_livros_df.drop(columns=['TEMPO_TOTAL_A_PE', 'TEMPO_DESLOCAMENTO_A_PE'])
        result_livros_df['TEMPO_DESLOCAMENTO'] = result_livros_df['TEMPO_DESLOCAMENTO'] + result_livros_df['TEMPO_DE_ENTRADA']

        # Escala diária sugerida (escolhe a menor escala que comporta TEMPO_TOTAL)
        result_livros_df['HORAS_DIARIAS'] = result_livros_df['MAX_TIME'] 
        result_livros_df['ESCALA'] = 'Full-Time'

        index = (result_livros_df['TEMPO_TOTAL'] <= max(escalas.values()))
        result_livros_df.loc[index, 'HORAS_DIARIAS'] = result_livros_df.loc[index,'TEMPO_TOTAL'].apply(lambda x: min([t for t in escalas.values() if t >= x]))
        result_livros_df.loc[index, 'ESCALA'] = result_livros_df.loc[index, 'HORAS_DIARIAS'].map({t:escala for escala, t in escalas.items()})
        result_livros_df['FTE'] = result_livros_df['HORAS_DIARIAS']/result_livros_df['MAX_TIME']
        
        # Anexa modal/escala às visitas (linhas detalhadas)
        result_df = result_df.merge(result_livros_df[['FILIAL', 'SUPERVISOR', 'PERIODO', 'LIVRO', 'MODAL', 'ESCALA']], on=['FILIAL', 'SUPERVISOR', 'PERIODO', 'LIVRO'], how='left')

        # Backup para compor tempos de deslocamento “Moto/Carro” (em minutos mais tarde)
        index = (result_df['MODAL'] == 'A Pé')
        backup_result_df=result_df.copy()            
        backup_result_df['TEMPO_DESLOCAMENTO'] = backup_result_df['TEMPO_DESLOCAMENTO'] + backup_result_df['TEMPO_DE_ENTRADA']
        backup_result_df['TEMPO_DESLOCAMENTO'] = round(backup_result_df['TEMPO_DESLOCAMENTO']/60, 1)

        # Para linhas “A Pé”: deslocamento calculado acima (já inclui entrada no bloco mais abaixo)
        result_df.loc[index, 'TEMPO_DESLOCAMENTO'] = result_df.loc[index, 'TEMPO_DESLOCAMENTO_A_PE']
        result_df['TEMPO_DESLOCAMENTO'] = result_df['TEMPO_DESLOCAMENTO'] + result_df['TEMPO_DE_ENTRADA']

        # Junta frequências originais para relatório final
        result_df = result_df.merge(freq_df,
                                    on=['FILIAL', 'PARCEIRO', 'PATRIMONIO'],
                                    how='left')
        
        # --------------------------------------------------------------------------
        # 4.2. DataFrames de saída para Excel (linhas detalhadas e resumo por livro)
        # --------------------------------------------------------------------------
        rename_dict = {
            'FILIAL':'Filial', 'PERIODO':'Dia', 'LIVRO':'Livro', 'ESCALA':'Escala Requerida', 'MODAL':'Modal de Transporte', 
            'VISITA':'# Visita', 'PARCEIRO':'Parceiro', 'PATRIMONIO':'Patrimônio', 
            'DIST':'Distância (km)', 'TEMPO_DESLOCAMENTO':'Tempo de Deslocamento (min)', 'TEMPO_SERVICO':'Tempo de Serviço (min)',
            'FREQUENCIA': 'Frequência (visitas/semana)'
        }
        xl_result_df = result_df[rename_dict.keys()].rename(columns=rename_dict).copy().sort_values(by=['Filial', 'Dia', 'Modal de Transporte', 'Escala Requerida', 'Livro', '# Visita'])

        # Conversões finais para unidades de relatório (min/km)
        xl_result_df['Tempo de Deslocamento (min)'] = round(xl_result_df['Tempo de Deslocamento (min)']/60, 1)
        xl_result_df['Tempo de Serviço (min)'] = round(xl_result_df['Tempo de Serviço (min)']/60, 1)
        xl_result_df['Distância (km)'] = round(xl_result_df['Distância (km)']/1000, 2)

        rename_dict = {
            'FILIAL':'Filial', 'PERIODO':'Dia', 'LIVRO':'Livro', 'ESCALA':'Escala Requerida', 'MODAL':'Modal de Transporte', 'HORAS_DIARIAS':'Horas Diárias', 'PATRIMONIO':'# Patrimônios', 
            'DIST':'Distância (km)', 'TEMPO_DESLOCAMENTO':'Tempo de Deslocamento (min)', 'TEMPO_SERVICO':'Tempo de Serviço (min)', 'PARCEIRO': '# Parceiros'
        }
        xl_result_livros_df = result_livros_df[rename_dict.keys()].rename(columns=rename_dict).copy().sort_values(by=['Filial', 'Dia','Modal de Transporte', 'Escala Requerida'])

        xl_result_livros_df['Horas Diárias'] = round(xl_result_livros_df['Horas Diárias']/3600, 1)
        xl_result_livros_df['Tempo de Deslocamento (min)'] = round(xl_result_livros_df['Tempo de Deslocamento (min)']/60, 1)
        xl_result_livros_df['Tempo de Serviço (min)'] = round(xl_result_livros_df['Tempo de Serviço (min)']/60, 1)
        xl_result_livros_df['Distância (km)'] = round(xl_result_livros_df['Distância (km)']/1000, 2)

        # --------------------------------------------------------------------------
        # 4.3. Salvando resultados intermediários em parquet (para diagnósticos)
        # --------------------------------------------------------------------------
        result_df = result_df[['FILIAL', 'PERIODO', 'LIVRO', 'ESCALA', 'MODAL', 'VISITA', 'PARCEIRO', 'PATRIMONIO', 'DIST', 'TEMPO_SERVICO', 'TEMPO_DESLOCAMENTO', 'GROUP']]
        result_df['TEMPO_SERVICO'] = round(result_df['TEMPO_SERVICO']/60, 1)
        result_df['TEMPO_DESLOCAMENTO'] = round(result_df['TEMPO_DESLOCAMENTO']/60, 1)
        result_df['DIST'] = round(result_df['DIST']/1000, 2)

        # Linha “VISITA=0” por livro para facilitar totalizações/quebras visuais
        aux_df = result_df.sort_values(by=['VISITA'], ascending=False).drop_duplicates(subset=['LIVRO'], keep='first')
        aux_df['VISITA'] = 0
        aux_df['TEMPO_SERVICO'] = 0
        aux_df['TEMPO_DESLOCAMENTO'] = 0
        aux_df['DIST'] = 0

        result_df = (
            pd.concat([result_df, aux_df])
            .sort_values(by=['FILIAL', 'PERIODO', 'LIVRO', 'VISITA'])
            .merge(
                depara_point_id[['FILIAL', 'PARCEIRO', 'LAT', 'LON']],
                on=['FILIAL', 'PARCEIRO'],
                how='left'
            )
        )

        # --------------------------------------------------------------------------
        # 4.4. Checagem e ajuste de capacidade semanal (apenas no modo 1:1)
        #      Caso exceda, reduz 5% e aumenta percentil de deslocamento, repetindo loop
        # --------------------------------------------------------------------------
        if rota_1_pra_1:
            check_livros = (
                result_df
                .merge(depara_livro,
                    on=['FILIAL', 'LIVRO'],
                    how='left')
                    .assign(TEMPO_DESLOCAMENTO=lambda x: x['TEMPO_DESLOCAMENTO']/60)
                    .assign(TEMPO_SERVICO=lambda x: x['TEMPO_SERVICO']/60)
                    .assign(TEMPO_OPERACAO=lambda x: x['TEMPO_DESLOCAMENTO'] + x['TEMPO_SERVICO'])
                    .assign(DIST=result_df['DIST'])
                    .groupby(['LIVRO_AGG'])
                    .agg({'VISITA':'nunique',
                            'TEMPO_DESLOCAMENTO':'sum',
                            'TEMPO_SERVICO':'sum',
                            'TEMPO_OPERACAO':'sum',
                            'DIST':'sum'})
                        .reset_index()
                )

            result_df_checks = result_df.copy()

            # Critério: nenhum livro agregado pode exceder “tempo_total_semana” (horas)
            restricao_semanal = len(check_livros.loc[lambda x: x['TEMPO_OPERACAO'] > tempo_total_semana]) == 0 
            if not restricao_semanal:
                tempo_total_semana_adj = int(np.ceil(tempo_total_semana_adj * 0.95))
                percentil_deslocamento = percentil_deslocamento + 5
                print('Capacidade semanal excedida, reduzindo a rota principal...')
        else:
            restricao_semanal = True


    # =====================
    # 5. Alocação de abastecedores e preparação final para Excel
    # =====================

    # Resumo final por livro (conversões de unidade)
    result_livros_df = result_livros_df[['FILIAL', 'SUPERVISOR', 'PERIODO', 'LIVRO', 'ESCALA', 'MODAL', 'HORAS_DIARIAS', 'FTE', 'PATRIMONIO', 'DIST', 'TEMPO_SERVICO', 'TEMPO_DESLOCAMENTO']].rename(columns={'PATRIMONIO':'PATRIMONIOS'})
    
    # Coordenadas médias por livro (centroide aproximado para relatórios)
    aux_df = result_df[result_df['VISITA'] != 0].groupby(['FILIAL', 'PERIODO', 'LIVRO']).agg({'LAT':'mean', 'LON':'mean'}).reset_index()
    result_livros_df = result_livros_df.merge(aux_df, on=['FILIAL', 'PERIODO', 'LIVRO'], how='left')

    # Conversão p/ unidades legíveis
    result_livros_df['HORAS_DIARIAS'] = round(result_livros_df['HORAS_DIARIAS']/3600, 1)
    result_livros_df['TEMPO_SERVICO'] = round(result_livros_df['TEMPO_SERVICO']/60, 1)
    result_livros_df['TEMPO_DESLOCAMENTO'] = round(result_livros_df['TEMPO_DESLOCAMENTO']/60, 1)
    result_livros_df['DIST'] = round(result_livros_df['DIST']/1000, 2)

    # Similaridade de livros (compartilhamento de patrimônios) — usado na alocação
    patr_sim = (
        pd.merge(
            result_df[['LIVRO', 'PATRIMONIO']].copy(),
            result_df[['LIVRO', 'PATRIMONIO']].copy(),
            on='PATRIMONIO',
            how='left',
            suffixes=('_I', '_J')
            )
        .groupby(['LIVRO_I', 'LIVRO_J'])
        .agg(PATR_COMUM = ('PATRIMONIO','nunique'))
        .reset_index()
    )

    # ---------------------
    # 5.1. Alocação
    # ---------------------
    # Se modo normal: heurística de emparelhamento de livros entre períodos; se 1:1:
    # abastecedor nomeado por livro agregado (depara_livro).
    if not rota_1_pra_1:
        alocation_dict = {}
        for supervisor in supervisores_filial.keys():
            filial = supervisores_filial[supervisor]
            filial_df = result_livros_df[((result_livros_df['FILIAL'] == filial) &
                                          (result_livros_df['SUPERVISOR'] == supervisor))].copy()

            livros = filial_df['LIVRO'].unique().tolist()
            periodos = filial_df['PERIODO'].unique().tolist()
            abastecedores = [i for i in range(len(livros))]
            pivot = {i:j for i,j in enumerate(livros)}

            periodo = filial_df.set_index('LIVRO')['PERIODO'].to_dict()
            max_dist = config[filial]['Dist Max']

            # Similaridade espacial/operacional entre livros para formar “pacotes” por pessoa
            sim_df = (
                filial_df[['LIVRO', 'LAT', 'LON', 'PERIODO', 'MODAL', 'HORAS_DIARIAS', 'PATRIMONIOS']]
                .assign(AUX=1)
                .merge(
                    filial_df[['LIVRO', 'LAT', 'LON', 'PERIODO', 'MODAL', 'HORAS_DIARIAS', 'PATRIMONIOS']]
                    .assign(AUX=1),
                    on='AUX',
                    how='left',
                    suffixes=('_I', '_J')
                )
                .drop(columns='AUX')
            )

            # Distância haversine aproximada (plano) em km entre centros dos livros
            sim_df["DIST"] = np.sqrt(
                    ((sim_df["LAT_J"] - sim_df["LAT_I"]) * 111.32) ** 2 +
                    ((sim_df["LON_J"] - sim_df["LON_I"]) * 111.32 *
                    np.cos(np.radians((sim_df["LAT_I"] + sim_df["LAT_J"]) / 2))) ** 2
                ).round(2)

            sim_df = sim_df.drop(columns = ['LAT_I','LAT_J','LON_I', 'LON_J'])

            # Normaliza “patrimônios em comum” e cria ordenação de emparelhamento
            sim_df = sim_df.merge(patr_sim, on=['LIVRO_I', 'LIVRO_J'], how='left').fillna(0)
            sim_df['PATR_COMUM'] = np.maximum(sim_df['PATR_COMUM']/sim_df['PATRIMONIOS_I'], sim_df['PATR_COMUM']/sim_df['PATRIMONIOS_J'])
            sim_df['SORT_2'] = (sim_df['HORAS_DIARIAS_I'] != sim_df['HORAS_DIARIAS_J'])
            sim_df['SORT_1'] = (sim_df['MODAL_J'] != sim_df['MODAL_I'])
            sim_df['SORT_4'] = -sim_df['HORAS_DIARIAS_I']
            sim_df['SORT_3'] = (sim_df['MODAL_I'] != 'Moto/Carro')
            sim_df['SORT_5'] = -(sim_df['PATR_COMUM'])
            sim_df['SORT_6'] = (sim_df['DIST'])
            sim_df = (
                sim_df[(sim_df['DIST'] <= max_dist) & (sim_df['PERIODO_I'] != sim_df['PERIODO_J'])]
                .sort_values(by=['SORT_1', 'SORT_2', 'SORT_3', 'SORT_4', 'SORT_5', 'SORT_6'], ascending=True)
                .drop(columns=['SORT_1', 'SORT_2', 'SORT_3', 'SORT_4', 'SORT_5', 'SORT_6'])
            )

            # Heurística de alocação: percorre livros base e “pega” melhor contraparte de outro dia
            x = {i:{j:0 for j in livros} for i in abastecedores}
            y = {i:0 for i in abastecedores}

            h_list = []
            for h in sim_df['LIVRO_I'].to_list():
                if not(h in h_list):
                    h_list.append(h)

            pivot_livro = {j:i for i,j in pivot.items()}
            not_alocated = [j for j in livros]
            for h in h_list:
                i = pivot_livro[h]
                if h in not_alocated:
                    x[i][h] = 1
                    y[i] = 1
                    not_alocated.remove(h)
                    p_list = [p for p in periodos if p != periodo[h]]
                    for p in p_list:
                        sim_list = sim_df[(sim_df['LIVRO_I'] == h) & (sim_df['LIVRO_J'].isin(not_alocated)) & (sim_df['PERIODO_J'] == p)]['LIVRO_J'].to_list()
                        if len(sim_list) > 0:
                            j = sim_list[0]
                            x[i][j] = 1
                            not_alocated.remove(j)

            # Nomeia abastecedores sequencialmente por filial/supervisor
            n = 1
            for i, value_y in y.items():
                if value_y == 1:
                    abast_name = f'ABASTECEDOR {filial} {supervisor} {n}'
                    for j, value_x in x[i].items():
                        if value_x == 1:
                            alocation_dict[j] = abast_name
                    n += 1
                    
            # Converte para dict abastecedor -> [rotas]
            abastecedor_dict = defaultdict(list)
            for rota, abastecedor in alocation_dict.items():
                abastecedor_dict[abastecedor].append(rota)
            abastecedor_rotas = dict(abastecedor_dict)

    else:
        # Modo 1:1: abastecedor é derivado do livro agregado (depara_livro)
        depara_livro['ABASTECEDOR'] = 'ABASTECEDOR ' + depara_livro['LIVRO_AGG'].str.extract(r'(\w+) - .*L(\d+)').agg(' '.join, axis=1)
        alocation_dict = depara_livro.set_index('LIVRO')['ABASTECEDOR'].to_dict()
        
        # Constrói dict abastecedor -> lista de livros (para facilitar rotulagem/relatório)
        abastecedor_dict={}
        for _,item in depara_livro.iterrows():
            abastecedor = item["ABASTECEDOR"]
            livro = item["LIVRO"]
    
            if abastecedor not in abastecedor_dict:
                abastecedor_dict[abastecedor] = []
            abastecedor_dict[abastecedor].append(livro)

    # Ajustes finais de escala/modal por abastecedor (se alguma rota do pacote for Full-Time)
    for livros in abastecedor_dict.values():
        moto_carro=False
        full_time=False
        moto_carro = xl_result_df.loc[xl_result_df["Livro"].isin(livros), "Modal de Transporte"].eq("Moto/Carro").any()
        full_time = xl_result_df.loc[xl_result_df["Livro"].isin(livros), "Escala Requerida"].eq("Full-Time").any()

        # Se qualquer rota do pacote demanda Full-Time, propaga Full-Time ao pacote
        if full_time:
            xl_result_df.loc[xl_result_df["Livro"].isin(livros), "Escala Requerida"] = "Full-Time"
            xl_result_livros_df.loc[xl_result_livros_df["Livro"].isin(livros), "Escala Requerida"] = "Full-Time"                
            xl_result_livros_df.loc[xl_result_livros_df["Livro"].isin(livros), "Horas Diárias"] = 8
            result_livros_df.loc[result_livros_df["LIVRO"].isin(livros), "HORAS_DIARIAS"] = 8
            result_livros_df.loc[result_livros_df["LIVRO"].isin(livros), "ESCALA"] = "Full-Time"
            result_livros_df.loc[result_livros_df["LIVRO"].isin(livros), "FTE"] = 1

    # Substitui tempos de deslocamento do relatório detalhado com valores “Moto/Carro” do backup
    mapa_tempo = backup_result_df.set_index(['LIVRO', 'PATRIMONIO'])['TEMPO_DESLOCAMENTO']
    condicao = xl_result_df['Modal de Transporte'] == 'Moto/Carro'
    xl_result_df.loc[condicao, 'Tempo de Deslocamento (min)'] = xl_result_df[condicao].set_index(['Livro', 'Patrimônio']).index.map(mapa_tempo)
    
    # Soma deslocamento por livro para refletir ajuste acima também no resumo
    tempo_por_livro = xl_result_df.groupby('Livro')['Tempo de Deslocamento (min)'].sum()
    xl_result_livros_df['Tempo de Deslocamento (min)'] = xl_result_livros_df['Livro'].map(tempo_por_livro)

    # Mapeia livro -> abastecedor (dicionário final de alocação)
    result_livros_df['ABASTECEDOR'] = result_livros_df['LIVRO'].map(alocation_dict)

    # Diagnóstico: quantos abastecedores excedem 44h semanais (soma de operação)
    print("Abastecedores com mais de 44 hrs semanais:", len(result_livros_df
          .assign(TEMPO_OPERACAO=(result_livros_df['TEMPO_SERVICO']+result_livros_df['TEMPO_DESLOCAMENTO'])/3600)
          .groupby(['ABASTECEDOR'])
          .agg({'TEMPO_OPERACAO':'sum'})
          .reset_index()
          .loc[lambda x: x['TEMPO_OPERACAO'] > 44]))
    
    # Uma linha por abastecedor com a escala/modal predominantes
    escala_df = result_livros_df.copy()
    escala_df['SORT_MODAL'] = escala_df['MODAL'] == 'Moto/Carro'
    escala_df = (escala_df
                 .sort_values(by=['SORT_MODAL', 'HORAS_DIARIAS'], ascending=False)
                 [['ABASTECEDOR', 'ESCALA', 'MODAL', 'HORAS_DIARIAS', 'FTE']]
                 .drop_duplicates(subset=['ABASTECEDOR']))

    # Pivot de alocação: linhas por abastecedor e colunas por período
    periodos = [i for i in range(1, n_p+1) if i in result_livros_df['PERIODO'].unique()]

    alocation_df = result_livros_df.pivot(index=['FILIAL','ABASTECEDOR'], columns='PERIODO', values='LIVRO').fillna('').reset_index().sort_values(by=['FILIAL','ABASTECEDOR'])
    alocation_df = alocation_df.merge(escala_df, on='ABASTECEDOR', how='left')[['FILIAL','ABASTECEDOR', 'ESCALA', 'MODAL', 'HORAS_DIARIAS', 'FTE'] + periodos]
    alocation_df.columns = ['Filial','Abastecedor', 'Escala Requerida', 'Modal', 'Horas Diárias', 'FTE'] + periodos

    # Tabela de alocação de patrimônios por abastecedor (marcação por dia)
    result_df['ABASTECEDOR'] = result_df['LIVRO'].map(alocation_dict)
    patrimonios_aloc = result_df[result_df['VISITA'] != 0].copy()

    patrimonios_aloc['ALOCACAO'] = 'X'
    patrimonios_aloc = (
        patrimonios_aloc
        .pivot(index=['FILIAL','ABASTECEDOR', 'PARCEIRO', 'PATRIMONIO'], columns='PERIODO', values='ALOCACAO')
        .fillna('').reset_index().sort_values(by=['FILIAL','ABASTECEDOR', 'PARCEIRO', 'PATRIMONIO'])
    )
    patrimonios_aloc.columns = ['Filial','Abastecedor', 'Parceiro', 'Patrimônio'] + periodos
    
    print('Processamento de resultados finalizado, tempo: {:.1f}s'.format(time.time() - section_start))
    
    # =====================
    # 5. Escrita em Excel (abas: Livros, Resumo de Livros, Alocação Sugerida)
    # =====================
    section_start = time.time()
    print('Salvando resultados...')

    # ---------------------
    # 5.1. Livros (detalhado)
    # ---------------------
    sheet = wb.sheets['Livros']  # Seleciona a aba específica

    # Limpa e escreve
    sheet.range("6:1048576").clear_contents()
    sheet.range("7:1048576").api.ClearFormats()
    sheet['B6'].options(index=False).value = xl_result_df

    rows = len(result_df)

    # Formatação (fonte/estilo)
    cell_range = sheet.range(f"B6:M{6 + rows}")
    cell_range.api.Font.Name = "Arial Narrow"

    cell_range = sheet.range(f"B7:I{6 + rows}")
    cell_range.number_format = "@"  # Define o formato de texto

    # Linhas auxiliares para bordas horizontais por mudança de livro/dia/filial
    livro_list = xl_result_df['Livro'].to_list()
    dia_list = xl_result_df['Dia'].to_list()
    filial_list = xl_result_df['Filial'].to_list()

    row = 7
    for i in range(len(livro_list)):

        if (i == (len(livro_list)-1)):
            pass
        
        elif (filial_list[i] != filial_list[i + 1]):
            cell_range = sheet.range(f"B{row}:M{row}")
            cell_range.api.Borders(9).Weight = 4  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xa5a5a5  # Preto
        
        elif dia_list[i] != dia_list[i+1]:
            cell_range = sheet.range(f"B{row}:M{row}")
            cell_range.api.Borders(9).Weight = 4  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xD2D2D2  # Preto

        elif livro_list[i] != livro_list[i+1]:
            cell_range = sheet.range(f"B{row}:M{row}")
            cell_range.api.Borders(9).Weight = 2  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xD2D2D2  # Preto

        row+=1

    # ---------------------
    # 5.2. Resumo de Livros
    # ---------------------
    sheet = wb.sheets['Resumo de Livros']  # Seleciona a aba específica

    sheet.range("6:1048576").clear_contents()
    sheet.range("7:1048576").api.ClearFormats()
    sheet['B6'].options(index=False).value = xl_result_livros_df

    rows = len(result_df)

    cell_range = sheet.range(f"B6:L{6 + rows}")
    cell_range.api.Font.Name = "Arial Narrow"

    cell_range = sheet.range(f"B7:F{6 + rows}")
    cell_range.number_format = "@"  # Define o formato de texto

    dia_list = xl_result_livros_df['Dia'].to_list()
    filial_list = xl_result_livros_df['Filial'].to_list()

    row = 7
    for i in range(len(dia_list)):

        if (i == (len(dia_list)-1)):
            pass
        
        elif (filial_list[i] != filial_list[i + 1]):
            cell_range = sheet.range(f"B{row}:L{row}")
            cell_range.api.Borders(9).Weight = 4  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xD2D2D25  # Preto
        
        elif dia_list[i] != dia_list[i+1]:
            cell_range = sheet.range(f"B{row}:L{row}")
            cell_range.api.Borders(9).Weight = 2  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xD2D2D2  # Preto

        row+=1

    # ---------------------
    # 5.3. Alocação Sugerida (Livros)
    # ---------------------
    sheet = wb.sheets['Alocação Sugerida (Livros)']  # Seleciona a aba específica

    sheet.range("6:1048576").clear_contents()
    sheet.range("7:1048576").api.ClearFormats()
    sheet['B6'].options(index=False).value = alocation_df

    rows = len(alocation_df)
    cols = len(alocation_df.columns)

    # Cabeçalho
    cell_range = sheet.range((6, 2), (6, 1 + cols))
    cell_range.api.Font.Name = "Arial Narrow"
    cell_range.number_format = "@"  # Define o formato de texto
    cell_range.api.Font.Bold = True

    borders = cell_range.api.Borders
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    cell_range.api.Interior.Color = 0xE6E6E6

    # Corpo
    cell_range = sheet.range((7, 2), (6 + rows, 1 + cols))
    cell_range.number_format = "@"  # Define o formato de texto
    cell_range.api.Font.Name = "Arial Narrow"
    
    borders = cell_range.api.Borders
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xD2D2D2  # Preto

    # Linha divisória por mudança de filial
    filial_list = alocation_df['Filial'].to_list()

    row = 7
    for i in range(len(filial_list)):
        if (i == (len(filial_list)-1)):
            pass
        elif (filial_list[i] != filial_list[i + 1]):
            cell_range = sheet.range((row, 2), (row, 1 + cols))
            cell_range.api.Borders(9).Weight = 4  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xD2D2D25  # Preto
        row+=1

    # ---------------------
    # 5.4. Alocação Sugerida (Patrimônios)
    # ---------------------
    sheet = wb.sheets['Alocação Sugerida (Patrimônios)']  # Seleciona a aba específica

    sheet.range("6:1048576").clear_contents()
    sheet.range("7:1048576").api.ClearFormats()
    sheet['B6'].options(index=False).value = patrimonios_aloc

    rows = len(patrimonios_aloc)
    cols = len(patrimonios_aloc.columns)

    # Cabeçalho
    cell_range = sheet.range((6, 2), (6, 1 + cols))
    cell_range.api.Font.Name = "Arial Narrow"
    cell_range.number_format = "@"  # Define o formato de texto
    cell_range.api.Font.Bold = True

    borders = cell_range.api.Borders
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    cell_range.api.Interior.Color = 0xE6E6E6

    # Corpo
    cell_range = sheet.range((7, 2), (6 + rows, 1 + cols))
    cell_range.number_format = "@"  # Define o formato de texto
    cell_range.api.Font.Name = "Arial Narrow"
    
    borders = cell_range.api.Borders
    borders(11).Weight = 3  # xlEdgeLeft, grossa
    borders(11).Color = 0xFFFFFF  # Branco
    borders(9).Weight = 2  # xlEdgeBottom, fina
    borders(9).Color = 0xD2D2D2  # Preto

    parceiro_list = patrimonios_aloc['Parceiro'].to_list()
    filial_list = patrimonios_aloc['Filial'].to_list()

    row = 7
    for i in range(len(parceiro_list)):

        if (i == (len(parceiro_list)-1)):
            pass
        
        elif (filial_list[i] != filial_list[i + 1]):
            cell_range = sheet.range((row, 2), (row, 1 + cols))
            cell_range.api.Borders(9).Weight = 4  # xlEdgeBottom, fina
            cell_range.api.Borders(9).Color = 0xD2D2D25  # Preto

        row+=1

    # ---------------------
    # 5.5. Persistência em parquet (auditoria/uso downstream)
    # ---------------------
    result_df.to_parquet(model_data_folder + 'result_livros.parquet', index=False)
    result_livros_df.to_parquet(model_data_folder + 'result_livros_resumo.parquet', index=False)
    alocation_df.to_parquet(model_data_folder + 'alocacao.parquet', index=False)
    patrimonios_aloc.to_parquet(model_data_folder + 'alocacao_patrimonios.parquet', index=False)

    print('Resultados Salvos, tempo: {:.1f}s'.format(time.time() - section_start))
    print('Otimização de Lotes Encerrada, tempo total: {:.1f}s\nPrecione Enter para continuar...'.format(time.time() - start))
    
    # Encerra o workbook conforme modo
    if developer:
        wb.close()
    else:
        input()

    # =====================
    # 6. Retorno opcional (métricas resumidas para “rodada”)
    # =====================
    if output_rodada:
        # Consolida quantidade de patrimônios por frequência
        df_aux = freq_df.groupby('FREQUENCIA').agg({'PATRIMONIO':'nunique'}).reset_index()
        dict_result = dict(zip(df_aux['FREQUENCIA'], df_aux['PATRIMONIO']))

        # Métricas de tempo (min) e headcount (por tipo de escala)
        tempo_deslocamento = xl_result_df['Tempo de Deslocamento (min)'].sum()
        tempo_servico = xl_result_df['Tempo de Serviço (min)'].sum()
        n_ftes = len(alocation_df.loc[alocation_df['Escala Requerida'] == 'Full-Time'])
        n_part_time = len(alocation_df.loc[alocation_df['Escala Requerida'] == 'Part-Time'])
        n_rpa3 = len(alocation_df.loc[alocation_df['Escala Requerida'] == 'RPA 3H'])
        n_rpa2 = len(alocation_df.loc[alocation_df['Escala Requerida'] == 'RPA 2H'])

        # Nota: multiplicação por 4 preservada conforme original
        return {
            'tempo_deslocamento':float(tempo_deslocamento)/60*4,
            'tempo_servico':float(tempo_servico)/60*4,
            'tempo_operação': float(tempo_deslocamento+tempo_servico)/60*4,
            'n_ftes':n_ftes,
            'n_part_time':n_part_time,
            'n_rpa3':n_rpa3,
            'n_rpa2':n_rpa2,
            'n_pat':dict_result}


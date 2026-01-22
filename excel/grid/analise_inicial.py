from os import getenv
from dotenv import load_dotenv
import pymysql
import pandas as pd
from datetime import date, datetime
import pyodbc
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name
from carrega_nfs import main as carrega_nfs
import numpy as np

def busca_nfs_mes_anterior(
    data_inicio: date = None, 
    data_fim: date = None, 
    margem_parceiros: float = 0.15, 
    margem: float = 0.27, 
    margem_maxima: float = 0.5
) -> pd.DataFrame | None:
    """Realiza busca das Notas fiscais que estão fora dos parâmetros de margem em determinado período

    Args:
        data_inicio (date): data de início do período. Default é 1º dia do mês anterior.
        data_fim (date): data de fim do período. Default é o último dia do mês anterior.
        margem_parceiros (float): parâmetro de margem para parceiros. Default é 0.15.
        margem (float): parâmetro de margem geral. Default é 0.27.
        margem_maxima (float): parâmetro de margem máxima para investigar margens muito altas. Default é 0.50.

    Returns:
        pd.DataFrame: Dataframe Pandas com as NFs que estão fora dos parâmetros.
        None: em caso de erro na execução da query
    """ 
    if not data_inicio:
        data_inicio = date.today().replace(day=1).fromordinal(date.today().replace(day=1).toordinal() - 1).replace(day=1).strftime('%Y-%m-%d')
    else:
        data_inicio = data_inicio.strftime('%Y-%m-%d')
        
    if not data_fim:
        data_fim = date.today().replace(day=1).fromordinal(date.today().replace(day=1).toordinal() - 1).strftime('%Y-%m-%d')
    else:
        data_fim = data_fim.strftime('%Y-%m-%d')
        
        
        
    print(f'Analisando NFs do período: {data_inicio} - {data_fim}')
    
    try:
        with pymysql.connect(
                host     = MYSQL_DB_HOST,
                user     = MYSQL_DB_USER,
                password = MYSQL_DB_PASSWORD,
                database = MYSQL_DB_DATABASE,
                port     = MYSQL_DB_PORT
            ) as con:
               with con.cursor() as cursor:
                    with open('grid/querys/busca_nfs_mes_anterior.sql', 'r', encoding='utf-8') as query:
                        # values = (data_inicio, data_fim, str(margem_parceiros), str(margem), str(margem_maxima))
                        values = (data_inicio, data_fim)
                        consulta = query.read()
                        cursor.execute(consulta, values)
                        rows = cursor.fetchall()
                        if rows:
                            columns = [desc[0] for desc in cursor.description]
                            return pd.DataFrame(rows, columns=columns)
                        else:
                            return pd.DataFrame()
    except Exception as e:
        print("Erro em busca_nfs_mes_anterior:", e)
        raise e

def busca_clientes_parceiros():
    try:
        with pymysql.connect(
                host     = MYSQL_DB_HOST,
                user     = MYSQL_DB_USER,
                password = MYSQL_DB_PASSWORD,
                database = MYSQL_DB_DATABASE,
                port     = MYSQL_DB_PORT
            ) as con:
                with con.cursor() as cursor:
                    with open('grid/querys/busca_clientes_parceiros.sql', 'r', encoding='utf-8') as query:
                        consulta = query.read()
                        cursor.execute(consulta)
                        rows = cursor.fetchall()
                        if rows:
                            columns = [desc[0] for desc in cursor.description]
                            return pd.DataFrame(rows, columns=columns)
                        else:
                            return pd.DataFrame()
    except Exception as e:
        print("Erro em busca_clientes_parceiros:", e)
        raise e
    

def formatar_aba(writer, df, sheet_name):
    if df.empty:
        return

    # Limpar os tipos
    df_clean = df.copy()
    for col in df_clean.columns:
        if pd.api.types.is_numeric_dtype(df_clean[col]):
            df_clean[col] = df_clean[col].astype(float)
    
    # Escreve os dados
    df_clean.to_excel(writer, sheet_name=sheet_name, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Formatos
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
    fmt_currency = workbook.add_format({'num_format': 'R$ #,##0.00'}) 
    fmt_percent = workbook.add_format({'num_format': '0.00%'})
    
    # Formato de Destaque (Apenas a cor de fundo)
    fmt_destaque = workbook.add_format({'bg_color': '#FFEB9C'})
    
    # 1. Aplica formatação de colunas (Largura e Números)
    for idx, col in enumerate(df_clean.columns):
        series = df_clean[col]
        max_len = max((series.astype(str).map(len).max(), len(str(col)))) + 2
        col_name_lower = col.lower()
        cell_format = None
        
        if any(x in col_name_lower for x in ['valor_', 'custo', 'vlr_', 'margem_bruta']): # Ajustei para pegar variações comuns
            cell_format = fmt_currency
        elif any(x in col_name_lower for x in ['percent', 'aliquota', 'diff_percentual']):
            cell_format = fmt_percent
            
        worksheet.set_column(idx, idx, max_len, cell_format)
        worksheet.write(0, idx, col, fmt_header)
    
    # 2. Lógica para colorir a linha inteira se tp_movimento == '010'
    if 'tp_movimento' in df_clean.columns:
        # Pega o índice numérico da coluna (ex: 3)
        col_idx = df_clean.columns.get_loc('tp_movimento')
        # Transforma em letra (ex: 'D') para usar na fórmula do Excel
        col_letter = xl_col_to_name(col_idx)
        
        # Define o intervalo total dos dados (começando da linha 1, excluindo cabeçalho)
        last_row = len(df_clean)
        last_col = len(df_clean.columns) - 1
        
        # Fórmula: =$D2="010" 
        # O $ antes da letra trava a coluna, garantindo que a linha inteira seja pintada
        # O 2 (sem $) indica que a verificação começa na segunda linha (pós cabeçalho) e desce
        formula = f'=${col_letter}2="010"'
        
        worksheet.conditional_format(1, 0, last_row, last_col, {
            'type':     'formula',
            'criteria': formula,
            'format':   fmt_destaque
        })

    worksheet.autofilter(0, 0, len(df_clean), len(df_clean.columns) - 1)

# def limpar_dados_para_excel(df):
#     """Converte colunas numéricas para float e strings para string limpa"""
#     if df.empty:
#         return df
        
#     for col in df.columns:
#         # Se parecer número (custo, preco, margem, diff), força conversão
#         col_lower = col.lower()
#         if any(x in col_lower for x in ['custo', 'preco', 'valor', 'total', 'margem', 'diff', 'percent']):
#             # Converte para numérico, transformando erros em NaN
#             df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
#     return df


if __name__ == '__main__':
    
    load_dotenv(override=True)
    
    MYSQL_DB_HOST = getenv('MYSQL_DB_HOST')
    MYSQL_DB_USER = getenv('MYSQL_DB_USER')
    MYSQL_DB_PASSWORD = getenv('MYSQL_DB_PASSWORD')
    MYSQL_DB_DATABASE = getenv('MYSQL_DB_DATABASE')
    MYSQL_DB_PORT = int(getenv('MYSQL_DB_PORT'))
    
    connectionString = f"DRIVER={getenv('PROTHEUS_ODBC_DRIVER')};SERVER={getenv('PROTHEUS_DB_HOST')};DATABASE={getenv('PROTHEUS_DB_DATABASE')};UID={getenv('PROTHEUS_DB_USER')};PWD={getenv('PROTHEUS_DB_PASSWORD')};TrustServerCertificate=yes"

    carrega_nfs()
    
    try:
        
        # nfs_mes_anterior = busca_nfs_mes_anterior(data_inicio=date(year=2025, month=9, day=1), data_fim=date(year=2025, month=9, day=30))
        
        nfs_mes_anterior = busca_nfs_mes_anterior()
        
        if nfs_mes_anterior is None or nfs_mes_anterior.empty:
            print("Nenhuma NF encontrada no período.")
        else:
            print(f"{len(nfs_mes_anterior)} NFs encontradas. Iniciando processamento...")
            
            df_revendas_final = pd.DataFrame()
            df_vendas_final = pd.DataFrame()

            # Conversão de tipos gerais
            cols_to_float = ['valor_contabil', 'custo', 'valor_unitario', 'valor_ipi', 'valor_imp5', 'valor_imp6', 'vlr_icms_difal', 'valor_icms', 'margem', 'margem_bruta', 'aliq_icms']
            for col in cols_to_float:
                if col in nfs_mes_anterior.columns:
                     nfs_mes_anterior[col] = pd.to_numeric(nfs_mes_anterior[col], errors='coerce')

            nfs_mes_anterior['aliq_icms'] = nfs_mes_anterior.apply(
                lambda x: x['aliq_icms'] / 100, 
                axis=1
            )
            
            # # Separação
            # vendas = nfs_mes_anterior[nfs_mes_anterior['cfop'].str.contains('5101|6101|5116|6116|6107', na=False)].copy()
            
            
            # # ========================================================
            # #                 Análise das Vendas (Produção)
            # # ========================================================
            # if not vendas.empty:
            #     produtos_venda = vendas['cod_produto'].unique().copy()
                
            #     with pyodbc.connect(connectionString) as con:
            #         with con.cursor() as cursor:
            #             with open('querys/preco_base.sql','r', encoding='utf-8') as query:
            #                 sql_query = query.read()
            #             precos_base = pd.read_sql(sql_query, con)
                
            #     precos_base = pd.DataFrame(precos_base, columns=['cod_produto', 'preco_base', 'tabela'])
                
            #     # Prepara tabela de preços
            #     precos_filtrados = precos_base[precos_base['cod_produto'].isin(produtos_venda)].copy()
            #     precos_filtrados['preco_base'] = pd.to_numeric(precos_filtrados['preco_base'], errors='coerce')

            #     # Merge
            #     vendas_analise = vendas.merge(precos_filtrados, on='cod_produto', how='left')
            #     vendas_analise['preco_base'] = vendas_analise['preco_base'].fillna(0.0)

            #     # Cálculos
            #     vendas_analise['diff_valor'] = vendas_analise['custo'] - vendas_analise['preco_base']
            #     vendas_analise['diff_percentual'] = vendas_analise.apply(
            #         lambda x: ((x['custo'] - x['preco_base']) / x['preco_base']) if x['preco_base'] > 0 else 0.0,
            #         axis=1
            #     )

            #     # Seleção de colunas
            #     cols_venda = ['nota', 'emissao', 'cliente', 'cod_produto', 'produto', 'tabela', 'custo', 'preco_base', 'diff_valor', 'diff_percentual']
            #     cols_finais_venda = [c for c in cols_venda if c in vendas_analise.columns]
            #     df_vendas_final = vendas_analise[cols_finais_venda]

            # ========================================================
            #                 Geração do Excel
            # ========================================================
            mes_ano_analise = date.today().replace(day=1).fromordinal(date.today().replace(day=1).toordinal() - 1).strftime('%Y-%m')
            nome_arquivo = f'analise_margem_grid_{mes_ano_analise}.xlsx'
            
            print(f"Gerando arquivo Excel: {nome_arquivo}...")
            
            # Preparação (Busca parceiros e cria chave match)
            set_parceiros = busca_clientes_parceiros()
            
            nfs_mes_anterior['chave_match'] = (
                nfs_mes_anterior['cod_cliente'].astype(str).str.strip() + 
                nfs_mes_anterior['loja'].astype(str).str.strip()
            )
            
            # ================================================================
            #                Status da Margem
            # ================================================================

            # Lista de produtos para ignorar (exceções)
            produtos_excecao = ['B0010046', 'E000H2P8']

            # Cria uma condição: Verdadeiro se o produto NÃO for uma exceção
            cond_produto_valido = ~nfs_mes_anterior['cod_produto'].isin(produtos_excecao)

            # --- Lógica de Identificação ---
            is_parceiro = nfs_mes_anterior['chave_match'].isin(set_parceiros)

            # Definição das regras isoladas
            cond_alta = (nfs_mes_anterior['margem_bruta_percentual'] > 0.50) 
            cond_baixa_parceiro = (is_parceiro) & (nfs_mes_anterior['margem_bruta_percentual'] < 0.17) 
            cond_baixa_comum = (~is_parceiro) & (nfs_mes_anterior['margem_bruta_percentual'] < 0.27) 

            # --- Combinação Final (Priorização) ---

            # Lista de Condições (a ordem importa: o Python avalia a primeira que for True)
            condicoes = [
                # 1. Caso Margem Alta (e produto válido)
                (cond_alta & cond_produto_valido),
                
                # 2. Caso Margem Baixa (Parceiro OU Comum) (e produto válido)
                ((cond_baixa_parceiro | cond_baixa_comum) & cond_produto_valido)
            ]

            # Lista de Resultados respectivos
            resultados = [
                'MARGEM ALTA',  # Resultado para condição 1
                'VERIFICAR'     # Resultado para condição 2
            ]

            # Aplica a lógica. Se nenhuma condição for atendida, usa o default 'OK'
            nfs_mes_anterior['status_margem'] = np.select(condicoes, resultados, default='OK')
                      
            
            with pd.ExcelWriter(f'grid/{nome_arquivo}', engine='xlsxwriter') as writer:
                workbook = writer.book
                # ================================================================
                #                Opções de Justificativa de Margem
                # ================================================================
                    
                # Cria a aba
                ws_dados = workbook.add_worksheet('Config')

                # --- Formatações (Estilos) ---
                # Estilo para o cabeçalho da lista (Azul escuro)
                fmt_header = workbook.add_format({
                    'bold': True,
                    'font_color': 'white',
                    'bg_color': '#44546A',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter'
                })

                # Estilo para os itens da lista (Borda simples)
                fmt_item = workbook.add_format({
                    'valign': 'vcenter'
                })

                # Estilo para o título das instruções (Azul e maior)
                fmt_info_title = workbook.add_format({
                    'bold': True,
                    'font_size': 12,
                    'font_color': '#2F75B5',
                    'valign': 'bottom'
                })

                # Estilo para o corpo do texto de instrução (Cinza e com quebra de linha)
                fmt_info_text = workbook.add_format({
                    'font_color': '#404040',
                    'text_wrap': True,
                    'valign': 'top',
                    'align': 'left'
                })

                # --- Dados ---
                opcoes = [
                    'OK. Margem Parceiro', 
                    'Tabela de Preço Errada', 
                    'Refaturado com Valor Correto',
                    'Desconto aplicado para concluir a venda',
                    'Venda Cancelada',
                ]

                # --- Escrita na Planilha ---

                # 1. Configura largura das colunas
                ws_dados.set_column('A:A', 40)  # Coluna da lista mais larga
                ws_dados.set_column('B:B', 2)   # Separador visual estreito
                ws_dados.set_column('C:F', 15)  # Espaço para o texto explicativo

                # 2. Escreve a Lista de Opções (Começando na linha 1 para deixar a 0 para cabeçalho)
                ws_dados.write('A1', "Justificativas Cadastradas", fmt_header)
                ws_dados.write_column('A2', opcoes, fmt_item)

                # 3. Escreve as Instruções ao lado (Colunas C a F)
                ws_dados.write('C2', "ℹ️ Como funciona esta lista?", fmt_info_title)

                texto_instrucao = (
                    "As opções listadas na coluna A alimentam automaticamente o menu "
                    "de seleção na aba de dados.\n\n"
                    "1. Análises de Margem: Todas as opções aqui são consideradas válidas como justificativas na tabela Todas NFs para vendas fora da margem.\n"
                    "2. Estatísticas: A aba 'Resumo' calcula os indicadores baseada nestes textos.\n\n"
                    "Para adicionar uma nova justificativa, basta escrever na próxima linha vazia "
                    "da coluna A."
                )

                # Mescla células para criar uma "caixa de texto" limpa
                ws_dados.merge_range('C3:F12', texto_instrucao, fmt_info_text)

                # (Opcional) Nota visual de rodapé
                ws_dados.write('C13', "Alterações aqui refletem imediatamente na validação de dados.", 
                            workbook.add_format({'italic': True, 'font_color': '#7F7F7F', 'font_size': 9}))
                
                
                # ================================================================
                # Todas as Notas Fiscais
                # ================================================================
                
                sheet_name = "Todas NFs"
                
                # Separa dados para exportar (sem a coluna auxiliar)
                cols_export = [c for c in nfs_mes_anterior.columns if c != 'chave_match']
                df_export = nfs_mes_anterior[cols_export]
                
                # 1. Escreve os dados base
                formatar_aba(writer, df_export, sheet_name)
                
                worksheet = writer.sheets[sheet_name]
                
                # 2. Definição de Cores
                cor_ruim = '#FFC7CE'   # Vermelho claro
                texto_ruim = '#9C0006' # Vermelho escuro
                cor_alta = '#FFEB9C'   # Amarelo claro
                texto_alta = '#9C6500' # Amarelo escuro

                # 3. Definição dos Formatos Combinados (Fundo + Tipo de Dado)
                # Precisamos criar variações para manter a formatação de R$ e %
                
                formats = {
                    'red': {
                        'geral': workbook.add_format({'bg_color': cor_ruim, 'font_color': texto_ruim}),

                        'money': workbook.add_format({'bg_color': cor_ruim, 'font_color': texto_ruim, 'num_format': 'R$ #,##0.00'}),
                        'percent': workbook.add_format({'bg_color': cor_ruim, 'font_color': texto_ruim, 'num_format': '0.00%'}),
                        'date': workbook.add_format({'bg_color': cor_ruim, 'font_color': texto_ruim, 'num_format': 'dd/mm/yyyy'})
                    },
                    'yellow': {
                        'geral': workbook.add_format({'bg_color': cor_alta, 'font_color': texto_alta}),
                        'money': workbook.add_format({'bg_color': cor_alta, 'font_color': texto_alta, 'num_format': 'R$ #,##0.00'}),
                        'percent': workbook.add_format({'bg_color': cor_alta, 'font_color': texto_alta, 'num_format': '0.00%'}),
                        'date': workbook.add_format({'bg_color': cor_alta, 'font_color': texto_alta, 'num_format': 'dd/mm/yyyy'})
                    },
                    'blank': {
                        'geral': workbook.add_format({}),
                        'money': workbook.add_format({'num_format': 'R$ #,##0.00'}),
                        'percent': workbook.add_format({'num_format': '0.00%'}),
                        'date': workbook.add_format({'num_format': 'dd/mm/yyyy'})
                    }
                }

                # 4. Iteração Linha a Linha
                try:
                    # Precisamos dos nomes das colunas para saber qual formato aplicar em cada célula
                    colunas = df_export.columns.tolist()
                    lotes_geradores_vendidos = []
                    revendas = []
                    abaixo_margem = 0
                    # Itera sobre o DataFrame de exportação (alinhado com as colunas escritas)
                    # e usa o DataFrame original para metadados como "chave_match" e "lote".
                    for row_idx, (orig_row, export_row) in enumerate(zip(nfs_mes_anterior.itertuples(index=False), df_export.itertuples(index=False)), start=1):
                        margem = getattr(orig_row, 'margem_bruta_percentual', 0)
                        parceiro = getattr(orig_row, 'chave_match', '') in set_parceiros
                        cod_produto = getattr(orig_row, 'cod_produto', None)
                        meta_minima = 0.17 if parceiro else 0.27
                                                
                        # Decide a cor da linha
                        tipo_destaque = 'blank' # 'red' ou 'yellow' ou None
                        
                        if cod_produto not in ('B0010046','E000H2P8'):
                            if margem < meta_minima:
                                tipo_destaque = 'red'
                                # lotes_geradores_vendidos.append(getattr(orig_row, 'lote', ''))
                                abaixo_margem += 1
                                
                                if not str(getattr(orig_row, 'cod_produto', '')).startswith('G0') \
                                    and getattr(orig_row, 'cfop', '') not in ('5101', '6101', '5116','6116','6107'):
                                    revendas.append(orig_row[:-2])
                                
                            elif margem > 0.50:
                                tipo_destaque = 'yellow'
                                # lotes_geradores_vendidos.append(getattr(orig_row, 'lote', ''))
                                if not str(getattr(orig_row, 'cod_produto', '')).startswith('G0') \
                                    and getattr(orig_row, 'cfop', '') not in ('5101', '6101', '5116','6116','6107'):
                                    revendas.append(orig_row[:-2])
                        
                        # Se precisar pintar a linha, percorre todas as colunas (usando os valores de export_row)
                        if tipo_destaque:
                            dict_formatos = formats[tipo_destaque]
                            
                            for col_idx, col_name in enumerate(colunas):
                                # Pega o valor da célula atual do DataFrame que corresponde ao Excel (df_export)
                                valor_celula = export_row[col_idx]
                                
                                # Decide qual formato usar baseado no NOME da coluna (mesma lógica do formatar_aba)
                                col_lower = col_name.lower()
                                formato_final = dict_formatos['geral'] # Padrão
                                
                                if any(x in col_lower for x in ['margem_percentual', 'margem_bruta_percentual']):
                                    formato_final = dict_formatos['percent']
                                elif any(x in col_lower for x in ['valor_contabil', 'custo', 'valor_unitario', 'valor_ipi', 'valor_imp5', 'valor_imp6', 'vlr_icms_difal', 'valor_icms', 'margem', 'margem_bruta']):
                                    formato_final = dict_formatos['money']
                                elif any(x in col_lower for x in ['data_emissao']):
                                    formato_final = dict_formatos['date']
                                
                                # Sobrescreve a célula com o valor e o formato colorido
                                worksheet.write(row_idx, col_idx, valor_celula, formato_final)
                                
                            # =B2 - C2 - D2 - F2 - G2 - H2 - SE(OU(G{row_idx}="5101"; G{row_idx}="6101"; G{row_idx}="5116"; G{row_idx}="6116"; G{row_idx}="6107"); AC{row_idx}*0,02; AC{row_idx})
                                                        
                        formula_margem = (
                            f'=V{row_idx+1}-W{row_idx+1}-Y{row_idx+1}-Z{row_idx+1}-AA{row_idx+1}-AB{row_idx+1}-'
                            f'IF(OR(G{row_idx+1}="5101", G{row_idx+1}="6101", G{row_idx+1}="5116", G{row_idx+1}="6116", G{row_idx+1}="6107"), '
                            f'AC{row_idx+1}*0.02, AC{row_idx+1})'
                        )

                        # Escreve a fórmula da Margem (Coluna AE / Índice 30)
                        worksheet.write_formula(row_idx, 30, formula_margem, dict_formatos['money'])

                        # Escreve a fórmula da Porcentagem (Coluna AF / Índice 31)
                        worksheet.write_formula(row_idx, 31, f'=AE{row_idx+1}/V{row_idx+1}', dict_formatos['percent'])
                    
                    worksheet.write(0, 33, 'justificativa', workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1}))
                
                    worksheet.data_validation(1, 33, len(nfs_mes_anterior), 33, {
                        'validate': 'list',
                        'source': "='Config'!$A$2:$A$50",
                        'input_title': 'Selecione a Justificativa',
                        'input_message': 'Por favor, escolha uma opção da lista. As justificativas podem ser editadas na aba Config'
                    })
                    
                    worksheet.autofit()
                    
            
                    # ========================================================
                    #                 Análise das Revendas
                    # ========================================================
                    
                    # revendas = nfs_mes_anterior[~nfs_mes_anterior['cfop'].str.contains('5101|6101|5116|6116|6107', na=False)].copy()
                    
                    revendas = pd.DataFrame(revendas, columns = [
                            'filial',
                            'nota',
                            'no_pedido',
                            'vendedor',
                            'data_emissao',
                            'lote',
                            'cfop',
                            'cfop_descri',
                            'atualiza_estoque',
                            'gera_duplicata',
                            'cod_produto',
                            'produto',
                            'tipo_produto',
                            'armazem',
                            'cod_cliente',
                            'loja',
                            'cliente',
                            'grp_amar_ctb',
                            'classificacao_produto',
                            'estado_destino',
                            'quantidade',
                            'valor_contabil',
                            'custo',
                            'valor_unitario',
                            'valor_ipi',
                            'valor_imp5',
                            'valor_imp6',
                            'vlr_icms_difal',
                            'valor_icms',
                            'aliq_icms',
                            'margem_bruta',
                            'margem_bruta_percentual',
                        ])
                    
                    if not revendas.empty:
                        produtos_revenda = tuple(revendas['cod_produto'].unique())
                        
                        query_revenda = f"""SELECT TRIM(B1_COD) AS B1_COD, B1_UPRC FROM SB1010 WHERE B1_COD IN {produtos_revenda} AND D_E_L_E_T_  <> '*' AND B1_FILIAL = '01'"""
                        
                        with pyodbc.connect(connectionString) as con:
                            with con.cursor() as cursor:
                                cursor.execute(query_revenda)
                                precos = cursor.fetchall()
                        
                        if precos:
                            dict_precos = {p.B1_COD: float(p.B1_UPRC) for p in precos} # Garante float
                            revendas['ultimo_preco_compra'] = revendas['cod_produto'].map(dict_precos).fillna(0.0)
                            
                            # Cálculo de diferença
                            revendas['diff_valor'] = revendas['custo'] - revendas['ultimo_preco_compra']
                            revendas['diff_percentual'] = revendas.apply(
                                lambda x: ((x['custo'] - x['ultimo_preco_compra']) / x['ultimo_preco_compra']) if x['ultimo_preco_compra'] > 0 else 0.0, axis=1
                            )

                            # Seleciona colunas relevantes para o relatório
                            cols_revenda = ['filial', 'nota', 'emissao', 'cliente', 'cod_produto', 'produto', 'grp_amar_ctb', 'classificacao_produto', 'custo', 'ultimo_preco_compra', 'diff_valor', 'diff_percentual', 'vendedor']
                            # Filtra apenas colunas que existem no df original + as novas
                            cols_finais = [c for c in cols_revenda if c in revendas.columns]
                            df_revendas_final = revendas[cols_finais]

                    else:
                        print("Revendas está vazio")
                    # ================================================================    
                    # Aba 2: Análise Revenda (Comparativo com Compra)
                    # ================================================================
                    
                    if not df_revendas_final.empty:
                        formatar_aba(writer, df_revendas_final, "Analise Revenda")
                        
                        # Formatação Condicional Extra para Revenda (Alerta se Custo < Compra)
                        workbook = writer.book
                        worksheet = writer.sheets["Analise Revenda"]
                        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
                        
                        # Localiza colunas para aplicar regra (Custo e Ultimo Preco)
                        # Exemplo: Pintar linha se Diferença Valor < 0 (Prejuízo sobre reposição)
                        col_diff_idx = df_revendas_final.columns.get_loc('diff_valor')
                        col_letter = xlsxwriter.utility.xl_col_to_name(col_diff_idx)
                        
                        worksheet.conditional_format(f'{col_letter}2:{col_letter}{len(df_revendas_final)+1}', {
                            'type': 'cell',
                            'criteria': '<',
                            'value': 0,
                            'format': red_format
                        })

                        worksheet.autofit()
                        
                    # # Aba 3: Análise Produção (Comparativo com Tabela)
                    # if not df_vendas_final.empty:
                    #     formatar_aba(writer, df_vendas_final, "Analise Producao")
                        
                    #     # Formatação Condicional para Produção (Divergência > 10%)
                    #     workbook = writer.book
                    #     worksheet = writer.sheets["Analise Producao"]
                    #     yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
                        
                    #     if 'diff_percentual' in df_vendas_final.columns:
                    #         col_pct_idx = df_vendas_final.columns.get_loc('diff_percentual')
                    #         col_letter = xlsxwriter.utility.xl_col_to_name(col_pct_idx)
                            
                    #         # Alerta se diferença for maior que 10% (negativo ou positivo)
                    #         worksheet.conditional_format(f'{col_letter}2:{col_letter}{len(df_vendas_final)+1}', {
                    #             'type': 'cell',
                    #             'criteria': 'not between',
                    #             'minimum': -0.10,
                    #             'maximum': 0.10,
                    #             'format': yellow_format
                    #         })
                        
                except Exception as e:
                    print(f"Erro ao aplicar formatação condicional nas linhas: {e}")
                    raise e
                
                # ================================================================
                # Iterar sobre as vendas de geradores e emitir os detalhamentos das OPS do período
                # ================================================================
                
                # geradores_vendidos_fora_margem = [f'{lote}' for lote in lotes_geradores_vendidos if lote.strip() != '']
                
                # query_op = "SELECT DISTINCT TRIM(D3_OP) AS D3_OP FROM SD3010 WHERE D_E_L_E_T_ <> '*' AND D3_ESTORNO <> 'S' AND TRIM(D3_LOTECTL) = ?"
                # # (
                # #     SELECT DISTINCT 
                # #         TRIM(D3_OP) AS D3_OP 
                # #     FROM SD3010
                # #     WHERE 
                # #         D_E_L_E_T_ <> '*'  
                # #         AND D3_ESTORNO <> 'S' 
                # #         AND D3_LOTECTL LIKE '7030'
                # # )

                # lista_final_ops = []
                
                # COLUNAS_DETALHAMENTO = [
                #     "filial",
                #     "produto",
                #     "armazem",
                #     "tp_movimento",
                #     "descricao_tm",
                #     "descr_prod",
                #     "unidade",
                #     "quantidade",
                #     "quant_2",
                #     "custo",
                #     "custo_2",
                #     "ord_producao",
                #     "lote",
                #     "os_ass_tecn.",
                #     "grupo",
                #     "descricao_grupo",
                #     "tipo_re_de",
                #     "ext_texto",
                #     "documento",
                #     "dt_emissao",
                #     "c_contabil",
                #     "descricao_da_conta",
                #     "centro_custo",
                #     "desc_centro_de_custo",
                #     "parc_total",
                #     "estornado",
                #     "sequencial",
                #     "tipo",
                #     "usuario",
                #     "nr_s_a",
                #     "item_s_a",
                #     "observacao"
                # ]
                
                # with pyodbc.connect(connectionString) as con:
                #     with con.cursor() as cursor:
                #         for lote in geradores_vendidos_fora_margem:
                #             lote_limpo = str(lote).strip() 
                #             cursor.execute(query_op, (lote_limpo,))
                            
                #             ops = cursor.fetchall()
                #             print(f"Buscando OPs para o lote {lote}...")
                #             if ops:
                #                 lista_ops = []
                #                 for op in ops:
                #                     if op.D3_OP.strip() != '':
                #                         lista_ops.append(op.D3_OP.strip())
                                        
                                
                #                 if lista_ops:
                #                     placeholders = ','.join('?' for _ in lista_ops)
                #                     query_detalhamento = f"""
                #                     SELECT
                #                         TRIM(D3_FILIAL) AS [Filial],
                #                         TRIM(D3_COD) AS [Produto],
                #                         TRIM(D3_LOCAL) AS [Armazem],
                #                         TRIM(D3_TM) AS [TP Movimento],
                #                         TRIM(F5_TEXTO) AS [Descrição TM],
                #                         TRIM(B1_DESC) AS [Descr. Prod],
                #                         TRIM(D3_UM) AS [Unidade],
                #                         D3_QUANT AS [Quantidade],
                #                         CASE WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_QUANT ELSE D3_QUANT * -1 END [Quant. 2],
                #                         D3_CUSTO1 AS [Custo],
                #                         CASE WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_CUSTO1 ELSE D3_CUSTO1 * -1 END [Custo 2],
                #                         TRIM(D3_OP) AS [Ord Producao],
                #                         TRIM(D3_LOTECTL) AS [Lote],
                #                         TRIM(D3_OSTEC) AS [OS Ass. Tecn.],
                #                         TRIM(D3_GRUPO) AS [Grupo],
                #                         TRIM(BM_DESC) AS [Descrição Grupo],
                #                         TRIM(D3_CF) AS [Tipo RE/DE], 
                #                         SUBSTRING(D3_CF, 1, 2) AS [Ext.texto],
                #                         TRIM(D3_DOC) AS [Documento],
                #                         CAST(D3_EMISSAO AS DATE) AS [DT Emissao],
                #                         TRIM(D3_CONTA) AS [C Contabil],
                #                         TRIM(D3_CONTA) + ' - ' + TRIM(CT1_DESC01) AS [Descrição da Conta],      --CT1 X 
                #                         TRIM(D3_CC) AS [Centro Custo],
                #                         TRIM(D3_CC) + ' - ' + TRIM(CTT_DESC01) AS [Desc Centro de Custo],    --CTT X 
                #                         TRIM(D3_PARCTOT) AS [Parc/Total],
                #                         TRIM(D3_ESTORNO) AS [Estornado],
                #                         TRIM(D3_NUMSEQ) AS [Sequencial],
                #                         TRIM(D3_TIPO) AS [Tipo],
                #                         TRIM(D3_USUARIO) AS [Usuario], 
                #                         TRIM(D3_NUMSA) AS [Nr.S.A.],
                #                         TRIM(D3_ITEMSA) AS [Item S.A.],
                #                         TRIM(D3_OBSERVA) AS [Observacao]
                #                     FROM SD3010

                #                     LEFT JOIN SB1010
                #                         ON SB1010.D_E_L_E_T_ <> '*'
                #                         AND B1_FILIAL = SUBSTRING(D3_FILIAL, 1, 2)
                #                         AND B1_COD = D3_COD

                #                     LEFT JOIN SBM010
                #                         ON SBM010.D_E_L_E_T_ <> '*'
                #                         AND BM_GRUPO = D3_GRUPO

                #                     LEFT JOIN CT1010 Contabil ON Contabil.CT1_CONTA = B1_CONTA
                #                         AND Contabil.D_E_L_E_T_ <> '*' 

                #                     LEFT JOIN NNR010
                #                         ON NNR010.D_E_L_E_T_ <> '*'
                #                         AND D3_FILIAL = NNR_FILIAL
                #                         AND NNR_CODIGO = D3_LOCAL

                #                     LEFT JOIN SF5010
                #                         ON SF5010.D_E_L_E_T_ <> '*'
                #                         AND F5_FILIAL = D3_FILIAL
                #                         AND F5_CODIGO = D3_TM

                #                     LEFT JOIN CTT010 
                #                         ON CTT010.D_E_L_E_T_ <> '*'
                #                         AND CTT_CUSTO = D3_CC
                #                         AND CTT_FILIAL = D3_FILIAL

                #                     WHERE 
                #                         SD3010.D_E_L_E_T_ <> '*'

                #                         AND D3_OP IN ({placeholders})

                #                         AND D3_ESTORNO <> 'S'

                #                     ORDER BY
                #                         D3_OP, D3_TM, D3_CUSTO1 DESC, TRIM(B1_DESC)"""
                            
                #                     cursor.execute(query_detalhamento, tuple(lista_ops))
                                                                
                #                     detalhamento = cursor.fetchall()
                                    
                #                     if detalhamento:
                #                         dados_convertidos = [tuple(row) for row in detalhamento]
            
                #                         op_detalhada = pd.DataFrame(dados_convertidos, columns=COLUNAS_DETALHAMENTO)
                                        
                #                         formatar_aba(writer, op_detalhada, f"OP {op_detalhada.iloc[0]['ord_producao'][:6].lstrip('0')} - Lote {lote_limpo}")
                                        
                #             else:
                #                 print(f"Nenhuma OP encontrada para o lote {lote}")      
                                
                                
                                
                # ================================================================
                #                           Resumo
                # ================================================================
                
                valor_contabil_total = nfs_mes_anterior['valor_contabil'].sum()
                
                custo_total = nfs_mes_anterior['custo'].sum()
                
                margem_bruta_total = nfs_mes_anterior['margem_bruta'].sum()
                
                margem_bruta_percentual_total = margem_bruta_total / valor_contabil_total
                
                total_abaixo_margem = abaixo_margem
                
                total_vendas = len(nfs_mes_anterior)
                
                percentual_margem_abaixo = total_abaixo_margem / total_vendas
                
                print(f'valor_contabil_total: {valor_contabil_total}')
                print(f'custo_total: {custo_total}')
                print(f'margem_bruta_total: {margem_bruta_total}')
                print(f'margem_bruta_percentual_total: {margem_bruta_percentual_total}')
                print(f'total_abaixo_margem: {total_abaixo_margem}')
                print(f'total_vendas: {total_vendas}')
                print(f'percentual_margem_abaixo: {percentual_margem_abaixo}')
                
                sheet_name = "Resumo"
                
                workbook = writer.book
                worksheet = workbook.add_worksheet(sheet_name)
                                
                dados = [
                    ['Valor Contábil Total', "=SUM('Todas NFs'!V:V)", formats['blank']['money']], 
                    ['Custo Total', "=SUM('Todas NFs'!W:W)", formats['blank']['money']],
                    ['Margem Bruta Total', "=SUM('Todas NFs'!AE:AE)", formats['blank']['money']],
                    ['Margem Bruta Percentual Total', "=B3/B1", formats['blank']['percent']],
                    ['Total de Vendas', "=COUNTA('Todas NFs'!A:A)-1", formats['blank']['geral']],
                    ['Total Abaixo da Margem', total_abaixo_margem, formats['blank']['geral']],
                    ['Percentual Abaixo da Margem', '=B6/B5', formats['blank']['percent']],
                ]

                # Escrevendo os dados e fórmulas linha a linha
                for row_idx, linha_dados in enumerate(dados):
                    rotulo, valor, formato = linha_dados
                    
                    worksheet.write(row_idx, 0, rotulo)
                    worksheet.write(row_idx, 1, valor, formato)
                
                
                # ================================================================
                #               Estatística das justificativas - Aba resumo
                # ================================================================

                # Escreve os cabeçalhos
                cabecalhos = ['Justificativa',  'Contagem', '% do Total', '% do Total Abaixo da Margem']
                for i, c in enumerate(cabecalhos):
                    worksheet.write(0, 4 + i, c , workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1}))

                # 1. UNIQUE geralmente não precisa de separador, mas se precisasse, seria vírgula.
                worksheet.write_formula('E2', f"=UNIQUE(FILTER('Config'!A2:A50, 'Config'!A2:A50<>\"\"))")

                # 2. CORREÇÃO AQUI: Trocado ; por ,
                formula_count = f"=COUNTIF('Todas NFs'!AG:AG, ANCHORARRAY(E2))"
                worksheet.write_formula("F2", formula_count)

                # 3. CORREÇÃO AQUI: Trocado ; por ,
                formula_percent_total = f'=IFERROR(ANCHORARRAY(F2) / $B$5, 0)'
                formula_percent_abaixo = f'=IFERROR(ANCHORARRAY(F2) / $B$6, 0)'

                
                worksheet.write_formula("G2", formula_percent_total)
                worksheet.write_formula("H2", formula_percent_abaixo)
                
                worksheet.set_column('A:A', 30) 
                worksheet.set_column('B:B', 25) 
                worksheet.set_column('E:E', 40) 
                worksheet.set_column('H:H',25,formats['blank']['percent'])
                worksheet.set_column('G:G',None,formats['blank']['percent'])
                
            print(f"Arquivo {nome_arquivo} gerado com sucesso!")

    except Exception as e:
        print(f"Erro fatal na execução: {e}")
        raise e # Descomente para debug completo
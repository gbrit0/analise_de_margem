from os import getenv
from dotenv import load_dotenv
import pymysql
import pandas as pd
from datetime import date, datetime
import pyodbc
import xlsxwriter

def busca_nfs_fora_da_margem(
    data_inicio: date = None, 
    data_fim: date = None, 
    margem_parceiros: float = 0.15, 
    margem: float = 0.27, 
    margem_maxima: float = 0.7
) -> pd.DataFrame | None:
    """Realiza busca das Notas fiscais que estão fora dos parâmetros de margem em determinado período

    Args:
        data_inicio (date): data de início do período. Default é 1º dia do mês anterior.
        data_fim (date): data de fim do período. Default é o último dia do mês anterior.
        margem_parceiros (float): parâmetro de margem para parceiros. Default é 0.15.
        margem (float): parâmetro de margem geral. Default é 0.27.
        margem_maxima (float): parâmetro de margem máxima para investigar margens muito altas. Default é 0.70.

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
                    with open('querys/busca_nfs_fora_da_margem.sql', 'r', encoding='utf-8') as query:
                        values = (data_inicio, data_fim, str(margem_parceiros), str(margem), str(margem_maxima))
                        consulta = query.read()
                        cursor.execute(consulta, values)
                        rows = cursor.fetchall()
                        if rows:
                            columns = [desc[0] for desc in cursor.description]
                            return pd.DataFrame(rows, columns=columns)
                        else:
                            return pd.DataFrame()
    except Exception as e:
        print("Erro em busca_nfs_fora_da_margem:", e)
        return None

def formatar_aba(writer, df, sheet_name):
    """Função auxiliar para aplicar formatação padrão nas abas do Excel"""
    if df.empty:
        return

    df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Formatos
    fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
    fmt_currency = workbook.add_format({'num_format': 'R$ #,##0.00'})
    fmt_percent = workbook.add_format({'num_format': '0.00%'})
    
    # Ajuste de colunas e aplicação de formatos
    for idx, col in enumerate(df.columns):
        series = df[col]
        max_len = max((series.astype(str).map(len).max(), len(str(col)))) + 2
        
        # Define formato baseado no nome da coluna (heurística simples)
        col_name_lower = col.lower()
        cell_format = None
        
        if any(x in col_name_lower for x in ['custo', 'preco', 'valor', 'total', 'liq']):
            cell_format = fmt_currency
        elif any(x in col_name_lower for x in ['percent', 'margem', 'aliquota']):
            cell_format = fmt_percent
            
        worksheet.set_column(idx, idx, max_len, cell_format)
        worksheet.write(0, idx, col, fmt_header)
    
    # Adicionar AutoFiltro
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

if __name__ == '__main__':
    
    load_dotenv(override=True)
    
    MYSQL_DB_HOST = getenv('MYSQL_DB_HOST')
    MYSQL_DB_USER = getenv('MYSQL_DB_USER')
    MYSQL_DB_PASSWORD = getenv('MYSQL_DB_PASSWORD')
    MYSQL_DB_DATABASE = getenv('MYSQL_DB_DATABASE')
    MYSQL_DB_PORT = int(getenv('MYSQL_DB_PORT'))
    
    connectionString = f"DRIVER={getenv('PROTHEUS_ODBC_DRIVER')};SERVER={getenv('PROTHEUS_DB_HOST')};DATABASE={getenv('PROTHEUS_DB_DATABASE')};UID={getenv('PROTHEUS_DB_USER')};PWD={getenv('PROTHEUS_DB_PASSWORD')};TrustServerCertificate=yes"

    try:
        nfs_fora_margem = busca_nfs_fora_da_margem()
        
        if nfs_fora_margem is None or nfs_fora_margem.empty:
            print("Nenhuma NF fora da margem encontrada no período.")
        else:
            print(f"{len(nfs_fora_margem)} NFs fora da margem encontradas. Iniciando processamento...")
            
            df_revendas_final = pd.DataFrame()
            df_vendas_final = pd.DataFrame()

            # Conversão de tipos gerais
            cols_to_float = ['custo', 'valor_total', 'margem_bruta_percentual']
            for col in cols_to_float:
                if col in nfs_fora_margem.columns:
                     nfs_fora_margem[col] = pd.to_numeric(nfs_fora_margem[col], errors='coerce')

            # Separação
            vendas = nfs_fora_margem[nfs_fora_margem['tipo_produto'].str.contains('PA|PI', na=False)].copy()
            revendas = nfs_fora_margem[~nfs_fora_margem['tipo_produto'].str.contains('PA|PI', na=False)].copy()
            
            # ========================================================
            #                 Análise das Revendas
            # ========================================================
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
                    cols_revenda = ['nota', 'emissao', 'cliente', 'cod_produto', 'produto', 'custo', 'ultimo_preco_compra', 'diff_valor', 'diff_percentual', 'vendedor']
                    # Filtra apenas colunas que existem no df original + as novas
                    cols_finais = [c for c in cols_revenda if c in revendas.columns]
                    df_revendas_final = revendas[cols_finais]

            # ========================================================
            #                 Análise das Vendas (Produção)
            # ========================================================
            if not vendas.empty:
                produtos_venda = vendas['cod_produto'].unique().copy()
                
                with pyodbc.connect(connectionString) as con:
                    with con.cursor() as cursor:
                        with open('querys/preco_base.sql','r', encoding='utf-8') as query:
                            sql_query = query.read()
                        precos_base = pd.read_sql(sql_query, con)
                
                precos_base = pd.DataFrame(precos_base, columns=['cod_produto', 'preco_base', 'tabela'])
                
                # Prepara tabela de preços
                precos_filtrados = precos_base[precos_base['cod_produto'].isin(produtos_venda)].copy()
                precos_filtrados['preco_base'] = pd.to_numeric(precos_filtrados['preco_base'], errors='coerce')

                # Merge
                vendas_analise = vendas.merge(precos_filtrados, on='cod_produto', how='left')
                vendas_analise['preco_base'] = vendas_analise['preco_base'].fillna(0.0)

                # Cálculos
                vendas_analise['diff_valor'] = vendas_analise['custo'] - vendas_analise['preco_base']
                vendas_analise['diff_percentual'] = vendas_analise.apply(
                    lambda x: ((x['custo'] - x['preco_base']) / x['preco_base']) if x['preco_base'] > 0 else 0.0,
                    axis=1
                )

                # Seleção de colunas
                cols_venda = ['nota', 'emissao', 'cliente', 'cod_produto', 'produto', 'tabela', 'custo', 'preco_base', 'diff_valor', 'diff_percentual', 'margem_bruta_percentual']
                cols_finais_venda = [c for c in cols_venda if c in vendas_analise.columns]
                df_vendas_final = vendas_analise[cols_finais_venda]

            # ========================================================
            #                 Geração do Excel
            # ========================================================
            data_hoje = date.today().strftime('%Y-%m-%d')
            nome_arquivo = f'analise_margem_{data_hoje}.xlsx'
            
            print(f"Gerando arquivo Excel: {nome_arquivo}...")
            
            with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as writer:
                # Aba 1: Resumo Geral (Dados Brutos do MySQL)
                formatar_aba(writer, nfs_fora_margem, "Todas NFs Fora Margem")
                
                # Aba 2: Análise Revenda (Comparativo com Compra)
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

                # Aba 3: Análise Produção (Comparativo com Tabela)
                if not df_vendas_final.empty:
                    formatar_aba(writer, df_vendas_final, "Analise Producao")
                    
                    # Formatação Condicional para Produção (Divergência > 10%)
                    workbook = writer.book
                    worksheet = writer.sheets["Analise Producao"]
                    yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
                    
                    if 'diff_percentual' in df_vendas_final.columns:
                        col_pct_idx = df_vendas_final.columns.get_loc('diff_percentual')
                        col_letter = xlsxwriter.utility.xl_col_to_name(col_pct_idx)
                        
                        # Alerta se diferença for maior que 10% (negativo ou positivo)
                        worksheet.conditional_format(f'{col_letter}2:{col_letter}{len(df_vendas_final)+1}', {
                            'type': 'cell',
                            'criteria': 'not between',
                            'minimum': -0.10,
                            'maximum': 0.10,
                            'format': yellow_format
                        })

            print(f"Arquivo {nome_arquivo} gerado com sucesso!")

    except Exception as e:
        print(f"Erro fatal na execução: {e}")
        # raise e # Descomente para debug completo
from os import getenv 
from dotenv import load_dotenv
import pyodbc
import pymysql
import pandas as pd

INSERT_COLUMNS = [
    "filial",
    "chave",
    "nota",
    "no_pedido",
    "vendedor",
    "data_emissao",
    "lote",
    "cfop",
    "cfop_descri",
    "atualiza_estoque",
    "gera_duplicata",
    "cod_produto",
    "produto",
    "tipo_produto",
    "armazem",
    "cod_cliente",
    "loja",
    "cliente",
    "grp_amar_ctb",
    "classificacao_produto",
    "estado_destino",
    "quantidade",
    "valor_contabil",
    "custo",
    "valor_unitario",
    "valor_ipi",
    "valor_imp5",
    "valor_imp6",
    "vlr_icms_difal",
    "valor_icms",
    "aliq_icms",
    "margem_bruta",
    "margem_bruta_percentual",
]
    
def busca_nfs() -> list[tuple]:
    """Realiza a busca das notas fiscais no banco do Protheus de acordo com a query padrão e retorna as linhas retornadas."""
    
    print("Buscando NFs no Protheus...")
    
    connectionString = f"DRIVER={getenv('PROTHEUS_ODBC_DRIVER')};SERVER={getenv('PROTHEUS_DB_HOST')};DATABASE={getenv('PROTHEUS_DB_DATABASE')};UID={getenv('PROTHEUS_DB_USER')};PWD={getenv('PROTHEUS_DB_PASSWORD')};TrustServerCertificate=yes"
    
    with pyodbc.connect(connectionString) as con:
        with con.cursor() as cursor:
            with open("grid/querys/analise.sql", "r", encoding="utf-8") as query:
                
                cursor.execute(query.read())
                
                rows = cursor.fetchall()
                
                return [tuple(row) for row in rows]
            
def insere_nfs(nfs: pd.DataFrame) -> bool:
    """Insere as notas fiscais buscadas no banco da aplicação (VPS) em lote."""
    
    if nfs.empty:
        print("Nenhuma NF nova encontrada para inserir.")
        return True
    
    missing_columns = [column for column in INSERT_COLUMNS if column not in nfs.columns]
    if missing_columns:
        print(f"As seguintes colunas estão ausentes nos dados retornados: {', '.join(missing_columns)}")
        return False
    
    placeholders = ", ".join(["%s"] * len(INSERT_COLUMNS))
    formatted_columns = ", ".join(INSERT_COLUMNS)
    sql = f"INSERT INTO analise ({formatted_columns}) VALUES ({placeholders})"
    
    con = None
    
    try:
        con = pymysql.connect(
            host     = getenv('MYSQL_DB_HOST'),
            user     = getenv('MYSQL_DB_USER'),
            password = getenv('MYSQL_DB_PASSWORD'),
            database = getenv('MYSQL_DB_DATABASE'),
            port     = int(getenv('MYSQL_DB_PORT'))
        )
        
        rows = list(nfs[INSERT_COLUMNS].itertuples(index=False, name=None))
        
        with con.cursor() as cursor:
            cursor.executemany(sql, rows)
        
        con.commit()
        return True
    except Exception as e:
        if con:
            con.rollback()
        print(f"Erro ao inserir NFs: {e}")
        return False
    finally:
        if con:
            con.close()
                    
def main():
    

    nfs = busca_nfs()
    
    # ========================================================
    #        Agregação de valores com Pandas
    # ========================================================
    
    nfs = pd.DataFrame(nfs, columns=INSERT_COLUMNS[:-2])  # Exclui as colunas de margem inicialmente
    
    # nfs['margem_bruta'] = nfs['valor_contabil']-nfs['custo']-nfs['valor_ipi']-(0.02*nfs['valor_icms'])-nfs['valor_imp5']-nfs['valor_imp6']-nfs['vlr_icms_difal']
     # Filtra vendas por CFOPs específicos (apenas vendas de produção/saída que usamos para cálculo diverso)
     
    cfop_values = {'5101', '6101', '5116', '6116', '6107'}
    
    # Normaliza cfop para string e verifica se está na lista solicitada
    mask = nfs['cfop'].fillna('').astype(str).str.strip().isin(cfop_values)

    # Calcula margem bruta de forma vetorizada (evita apply que por algum motivo estava retornando múltiplas colunas)
    # Algumas linhas podem ter desc_tes como NaN, então usamos fillna('') antes de startswith
    # mask = nfs['desc_tes'].fillna('').str.startswith('VENDA DE PRODUCAO')

    nfs.loc[mask, 'margem_bruta'] = (
        nfs.loc[mask, 'valor_contabil']
        - nfs.loc[mask, 'custo']
        - nfs.loc[mask, 'valor_ipi']
        - (0.02 * nfs.loc[mask, 'valor_icms'])
        - nfs.loc[mask, 'valor_imp5']
        - nfs.loc[mask, 'valor_imp6']
        - nfs.loc[mask, 'vlr_icms_difal']
    )

    nfs.loc[~mask, 'margem_bruta'] = (
        nfs.loc[~mask, 'valor_contabil']
        - nfs.loc[~mask, 'custo']
        - nfs.loc[~mask, 'valor_ipi']   
        - nfs.loc[~mask, 'valor_icms']
        - nfs.loc[~mask, 'valor_imp5']
        - nfs.loc[~mask, 'valor_imp6']
        - nfs.loc[~mask, 'vlr_icms_difal']
    )
    
    nfs['margem_bruta_percentual'] = nfs['margem_bruta'] / nfs['valor_contabil']
    
    # ========================================================
    #    Inserção das NF com as margens calculadas no banco   
    # ========================================================
    
    if len(nfs) > 0:
        print("Inserindo NFs no banco...")
        sucesso = insere_nfs(nfs)
        if sucesso:
            print(f"{len(nfs)} NFs inseridas com sucesso!")
        else:
            print("Falha ao inserir NFs.")
    else:
        print("Nenhuma NF nova encontrada para inserir.")

if __name__ == '__main__':
    load_dotenv(override=True)

    main()
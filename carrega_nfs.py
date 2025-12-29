from os import getenv 
from dotenv import load_dotenv
import pyodbc
import pymysql
import pandas as pd

INSERT_COLUMNS = [
    "chave",
    "nota",
    "no_pedido",
    "vendedor",
    "data_emissao",
    "lote",
    "tes",
    "desc_tes",
    "atualiza_estoque",
    "gera_duplicata",
    "cod_produto",
    "produto",
    "tipo_produto",
    "cod_cliente",
    "loja",
    "cliente",
    "grp_amar_ctb",
    "classificacao_produto",
    "valor_contabil",
    "custo",
    "valor_bruto",
    "valor_ipi",
    "valor_imp5",
    "valor_imp6",
    "vlr_icms_difal",
    "valor_icms",
    "margem",
    "margem_percentual",
    "margem_bruta",
    "margem_bruta_percentual",
]

def busca_ultima_nf() -> int:
    """Conecta no banco da aplicação (VPS) e busca a última NF carregada para usar de filtro na busca de atualizações"""
    
    with pymysql.connect(
        host     = MYSQL_DB_HOST,
        user     = MYSQL_DB_USER,
        password = MYSQL_DB_PASSWORD,
        database = MYSQL_DB_DATABASE,
        port     = MYSQL_DB_PORT
    ) as con:
        with con.cursor() as cursor:
            cursor.execute("""SELECT MAX(nota) FROM analise""")
            
            return cursor.fetchone()[0] or 0  
    
def busca_nfs() -> list[tuple]:
    """Realiza a busca das notas fiscais no banco do Protheus de acordo com a query padrão e retorna as linhas retornadas."""
    
    print("Buscando NFs no Protheus...")
    
    connectionString = f"DRIVER={getenv('PROTHEUS_ODBC_DRIVER')};SERVER={getenv('PROTHEUS_DB_HOST')};DATABASE={getenv('PROTHEUS_DB_DATABASE')};UID={getenv('PROTHEUS_DB_USER')};PWD={getenv('PROTHEUS_DB_PASSWORD')};TrustServerCertificate=yes"
    
    ultima_nf = busca_ultima_nf()
    print(f"Última NF inserida: {ultima_nf}")
    
    with pyodbc.connect(connectionString) as con:
        with con.cursor() as cursor:
            with open("querys/analise.sql", "r", encoding="utf-8") as query:
                
                values = (int(ultima_nf),)
                cursor.execute(query.read(), values)
                
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
            host     = MYSQL_DB_HOST,
            user     = MYSQL_DB_USER,
            password = MYSQL_DB_PASSWORD,
            database = MYSQL_DB_DATABASE,
            port     = MYSQL_DB_PORT
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
                    
if __name__ == '__main__':
    
    load_dotenv(override=True)
    
    MYSQL_DB_HOST = getenv('MYSQL_DB_HOST')
    MYSQL_DB_USER = getenv('MYSQL_DB_USER')
    MYSQL_DB_PASSWORD = getenv('MYSQL_DB_PASSWORD')
    MYSQL_DB_DATABASE = getenv('MYSQL_DB_DATABASE')
    MYSQL_DB_PORT = int(getenv('MYSQL_DB_PORT'))

    nfs = busca_nfs()
    
    # ========================================================
    #        Agregação de valores com Pandas
    # ========================================================
    
    nfs = pd.DataFrame(nfs, columns=INSERT_COLUMNS[:-4])  # Exclui as colunas de margem inicialmente
    
    nfs['margem'] = nfs['valor_contabil'] - nfs['custo']
    nfs['margem_percentual'] = nfs['margem'] / nfs['valor_contabil']
    nfs['margem_bruta'] = nfs['valor_contabil']-nfs['custo']-nfs['valor_ipi']-(0.02*nfs['valor_icms'])-nfs['valor_imp5']-nfs['valor_imp6']-nfs['vlr_icms_difal']
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

import pyodbc

conn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=geradores.datasetsolucoes.com.br,37001;"
    "DATABASE=PROTHEUS12_HOLDING_PROD;"
    "UID=gabriel.brito;"
    "PWD=a1B2c3D4e5F6;"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
    "Connection Timeout=30;"
)

conn = pyodbc.connect(conn_str)

cursor = conn.cursor()
cursor.execute("SELECT @@VERSION")

print(cursor.fetchone())

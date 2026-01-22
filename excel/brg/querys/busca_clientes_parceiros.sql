-- =========================================================
-- Busca clientes parceiros para verificação da margem
-- Consulta será executada com pymysql que utiliza %%s como placeholder ao invés de ? como no pyodbc
-- =========================================================

SELECT
    cod_cliente + loja_cliente as cliente
FROM
    cliente_parceiro
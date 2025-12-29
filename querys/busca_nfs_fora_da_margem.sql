-- =========================================================
-- Busca Notas Fiscais fora da margem esperada
-- Consulta será executada com pymysql que utiliza %%s como placeholder ao invés de ? como no pyodbc
-- =========================================================

SELECT
    *
FROM
    analise
WHERE
    cod_produto NOT IN ('B0010046','E000H2P8')  -- Ignora Conjunto de Reparo e Conjunto de Painéis
    AND data_emissao BETWEEN %s AND %s
    AND (
        (
            cod_cliente + loja IN (SELECT cod_cliente + loja_cliente FROM cliente_parceiro) 
            AND margem_bruta_percentual < %s -- 0.15
        )
        OR margem_bruta_percentual < %s -- 0.27
        OR margem_bruta_percentual > %s -- 0.7
    )
    


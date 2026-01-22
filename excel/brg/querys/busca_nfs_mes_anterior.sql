-- =========================================================
-- Busca Notas Fiscais do mes anterior
-- Consulta será executada com pymysql que utiliza %%s como placeholder ao invés de ? como no pyodbc
-- =========================================================

SELECT
    filial,
    nota,
    no_pedido,
    vendedor,
    data_emissao,
    lote,
    cfop,
    cfop_descri,
    atualiza_estoque,
    gera_duplicata,
    cod_produto,
    produto,
    tipo_produto,
    armazem,
    cod_cliente,
    loja,
    cliente,
    grp_amar_ctb,
    classificacao_produto,
    estado_destino,
    quantidade,
    valor_contabil,
    custo,
    valor_unitario,
    valor_ipi,
    valor_imp5,
    valor_imp6,
    vlr_icms_difal,
    valor_icms,
    aliq_icms,
    margem_bruta,
    margem_bruta_percentual
FROM
    analise
WHERE
    -- Comentado pois, por mais que não sejam analisados o conjunto de painéis e conjunto de reparo devem aparecer na lista de todas as nfs. cod_produto NOT IN ('B0010046','E000H2P8')  -- Ignora Conjunto de Reparo e Conjunto de Painéis
    cfop NOT IN ('5922', '6922') -- ignora vendas futuras
    AND data_emissao BETWEEN %s AND %s

    -- =========================================================
    -- Comentada essa parte para retornar todas as vendas e 
    -- realizar a filtragem no código apenas destacando as margens
    -- divergentes sem excluir as margens OK
    -- =========================================================
    
    -- AND (
    --     (
    --         cod_cliente + loja IN (SELECT cod_cliente + loja_cliente FROM cliente_parceiro) 
    --         AND margem_bruta_percentual < %%s -- 0.15
    --     )
    --     OR margem_bruta_percentual < %%s -- 0.27
    --     OR margem_bruta_percentual > %%s -- 0.7
    -- )
    

-- analise_margem.analise definition
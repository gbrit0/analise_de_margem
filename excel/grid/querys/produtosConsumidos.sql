-- =========================================================
-- Registro de produtos baixados em uma ordem de produção
-- =========================================================

WITH ProdutosIntermediarios
AS (
    SELECT DISTINCT
        D3_OP --,
        -- X5_DESCRI,
        -- *
    FROM SD3010 D3
    LEFT JOIN SX5010 X5 ON
        X5.D_E_L_E_T_  <> '*'
        AND X5_TABELA = '02'
        AND X5_CHAVE = D3_TIPO
    WHERE
        D3.D_E_L_E_T_ <> '*'
        AND D3_OP LIKE '%7030%'
        AND D3_TIPO = 'PA'
        AND D3_TM = '010'
)
-- , ProdutosConsumidos AS (
    SELECT 
        -- CASE
        --     WHEN D3_DEV.D3_COD IS NOT NULL THEN 1 ELSE 0 
        -- END AS DEVOLUCAO,
        F5_TEXTO,
        TRIM(B1_PI.B1_COD) AS [PRODUTO INTERMEDIARIO],
        TRIM(D3.D3_OP) AS [COD_PRODUTO],
        TRIM(B1.B1_DESC) AS [PRODUTO],
        D3.D3_QUANT AS [QUANTIDADE],
        D3.D3_CUSTO1 AS [CUSTO],
        D3.D3_ESTORNO,
        D3.D3_TM
        -- D3.D3_TM

    FROM
        SD3010 D3

        -- SF5 (Tipos de Movimentacao)	SD3 (Movimentacoes Internas)	F5_CODIGO	D3_TM
        LEFT JOIN SF5010 F5
            ON F5.D_E_L_E_T_ <> '*'
            AND F5_CODIGO = D3_TM
            AND F5_FILIAL = D3_FILIAL

        -- SB1 (Descricao Generica do Produto)	SD3 (Movimentacoes Internas)	B1_COD	D3_COD
        LEFT JOIN SB1010 B1 ON
            B1.D_E_L_E_T_ <> '*'
            AND B1_FILIAL = SUBSTRING(D3_FILIAL, 1, 2)
            AND B1_COD = D3_COD

        LEFT JOIN SD3010 D3_DEV
            ON D3_DEV.D_E_L_E_T_ <> '*'
            AND D3.D3_COD = D3_DEV.D3_COD
            AND D3.D3_FILIAL = D3_DEV.D3_FILIAL
            AND D3.D3_OP = D3_DEV.D3_OP
            AND D3_DEV.D3_TM = '050' -- DEVOLUCAO
            AND D3.D3_CUSTO1 = D3_DEV.D3_CUSTO1

        LEFT JOIN SD3010 AS D3_PI
            ON D3_PI.D_E_L_E_T_ <> '*'
            AND D3_PI.D3_DOC = SUBSTRING(D3.D3_OP, 1, 9)
            AND D3_PI.D3_TM = '010' -- PRODUCAO
            AND D3_PI.D3_FILIAL = D3.D3_FILIAL

        LEFT JOIN SB1010 B1_PI ON
            B1_PI.D_E_L_E_T_ <> '*'
            AND B1_PI.B1_FILIAL = SUBSTRING(D3_PI.D3_FILIAL, 1, 2)
            AND B1_PI.B1_COD = D3_PI.D3_COD
    WHERE
        D3.D_E_L_E_T_ <> '*'
        AND D3.D3_ESTORNO <> 'S'
        AND D3.D3_TM = '010' -- PRODUCAO
        AND D3.D3_OP IN (
            SELECT
                *
            FROM 
                ProdutosIntermediarios
        )
        -- AND D3.D3_TM <> '999'
        -- AND D3_DEV.D3_COD IS NULL
    ORDER BY 1, 3

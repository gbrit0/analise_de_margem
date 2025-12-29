WITH D1 AS (
    SELECT DISTINCT
        D1_FILIAL AS Filial,
        -- TRIM(D1_DOC) as "Nº Documento",
        -- D1_SERIE as "Série",
        -- D1_FORNECE as "Cod Fornecedor",
        -- D1_LOJA as "Loja Fornecedor", 
        -- D1_ITEM as "Item",
        D1_COD as "Cod Produto",
        CAST(D1_VUNIT AS FLOAT) as "Valor Unitário"

    FROM SD1010

    LEFT JOIN SF4010 
        ON (F4_CODIGO) = (D1_TES)
        AND (F4_FILIAL) = (D1_FILIAL)
        AND SF4010.D_E_L_E_T_ <> '*'

    WHERE 
        SD1010.D_E_L_E_T_ <> '*'
        -- Todas as TES da SD1 são menores que 500 (entradas de estoque)    
        
        AND TRIM(D1_TIPO) = 'N' -- Nota de Entrada Normal
        AND TRIM(F4_ESTOQUE) = 'S'-- Atualiza Estoque = SIM 
        AND TRIM(F4_DUPLIC) = 'S'-- Gera Financeiro = SIM 
        AND TRIM(F4_PODER3) = 'N'-- Poder de Terceiros = NÃO
)

SELECT
    TRIM(B1_COD) AS [cod_produto],
    CASE 
        WHEN SUBSTRING(B1_UCOM, 1, 4) = CAST(YEAR(GETDATE()) AS VARCHAR)
            THEN B1_UPRC
        ELSE
            CASE 
                WHEN COALESCE(MAX(CAST(D1.[Valor Unitário] AS FLOAT)), 0) > COALESCE(MAX(CAST(B2_CM1 AS FLOAT)), 0) THEN MAX(CAST(D1.[Valor Unitário] AS FLOAT))
                WHEN COALESCE(MAX(CAST(D1.[Valor Unitário] AS FLOAT)), 0) < COALESCE(MAX(CAST(B2_CM1 AS FLOAT)), 0) THEN MAX(CAST(B2_CM1 AS FLOAT))
                WHEN COALESCE(MAX(CAST(D1.[Valor Unitário] AS FLOAT)), 0) = COALESCE(MAX(CAST(B2_CM1 AS FLOAT)), 0) AND MAX(CAST(D1.[Valor Unitário] AS FLOAT)) > 0 THEN MAX(CAST(B2_CM1 AS FLOAT))
                ELSE 0
            END    
    END AS [preco_base],
    CASE 
        WHEN SUBSTRING(B1_UCOM, 1, 4) = CAST(YEAR(GETDATE()) AS VARCHAR)
            THEN 'SB1'
        ELSE
            CASE 
                WHEN COALESCE(MAX(CAST(D1.[Valor Unitário] AS FLOAT)), 0) > COALESCE(MAX(CAST(B2_CM1 AS FLOAT)), 0) THEN 'SD1'
                WHEN COALESCE(MAX(CAST(D1.[Valor Unitário] AS FLOAT)), 0) < COALESCE(MAX(CAST(B2_CM1 AS FLOAT)), 0) THEN 'SB2'
                WHEN COALESCE(MAX(CAST(D1.[Valor Unitário] AS FLOAT)), 0) = COALESCE(MAX(CAST(B2_CM1 AS FLOAT)), 0) AND MAX(CAST(D1.[Valor Unitário] AS FLOAT)) > 0 THEN 'SD1'
                ELSE 'NENHUMA TABELA'
            END    
    END AS [tabela]

FROM  
    SB1010

LEFT JOIN SB2010
    ON SUBSTRING(B2_FILIAL, 1, 2) = B1_FILIAL
    AND B2_COD = B1_COD
    AND SB2010.D_E_L_E_T_ <> '*'

LEFT JOIN D1
    ON D1.Filial = B2_FILIAL
    AND D1.[Cod Produto] = B2_COD

WHERE 
    SB1010.D_E_L_E_T_ <> '*'
    AND B1_MSBLQL = '2'
    AND B1_TIPO IN ('PA', 'PI')
    AND B2_FILIAL IS NOT NULL
    AND TRIM(B2_FILIAL) =  '0101'

GROUP BY
    TRIM(B1_COD), B1_UCOM, B1_UPRC
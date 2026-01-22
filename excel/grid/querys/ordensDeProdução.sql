-- =========================================================
-- Buscar os produtos consumidos nas OPs
-- =========================================================

-- SC2 (Ordens de Producao)	SD3 (Movimentacoes Internas)	C2_NUM + C2_ITEM + C2_SEQUEN + C2_ITEMGRD	D3_OP

SELECT DISTINCT TOP 100 
    SD3.D3_OP,
    SC2.*
FROM 
    SC2010 AS SC2 -- Ordens de Producao

LEFT JOIN SD3010 AS SD3 -- Movimentacoes Internas
    ON SD3.D_E_L_E_T_ <> '*'
    AND C2_FILIAL = D3_FILIAL
    AND TRIM(C2_NUM) + TRIM(C2_ITEM) + TRIM(C2_SEQUEN) + TRIM(C2_ITEMGRD) = D3_OP

WHERE 
    SC2.D_E_L_E_T_ <> '*'
AND D3_OP LIKE '010027%'


SELECT
    TRIM(D3_FILIAL) filial,
    TRIM(D3_COD) produto,
    TRIM(D3_LOCAL) armazem,
    TRIM(D3_TM) tp_movimento,
    COALESCE(TRIM(F5_TEXTO), '-') descricao_tm,
    TRIM(B1_DESC) descr_prod,
    TRIM(D3_UM) unidade,
    D3_QUANT quantidade,
    CASE 
        -- WHEN TRIM(D3_TM) = '010' THEN 0 
        WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_QUANT 
        ELSE D3_QUANT * -1 
    END quant_2,
    D3_CUSTO1 custo,
    CASE 
        -- WHEN TRIM(D3_TM) = '010' THEN 0
            -- CASE 
            --     WHEN COALESCE(MAX(CAST(B2_CMFIM1 AS FLOAT)), 0) = 0 THEN MAX(CAST(B2_CM1 AS FLOAT))
            --     ELSE MAX(CAST(B2_CMFIM1 AS FLOAT))
            -- END
        
        WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN (D3_CUSTO1) 
        ELSE (D3_CUSTO1) * -1 
    END custo_2,
    TRIM(D3_OP) ord_producao,
    TRIM(D3_LOTECTL) lote,
    TRIM(D3_OSTEC) os_ass_tecn,
    TRIM(D3_GRUPO) grupo,
    COALESCE(TRIM(BM_DESC), '-') descricao_grupo,
    TRIM(D3_CF) tipo_re_de,
    SUBSTRING(D3_CF, 1, 2) ext_texto,
    TRIM(D3_DOC) documento,
    CAST(D3_EMISSAO AS DATE) dt_emissao,
    TRIM(D3_CONTA) c_contabil,
    COALESCE(TRIM(CT1_DESC01), '-') descricao_da_conta,
    TRIM(D3_CC) centro_custo,
    COALESCE(TRIM(CTT_DESC01), '-') desc_centro_de_custo,
    TRIM(D3_PARCTOT) parc_total,
    TRIM(D3_ESTORNO) estornado,
    TRIM(D3_NUMSEQ) sequencial,
    TRIM(D3_TIPO) tipo,
    TRIM(D3_USUARIO) usuario,
    TRIM(D3_NUMSA) nr_s_a,
    TRIM(D3_ITEMSA) item_s_a
    
FROM SD3010

LEFT JOIN SB1010
    ON SB1010.D_E_L_E_T_ <> '*'
    AND B1_FILIAL = SUBSTRING(D3_FILIAL, 1, 2)
    AND B1_COD = D3_COD

LEFT JOIN SBM010
    ON SBM010.D_E_L_E_T_ <> '*'
    AND BM_GRUPO = D3_GRUPO

LEFT JOIN CT1010 Contabil ON Contabil.CT1_CONTA = B1_CONTA
    AND Contabil.D_E_L_E_T_ <> '*' 

LEFT JOIN NNR010
    ON NNR010.D_E_L_E_T_ <> '*'
    AND D3_FILIAL = NNR_FILIAL
    AND NNR_CODIGO = D3_LOCAL

LEFT JOIN SF5010
    ON SF5010.D_E_L_E_T_ <> '*'
    AND F5_FILIAL = D3_FILIAL
    AND F5_CODIGO = D3_TM

LEFT JOIN CTT010 
    ON CTT010.D_E_L_E_T_ <> '*'
    AND CTT_CUSTO = D3_CC
    AND CTT_FILIAL = D3_FILIAL

LEFT JOIN SB2010
    ON SB2010.D_E_L_E_T_ <> '*'
    AND B2_FILIAL = D3_FILIAL
    AND B2_COD = D3_COD

WHERE 
    SD3010.D_E_L_E_T_ <> '*'

    AND D3_OP IN (
        SELECT 
            DISTINCT TRIM(D3_OP) AS D3_OP 
        FROM SD3010 
        WHERE 
            D_E_L_E_T_ <> '*' 
            -- AND D3_ESTORNO <> 'S' 
            AND TRIM(D3_LOTECTL) = ? --'7166' --
    )


    AND D3_ESTORNO <> 'S'
    and D3_COD <> 'MO1501MOD'
GROUP BY
    D3_FILIAL, D3_COD, D3_LOCAL, D3_TM, F5_TEXTO, B1_DESC, D3_UM, D3_QUANT, D3_CUSTO1, D3_OP, D3_LOTECTL, D3_OSTEC, D3_GRUPO, BM_DESC, D3_CF, D3_DOC, D3_EMISSAO, D3_CONTA, CT1_DESC01, D3_CC, CTT_DESC01, D3_PARCTOT, D3_ESTORNO, D3_NUMSEQ, D3_TIPO, D3_USUARIO, D3_NUMSA, D3_ITEMSA

ORDER BY
    D3_OP, D3_TM, D3_CUSTO1 DESC, TRIM(B1_DESC);


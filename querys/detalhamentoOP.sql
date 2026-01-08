SELECT
    TRIM(D3_FILIAL) AS [Filial],
    TRIM(D3_COD) AS [Produto],
    TRIM(D3_LOCAL) AS [Armazem],
    TRIM(D3_TM) AS [TP Movimento],
    TRIM(F5_TEXTO) AS [Descrição TM],
    TRIM(B1_DESC) AS [Descr. Prod],
    TRIM(D3_UM) AS [Unidade],
    D3_QUANT AS [Quantidade],
    CASE WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_QUANT ELSE D3_QUANT * -1 END [Quant. 2],
    D3_CUSTO1 AS [Custo],
    CASE WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_CUSTO1 ELSE D3_CUSTO1 * -1 END [Custo 2],
    TRIM(D3_OP) AS [Ord Producao],
    TRIM(D3_LOTECTL) AS [Lote],
    TRIM(D3_OSTEC) AS [OS Ass. Tecn.],
    TRIM(D3_GRUPO) AS [Grupo],
    TRIM(BM_DESC) AS [Descrição Grupo],
    TRIM(D3_CF) AS [Tipo RE/DE], 
    SUBSTRING(D3_CF, 1, 2) AS [Ext.texto],
    TRIM(D3_DOC) AS [Documento],
    CAST(D3_EMISSAO AS DATE) AS [DT Emissao],
    TRIM(D3_CONTA) AS [C Contabil],
    TRIM(D3_CONTA) + ' - ' + TRIM(CT1_DESC01) AS [Descrição da Conta],  
    TRIM(D3_CC) AS [Centro Custo],
    TRIM(D3_CC) + ' - ' + TRIM(CTT_DESC01) AS [Desc Centro de Custo],    
    TRIM(D3_PARCTOT) AS [Parc/Total],
    TRIM(D3_ESTORNO) AS [Estornado],
    TRIM(D3_NUMSEQ) AS [Sequencial],
    TRIM(D3_TIPO) AS [Tipo],
    TRIM(D3_USUARIO) AS [Usuario], 
    TRIM(D3_NUMSA) AS [Nr.S.A.],
    TRIM(D3_ITEMSA) AS [Item S.A.],
    TRIM(D3_OBSERVA) AS [Observacao]
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

WHERE 
    SD3010.D_E_L_E_T_ <> '*'

    -- ===============================================================================
    -- Essa seção filtra as OPs relacionadas a um determinado lote. Será usada para
    -- detalhar as OPs de geradores vendidos abaixo da margem esperada.
    -- ===============================================================================
    AND D3_OP IN ?
    -- (
    --     SELECT DISTINCT 
    --         D3_OP 
    --     FROM SD3010
    --     WHERE 
    --         D_E_L_E_T_ <> '*'  
    --         AND D3_ESTORNO <> 'S' 
    --         AND D3_LOTECTL LIKE '7030'
    -- )
    -- ===============================================================================

    AND D3_ESTORNO <> 'S'

ORDER BY
    D3_OP, D3_TM, D3_CUSTO1 DESC, TRIM(B1_DESC)

SELECT DISTINCT TRIM(D3_OP) AS D3_OP FROM SD3010 WHERE D_E_L_E_T_ <> '*' AND D3_ESTORNO <> 'S' AND D3_LOTECTL LIKE '6386'             
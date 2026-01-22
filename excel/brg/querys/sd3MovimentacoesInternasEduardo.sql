SELECT
    D3_FILIAL AS Filial,
    D3_COD AS Produto, 
    D3_LOCAL AS Armazem,
    D3_TM AS "TP Movimento",
    F5_TEXTO AS "Descrição TM",    --SF5 X
    B1_DESC AS "Descr. Prod",     --SB1 X
    D3_UM AS Unidade,
    D3_QUANT AS Quantidade,
    CASE WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_QUANT ELSE D3_QUANT * -1 END "Quant. 2",
    D3_CUSTO1 AS Custo, 
    CASE WHEN SUBSTRING(D3_CF, 1, 2) = 'RE' THEN D3_CUSTO1 ELSE D3_CUSTO1 * -1 END " Custo 2",
    D3_OP AS "Ord Producao",
    D3_LOTECTL AS Lote, 
    D3_OSTEC AS "OS Ass.Tecn.",
    D3_GRUPO AS Grupo,
    BM_DESC AS "Descrição Grupo",     --SBM X
    D3_CF AS "Tipo RE/DE", 
    SUBSTRING(D3_CF, 1, 2) AS "Ext.texto",
    D3_DOC AS Documento,
    D3_EMISSAO AS "DT Emissao",
    D3_CONTA AS "C Contabil",
    LTRIM(RTRIM(D3_CONTA)) + ' - ' + LTRIM(RTRIM(CT1_DESC01)) AS "Descrição da Conta",      --CT1 X 
    D3_CC AS "Centro Custo",
    LTRIM(RTRIM(D3_CC)) + ' - ' + LTRIM(RTRIM(CTT_DESC01)) AS "Desc Centro de Custo",    --CTT X 
    D3_PARCTOT AS "Parc/Total",
    D3_ESTORNO AS Estornado,
    D3_NUMSEQ AS Sequencial,
    D3_TIPO AS Tipo,
    D3_USUARIO AS Usuario, 
    D3_NUMSA AS "Nr.S.A.",
    D3_ITEMSA AS "Item S.A.",
    D3_OBSERVA AS Observacao

FROM SD3010

LEFT JOIN SF5010 ON F5_CODIGO = D3_TM
    AND SF5010.D_E_L_E_T_ <> '*'
    AND D3_FILIAL = F5_FILIAL

LEFT JOIN SB1010 ON B1_COD = D3_COD
    AND SB1010.D_E_L_E_T_ <> '*'
    AND B1_FILIAL = SUBSTRING(D3_FILIAL, 1,2) 

LEFT JOIN SBM010 ON BM_GRUPO = D3_GRUPO
    AND SBM010.D_E_L_E_T_ <> '*'
    
LEFT JOIN CT1010 ON CT1_CONTA = D3_CONTA
    AND CT1010.D_E_L_E_T_ <> '*'
    AND CT1_FILIAL = ' '

LEFT JOIN CTT010 ON CTT_CUSTO = D3_CC
    AND CTT010.D_E_L_E_T_ <> '*'
    AND CTT_FILIAL = D3_FILIAL


WHERE SD3010.D_E_L_E_T_ <> '*'
AND D3_FILIAL IN ('0101','0501','0502','0503','1001')
AND D3_OP <> ' '
AND D3_ESTORNO <> 'S'
AND D3_EMISSAO >= '20220101'
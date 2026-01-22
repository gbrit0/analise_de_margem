-- =========================================================
-- Apuração de Custo no momento da emissão da nota de saída
-- =========================================================

SELECT
    TRIM(D2_FILIAL) AS [filial],
    TRIM(D2_FILIAL) + TRIM(D2_DOC) + TRIM(D2_SERIE) + TRIM(D2_CLIENTE) + TRIM(D2_LOJA) + TRIM(D2_ITEM) AS [chave],
    LTRIM(D2_DOC, 0) AS [nota],
    D2_PEDIDO AS [no_pedido],
    TRIM(A3_NOME) AS [vendedor],
    CAST(F2_EMISSAO AS DATE) AS [data_emissao],
    TRIM(D2_LOTECTL) AS [lote],
    TRIM(D2_CF) AS [cfop],
    TRIM(X5_DESCRI) AS [cfop_descri],
    TRIM(F4_ESTOQUE) AS [atualiza_estoque],
    TRIM(F4_DUPLIC) AS [gera_duplicata],
    TRIM(B1_COD) AS [cod_produto],
    TRIM(SUBSTRING(B1_DESC, 1, 50)) AS [produto],
    TRIM(B1_TIPO) AS [tipo_produto],
    TRIM(D2_LOCAL) + ' - ' + TRIM(NNR_DESCRI) AS [armazem],
    TRIM(F2_CLIENTE) AS [cod_cliente],
    TRIM(F2_LOJA) AS [loja],
    TRIM(A1_NOME) AS [cliente],
    LTRIM(B1_XGRPCTB, 0) AS [grp_amar_ctb],
    TRIM(ZC2_DESCR) AS [classificacao_produto],
    TRIM(D2_EST) AS [estado_destino],
    D2_QUANT AS [quantidade],
    D2_VALBRUT AS [valor_contabil],
    D2_CUSTO1 AS [custo],
    D2_PRCVEN AS [valor_unitario],
    D2_VALIPI AS [valor_ipi],
    D2_VALIMP5 AS [valor_imp5],
    D2_VALIMP6 AS [valor_imp6],
    D2_DIFAL AS [vlr_icms_difal],
    D2_VALICM AS [valor_icms],
    D2_PICM AS [aliq_icms]

FROM SD2010 AS D2 -- Itens de Venda da NF

    INNER JOIN SF4010 F4 ON -- Tipos de Entrada e Saída
        F4.D_E_L_E_T_ <> '*'
        AND F4_CODIGO = D2_TES
        AND F4_FILIAL = D2_FILIAL
        AND F4_TEXTO LIKE 'VENDA%'

    LEFT JOIN SB1010 B1 ON -- Cadastro de Produtos
        B1.D_E_L_E_T_ <> '*'
        AND B1_FILIAL = SUBSTRING(D2_FILIAL, 1, 2)
        AND B1_COD = D2_COD

    LEFT JOIN SF2010 AS F2 ON --  Cabeçalho das NF de Saída
        F2.D_E_L_E_T_ <> '*'
        AND TRIM(F2_DOC) = TRIM(D2_DOC)
        AND TRIM(F2_SERIE) = TRIM(D2_SERIE)
        AND TRIM(F2_CLIENTE) = TRIM(D2_CLIENTE)
        AND TRIM(F2_LOJA) = TRIM(D2_LOJA)
        AND TRIM(F2_FILIAL) = TRIM(D2_FILIAL)

    LEFT JOIN SA1010 AS A1 ON -- Clientes
        A1.D_E_L_E_T_ <> '*'
        AND TRIM(A1_COD) = TRIM(F2_CLIENTE)
        AND TRIM(A1_LOJA) = TRIM(F2_LOJA)
        AND A1_FILIAL = SUBSTRING(F2_FILIAL, 1, 2)

    LEFT JOIN SA3010 AS A3 ON -- Vendedores
        A3.D_E_L_E_T_ <> '*'
        AND A3_COD = F2_VEND1
        AND A3_FILIAL = F2_FILIAL

    -- Junção com a tabela ZC2 para obter a descrição do grupo de amarração contábil
    LEFT JOIN ZC2010 AS ZC2 ON 
        ZC2.D_E_L_E_T_ <> '*'
        AND ZC2_GRP = B1_XGRPCTB
        AND ZC2_FILIAL = B1_FILIAL

    -- SX5 (Tabelas)	SD2 (Itens de Venda da NF)	X5_TABELA + X5_CHAVE	'13' + D2_CF
    LEFT JOIN SX5010 ON
        SX5010.D_E_L_E_T_ <> '*'
        AND X5_TABELA = '13'
        AND X5_FILIAL = D2_FILIAL
        AND X5_CHAVE = D2_CF

    -- NNR (Locais de Estoque)	SD2 (Itens de Venda da NF)	NNR_CODIGO	D2_LOCAL
    LEFT JOIN NNR010 ON
        NNR010.D_E_L_E_T_ <> '*'
        AND D2_FILIAL = NNR_FILIAL 
        AND D2_LOCAL = NNR_CODIGO

WHERE
    D2.D_E_L_E_T_ <> '*'
    AND D2_EMISSAO >= 20250901 -- AND 20250930
    AND D2_FILIAL = '0101'
    -- AND B1_COD NOT IN ('B0010046', 'E000H2P8')
    AND D2_DOC > ?
    -- AND TRIM(D2_LOTECTL) = '6943'
    -- AND D2_PEDIDO LIKE '%10672%'
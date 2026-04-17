SELECT DISTINCT
    TRIM(D2_FILIAL) + TRIM(D2_DOC) + TRIM(D2_SERIE) + TRIM(D2_CLIENTE) + TRIM(D2_LOJA) + TRIM(D2_ITEM) AS [chave],
    TRIM(D2_FILIAL) AS [filial],
    TRIM(M0_FILIAL) AS [nome_filial],
    LTRIM(D2_DOC, 0) AS [nota],
    LTRIM(D2_ITEM, 0) as [item],
    D2_PEDIDO AS [no_pedido],
    TRIM(A3_NOME) AS [vendedor],
    CAST(F2_EMISSAO AS DATE) AS [data_emissao],
    TRIM(D2_LOTECTL) AS [lote],
    TRIM(D2_CF) AS [cfop],
    TRIM(X5_CF.X5_DESCRI) AS [cfop_descri],
    TRIM(F4_ESTOQUE) AS [atualiza_estoque],
    TRIM(F4_DUPLIC) AS [gera_duplicata],
    TRIM(B1_COD) AS [cod_produto],
    TRIM(SUBSTRING(B1_DESC, 1, 50)) AS [produto],
    TRIM(B1_TIPO) AS [tipo_produto],
    TRIM(X5.X5_DESCRI) AS [desc_tipo_produto],
    TRIM(D2_LOCAL) + ' - ' + TRIM(NNR_DESCRI) AS [armazem],
    TRIM(F2_CLIENTE) AS [cod_cliente],
    TRIM(F2_LOJA) AS [loja],
    TRIM(A1_NOME) AS [cliente],
    LTRIM(B1_XGRPCTB, 0) AS [grp_amar_ctb],
    TRIM(ZC2_DESCR) AS [classificacao_produto],
    TRIM(D2_EST) AS [estado_destino],
    D2_QUANT AS [quantidade],
    CASE 
        WHEN A1_TABELA IS NULL THEN '-'
        ELSE TRIM(A1_TABELA) + ' - ' + TRIM(DA0.DA0_DESCRI)
    END AS [tabela_preco],
    CASE WHEN 
        DA1.DA1_PRCVEN IS NULL THEN 0
        ELSE DA1.DA1_PRCVEN * D2_QUANT 
    END AS [preco_tabela],
    D2_VALBRUT AS [valor_contabil],
    D2_CUSTO1 AS [custo],
    D2_PRCVEN AS [valor_unitario],
    D2_VALIPI AS [valor_ipi],
    D2_VALIMP5 AS [valor_imp5],
    D2_VALIMP6 AS [valor_imp6],
    D2_DIFAL AS [vlr_icms_difal],
    D2_VALICM AS [valor_icms],
    D2_PICM AS [aliq_icms],
    D2.R_E_C_N_O_ as [recno],
    C5.C5_COMENT AS [comentario]

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

    LEFT JOIN DA1010 AS DA1 ON -- Itens da Tabela de preço
        DA1.DA1_CODTAB  = A1_TABELA
        AND DA1.DA1_CODPRO = B1.B1_COD
        AND DA1.DA1_FILIAL = D2_FILIAL
        AND DA1.D_E_L_E_T_ <> '*'
    
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
    LEFT JOIN SX5010 AS X5_CF ON -- SX5 (Tabelas)	
        X5_CF.D_E_L_E_T_ <> '*'
        AND X5_CF.X5_TABELA = '13'
        AND X5_CF.X5_FILIAL = D2_FILIAL
        AND X5_CF.X5_CHAVE = D2_CF
    
    LEFT JOIN NNR010 ON -- NNR (Locais de Estoque)	
        NNR010.D_E_L_E_T_ <> '*'
        AND D2_FILIAL = NNR_FILIAL 
        AND D2_LOCAL = NNR_CODIGO

    LEFT JOIN SYS_COMPANY AS C ON -- TABELA DE FILIAIS
        C.D_E_L_E_T_ <> '*'
        AND C.M0_CODFIL = D2_FILIAL
        AND TRIM(M0_NOME) = 'BRG Geradores'

    LEFT JOIN SX5010 AS X5 ON  -- SX5 (Tabelas)	
        X5.D_E_L_E_T_ <> '*'
        AND X5.X5_TABELA = '02'
        AND X5.X5_CHAVE = B1_TIPO
        AND SUBSTRING(X5.X5_FILIAL, 1, 2) = B1_FILIAL
	
    LEFT JOIN DA0010 AS DA0 ON -- Tabela DA0 - Tabela de Precos
    	DA0.D_E_L_E_T_ <> '*'
    	AND DA0.DA0_CODTAB = A1.A1_TABELA
    	-- AND DA0.DA0_FILIAL = A1.A1_FILIAL

    LEFT JOIN SC5010 AS C5 ON
    	C5.D_E_L_E_T_ <> '*'
    	AND C5_FILIAL = D2_FILIAL
    	AND C5_NOTA = D2_DOC
    	AND C5_SERIE = D2_SERIE

WHERE
    D2.D_E_L_E_T_ <> '*'
    AND D2_EMISSAO >= 20250901 -- AND 20250930
    AND TRIM(D2_FILIAL) IN ('0101', '0501', '0502', '0503') -- BRG MATRIZ, GRID GO, GRID MG E GRID PA
    AND TRIM(B1.B1_COD) NOT IN ('B0010046', 'E000H2P8') 
    AND TRIM(D2_CF) NOT IN ('5922', '6922')

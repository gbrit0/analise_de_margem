/*
 * Esta é uma CTE (Common Table Expression) Recursiva.
 * Ela é dividida em duas partes:
 * 1. MEMBRO ÂNCORA: A consulta inicial que "planta a semente" (Nível 1: filhos diretos).
 * 2. MEMBRO RECURSIVO: A consulta que se auto-referencia para "descer" os níveis (netos, bisnetos...).
 */
WITH EstruturaBOM (
    Nivel,
    CodigoPai,
    CodigoFilho,
    DescFilho,
    TipoFilho,
    DescTipo,
    PrecoCompraFilho,
    Quantidade_Pai,         -- Qtde necessária para o pai IMEDIATO
    Quantidade_Acumulada    -- Qtde TOTAL necessária para o pai PRINCIPAL (GP010094)
)
AS (
    -- ==================================================================
    -- 1. MEMBRO ÂNCORA (Busca o Nível 1: Filhos diretos do produto-pai)
    -- ==================================================================
    SELECT
        1 AS Nivel,
        SG.G1_COD AS CodigoPai,
        SG.G1_COMP AS CodigoFilho,
        COMP.B1_DESC AS DescFilho,
        COMP.B1_TIPO AS TipoFilho,
        X5.X5_DESCRI AS DescTipo,
        COMP.B1_UPRC AS PrecoCompraFilho,
        SG.G1_QUANT AS Quantidade_Pai,
        SG.G1_QUANT AS Quantidade_Acumulada -- No nível 1, a qtde acumulada é a própria qtde
    FROM
        SG1010 AS SG
    INNER JOIN
        SB1010 AS COMP ON COMP.B1_COD = SG.G1_COMP
                      AND SUBSTRING(SG.G1_FILIAL, 1, 2) = COMP.B1_FILIAL
                      AND COMP.D_E_L_E_T_ <> '*'
    -- SX5 (Tabelas)	SB1 (Descricao Generica do Produto)	X5_TABELA + X5_CHAVE	'02' + B1_TIPO
    INNER JOIN SX5010 AS X5 ON  
        X5.D_E_L_E_T_ <> '*'
        AND SUBSTRING(X5_FILIAL, 1, 2) = B1_FILIAL
        AND X5_TABELA = '02'
        AND X5_CHAVE = B1_TIPO
    WHERE
        SG.G1_COD = 'D0000505'  -- << PRODUTO-PAI PRINCIPAL AQUI
        AND SG.G1_FILIAL = '0101'
        AND SG.D_E_L_E_T_ <> '*'

    UNION ALL

    -- ==================================================================
    -- 2. MEMBRO RECURSIVO (Busca os Níveis 2+: Netos, Bisnetos, etc.)
    -- ==================================================================
    SELECT
        BOM.Nivel + 1,          -- Incrementa o nível
        Filho.G1_COD AS CodigoPai, -- O "Filho" da rodada anterior é o "Pai" desta
        Filho.G1_COMP AS CodigoFilho,
        COMP.B1_DESC AS DescFilho,
        COMP.B1_TIPO AS TipoFilho,
        X5.X5_DESCRI AS DescTipo,
        COMP.B1_UPRC AS PrecoCompraFilho,
        Filho.G1_QUANT AS Quantidade_Pai,
        (BOM.Quantidade_Acumulada * Filho.G1_QUANT) AS Quantidade_Acumulada 
    FROM
        SG1010 AS Filho
    INNER JOIN
        EstruturaBOM AS BOM ON Filho.G1_COD = BOM.CodigoFilho 
    INNER JOIN
        SB1010 AS COMP ON COMP.B1_COD = Filho.G1_COMP
                      AND SUBSTRING(Filho.G1_FILIAL, 1, 2) = COMP.B1_FILIAL
                      AND COMP.D_E_L_E_T_ <> '*'
    -- SX5 (Tabelas)	SB1 (Descricao Generica do Produto)	X5_TABELA + X5_CHAVE	'02' + B1_TIPO
    INNER JOIN SX5010 AS X5 ON  
        X5.D_E_L_E_T_ <> '*'
        AND SUBSTRING(X5_FILIAL, 1, 2) = B1_FILIAL
        AND X5_TABELA = '02'
        AND X5_CHAVE = B1_TIPO

    WHERE
        Filho.D_E_L_E_T_ <> '*'
)

-- ==================================================================
-- RESULTADO FINAL
-- Seleciona os dados da CTE, agora em formato de linhas
-- ==================================================================
SELECT DISTINCT
    E.Nivel,
    E.CodigoPai,
    E.CodigoFilho,
    TRIM(E.DescFilho) AS DescricaoCompletaFilho,
    E.TipoFilho,
    E.DescTipo,
    E.Quantidade_Pai,
    E.Quantidade_Acumulada,
    E.PrecoCompraFilho,
    -- Este é o cálculo de custo multiplicativo
    (E.Quantidade_Acumulada * E.PrecoCompraFilho) AS CustoTotalComponente
FROM
    EstruturaBOM AS E
ORDER BY
    E.Nivel, E.CodigoPai, E.CodigoFilho, TRIM(E.DescFilho)

--  opção para evitar o limite de recursão padrão de 100 níveis
OPTION (MAXRECURSION 0);
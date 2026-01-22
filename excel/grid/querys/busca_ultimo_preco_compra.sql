SELECT 
    B1_COD, 
    B1_UPRC
FROM SB1010
WHERE 
    B1_COD IN ? 
    AND D_E_L_E_T_  <> '*' 
    AND B1_UPRC > 0 
    AND B1_FILIAL = '01'
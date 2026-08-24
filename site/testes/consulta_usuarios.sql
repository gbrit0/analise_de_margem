SELECT 
    CONVERT('Site dos BOTS' USING utf8mb4) AS sistema, 
    id, 
    CONVERT(username USING utf8mb4) AS username, 
    CONVERT(first_name USING utf8mb4) AS first_name, 
    CONVERT(last_name USING utf8mb4) AS last_name, 
    CONVERT(email USING utf8mb4) AS email, 
    is_staff, 
    is_active, 
    is_superuser 
FROM automacoes.usuarios_usuario

UNION ALL

SELECT 
    CONVERT('Salas SDO' USING utf8mb4), 
    id, 
    CONVERT(username USING utf8mb4), 
    CONVERT(first_name USING utf8mb4), 
    CONVERT(last_name USING utf8mb4), 
    CONVERT(email USING utf8mb4), 
    is_staff, 
    is_active, 
    is_superuser 
FROM salas_sdo.usuarios_usuario

UNION ALL

SELECT 
    CONVERT('Salas BRG' USING utf8mb4), 
    id, 
    CONVERT(username USING utf8mb4), 
    CONVERT(first_name USING utf8mb4), 
    CONVERT(last_name USING utf8mb4), 
    CONVERT(email USING utf8mb4), 
    is_staff, 
    is_active, 
    is_superuser 
FROM salas_brg.usuarios_usuario

UNION ALL

SELECT 
    CONVERT('Comercial BRG' USING utf8mb4), 
    id, 
    CONVERT(username USING utf8mb4), 
    CONVERT(first_name USING utf8mb4), 
    CONVERT(last_name USING utf8mb4), 
    CONVERT(email USING utf8mb4), 
    is_staff, 
    is_active, 
    is_superuser 
FROM proposta_brg.auth_user

UNION ALL

SELECT 
    CONVERT('RH - Contratação Digital' USING utf8mb4), 
    id, 
    CONVERT(username USING utf8mb4), 
    CONVERT(first_name USING utf8mb4), 
    CONVERT(last_name USING utf8mb4), 
    CONVERT(email USING utf8mb4), 
    is_staff, 
    is_active, 
    is_superuser 
FROM contratacao_digital.auth_user

UNION ALL

SELECT 
    CONVERT('Cobrança BRG' USING utf8mb4), 
    id, 
    CONVERT(username USING utf8mb4), 
    CONVERT(first_name USING utf8mb4), 
    CONVERT(last_name USING utf8mb4), 
    CONVERT(email USING utf8mb4), 
    is_staff, 
    is_active, 
    is_superuser 
FROM automacao_cobranca_novo.auth_user

UNION ALL

SELECT 
    CONVERT('Análise de Margem' USING utf8mb4), 
    id, 
    CONVERT(username USING utf8mb4), 
    CONVERT(first_name USING utf8mb4), 
    CONVERT(last_name USING utf8mb4), 
    CONVERT(email USING utf8mb4), 
    is_staff, 
    is_active, 
    is_superuser 
FROM analise_margem.users_customuser
;
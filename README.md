# Análise de Margem

Projeto para automação do processo de análise de margem da venda de geradores.

Baseado na planilha excel passada pela Giuliana foi gerada a consulta [analise.sql](querys/analise.sql) que captura os campos iniciais utilizados na análise e alimenta o processo a ser automatizado.

A partir da planilha de análise, os campos de margem, margem percentual, margem bruta (com impostos), margem bruta (com impostos) percentual e valor do ICMS com benefício pró-Goiás são calculados de acordo com as fórmulas aplicadas no excel.

A carga é executada no [carrega_nfs.py](carrega_nfs.py).

## Análise inicial

O script [analise_inicial.py](analise_inicial.py) busca Notas Fiscais (NFs) de um período definido e identifica registros cuja margem está fora de parâmetros definidos (margem de parceiros, margem geral e margem máxima). Em seguida separa os registros por tipo de produto (revenda vs. produção), faz consultas complementares nas bases (Protheus / SQL Server e MySQL), calcula divergências entre custo e preços de referência e gera um arquivo Excel com 3 abas principais:

- "Todas NFs Fora Margem" — dados brutos retornados pela query base do MySQL.
- "Analise Revenda" — comparação do custo com o último preço de compra (Protheus).
- "Analise Producao" — comparação do custo com tabela de preço base (Protheus / query `preco_base.sql`).

Arquivo gerado: `analise_margem_YYYY-MM-DD.xlsx` (data de execução).

## Contrato (Inputs / Outputs / Regras)

- Inputs:
  - Variáveis de ambiente (obrigatórias):
    - MYSQL_DB_HOST, MYSQL_DB_USER, MYSQL_DB_PASSWORD, MYSQL_DB_DATABASE, MYSQL_DB_PORT
    - PROTHEUS_ODBC_DRIVER, PROTHEUS_DB_HOST, PROTHEUS_DB_DATABASE, PROTHEUS_DB_USER, PROTHEUS_DB_PASSWORD
  - Arquivos SQL usados (em `querys/`):
    - `busca_nfs_fora_da_margem.sql` — query principal (MySQL) que retorna as NFs fora da margem.
    - `preco_base.sql` — query para obter preços base (usada para produtos de produção).
  - Parâmetros de função (opcionais): data_inicio, data_fim, margem_parceiros, margem, margem_maxima. Se não informados, o script calcula o período como o mês anterior.

- Outputs:
  - Excel (`analise_margem_YYYY-MM-DD.xlsx`) com as abas descritas acima e formatações/condicionais para chamar atenção a divergências.

- Regras e cálculos principais:
  - Conversão de colunas numéricas (`custo`, `valor_total`, `margem_bruta_percentual`) para float.
  - Classificação de linhas em `vendas` (produção) se `tipo_produto` contém `PA|PI`, caso contrário `revendas`.
  - Para revendas: busca último preço de compra em `SB1010` (Protheus) e calcula `diff_valor` = custo - ultimo_preco_compra e `diff_percentual` relativo.
  - Para produção: usa `preco_base` da query `preco_base.sql`, faz merge e calcula `diff_valor` e `diff_percentual` relativo.

## Fluxo detalhado (passo a passo)

1. Carrega variáveis de ambiente via `dotenv`.
2. Calcula `data_inicio` e `data_fim` (mês anterior) caso não informados.
3. Conecta ao MySQL usando `pymysql` e executa `querys/busca_nfs_fora_da_margem.sql` com parâmetros de data e margens.
4. Recebe resultado em um DataFrame `nfs_fora_margem`.
   - Se vazio, encerra com mensagem "Nenhuma NF fora da margem encontrada no período.".
5. Normaliza tipos (colunas numéricas) e separa em `vendas` e `revendas` via `tipo_produto`.
6. Para `revendas`:
   - Monta query para `SB1010` no Protheus obtendo `B1_UPRC` (último preço).
   - Preenche `ultimo_preco_compra`, calcula diferenças e prepara `df_revendas_final` com colunas relevantes.
7. Para `vendas` (produção):
   - Executa `querys/preco_base.sql` no Protheus e constrói `precos_base`.
   - Filtra preços para os produtos em análise, faz merge e calcula diferenças; produz `df_vendas_final`.
8. Gera o Excel com `pandas.ExcelWriter` + `xlsxwriter`.
   - Aplica formatação por coluna com heurística (nomes como 'custo', 'preco', 'valor' → formato currency; 'margem' → percentual).
   - Insere formatação condicional (ex.: diffs negativos em revenda coloridos em vermelho; divergências > 10% em produção destacadas em amarelo).

## Arquivos / Queries envolvidos

- `querys/busca_nfs_fora_da_margem.sql` — query principal que alimenta o conjunto de NFs a serem analisadas.
- `querys/preco_base.sql` — query para construir tabela de preços base por produto.

Observação: confirmar com equipe de dados se os filtros e joins dessas queries estão atualizados (filiais, flags de exclusão, etc.).

## Pontos de validação para stakeholders (checklist)

1. Execução:
   - [ ] O script roda com `python3 analise_inicial.py` e gera arquivo `analise_margem_YYYY-MM-DD.xlsx` sem exceptions.
2. Conectividade e credenciais:
   - [ ] As variáveis de ambiente estão definidas (.env) e as conexões MySQL/Protheus retornam dados.
3. Cobertura dos dados:
   - [ ] A aba "Todas NFs Fora Margem" contém as colunas esperadas e o número total de linhas condiz com a query de origem.
   - [ ] A aba "Analise Revenda" contém produtos com `ultimo_preco_compra` preenchido quando existe correspondência; divergências podem ser verificadas em amostras.
   - [ ] A aba "Analise Producao" contém `preco_base` para os produtos encontrados e as colunas `diff_valor` e `diff_percentual` calculadas.
4. Regras de negócio:
   - [ ] Validar que a separação entre `vendas` (PA|PI) e `revendas` está correta com amostras reais.
   - [ ] Validar valores de parâmetro padrão de margem: parceiros=0.15, margem=0.27, margem_maxima=0.7.
5. Formatação e alertas:
   - [ ] Verificar formatação condicional (linhas pintadas conforme regra) e se as colunas alvo existem.

## Exemplos de testes rápidos (execução local)

1. Configurar `.env` com as variáveis necessárias. Exemplo mínimo:

```
MYSQL_DB_HOST=meu_mysql
MYSQL_DB_USER=usuario
MYSQL_DB_PASSWORD=senha
MYSQL_DB_DATABASE=banco
MYSQL_DB_PORT=3306

PROTHEUS_ODBC_DRIVER={ODBC Driver}
PROTHEUS_DB_HOST=meu_protheus
PROTHEUS_DB_DATABASE=dbp
PROTHEUS_DB_USER=usr
PROTHEUS_DB_PASSWORD=senha
```

2. Instalar dependências (arquivo `requirements.txt` já presente):

```bash
pip install -r requirements.txt
```

3. Rodar o script:

```bash
python3 analise_inicial.py
```

4. Conferir arquivo `analise_margem_YYYY-MM-DD.xlsx` gerado na pasta do script.

## Possíveis erros / riscos conhecidos

- Falha de conexão com MySQL/Protheus (credenciais ou rede) → script captura exceções e imprime erro; para debug, pode-se descomentar `raise e` no final.
- `MYSQL_DB_PORT` esperado como inteiro; var env ausente ou mal formatada causa erro ao converter.
- Dados faltantes nas queries externas (preço não encontrado) resultam em `0.0` no `ultimo_preco_compra` / `preco_base` — isso influencia `diff_percentual` (evita divisão por zero).
- Heurísticas de formatação por nome de coluna podem não cobrir todos os casos; revisar nomes esperados.

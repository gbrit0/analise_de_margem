# Site Análise de Margem

O objetivo deste site é realizar uma implementação online da tabela de análise de margem que centralize as análises, permita a edição de custos, gere estatísticas, registre justificativas e seja rastreável.

## 🚀 Funcionalidades Principais

- **Tabela de Notas e Análise de Margem:** Visualização de notas fiscais de venda com indicadores de margem.
- **Gestão de Custos:** Atualização dinâmica de custos por nota/item, refletindo automaticamente na margem calculada.
- **Registro de Justificativas:** Inserção e edição de justificativas para vendas com margens fora do padrão aprovado.
- **Estatísticas e Dashboard:** Tela de dashboard centralizado (`/estatisticas/`) com visões por período, filial e justificativas de vendas.
- **Exportação Interativa:** Exportação formatada para Excel refletindo fielmente os filtros aplicados pelo usuário no frontend, incluindo formatações condicionais e larguras de coluna ajustadas.
- **Painel de Administração de Justificativas:** Interface customizada (`/justificativas/`) para criação e ativação/desativação de opções de justificativa.
- **Rastreabilidade de Produção (OPs):** Visualização de Ordens de Produção associadas aos lotes.

## 🛠️ Stack Tecnológica

- **Backend:** Python 3.10.12 + Django 5.2.10
- **Banco de Dados Principal (Aplicação):** MySQL (`default`)
- **Banco de Dados ERP (Integração):** SQL Server (`protheus`)
- **Cache Múltiplo:** Redis 
- **Frontend:** HTML, CSS, JavaScript (com bibliotecas como DataTables e AJAX via Fetch API).
- **Exportação para Excel:** `openpyxl`

## 📂 Estrutura do Projeto

- `setup/`: Configurações principais do Django (`settings.py`, `urls.py`).
- `notas/`: App principal contendo os modelos e lógicas de negócios sobre Notas, Custos, Margens, Ordens de Produção (OPs) e Justificativas.
- `users/`: App para controle customizado de usuários e autenticação.
- `templates/`: Arquivos HTML do projeto contendo as interfaces de tabelas, dashboards e formulários.
- `static/`: Recursos estáticos como CSS, scripts JS, imagens.

## ⚙️ Instalação e Configuração

1. Clone o repositório e crie um ambiente virtual:
```bash
python -m venv venv
source venv/bin/activate  # ou venv\\Scripts\\activate no Windows
```

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

3. Configure o arquivo `.env`:
Copie `.env.example` para `.env` e preencha as variáveis de banco de dados (MySQL e MSSQL/Protheus) e as portas do Redis.

4. Execute as migrações:
```bash
python manage.py makemigrations
python manage.py migrate
```

5. Rode o servidor de desenvolvimento:
```bash
python manage.py runserver 0.0.0.0:30035
```

## 🔐 Integração e Acessos
- Integração de dados de faturamento e custos vem diretamente da base Protheus configurada.
- Restrições de acesso (exemplo: admin customizado de justificativas) são validadas através da modelagem de `CustomUser`.

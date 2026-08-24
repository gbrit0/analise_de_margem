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
python manage.py runserver
```

## 🔐 Integração e Acessos
- Integração de dados de faturamento e custos vem diretamente da base Protheus configurada.
- Restrições de acesso (exemplo: admin customizado de justificativas) são validadas através da modelagem de `CustomUser`.

## 🔑 Autenticação (SSO com o portal unificado)

O sistema é servido sob `https://portal.brggeradores.com.br/analise_de_margem/` — a
mesma origem do portal. Ao autenticar, o portal grava o access token do Keycloak
no cookie `brg_access_token`, que portanto também chega até aqui.

Há dois caminhos de entrada:

| Caminho | Quando ocorre | Backend |
|---|---|---|
| **Token do portal** | Usuário clica no card no portal | `notas.auth_backend.KeycloakPortalBackend` |
| **OIDC tradicional** | Acesso direto ao sistema, sem passar pelo portal | `notas.auth_backend.KeycloakOIDCBackend` |

### Fluxo via portal

`setup.middleware.KeycloakPortalAuthMiddleware` lê o token (header
`Authorization: Bearer` ou o cookie), valida assinatura/`exp`/`aud`/`iss` contra
o JWKS do Keycloak e abre a sessão do Django — sem segunda tela de login.

O token é revalidado apenas quando muda (comparação pelo `jti`); nos demais
requests a sessão do Django assume. Isso evita revalidar o JWT a cada clique e,
principalmente, impede que o usuário seja derrubado quando o access token expira
(o portal não renova o cookie).

Se não houver token, ou ele estiver inválido/expirado, o middleware não interfere:
a sessão existente continua valendo e, na ausência dela, o fluxo OIDC assume.
O HTTP 403 fica reservado ao caso de token válido sem permissão.

> A sessão só é aberta em métodos seguros (GET/HEAD/OPTIONS/TRACE). `login()`
> gira o segredo do CSRF, o que invalidaria o token enviado num POST.

### Permissões

Derivadas das roles do Keycloak, seguindo o padrão do `auth_validador`
(`resource_access.analise_de_margem.roles`) e sincronizadas a cada autenticação:

| Role no client | Efeito no Django |
|---|---|
| `is_active` | libera o acesso ao sistema (sem ela → 403) |
| `is_staff` | `user.is_staff = True` |
| `is_superuser` | `user.is_superuser = True` (e `is_staff`) |

Roles de realm `admin` / `global_admin` / `superusuario` concedem acesso e
superusuário (administradores globais do portal).

O **Client ID no Keycloak deve ser idêntico ao prefixo da rota** (`analise_de_margem`),
e o token precisa trazer esse client na audiência (`aud`) — exigência tanto do
`auth_validador` no Nginx quanto deste middleware.

### Logout

`notas.auth_backend.provider_logout` decide o destino: sessões vindas do portal
são encerradas no próprio portal (que apaga o cookie `brg_access_token`); as demais,
no endpoint de logout do Keycloak.

### Testes

O usuário do MySQL de produção não pode criar a base de testes, então use o
settings de teste (SQLite em memória):

```bash
python manage.py test notas.tests_portal_sso --settings=setup.settings_test
```

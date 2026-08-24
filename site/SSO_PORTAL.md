# SSO com o portal unificado — Análise de Margem (17/08/2026)

Réplica do que já está em produção no Portal de Medição. O desenho completo, o
porquê de cada decisão e as armadilhas encontradas estão em
`/DESENVOLVEDORES/unificacao_sistemas/locadores/SSO_PORTAL.md`; aqui fica só o
que é específico deste sistema.

## O que mudou

O sistema já lia o cookie do portal, mas ainda tinha dois furos:

1. **Ainda mandava para o Keycloak.** Quem chegasse sem sessão batia no
   `@login_required` das views de `notas/urls.py` e era enviado ao `LOGIN_URL`,
   que era `/analise_de_margem/oidc/authenticate/` — ou seja, a tela do
   Keycloak, um segundo login depois de a pessoa já ter entrado no portal.
2. **Não renovava o cookie.** O access token vive **5 minutos**
   (`exp - iat = 300s` nos tokens reais do `auth_validador/debug.log`), e quem
   barra a requisição é o nginx, *antes* de o Django ser chamado: a pessoa
   estava conferindo margem e do nada levava "acesso negado" mesmo tendo
   permissão.

Agora:

```
[navegador]
   │  cookie brg_access_token (JWT do Keycloak)
   ▼
[nginx]  auth_request → auth_validador  (aud + role is_active)
   │
   ▼
[Django]
   ├── ProxyPrefixMiddleware               prefixo /analise_de_margem
   ├── AuthenticationMiddleware
   ├── KeycloakPortalAuthMiddleware  ◄──── valida o JWT (JWKS) e ABRE A SESSÃO
   └── ForceSSOMiddleware            ◄──── sem sessão? volta para o PORTAL
```

- **Nunca mais redireciona para o Keycloak.** Sem sessão, ou com token vencido,
  a pessoa vai para `PORTAL_URL`. `LOGIN_URL` passou a ser o portal, então
  qualquer view nova com `@login_required` também cai lá.
- **`SessionRefresh` foi removido** do `MIDDLEWARE`: a única coisa que ele fazia
  era renovar a sessão contra o Keycloak (redirect para o `/auth` com
  `prompt=none`), exatamente o que não deve mais acontecer.
- **A tela pinga `/auth/refresh` a cada 60s** (inline no `base.html`, e também
  ao voltar de segundo plano). O endpoint é do portal e serve a todos os
  sistemas — este projeto só o chama.
- **Chamada de fundo não é redirecionada.** As APIs de custo, justificativa,
  comentário e estatísticas são chamadas por `fetch` a partir das telas; com a
  sessão vencida elas recebem **401 JSON**. Redirecionar devolveria o HTML da
  tela de login dentro do `fetch`.
- **Token válido sem a role `is_active`** recebe uma página 403 dizendo qual
  role falta e em qual client, em vez do 403 seco do nginx.
- **"Sair"** encerra a sessão **só deste app** e devolve a pessoa à tela inicial
  do portal (`notas/auth_backend.provider_logout`). Ver
  ["Sair" é voltar ao portal](#sair-é-voltar-ao-portal), abaixo.
- **Permissões só promovem** por padrão (`KEYCLOAK_ESPELHA_PERMISSOES=False`),
  para não tirar o acesso do superusuário local que administra as
  justificativas e ainda não tem a role no client.
- O fluxo OIDC clássico continua registrado em `/analise_de_margem/oidc/` para
  acesso direto, fora do portal — mas nada mais leva ninguém para lá
  automaticamente.

## "Sair" é voltar ao portal

Num hub, sair de **um** sistema é voltar ao hub — não deslogar de todos.

O caminho anterior mandava para o `PORTAL_LOGOUT_URL`, que apaga o cookie
`brg_access_token`: quem clicasse em "Sair" na Análise de Margem era deslogado
de Medição, Cobrança, Propostas e de tudo o mais ao mesmo tempo. Agora
`provider_logout` devolve `PORTAL_URL`: `logout()` limpa a sessão do Django
deste app e a pessoa cai na tela inicial do portal, com o resto intacto.

**A contrapartida, que é preciso conhecer:** o cookie do portal continua de pé,
então clicar no card de novo reentra na hora, sem pedir nada — o
`KeycloakPortalAuthMiddleware` abre uma sessão nova no primeiro GET. Quem quiser
encerrar tudo usa o "Sair" do próprio portal. Para voltar ao logout global,
basta devolver `settings.PORTAL_LOGOUT_URL` em `provider_logout`; a chave
continua configurada.

Sessão que **não** veio do portal (fluxo OIDC direto) continua encerrando no
Keycloak: essa nasceu aqui, com sessão de navegador no próprio Keycloak, e é
aqui que se encerra.

## Arquivos

| Arquivo | Papel |
|---|---|
| `notas/keycloak.py` | extrai o token (header ou cookie), valida contra o JWKS, traduz roles em permissões |
| `notas/auth_backend.py` | `KeycloakPortalBackend` (sessão vinda do portal), backend OIDC clássico e o `provider_logout` |
| `setup/middleware.py` | `KeycloakPortalAuthMiddleware` (token → sessão) e `ForceSSOMiddleware` (sem sessão → portal) |
| `setup/context.py` | alimenta o ping de renovação no `base.html` |
| `templates/base.html` | o ping de 60s, inline |
| `templates/notas/sem_permissao.html` | 403 explicado, com link de volta ao portal |
| `notas/tests_portal_sso.py` | 35 testes do fluxo |

## Chaves do `.env`

Todas têm padrão correto em `setup/settings.py`; ficam comentadas no `.env` para
documentar o que se pode virar sem mexer no código. As principais:

| Variável | Padrão | Para que serve |
|---|---|---|
| `PORTAL_URL` | `https://portal.brggeradores.com.br/` | para onde vai quem está sem sessão ou com token vencido |
| `PORTAL_LOGOUT_URL` | `.../logout` | logout GLOBAL do hub. Não é usado pelo "Sair" daqui (que só volta ao portal); fica para quem precisar derrubar tudo |
| `PORTAL_REFRESH_URL` | `/auth/refresh` | endpoint do portal que renova o cookie |
| `PORTAL_REFRESH_INTERVALO_SEGUNDOS` | `60` | de quanto em quanto tempo a tela pinga |
| `KEYCLOAK_PORTAL_SSO_ENABLED` | `True` | desliga o SSO por cookie |
| `KEYCLOAK_PORTAL_COOKIE_NAME` | `brg_access_token` | tem de ser o mesmo do portal e do `auth_validador` |
| `KEYCLOAK_ESPELHA_PERMISSOES` | `False` | `False`: roles só promovem. `True`: Keycloak é a única fonte de verdade |
| `KEYCLOAK_VERIFY_AUDIENCE` / `_ISSUER` | `True` | só desligue se faltar o audience mapper no client |

> [!IMPORTANT]
> O JWKS é baixado com **User-Agent próprio** (`KEYCLOAK_JWKS_USER_AGENT`). O
> Cloudflare, na frente do Keycloak, responde **403 ao User-Agent padrão do
> urllib**. Sem esse header nenhum token é validado.

## Rotas que continuam abertas

O sistema não tem link público de cliente nem webhook: toda view de
`notas/urls.py` já era `@login_required`. Ficam fora da exigência de sessão:

- `/analise_de_margem/static/` e `/media/` — estáticos;
- `/analise_de_margem/oidc/` — o fluxo OIDC de acesso direto e o `oidc_logout`;
- `/analise_de_margem/admin/` — o Django admin tem login local próprio (não leva
  ninguém ao Keycloak) e é a saída de emergência para superusuário sem role no
  client.

## Rodando os testes

Nunca contra o MySQL de produção — o `settings_test` usa SQLite em memória:

```bash
docker exec unificacao_analise_de_margem \
    python manage.py test notas.tests_portal_sso --settings=setup.settings_test -v2
```

## nginx

No `location /analise_de_margem/` de `/etc/nginx/sites-available/portal.conf`:

```nginx
error_page 401 = @volta_ao_portal;   # token ausente/vencido -> tela inicial do portal
error_page 403 =302 /403;            # autenticado, sem a role -> acesso negado
```

`error_page` no `location` **substitui** o do `server`, por isso as duas linhas.
O named location `@volta_ao_portal` já existia desde 17/08/2026. Backup do
arquivo em `portal.conf.bak-20260817-sso-analise_de_margem`.

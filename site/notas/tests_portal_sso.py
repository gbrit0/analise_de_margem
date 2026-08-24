"""
Testes do SSO com o portal unificado.

Rodar (NUNCA contra o MySQL de produção — o settings_test usa sqlite em memória):

    docker exec unificacao_analise_de_margem \
        python manage.py test notas.tests_portal_sso --settings=setup.settings_test -v2

Duas camadas são exercitadas aqui:

1. `ValidacaoTokenTest` / `PermissoesTest` — a validação criptográfica de
   verdade. Um par de chaves RSA é gerado em memória e o cliente JWKS é
   substituído, de modo que assinatura, `exp`, `aud` e `iss` são conferidos pela
   PyJWT como em produção.

2. `SsoDoPortalTest` / `TokenVencidoTest` / `SemSessaoTest` — a decisão do
   middleware, pelo `django.test.Client`, com a cadeia de middlewares real.
   Aqui `validar_token` é dublado: o que está sob teste é o comportamento
   (redireciona? 401? 403? refaz o login?), não a criptografia.

O que estes testes protegem, que é justamente o que quebrava antes:
  - cookie do portal com token válido entra SEM segunda tela de login;
  - token vencido/ausente manda para a URL INICIAL DO PORTAL, nunca para o
    Keycloak;
  - chamada de fundo (as APIs de custo/justificativa/estatísticas, que a tela
    dispara por fetch) recebe 401 em vez de redirecionamento;
  - token válido sem a role `is_active` recebe 403 explicando o que falta;
  - o cookie renovado a cada 60s não refaz o login (e não gira o CSRF).
"""

import json
import time
from unittest import mock

import jwt
from cryptography.hazmat.primitives.asymmetric import rsa
from django.conf import settings
from django.contrib import admin
from django.contrib.auth import get_user_model
from django.shortcuts import render
from django.test import TestCase, RequestFactory, override_settings
from django.urls import include, path

from notas import keycloak
from setup.middleware import KeycloakPortalAuthMiddleware

CHAVE_PRIVADA = rsa.generate_private_key(public_exponent=65537, key_size=2048)
CHAVE_PUBLICA = CHAVE_PRIVADA.public_key()

PORTAL = 'https://portal.teste.local/'


def gerar_token(roles_client=('is_active',), roles_realm=(), expirado=False, **extra):
    agora = int(time.time())
    claims = {
        'iss': settings.KEYCLOAK_ISSUER,
        'aud': [settings.OIDC_RP_CLIENT_ID, 'account'],
        'azp': 'admin-api-client',
        'sub': 'uuid-do-usuario',
        'jti': extra.pop('jti', 'token-1'),
        'iat': agora - 60,
        'exp': agora - 30 if expirado else agora + 300,
        'preferred_username': 'joao.silva',
        'email': 'joao.silva@brggeradores.com.br',
        'given_name': 'João',
        'family_name': 'Silva',
        'realm_access': {'roles': list(roles_realm)},
        'resource_access': {settings.OIDC_RP_CLIENT_ID: {'roles': list(roles_client)}},
    }
    claims.update(extra)
    return jwt.encode(claims, CHAVE_PRIVADA, algorithm='RS256')


class JWKFalso:
    key = CHAVE_PUBLICA

    def get_signing_key_from_jwt(self, token):
        return self


class BasePortalTest(TestCase):
    def setUp(self):
        self.factory = RequestFactory()
        patch = mock.patch.object(keycloak, '_get_jwk_client', return_value=JWKFalso())
        patch.start()
        self.addCleanup(patch.stop)


# ---------------------------------------------------------------------------
# 1. Validação criptográfica (PyJWT de verdade contra um JWKS falso)
# ---------------------------------------------------------------------------

class ValidacaoTokenTest(BasePortalTest):
    def test_token_valido_devolve_claims(self):
        claims = keycloak.validar_token(gerar_token())
        self.assertEqual(claims['preferred_username'], 'joao.silva')

    def test_token_expirado_e_rejeitado(self):
        with self.assertRaises(keycloak.TokenInvalido):
            keycloak.validar_token(gerar_token(expirado=True))

    def test_assinatura_invalida_e_rejeitada(self):
        outra_chave = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        token = jwt.encode({'sub': 'x'}, outra_chave, algorithm='RS256')
        with self.assertRaises(keycloak.TokenInvalido):
            keycloak.validar_token(token)

    def test_audiencia_de_outro_sistema_e_rejeitada(self):
        with self.assertRaises(keycloak.TokenInvalido):
            keycloak.validar_token(gerar_token(aud=['outro_sistema']))

    def test_emissor_diferente_e_rejeitado(self):
        with self.assertRaises(keycloak.TokenInvalido):
            keycloak.validar_token(gerar_token(iss='https://keycloak.invasor.com/realms/x'))

    def test_identificador_le_o_jti_sem_validar(self):
        """O `jti` é lido sem assinatura de propósito: sozinho não concede nada,
        só evita revalidar o JWT a cada clique."""
        self.assertEqual(
            keycloak.identificador_do_token(gerar_token(jti='abc', expirado=True)), 'abc')


class PermissoesTest(BasePortalTest):
    def test_is_active_libera_acesso(self):
        self.assertTrue(keycloak.tem_acesso(keycloak.validar_token(gerar_token())))

    def test_sem_is_active_bloqueia(self):
        claims = keycloak.validar_token(gerar_token(roles_client=()))
        self.assertFalse(keycloak.tem_acesso(claims))

    def test_admin_global_do_realm_libera_acesso(self):
        claims = keycloak.validar_token(gerar_token(roles_client=(), roles_realm=('global_admin',)))
        self.assertTrue(keycloak.tem_acesso(claims))

    def test_roles_viram_permissoes_do_django(self):
        claims = keycloak.validar_token(gerar_token(roles_client=('is_active', 'is_superuser')))
        user = keycloak.obter_ou_criar_usuario(claims)
        self.assertTrue(user.is_superuser)
        self.assertTrue(user.is_staff)
        self.assertEqual(user.username, 'joao.silva')
        self.assertEqual(user.first_name, 'João')

    def test_usuario_comum_nao_vira_superusuario(self):
        claims = keycloak.validar_token(gerar_token())
        user = keycloak.obter_ou_criar_usuario(claims)
        self.assertFalse(user.is_superuser)
        self.assertFalse(user.is_staff)

    def test_conta_existente_e_reaproveitada_pelo_email(self):
        """Conta anterior ao SSO, com username diferente do Keycloak, não pode
        virar uma conta nova — a pessoa perderia de vista o que já era dela."""
        User = get_user_model()
        antiga = User.objects.create_user(
            username='jsilva', email='joao.silva@brggeradores.com.br')
        # as migrações já semeiam usuários (USER_SYSTEM/USER_ADMIN): o que
        # importa é que NENHUMA conta nova nasça deste login
        antes = User.objects.count()

        user = keycloak.obter_ou_criar_usuario(keycloak.validar_token(gerar_token()))

        self.assertEqual(user.pk, antiga.pk)
        self.assertEqual(User.objects.count(), antes)

    def test_sem_permissao_levanta_excecao(self):
        claims = keycloak.validar_token(gerar_token(roles_client=()))
        with self.assertRaises(keycloak.SemPermissao):
            keycloak.obter_ou_criar_usuario(claims)

    @override_settings(KEYCLOAK_ESPELHA_PERMISSOES=False)
    def test_por_padrao_nao_rebaixa_admin_local(self):
        """O superusuário que administra as justificativas pode não ter a role
        no client ainda; espelhar de cara tiraria o acesso dele."""
        user = get_user_model().objects.create_user(username='admin.local')
        user.is_staff = True
        user.is_superuser = True
        user.save()

        keycloak.aplicar_permissoes(user, keycloak.validar_token(gerar_token()))

        user.refresh_from_db()
        self.assertTrue(user.is_staff)
        self.assertTrue(user.is_superuser)

    @override_settings(KEYCLOAK_ESPELHA_PERMISSOES=True)
    def test_espelhando_o_keycloak_manda(self):
        user = get_user_model().objects.create_user(username='admin.local2')
        user.is_staff = True
        user.is_superuser = True
        user.save()

        keycloak.aplicar_permissoes(user, keycloak.validar_token(gerar_token()))

        user.refresh_from_db()
        self.assertFalse(user.is_staff)
        self.assertFalse(user.is_superuser)


# ---------------------------------------------------------------------------
# 2. Comportamento do middleware, com a cadeia real, pelo test client
# ---------------------------------------------------------------------------

CLAIMS_OK = {
    'jti': 'token-1',
    'preferred_username': 'gabriel.brito',
    'email': 'gabriel.brito@brggeradores.com.br',
    'given_name': 'Gabriel',
    'family_name': 'Brito',
    'aud': ['analise_de_margem'],
    'realm_access': {'roles': []},
    'resource_access': {'analise_de_margem': {'roles': ['is_active', 'is_superuser']}},
}

CLAIMS_SEM_ROLE = {
    'jti': 'token-2',
    'preferred_username': 'sem.acesso',
    'email': 'sem.acesso@brggeradores.com.br',
    'aud': ['analise_de_margem'],
    'realm_access': {'roles': []},
    'resource_access': {'analise_de_margem': {'roles': ['uma_protection']}},
}


def jwt_falso(claims):
    """JWT com assinatura de brinquedo.

    A assinatura não importa (`validar_token` está dublado nestas classes), mas
    o token PRECISA ser um JWT de verdade na forma: o middleware lê o `jti` do
    próprio token, sem verificar assinatura, para saber se a sessão atual já
    nasceu dele.
    """
    return jwt.encode(claims, 'segredo-de-teste-com-32-bytes-ou-mais', algorithm='HS256')


def _tela(request):
    """Tela qualquer que estende o base.html — é onde o ping é injetado."""
    return render(request, 'base.html')


def _api_de_fundo(request):
    from django.http import JsonResponse
    return JsonResponse({'ok': True})


# urlconf de teste: mantém as rotas reais (o base.html usa {% url %} de
# 'lista_notas', 'estatisticas', 'admin_justificativas' e 'oidc_logout') e
# acrescenta duas views baratas, para não arrastar as consultas ao Protheus
# para dentro de um teste de autenticação.
urlpatterns = [
    path('tela/', _tela),
    path('api/fundo/', _api_de_fundo),
    path('admin/', admin.site.urls),
    path('', include('notas.urls')),
    path('oidc/', include('mozilla_django_oidc.urls')),
]

URLCONF_TESTE = __name__

# Cabeçalhos de uma navegação de verdade (é o que autoriza redirecionamento).
NAVEGACAO = {'HTTP_ACCEPT': 'text/html', 'HTTP_SEC_FETCH_MODE': 'navigate'}
# Cabeçalhos do fetch() de fundo da tela (salvar custo, buscar estatísticas...).
FUNDO = {'HTTP_ACCEPT': 'application/json', 'HTTP_SEC_FETCH_MODE': 'same-origin'}


@override_settings(ROOT_URLCONF=URLCONF_TESTE, PORTAL_URL=PORTAL)
class SsoDoPortalTest(TestCase):
    """O caminho normal: quem já entrou no portal não loga de novo."""

    def setUp(self):
        self.User = get_user_model()

    def _com_token(self, claims):
        """Faz o token do cookie valer as claims informadas."""
        patcher = mock.patch.object(keycloak, 'validar_token', return_value=claims)
        self.addCleanup(patcher.stop)
        return patcher.start()

    def test_cookie_do_portal_autentica_sem_segunda_tela_de_login(self):
        self._com_token(CLAIMS_OK)
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)

        resposta = self.client.get('/tela/', **NAVEGACAO)

        # não houve redirecionamento para login algum: a página abriu
        self.assertEqual(resposta.status_code, 200)
        user = self.User.objects.get(username='gabriel.brito')
        self.assertEqual(int(self.client.session['_auth_user_id']), user.pk)
        self.assertTrue(self.client.session[settings.KEYCLOAK_PORTAL_SESSION_KEY])
        self.assertEqual(self.client.session[settings.KEYCLOAK_PORTAL_TOKEN_KEY], 'token-1')
        # roles do client viraram permissões do Django
        self.assertTrue(user.is_superuser)
        self.assertEqual(user.first_name, 'Gabriel')

    def test_a_tela_sai_com_o_ping_de_renovacao(self):
        """Sem o ping o token vence em 5 min e o nginx barra quem está trabalhando."""
        self._com_token(CLAIMS_OK)
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)

        conteudo = self.client.get('/tela/', **NAVEGACAO).content.decode()

        self.assertIn(settings.PORTAL_REFRESH_URL, conteudo)
        self.assertIn('visibilitychange', conteudo)

    def test_sessao_fora_do_portal_nao_recebe_o_ping(self):
        """Não há cookie do portal a renovar; o script só polui a página."""
        user = self.User.objects.create_user(username='direto')
        self.client.force_login(user)

        conteudo = self.client.get('/tela/', **NAVEGACAO).content.decode()

        self.assertNotIn(settings.PORTAL_REFRESH_URL, conteudo)

    def test_token_valido_sem_role_recebe_403_explicado(self):
        self._com_token(CLAIMS_SEM_ROLE)
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_SEM_ROLE)

        resposta = self.client.get('/tela/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 403)
        conteudo = resposta.content.decode()
        self.assertIn('is_active', conteudo)
        self.assertIn('analise_de_margem', conteudo)
        self.assertIn(PORTAL, conteudo)
        self.assertNotIn('_auth_user_id', self.client.session)

    def test_segundo_request_nao_revalida_o_mesmo_token(self):
        """Validar JWT (e ir ao banco) a cada clique seria desperdício: o `jti`
        guardado na sessão evita isso."""
        validar = self._com_token(CLAIMS_OK)
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)

        self.client.get('/tela/', **NAVEGACAO)
        self.assertEqual(validar.call_count, 1)

        self.client.get('/tela/', **NAVEGACAO)
        self.assertEqual(validar.call_count, 1)

    def test_cookie_renovado_nao_refaz_o_login(self):
        """O portal renova o cookie a cada 60s (/auth/refresh).

        Se cada token novo refizesse o login, `rotate_token()` trocaria o
        segredo do CSRF e o formulário já aberto na tela passaria a dar 403 no
        meio do trabalho. Token novo da MESMA pessoa só atualiza o `jti`.
        """
        validar = self._com_token(CLAIMS_OK)
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)
        self.client.get('/tela/', **NAVEGACAO)

        chave_da_sessao = self.client.session.session_key
        segredo_csrf = self.client.cookies[settings.CSRF_COOKIE_NAME].value

        renovado = dict(CLAIMS_OK, jti='token-renovado')
        validar.return_value = renovado
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(renovado)
        resposta = self.client.get('/tela/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 200)
        self.assertEqual(self.client.session.session_key, chave_da_sessao)
        self.assertEqual(
            self.client.session[settings.KEYCLOAK_PORTAL_TOKEN_KEY], 'token-renovado')
        self.assertEqual(self.client.cookies[settings.CSRF_COOKIE_NAME].value, segredo_csrf)

    def test_promocao_de_permissoes_no_token_renovado(self):
        """Ganhar a role no Keycloak reflete aqui sem precisar sair e entrar."""
        validar = self._com_token(
            dict(CLAIMS_OK, resource_access={'analise_de_margem': {'roles': ['is_active']}}))
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)
        self.client.get('/tela/', **NAVEGACAO)
        self.assertFalse(self.User.objects.get(username='gabriel.brito').is_superuser)

        promovido = dict(CLAIMS_OK, jti='token-3')
        validar.return_value = promovido
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(promovido)
        self.client.get('/tela/', **NAVEGACAO)

        self.assertTrue(self.User.objects.get(username='gabriel.brito').is_superuser)


@override_settings(ROOT_URLCONF=URLCONF_TESTE, PORTAL_URL=PORTAL,
                   PORTAL_LOGOUT_URL=PORTAL + 'logout')
class LogoutTest(TestCase):
    """"Sair" encerra a sessão SÓ deste app e devolve a pessoa ao portal."""

    def _sessao_do_portal(self):
        with mock.patch.object(keycloak, 'validar_token', return_value=CLAIMS_OK):
            self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)
            self.client.get('/tela/', **NAVEGACAO)
        self.assertIn('_auth_user_id', self.client.session)

    def test_sair_volta_ao_portal_sem_derrubar_a_sessao_global(self):
        """Mandar para o PORTAL_LOGOUT_URL apagaria o cookie do portal e
        deslogaria a pessoa de TODOS os sistemas do hub de uma vez."""
        self._sessao_do_portal()

        resposta = self.client.post('/oidc/logout/')

        self.assertEqual(resposta.status_code, 302)
        self.assertEqual(resposta['Location'], PORTAL)
        self.assertNotIn('logout', resposta['Location'])
        self.assertNotIn('keycloak', resposta['Location'])
        # a sessão local caiu...
        self.assertNotIn('_auth_user_id', self.client.session)
        self.assertNotIn(settings.KEYCLOAK_PORTAL_SESSION_KEY, self.client.session)
        # ...mas o cookie do portal continua de pé (é o que mantém os outros
        # sistemas logados; a contrapartida é reentrar na hora pelo card)
        self.assertIn(settings.KEYCLOAK_PORTAL_COOKIE_NAME, self.client.cookies)

    def test_sessao_fora_do_portal_ainda_encerra_no_keycloak(self):
        """Quem entrou pelo /oidc/ criou uma sessão de navegador NO Keycloak;
        essa nasceu aqui e é aqui que se encerra."""
        user = get_user_model().objects.create_user(username='direto')
        self.client.force_login(user)

        resposta = self.client.post('/oidc/logout/')

        self.assertEqual(resposta.status_code, 302)
        self.assertIn('keycloak', resposta['Location'])
        self.assertNotIn('_auth_user_id', self.client.session)


@override_settings(ROOT_URLCONF=URLCONF_TESTE, PORTAL_URL=PORTAL)
class TokenVencidoTest(TestCase):
    """Token vencido vai para o portal — nunca para a tela do Keycloak."""

    def setUp(self):
        self.User = get_user_model()

    def _sessao_do_portal(self):
        """Deixa o navegador com uma sessão aberta pelo portal."""
        with mock.patch.object(keycloak, 'validar_token', return_value=CLAIMS_OK):
            self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(CLAIMS_OK)
            self.client.get('/tela/', **NAVEGACAO)
        self.assertIn('_auth_user_id', self.client.session)

    def _token_vencido(self):
        self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = jwt_falso(
            dict(CLAIMS_OK, jti='token-vencido'))
        return mock.patch.object(
            keycloak, 'validar_token', side_effect=keycloak.TokenInvalido('Token expirado'))

    def test_navegacao_com_token_vencido_volta_para_o_portal(self):
        self._sessao_do_portal()

        with self._token_vencido():
            resposta = self.client.get('/tela/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 302)
        self.assertEqual(resposta['Location'], PORTAL)
        self.assertNotIn('keycloak', resposta['Location'])
        # a sessão local foi encerrada junto
        self.assertNotIn('_auth_user_id', self.client.session)

    def test_cookie_apagado_no_portal_encerra_a_sessao_daqui(self):
        """Sair no portal apaga o cookie; a sessão daqui não pode continuar, ou
        'Sair' não sairia de nada."""
        self._sessao_do_portal()
        del self.client.cookies[settings.KEYCLOAK_PORTAL_COOKIE_NAME]

        resposta = self.client.get('/tela/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 302)
        self.assertEqual(resposta['Location'], PORTAL)
        self.assertNotIn('_auth_user_id', self.client.session)

    def test_chamada_de_fundo_recebe_401_e_nao_redirecionamento(self):
        """A tela salva custo/justificativa e busca estatísticas por fetch.
        Redirecionar essas chamadas devolveria HTML dentro do fetch e roubaria o
        destino da navegação."""
        self._sessao_do_portal()

        with self._token_vencido():
            resposta = self.client.get('/api/fundo/', **FUNDO)

        self.assertEqual(resposta.status_code, 401)
        self.assertIn('erro', json.loads(resposta.content))

    def test_sessao_de_fluxo_oidc_direto_nao_e_derrubada(self):
        """Quem entrou pelo /oidc/ (fora do portal) — ou tem sessão antiga, de
        antes desta mudança — não pode ser derrubado por um cookie do portal que
        nunca foi dele. É o que garante que o deploy não desloga ninguém."""
        user = self.User.objects.create_user(username='direto', email='d@brg.com')
        self.client.force_login(user)

        with self._token_vencido():
            resposta = self.client.get('/tela/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 200)
        self.assertIn('_auth_user_id', self.client.session)


@override_settings(ROOT_URLCONF=URLCONF_TESTE, PORTAL_URL=PORTAL)
class SemSessaoTest(TestCase):
    """Ninguém mais é enviado ao Keycloak por falta de sessão."""

    def test_sem_cookie_e_sem_sessao_vai_para_o_portal(self):
        resposta = self.client.get('/tela/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 302)
        self.assertEqual(resposta['Location'], PORTAL)
        self.assertNotIn('keycloak', resposta['Location'])

    def test_chamada_de_fundo_sem_sessao_recebe_401(self):
        resposta = self.client.get('/api/fundo/', **FUNDO)

        self.assertEqual(resposta.status_code, 401)

    def test_login_url_aponta_para_o_portal(self):
        """`@login_required` das views de notas.urls não pode cair no Keycloak.

        Comparado com o valor REAL de configuração (o `override_settings` desta
        classe só troca o PORTAL_URL usado nos redirecionamentos; o LOGIN_URL é
        resolvido na importação do settings, que é exatamente o que precisa
        estar certo em produção).
        """
        from setup import settings as settings_reais

        self.assertEqual(settings.LOGIN_URL, settings_reais.PORTAL_URL)
        self.assertNotIn('keycloak', settings.LOGIN_URL)
        self.assertNotIn('/oidc/', settings.LOGIN_URL)

    def test_admin_local_continua_alcancavel(self):
        """O Django admin tem login próprio (não leva ao Keycloak) e é a saída
        de emergência para superusuário sem role no client."""
        resposta = self.client.get('/admin/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 302)
        self.assertIn('/admin/login/', resposta['Location'])
        self.assertNotIn(PORTAL, resposta['Location'])

    def test_rota_oidc_continua_aberta_para_acesso_direto(self):
        """O fluxo OIDC segue registrado para quem acessa fora do portal — mas
        nada mais leva ninguém para lá automaticamente."""
        resposta = self.client.get('/oidc/authenticate/', **NAVEGACAO)

        self.assertEqual(resposta.status_code, 302)
        self.assertNotEqual(resposta['Location'], PORTAL)


@override_settings(ROOT_URLCONF=URLCONF_TESTE, PORTAL_URL=PORTAL)
class MetodosInsegurosTest(BasePortalTest):
    """POST não pode abrir sessão: `login()` gira o segredo do CSRF antes de o
    CsrfViewMiddleware validar o formulário."""

    def _rodar(self, request):
        chamou = {'view': False}

        def get_response(_req):
            from django.http import HttpResponse
            chamou['view'] = True
            return HttpResponse('ok')

        resposta = KeycloakPortalAuthMiddleware(get_response)(request)
        return resposta, chamou['view']

    def _post(self, token):
        from django.contrib.auth.models import AnonymousUser
        request = self.factory.post('/api/atualizar-custo/')
        request.session = self.client.session
        request.user = AnonymousUser()
        request.COOKIES[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = token
        return request

    def test_post_nao_abre_sessao(self):
        request = self._post(gerar_token())

        _, chamou_view = self._rodar(request)

        self.assertTrue(chamou_view)
        self.assertFalse(request.user.is_authenticated)

    def test_post_sem_permissao_continua_bloqueado(self):
        request = self._post(gerar_token(roles_client=()))

        resposta, chamou_view = self._rodar(request)

        self.assertEqual(resposta.status_code, 403)
        self.assertFalse(chamou_view)

    def test_post_com_sessao_valida_passa_direto(self):
        user = get_user_model().objects.create_user(username='joao.silva')
        request = self._post(gerar_token())
        request.user = user

        resposta, chamou_view = self._rodar(request)

        self.assertTrue(chamou_view)
        self.assertEqual(resposta.status_code, 200)

    def test_rotas_ignoradas_nao_passam_pelo_token(self):
        from django.contrib.auth.models import AnonymousUser
        request = self.factory.get('/oidc/authenticate/')
        request.session = self.client.session
        request.user = AnonymousUser()
        request.COOKIES[settings.KEYCLOAK_PORTAL_COOKIE_NAME] = gerar_token()

        self._rodar(request)

        self.assertFalse(request.user.is_authenticated)

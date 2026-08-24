import logging

from django.conf import settings
from django.contrib.auth import authenticate, login, logout
from django.http import JsonResponse
from django.shortcuts import redirect, render
from django.urls import set_script_prefix

from notas import keycloak

logger = logging.getLogger(__name__)


class ProxyPrefixMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        prefix = request.META.get('HTTP_X_FORWARDED_PREFIX') or '/analise_de_margem'
        prefix = prefix.rstrip('/')
        if prefix:
            set_script_prefix(prefix + '/')
            if request.path_info.startswith(prefix):
                request.META['SCRIPT_NAME'] = prefix
                new_path_info = request.path_info[len(prefix):]
                if not new_path_info.startswith('/'):
                    new_path_info = '/' + new_path_info
                request.path_info = new_path_info
                request.META['PATH_INFO'] = new_path_info
            elif not request.META.get('SCRIPT_NAME'):
                request.META['SCRIPT_NAME'] = prefix
        return self.get_response(request)


def _e_navegacao(request):
    """
    Só navegação de verdade pode ser REDIRECIONADA para fora do sistema.

    As telas da Análise de Margem vivem de chamada de fundo: a lista de notas
    salva custo/justificativa/comentário por `fetch`, o dashboard busca
    `/api/estatisticas/` e `/api/estatisticas/vendedor/notas/` enquanto a pessoa
    navega. Redirecionar essas chamadas para o portal devolveria o HTML da tela
    de login dentro do `fetch` (ou, pior, sequestraria o destino do login e a
    pessoa terminaria em cima de um JSON). Chamada de fundo recebe 401 e o JS a
    ignora; quem redireciona é só a navegação.
    """
    if request.method not in ('GET', 'HEAD'):
        return False
    if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        return False
    # fetch() manda Sec-Fetch-Mode: cors/same-origin; navegação manda navigate
    modo = request.headers.get('Sec-Fetch-Mode')
    if modo and modo != 'navigate':
        return False
    return 'text/html' in request.headers.get('Accept', 'text/html')


def _volta_para_o_portal(request, motivo):
    """
    Resposta para quem chegou sem sessão válida.

    Navegação vai para a URL INICIAL DO PORTAL (`PORTAL_URL`) — nunca para a
    tela do Keycloak. É o portal quem tem a interface de login com a nossa
    identidade; o Keycloak fica sendo só o backend de autenticação, sem aparecer
    para o usuário.
    """
    if not _e_navegacao(request):
        return JsonResponse(
            {'erro': 'sessao expirada, recarregue a pagina', 'motivo': motivo},
            status=401,
        )
    return redirect(settings.PORTAL_URL)


class KeycloakPortalAuthMiddleware:
    """
    Autentica o usuário a partir do token que o portal unificado já obteve do
    Keycloak, dispensando um segundo login.

    Fluxo:
      1. Procura o token no header `Authorization: Bearer` ou no cookie
         `brg_access_token` gravado pelo portal.
      2. Se o token for o mesmo que já abriu a sessão atual (mesmo `jti`),
         não faz nada — a sessão do Django assume dali em diante e não se
         revalida JWT (nem se vai ao banco) a cada clique.
      3. Caso contrário, valida o token contra o JWKS do Keycloak e faz o
         login, sincronizando as permissões a partir das roles.

    Quando o token vence ou o cookie some:
      - sessão que veio do portal é encerrada e a pessoa volta para a URL
        inicial do portal (nunca para o Keycloak);
      - sessão de outra origem (fluxo OIDC direto, login local do admin) não é
        tocada — um cookie que nunca foi dela não pode derrubá-la.

    Deve vir DEPOIS de `AuthenticationMiddleware` (precisa de `request.user`) e
    ANTES do `ForceSSOMiddleware`.
    """

    #: caminhos que nunca precisam da autenticação por token. `/admin/` fica de
    #: fora porque o Django admin tem login local próprio (não leva ninguém ao
    #: Keycloak) e é a saída de emergência para superusuário sem role no client.
    CAMINHOS_IGNORADOS = ('/oidc/', '/static/', '/media/', '/admin/')

    #: A sessão só é aberta em métodos seguros. `django.contrib.auth.login`
    #: chama `rotate_token()`, que troca o segredo do CSRF — e isso acontece
    #: antes de `CsrfViewMiddleware.process_view` validar o token enviado no
    #: formulário, o que faria todo POST que disparasse o login falhar com 403.
    #: Na prática não há perda: o usuário sempre chega ao sistema por um GET
    #: (o card do portal), e os POSTs seguintes usam a sessão já estabelecida.
    METODOS_SEGUROS = ('GET', 'HEAD', 'OPTIONS', 'TRACE')

    def __init__(self, get_response):
        self.get_response = get_response
        self.ativo = getattr(settings, 'KEYCLOAK_PORTAL_SSO_ENABLED', True)
        self.chave_sessao = settings.KEYCLOAK_PORTAL_SESSION_KEY
        self.chave_token = settings.KEYCLOAK_PORTAL_TOKEN_KEY

    def __call__(self, request):
        if self.ativo and not self._ignorar(request):
            resposta = self._autenticar(request)
            if resposta is not None:
                return resposta
        return self.get_response(request)

    def _ignorar(self, request):
        return any(request.path_info.startswith(p) for p in self.CAMINHOS_IGNORADOS)

    def _autenticar(self, request):
        """Devolve uma resposta para interromper o request, ou None para seguir."""
        token = keycloak.extrair_token(request)
        if not token:
            # Sem cookie: ou a pessoa nunca passou pelo portal, ou saiu de lá
            # (o /logout do portal apaga o cookie). Quem entrou pelo portal não
            # pode continuar aqui com a sessão órfã.
            return self._encerrar_se_veio_do_portal(request, 'sem token do portal')

        identificador = keycloak.identificador_do_token(token)

        # Sessão já aberta por este mesmo token: nada a fazer. Evita revalidar
        # o JWT (e ir ao banco) a cada request.
        if (
            request.user.is_authenticated
            and identificador is not None
            and request.session.get(self.chave_token) == identificador
        ):
            return None

        try:
            claims = keycloak.validar_token(token)
        except keycloak.TokenInvalido as erro:
            return self._encerrar_se_veio_do_portal(request, str(erro))

        if not keycloak.tem_acesso(claims):
            logger.warning(
                "Acesso negado a '%s': sem a role '%s' no client '%s'",
                claims.get('preferred_username'), keycloak.ROLE_ACESSO,
                settings.OIDC_RP_CLIENT_ID,
            )
            return self._sem_permissao(request, claims)

        if request.method not in self.METODOS_SEGUROS:
            # Não abre sessão aqui (ver METODOS_SEGUROS). Quem já está
            # autenticado segue com a sessão atual; quem não está cai no
            # tratamento normal de não autenticado.
            return None

        user = authenticate(request, claims=claims)
        if user is None:
            return None

        if request.user.is_authenticated and request.user.pk == user.pk:
            # Mesma pessoa, token NOVO — é o portal renovando o cookie
            # (/auth/refresh, a cada 60s). Aqui NÃO se refaz o login:
            # `login()` chama `rotate_token()`, que troca o segredo do CSRF, e o
            # formulário de custo/justificativa já aberto na tela passaria a dar
            # 403 no meio do uso — a cada poucos minutos, por causa do ping.
            # Só se anota qual token vale agora; as permissões já foram
            # ressincronizadas pelo authenticate() acima.
            request.session[self.chave_sessao] = True
            request.session[self.chave_token] = identificador
            return None

        # Troca de usuário no mesmo navegador: encerra a sessão anterior antes.
        if request.user.is_authenticated:
            logout(request)

        login(request, user, backend='notas.auth_backend.KeycloakPortalBackend')
        request.session[self.chave_sessao] = True
        request.session[self.chave_token] = identificador
        logger.info("Usuário '%s' autenticado via portal unificado", user.username)
        return None

    def _encerrar_se_veio_do_portal(self, request, motivo):
        """
        Token vencido/ausente derruba SÓ quem entrou pelo portal.

        Sessão aberta por outro caminho (fluxo OIDC direto, login local do
        admin) não pode ser encerrada por causa de um cookie que nunca foi dela.
        Sessão antiga, criada antes desta mudança, não tem a chave
        `portal_sso` — logo é tratada como "não veio do portal" e segue valendo:
        ninguém é derrubado pelo deploy.
        """
        if not request.session.get(self.chave_sessao):
            return None

        logger.info(
            "Sessão do portal encerrada para '%s': %s",
            getattr(request.user, 'username', '?'), motivo,
        )
        if request.user.is_authenticated:
            logout(request)
        return _volta_para_o_portal(request, motivo)

    def _sem_permissao(self, request, claims):
        """
        Token válido, usuário sem a role deste sistema. Página explicando o que
        falta (e para quem pedir), em vez do 403 seco do nginx: aqui já se sabe
        QUEM é a pessoa e QUAL role falta.
        """
        contexto = {
            'usuario': claims.get('preferred_username') or claims.get('email') or '',
            'sistema': settings.OIDC_RP_CLIENT_ID,
            'role': keycloak.ROLE_ACESSO,
            'portal_url': settings.PORTAL_URL,
        }
        return render(request, 'notas/sem_permissao.html', contexto, status=403)


class ForceSSOMiddleware:
    """
    Exige sessão para as telas internas; quem não tem volta para o PORTAL.

    É esta a peça que garante o critério "nada mais manda o usuário para a tela
    do Keycloak". Antes, quem chegasse sem sessão batia no `@login_required` das
    views e era mandado para o `LOGIN_URL`, que era
    `/analise_de_margem/oidc/authenticate/` — ou seja, a tela do Keycloak, um
    segundo login depois de a pessoa já ter entrado no portal.

    O fluxo OIDC continua registrado em `/analise_de_margem/oidc/` para acesso
    direto ao sistema, fora do portal, mas nada mais leva ninguém para lá
    automaticamente.

    Deve vir DEPOIS do `KeycloakPortalAuthMiddleware`, que é quem transforma o
    token do portal em sessão.
    """

    #: `path_info` é o caminho SEM o prefixo da subrota (o ProxyPrefixMiddleware
    #: já o recortou). Comparar com `request.path`, que ainda contém
    #: `/analise_de_margem`, nunca casaria com '/oidc/' e bloquearia a própria
    #: tela de login e o callback do Keycloak.
    CAMINHOS_LIBERADOS = ('/oidc/', '/static/', '/media/', '/admin/')

    def __init__(self, get_response):
        self.get_response = get_response

    # exposto para quem quiser reaproveitar a mesma decisão
    _e_navegacao = staticmethod(_e_navegacao)

    def __call__(self, request):
        if not request.user.is_authenticated:
            caminho = request.path_info
            if not any(caminho.startswith(p) for p in self.CAMINHOS_LIBERADOS):
                return _volta_para_o_portal(request, 'sem sessao')

        return self.get_response(request)

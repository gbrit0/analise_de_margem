"""
Backends de autenticação da Análise de Margem.

Dois caminhos coexistem:

1. `KeycloakPortalBackend` — usado quando o usuário chega pelo portal
   unificado. O token JWT já emitido pelo Keycloak (cookie `brg_access_token`
   ou header `Authorization: Bearer`) é validado e o usuário entra direto, sem
   nova tela de login. Ver `notas/middleware`-equivalente em
   `setup.middleware.KeycloakPortalAuthMiddleware`.

2. `KeycloakOIDCBackend` — fluxo OIDC tradicional (`mozilla-django-oidc`),
   mantido para acesso direto ao sistema fora do portal.

Ambos derivam as permissões do Django das mesmas roles do Keycloak
(`notas.keycloak.aplicar_permissoes`).
"""

import logging
import urllib.parse

from django.conf import settings
from django.contrib.auth.backends import BaseBackend
from django.contrib.auth import get_user_model
from mozilla_django_oidc.auth import OIDCAuthenticationBackend

from notas import keycloak

logger = logging.getLogger(__name__)


class KeycloakPortalBackend(BaseBackend):
    """
    Autentica a partir das claims de um token do Keycloak já validado.

    Não herda de `OIDCAuthenticationBackend` de propósito: assim o middleware
    `mozilla_django_oidc.middleware.SessionRefresh` ignora as sessões abertas
    pelo portal e não tenta renová-las contra o Keycloak (o que provocaria um
    redirect indesejado para a tela de login).
    """

    def authenticate(self, request, claims=None, **kwargs):
        if not claims:
            return None
        try:
            return keycloak.obter_ou_criar_usuario(claims)
        except keycloak.SemPermissao as erro:
            logger.warning("Acesso negado via portal: %s", erro)
            return None
        except keycloak.TokenInvalido as erro:
            logger.warning("Claims inválidas na autenticação via portal: %s", erro)
            return None

    def get_user(self, user_id):
        UserModel = get_user_model()
        try:
            return UserModel.objects.get(pk=user_id)
        except UserModel.DoesNotExist:
            return None


class KeycloakOIDCBackend(OIDCAuthenticationBackend):
    """Fluxo OIDC clássico, para quem acessa o sistema sem passar pelo portal."""

    def get_userinfo(self, access_token, id_token, payload):
        """
        O endpoint `userinfo` do Keycloak não devolve `realm_access` /
        `resource_access` a menos que existam mappers específicos. As roles,
        porém, sempre viajam dentro do access token — que acabou de ser obtido
        por nós, via TLS, direto do token endpoint. Complementamos as claims
        com elas para que as permissões sejam aplicadas também neste fluxo.
        """
        userinfo = super().get_userinfo(access_token, id_token, payload)

        if 'resource_access' not in userinfo or 'realm_access' not in userinfo:
            try:
                import jwt

                claims_token = jwt.decode(
                    access_token,
                    options={'verify_signature': False, 'verify_aud': False},
                )
                userinfo.setdefault('realm_access', claims_token.get('realm_access', {}))
                userinfo.setdefault('resource_access', claims_token.get('resource_access', {}))
            except Exception as erro:
                logger.warning("Não foi possível ler as roles do access token: %s", erro)

        return userinfo

    def verify_claims(self, claims):
        """
        O padrão da biblioteca exige `email`; no Keycloak nem todo usuário tem.
        `preferred_username` (ou `sub`) já identifica a pessoa.
        """
        return bool(claims.get('email') or claims.get('preferred_username') or claims.get('sub'))

    def filter_users_by_claims(self, claims):
        email = claims.get('email')
        username = keycloak.username_do_token(claims)

        if username:
            users = self.UserModel.objects.filter(username__iexact=username)
            if users.exists():
                return users

        if email:
            users = self.UserModel.objects.filter(email__iexact=email)
            if users.exists():
                return users

        return self.UserModel.objects.none()

    def create_user(self, claims):
        """Executado no primeiro acesso do usuário via Keycloak."""
        email = claims.get('email', '')
        username = keycloak.username_do_token(claims) or 'user_oidc'

        user = self.UserModel.objects.create_user(username=username, email=email)
        return self.update_user(user, claims)

    def update_user(self, user, claims):
        """Sincroniza as roles do Keycloak com as permissões do Django a cada login."""
        return keycloak.aplicar_permissoes(user, claims)


def provider_logout(request):
    """
    Define para onde o usuário vai ao clicar em "Sair"
    (`OIDC_OP_LOGOUT_URL_METHOD`, chamado pela view `oidc_logout`).

    Sessão que veio do portal: encerra **apenas aqui** e devolve a pessoa à tela
    inicial do portal (`PORTAL_URL`).

    A escolha é deliberada. O caminho anterior mandava para o `PORTAL_LOGOUT_URL`,
    que apaga o cookie `brg_access_token` e derruba a sessão do portal INTEIRA —
    ou seja, sair da Análise de Margem deslogava a pessoa de todos os outros
    sistemas do hub de uma vez, o que ninguém espera ao fechar um sistema só.
    Num hub, "Sair" de um sistema é voltar ao hub.

    A contrapartida, que é preciso conhecer: como o cookie do portal continua de
    pé, clicar no card de novo reentra na hora, sem pedir nada. `logout()` limpa
    a sessão do Django deste app (inclusive o que estiver guardado nela) e o
    `KeycloakPortalAuthMiddleware` abre uma sessão nova no próximo GET. Quem
    quiser encerrar tudo usa o "Sair" do próprio portal.

    Para voltar ao logout global, basta devolver `settings.PORTAL_LOGOUT_URL`
    aqui — a chave continua configurada.
    """
    if request.session.get(settings.KEYCLOAK_PORTAL_SESSION_KEY):
        return settings.PORTAL_URL

    script_name = request.META.get('SCRIPT_NAME', '')
    redirect_uri = request.build_absolute_uri(script_name + '/')
    params = {
        'post_logout_redirect_uri': redirect_uri,
        'client_id': settings.OIDC_RP_CLIENT_ID,
    }
    return f"{settings.OIDC_OP_BASE_URL}/logout?{urllib.parse.urlencode(params)}"

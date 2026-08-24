"""
Leitura e validação do token JWT emitido pelo Keycloak e repassado pelo portal
unificado (Hub Central).

O portal, ao autenticar o usuário, grava o access token em um cookie
(`brg_access_token`) no domínio do portal. Como a Análise de Margem é servida
sob a mesma origem (`portal.brggeradores.com.br/analise_de_margem/`), o cookie
chega até aqui e pode ser usado para autenticar o usuário sem um segundo login.

Este módulo concentra:
  - extração do token (header `Authorization: Bearer` ou cookie do portal);
  - validação criptográfica contra o JWKS do Keycloak (assinatura, exp, aud, iss);
  - tradução das roles do Keycloak em permissões do Django.

As convenções de role seguem o padrão adotado pelo `auth_validador`:
`resource_access.<client_id>.roles` contendo `is_active`, `is_staff` e
`is_superuser`.
"""

import logging
import threading

import jwt
from jwt import PyJWKClient
from django.conf import settings

logger = logging.getLogger(__name__)

# Role mínima exigida pelo auth_validador para liberar o acesso ao sistema.
ROLE_ACESSO = 'is_active'

# Roles (do client) que concedem acesso ao Django admin / áreas restritas.
ROLES_STAFF = {'is_staff', 'is_superuser', 'superusuario', 'admin'}

# Roles (do client) que concedem superusuário dentro da Análise de Margem.
ROLES_SUPERUSER = {'is_superuser', 'superusuario', 'admin'}

# Roles de realm que identificam o administrador global do portal.
ROLES_ADMIN_GLOBAL = {'admin', 'global_admin', 'superusuario'}


class TokenInvalido(Exception):
    """Token ausente, malformado, expirado ou com assinatura inválida."""


class SemPermissao(Exception):
    """Token válido, porém o usuário não possui acesso a esta aplicação."""


_jwk_client = None
_jwk_lock = threading.Lock()


def _get_jwk_client():
    """Cliente JWKS com cache das chaves públicas (evita ida ao Keycloak a cada request)."""
    global _jwk_client
    if _jwk_client is None:
        with _jwk_lock:
            if _jwk_client is None:
                _jwk_client = PyJWKClient(
                    settings.OIDC_OP_JWKS_ENDPOINT,
                    cache_keys=True,
                    lifespan=getattr(settings, 'KEYCLOAK_JWKS_CACHE_SECONDS', 3600),
                    # O Keycloak está atrás do Cloudflare, que devolve 403 para
                    # o User-Agent padrão do urllib ("Python-urllib/3.x").
                    headers={'User-Agent': getattr(
                        settings, 'KEYCLOAK_JWKS_USER_AGENT', 'analise_de_margem/1.0')},
                    timeout=getattr(settings, 'KEYCLOAK_JWKS_TIMEOUT', 10),
                )
    return _jwk_client


def extrair_token(request):
    """
    Recupera o token na mesma ordem usada pelo auth_validador:
    header `Authorization: Bearer <token>` e, na ausência dele, o cookie do portal.
    """
    autorizacao = request.META.get('HTTP_AUTHORIZATION', '')
    if autorizacao.startswith('Bearer '):
        token = autorizacao[len('Bearer '):].strip()
        if token:
            return token

    cookie = getattr(settings, 'KEYCLOAK_PORTAL_COOKIE_NAME', 'brg_access_token')
    return request.COOKIES.get(cookie) or None


def validar_token(token):
    """
    Valida o token contra o Keycloak e devolve as claims.

    Levanta `TokenInvalido` em qualquer falha de assinatura, validade,
    audiência ou emissor.
    """
    client_id = settings.OIDC_RP_CLIENT_ID
    try:
        chave = _get_jwk_client().get_signing_key_from_jwt(token)
        return jwt.decode(
            token,
            chave.key,
            algorithms=[settings.OIDC_RP_SIGN_ALGO],
            audience=client_id,
            issuer=settings.KEYCLOAK_ISSUER,
            leeway=getattr(settings, 'KEYCLOAK_LEEWAY', 30),
            options={
                'verify_aud': getattr(settings, 'KEYCLOAK_VERIFY_AUDIENCE', True),
                'verify_iss': getattr(settings, 'KEYCLOAK_VERIFY_ISSUER', True),
            },
        )
    except jwt.ExpiredSignatureError:
        # Caminho MAIS comum em produção: o access token do realm vive 5 minutos
        # (exp - iat = 300s nos tokens reais). Quem trata isso é o middleware,
        # mandando a pessoa de volta ao portal — nunca para a tela do Keycloak.
        # O ping de /auth/refresh no base.html existe justamente para que este
        # caminho raramente seja alcançado com a tela aberta.
        logger.info('Token do portal expirado')
        raise TokenInvalido('Token expirado')
    except jwt.InvalidAudienceError:
        # Erro de configuração no Keycloak é o suspeito nº 1 aqui: falta o
        # audience mapper do client no token emitido para o portal.
        logger.warning(
            "Token rejeitado: audiência não contém '%s'. Audiências no token: %s",
            client_id, (claims_sem_validar(token) or {}).get('aud'),
        )
        raise TokenInvalido(f"Token sem audiência '{client_id}'")
    except jwt.PyJWTError as erro:
        logger.info("Token do portal rejeitado: %s", erro)
        raise TokenInvalido(str(erro))
    except Exception as erro:  # falha ao buscar o JWKS, rede fora, etc.
        logger.warning("Falha ao validar token do portal: %s", erro)
        raise TokenInvalido(str(erro))


def claims_sem_validar(token):
    """
    Claims SEM verificar a assinatura — só para log de diagnóstico e para ler o
    `jti`. Nunca use o retorno daqui para decidir acesso.
    """
    try:
        return jwt.decode(token, options={'verify_signature': False, 'verify_aud': False})
    except Exception:
        return None


def identificador_do_token(token):
    """
    `jti` (ou `sid`) do token, usado como chave de comparação com o token que
    abriu a sessão atual.

    Lido sem verificar assinatura de propósito: sozinho não concede nada, só
    serve para decidir se vale revalidar o JWT (e ir ao banco) a cada clique.
    Forjar um `jti` igual exigiria, antes, a posse do cookie de sessão do
    Django — que por si já autentica.
    """
    claims = claims_sem_validar(token)
    if not claims:
        return None
    return claims.get('jti') or claims.get('sid')


def roles_do_token(claims):
    """Devolve (roles do client desta aplicação, roles do realm)."""
    client_id = settings.OIDC_RP_CLIENT_ID
    roles_client = claims.get('resource_access', {}).get(client_id, {}).get('roles') or []
    roles_realm = claims.get('realm_access', {}).get('roles') or []
    return set(roles_client), set(roles_realm)


def tem_acesso(claims):
    """
    Acesso liberado quando o usuário possui a role `is_active` no client
    `analise_de_margem` ou quando é administrador global do portal.
    """
    roles_client, roles_realm = roles_do_token(claims)
    return ROLE_ACESSO in roles_client or bool(roles_realm & ROLES_ADMIN_GLOBAL)


def aplicar_permissoes(user, claims):
    """
    Sincroniza `is_staff`/`is_superuser`/`is_active` e os dados cadastrais do
    usuário a partir das roles do Keycloak. Chamado a cada nova autenticação,
    de modo que mudanças de papel no Keycloak refletem no Django.

    Sobre REBAIXAR: por padrão as roles do Keycloak só PROMOVEM
    (`KEYCLOAK_ESPELHA_PERMISSOES=False`). O motivo é concreto: a tela de
    administração das justificativas (`/justificativas/`) exige `is_superuser`,
    e há superusuários criados à mão (`USER_ADMIN`, `createsuperuser`) que ainda
    não têm `is_superuser` no client do Keycloak — espelhar de cara tiraria o
    acesso deles no primeiro clique, sem ninguém entender o porquê. Depois que
    as roles estiverem conferidas no Keycloak, ligue
    `KEYCLOAK_ESPELHA_PERMISSOES=True` no .env e o Keycloak passa a ser a única
    fonte de verdade (aí sim revogar lá revoga aqui).
    """
    roles_client, roles_realm = roles_do_token(claims)
    admin_global = bool(roles_realm & ROLES_ADMIN_GLOBAL)

    staff = bool(roles_client & ROLES_STAFF) or admin_global
    superusuario = bool(roles_client & ROLES_SUPERUSER) or admin_global

    if not getattr(settings, 'KEYCLOAK_ESPELHA_PERMISSOES', False):
        staff = staff or user.is_staff
        superusuario = superusuario or user.is_superuser

    user.is_active = True
    user.is_staff = staff
    user.is_superuser = superusuario

    email = claims.get('email') or ''
    nome = claims.get('given_name') or ''
    sobrenome = claims.get('family_name') or ''
    if email:
        user.email = email
    if nome:
        user.first_name = nome
    if sobrenome:
        user.last_name = sobrenome

    user.save()
    return user


def username_do_token(claims):
    """Identificador estável do usuário, na mesma ordem usada pelo backend OIDC."""
    username = claims.get('preferred_username') or claims.get('sub')
    if not username:
        email = claims.get('email') or ''
        username = email.split('@')[0] if email else None
    return username


def obter_ou_criar_usuario(claims):
    """
    Localiza o usuário pelo `preferred_username` (ou e-mail) e o cria caso ainda
    não exista, aplicando as permissões vindas do token.
    """
    from django.contrib.auth import get_user_model

    if not tem_acesso(claims):
        raise SemPermissao(
            f"Usuário sem a role '{ROLE_ACESSO}' no client '{settings.OIDC_RP_CLIENT_ID}'"
        )

    UserModel = get_user_model()
    username = username_do_token(claims)
    if not username:
        raise TokenInvalido('Token sem preferred_username/sub/email')

    email = claims.get('email') or ''

    user = UserModel.objects.filter(username__iexact=username).first()
    if user is None and email:
        user = UserModel.objects.filter(email__iexact=email).first()
    if user is None:
        user = UserModel.objects.create_user(username=username, email=email)

    return aplicar_permissoes(user, claims)

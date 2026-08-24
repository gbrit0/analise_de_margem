"""Contexto de template da integração com o portal unificado."""

from django.conf import settings


def portal(request):
    """
    Dados que o `base.html` usa para manter o cookie do portal vivo.

    `portal_renovacao_url` só sai preenchida quando a sessão nasceu no portal —
    é o único caso em que existe cookie a renovar. Para quem entrou pelo fluxo
    OIDC direto (ou pelo login local do admin) o script simplesmente não é
    incluído na página.
    """
    veio_do_portal = False
    if hasattr(request, 'session'):
        veio_do_portal = bool(request.session.get(settings.KEYCLOAK_PORTAL_SESSION_KEY))

    return {
        'portal_url': settings.PORTAL_URL,
        'portal_renovacao_url': settings.PORTAL_REFRESH_URL if veio_do_portal else '',
        'portal_renovacao_intervalo': settings.PORTAL_REFRESH_INTERVALO_SEGUNDOS,
    }

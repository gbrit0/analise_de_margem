import os
import io
import django
from django.conf import settings
from django.core.handlers.wsgi import WSGIRequest
from django.shortcuts import resolve_url
from django.urls import reverse, set_script_prefix

settings.configure(
    DEBUG=True,
    ROOT_URLCONF='setup.urls',
    ALLOWED_HOSTS=['*'],
    USE_X_FORWARDED_HOST=True,
    SECURE_PROXY_SSL_HEADER=('HTTP_X_FORWARDED_PROTO', 'https'),
    INSTALLED_APPS=[
        'django.contrib.admin',
        'django.contrib.auth',
        'django.contrib.contenttypes',
        'django.contrib.sessions',
        'django.contrib.messages',
        'django.contrib.staticfiles',
        'mozilla_django_oidc',
        'django_extensions',
        'django_filters',
        'users',
        'notas',
    ]
)
django.setup()

environ = {
    'REQUEST_METHOD': 'GET',
    'HTTP_HOST': 'backend:8000',
    'HTTP_X_FORWARDED_HOST': 'margem.brggeradores.com.br',
    'HTTP_X_FORWARDED_PROTO': 'https',
    'SERVER_NAME': 'testserver',
    'SERVER_PORT': '80',
    'PATH_INFO': '/analise_de_margem/',
    'HTTP_X_FORWARDED_PREFIX': '/analise_de_margem',
    'wsgi.url_scheme': 'http',
    'wsgi.input': io.BytesIO(b""),
}
request = WSGIRequest(environ)

from setup.middleware import ProxyPrefixMiddleware
def get_response(request): return request
middleware = ProxyPrefixMiddleware(get_response)
request = middleware(request)

print("reverse:", reverse('oidc_authentication_callback'))
print("build_absolute_uri:", request.build_absolute_uri(reverse('oidc_authentication_callback')))

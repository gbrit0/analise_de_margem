"""
Settings para execução da suíte de testes.

O usuário do MySQL de produção não tem permissão para criar a base
`test_analise_margem`, então os testes rodam sobre SQLite em memória.

    python manage.py test --settings=setup.settings_test
"""

from setup.settings import *  # noqa: F401,F403

DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': ':memory:',
    },
    'protheus': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': ':memory:',
    },
}

CACHES = {
    'default': {
        'BACKEND': 'django.core.cache.backends.locmem.LocMemCache',
    }
}

PASSWORD_HASHERS = ['django.contrib.auth.hashers.MD5PasswordHasher']

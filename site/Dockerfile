FROM python:3.10-slim

# 1. Configura as variáveis de ambiente (Locale e SQL Server)
ENV LANG=pt_BR.UTF-8 \
    LANGUAGE=pt_BR:pt \
    LC_ALL=pt_BR.UTF-8 \
    PATH="$PATH:/opt/mssql-tools/bin" \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Trecho RUN para instalar as dependências e os drivers da Microsoft
RUN apt-get update && \
    apt-get install -y \
        curl \
        apt-transport-https \
        gnupg2 \
        default-libmysqlclient-dev \
        build-essential \
        pkg-config \
        locales \
    && sed -i -e 's/# pt_BR.UTF-8 UTF-8/pt_BR.UTF-8 UTF-8/' /etc/locale.gen \
    && locale-gen \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && export DEBIAN_VERSION=$(grep VERSION_ID /etc/os-release | cut -d '"' -f 2 | cut -d '.' -f 1) \
    && curl -sSL -O https://packages.microsoft.com/config/debian/$DEBIAN_VERSION/packages-microsoft-prod.deb \
    && dpkg -i packages-microsoft-prod.deb \
    && rm packages-microsoft-prod.deb \
    && apt-get update \
    && ACCEPT_EULA=Y apt-get install -y \
        msodbcsql17="17.11.*" \
        mssql-tools \
        unixodbc-dev \
        libgssapi-krb5-2 \
    # && sed -i 's/MinProtocol = TLSv1.2/MinProtocol = TLSv1.0/g' /etc/ssl/openssl.cnf || true \
    # && sed -i 's/CipherString = DEFAULT@SECLEVEL=2/CipherString = DEFAULT@SECLEVEL=1/g' /etc/ssl/openssl.cnf || true \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*
    

WORKDIR /analise_de_margem

COPY requirements.txt /analise_de_margem/    

RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r /analise_de_margem/requirements.txt

COPY . /analise_de_margem

RUN mkdir -p /analise_de_margem/static

# # Cria um usuário não-root e muda a propriedade dos arquivos
# RUN useradd --no-create-home django-user \
#     && chown -R django-user:django-user /analise_de_margem/

# # Alterna para o usuário não-root
# USER django-user

# RUN python manage.py migrate --noinput
# RUN python manage.py collectstatic --noinput

RUN chmod +x /analise_de_margem/entrypoint.sh

# 9. Expõe a porta que o Gunicorn vai usar
EXPOSE 30117

ENTRYPOINT ["/analise_de_margem/entrypoint.sh"]

# 10. Comando para rodar a aplicação usando Gunicorn
#
# --limit-request-field_size: o padrão do gunicorn (8190 bytes por campo de
#   header) é pequeno demais para o domínio do portal, onde o header Cookie
#   acumula os cookies de todos os sistemas + o brg_access_token (JWT) + a
#   sessão do Flask, chegando a ~11 KB. Estourar esse limite faz o gunicorn
#   responder "Request Header Fields Too Large" antes do Django ser chamado.
#   O nginx continua sendo o limite efetivo (large_client_header_buffers 16k).
CMD ["gunicorn", "--bind", "0.0.0.0:30117", "--workers", "3", "--timeout", "120", \
     "--limit-request-field_size", "32768", "--limit-request-fields", "200", \
     "setup.wsgi:application"]
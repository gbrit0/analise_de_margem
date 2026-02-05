# Site Análise de Margem

O objetivo deste site é realizar uma implementação online da tabela de análise de margem que centralize as análises, permita a edição de custos, gere estatísticas, registre justificativas e seja rastreável.

# Instalação

Para instalar o projeto em um novo servidor siga os seguintes passos:

1 - Clone o repositório git: 

```bash
git clone https://github.com/gbrit0/analise_de_margem.git
```

2 - Acesse o repositório clonado, crie o ambiente virtual e instale as dependências:

```bash
cd analise_de_margem/site
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

3 - Copie o arquivo [.env.example](/.env.example) e altere os valores de exemplo para valores reais:

```bash
cp .env.example .env
```

4 - Realize as migrações do django:

```bash
python manage.py makemigrations
python manage.py migrate 
```

5 - Crie o superusuário:

```bash
python manage.py createsuperuser
# Preencha os dados solicitados
```

6 - Inicialize o projeto:

```bash
python manage.py runserver
```

7 - Crie o usuário do systema

Com a aplicação rodando acesse o /admin, logue com o superuser criado anteriormente e crie um novo usuário system. Esse usuário será destinado para execuçoes padrão, como um bot, dentro do site e seu id deve ser id=2.

8 - Após a cração do usuário o sistema encontra-se pronto para ter dados carregados:

```bash
python manage.py sync_protheus
```

9 - Rodar novamente a aplicação

Após realizar a carga inicial dos dados, rodar novamente o servidor.

```bash
python manage.py runserver
```

Nesse ponto a aplicação estará funcional apesar de ser altamente recomendado subir a aplicação por meio de um wsgi como uvicorn ou gunicorn.

10 - Agendamento da sincronização do banco ao Protheus

A tabela de Notas desta aplicação é alimentada por query direta ao banco do Protheus de modo que é necessário adicionar uma regra de crontab que periodicamente chame o script de sincronização:

```bash
crontab -e
```

E em seguida adicione o seguinte texto ao final do arquivo:

```bash
*/15 * * * 1-5 /path/to/venv/bin/python manage.py sync_protheus # Altere /path/to/venv para o caminho real da instalação do projeto
```

from django.core.management.base import BaseCommand
from django.conf import settings
from users.models import CustomUser as User

class Command(BaseCommand):
    help='Cria o usuário padrão do sistema'
    
    def handle(self, *args, **options):
        self.stdout.write("Criando o usuário padrão")
        
        try:
            User.objects.create(
                username=settings.USER_SYSTEM,
                password=settings.PASS_SYSTEM
            )
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Erro ao criar usuário: {e}'))
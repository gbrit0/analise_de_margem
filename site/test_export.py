import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'setup.settings')
django.setup()

from notas.views import exportar_op_excel
from django.test import RequestFactory
from users.models import CustomUser

user = CustomUser.objects.first()
request = RequestFactory().get('/ops/0000000001/exportar/')
request.user = user

try:
    response = exportar_op_excel(request, '0000000001')
    print("SUCCESS", response.status_code)
except Exception as e:
    import traceback
    traceback.print_exc()

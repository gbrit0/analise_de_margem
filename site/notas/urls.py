from django.urls import path
from notas.views import NotasListView, atualizar_custo_api, atualizar_justificativa_api
from .models import Nota

urlpatterns = [
    path('', NotasListView.as_view(model=Nota), name='lista_notas'),
    path('api/atualizar-custo/', atualizar_custo_api, name='api_atualizar_custo'),
    path('api/atualizar-justificativa/', atualizar_justificativa_api, name='api_atualizar_justificativa'),
]
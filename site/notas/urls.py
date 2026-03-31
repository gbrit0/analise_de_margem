from django.urls import path

from notas.views import (
    NotasListView, 
    atualizar_custo_api, 
    atualizar_justificativa_api, 
    dashboard_view,
    dados_vendas_api,
    op_list_view,
    justificativa_admin_view,
    justificativa_save,
    justificativa_toggle_status
)

from django.contrib.auth.decorators import login_required
from .models import Nota, OP, Justificativa

urlpatterns = [
    path('notas/', login_required(NotasListView.as_view(model=Nota)), name='lista_notas'),
    path('api/atualizar-custo/', atualizar_custo_api, name='api_atualizar_custo'),
    path('api/atualizar-justificativa/', atualizar_justificativa_api, name='api_atualizar_justificativa'),
    path('estatisticas/', dashboard_view, name='estatisticas'),
    path('api/estatisticas/', dados_vendas_api, name='dados_vendas_api'),
    path('ops/<str:lote>/', op_list_view, name='lista_ops'),
    path('justificativas/', justificativa_admin_view, name='admin_justificativas'),
    path('api/justificativa/salvar/', justificativa_save, name='api_justificativa_salvar'),
    path('api/justificativa/toggle/', justificativa_toggle_status, name='api_justificativa_toggle'),
]
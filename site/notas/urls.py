from django.urls import path

from notas.views import (
    NotasListView, 
    atualizar_custo_api, 
    atualizar_justificativa_api, 
    # estatisticas, 
    dashboard_view,
    dados_vendas_api,
    # OPListView, 
    op_list_view, JustificativaListView
)

from django.contrib.auth.decorators import login_required
from .models import Nota, OP, Justificativa

urlpatterns = [
    path('notas/', login_required(NotasListView.as_view(model=Nota)), name='lista_notas'),
    path('api/atualizar-custo/', atualizar_custo_api, name='api_atualizar_custo'),
    path('api/atualizar-justificativa/', atualizar_justificativa_api, name='api_atualizar_justificativa'),
    path('estatisticas/', dashboard_view, name='estatisticas'),
    path('api/estatisticas/', dados_vendas_api, name='dados_vendas_api'),
    # path('ops/<str:lote>/', login_required(OPListView.as_view(model=OP)), name='lista_ops'),
    path('ops/<str:lote>/', op_list_view, name='lista_ops'),
    path('justificativas/', login_required(JustificativaListView.as_view(model=Justificativa)), name='lista_justificativas'),
]
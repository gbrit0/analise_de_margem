from django.urls import path

from notas.views import (
    NotasListView, 
    atualizar_custo_api, 
    atualizar_justificativa_api, 
    atualizar_comentario_api,
    dashboard_view,
    dados_vendas_api,
    op_list_view,
    justificativa_admin_view,
    justificativa_save,
    justificativa_toggle_status,
    salvar_preferencia_colunas_nota,
    toggle_bloqueio_mes_api,
    exportar_excel,
    exportar_estatisticas_excel,
    exportar_op_excel,
    atualizar_custo2_op_api,
    # atualizar_custo_op_api
)

from django.contrib.auth.decorators import login_required
from .models import Nota, OP, Justificativa

urlpatterns = [
    path('notas/', login_required(NotasListView.as_view(model=Nota)), name='lista_notas'),
    path('api/atualizar-custo/', atualizar_custo_api, name='api_atualizar_custo'),
    path('api/atualizar-custo2-op/', atualizar_custo2_op_api, name='api_atualizar_custo2_op'),
    # path('api/atualizar-custo-op/', atualizar_custo_op_api, name='api_atualizar_custo_op'),
    path('api/atualizar-justificativa/', atualizar_justificativa_api, name='api_atualizar_justificativa'),
    path('api/atualizar-comentario/', atualizar_comentario_api, name='api_atualizar_comentario'),
    path('estatisticas/', dashboard_view, name='estatisticas'),
    path('api/estatisticas/', dados_vendas_api, name='dados_vendas_api'),
    path('ops/<str:lote>/exportar/', exportar_op_excel, name='exportar_op_excel'),
    path('ops/<str:lote>/<str:cod_produto>/', op_list_view, name='lista_ops'),
    path('justificativas/', justificativa_admin_view, name='admin_justificativas'),
    path('api/justificativa/salvar/', justificativa_save, name='api_justificativa_salvar'),
    path('api/justificativa/toggle/', justificativa_toggle_status, name='api_justificativa_toggle'),
    path('api/notas/colunas/salvar/', salvar_preferencia_colunas_nota, name='api_salvar_preferencia_colunas_nota'),
    path('api/mes-bloqueado/toggle/', toggle_bloqueio_mes_api, name='api_mes_bloqueado_toggle'),
    path('exportar/', exportar_excel, name='exportar_excel'),
    path('exportar-estatisticas/', exportar_estatisticas_excel, name='exportar_estatisticas_excel'),
]

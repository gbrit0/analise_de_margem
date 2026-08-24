from django.urls import path
from django.contrib.auth.decorators import login_required

from notas.views import (
    NotasListView, 
    atualizar_custo_api, 
    atualizar_justificativa_api, 
    atualizar_comentario_api,
    dashboard_view,
    dados_vendas_api,
    vendedor_notas_api,
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

from .models import Nota, OP, Justificativa

urlpatterns = [
    path('notas/', login_required(NotasListView.as_view(model=Nota)), name='lista_notas'),
    path('api/atualizar-custo/', login_required(atualizar_custo_api), name='api_atualizar_custo'),
    path('api/atualizar-custo2-op/', login_required(atualizar_custo2_op_api), name='api_atualizar_custo2_op'),
    path('api/atualizar-justificativa/', login_required(atualizar_justificativa_api), name='api_atualizar_justificativa'),
    path('api/atualizar-comentario/', login_required(atualizar_comentario_api), name='api_atualizar_comentario'),
    path('estatisticas/', login_required(dashboard_view), name='estatisticas'),
    path('api/estatisticas/', login_required(dados_vendas_api), name='dados_vendas_api'),
    path('api/estatisticas/vendedor/notas/', login_required(vendedor_notas_api), name='vendedor_notas_api'),
    path('ops/<str:lote>/exportar/', login_required(exportar_op_excel), name='exportar_op_excel'),
    path('ops/<str:lote>/<str:cod_produto>/', login_required(op_list_view), name='lista_ops'),
    path('justificativas/', login_required(justificativa_admin_view), name='admin_justificativas'),
    path('api/justificativa/salvar/', login_required(justificativa_save), name='api_justificativa_salvar'),
    path('api/justificativa/toggle/', login_required(justificativa_toggle_status), name='api_justificativa_toggle'),
    path('api/notas/colunas/salvar/', login_required(salvar_preferencia_colunas_nota), name='api_salvar_preferencia_colunas_nota'),
    path('api/mes-bloqueado/toggle/', login_required(toggle_bloqueio_mes_api), name='api_mes_bloqueado_toggle'),
    path('exportar/', login_required(exportar_excel), name='exportar_excel'),
    path('exportar-estatisticas/', login_required(exportar_estatisticas_excel), name='exportar_estatisticas_excel'),
]

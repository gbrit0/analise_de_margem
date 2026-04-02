from setup import settings
from .filters import NotaFilter
from .models import (
    Justificativa, 
    Nota, 
    Custo, 
    Margem, 
    Nf_Has_Justificativa, 
    OP
)

from users.models import CustomUser

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

import os
import json
import locale
import pyodbc
import pymysql
import datetime
from decimal import Decimal
from dbutils.pooled_db import PooledDB
from dateutil.relativedelta import relativedelta

from django.utils import timezone
from django.shortcuts import render
from django.views.generic import ListView
from django_filters.views import FilterView
from django.utils.dateparse import parse_date
from django.db.models import Sum, Count, Q, Avg
from django.db.models.functions import TruncMonth
from django.http import HttpResponse, JsonResponse
from django.views.decorators.cache import cache_page
from django.shortcuts import render, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib.auth.mixins import LoginRequiredMixin
from django.db.models.functions import Coalesce, ExtractYear, ExtractMonth
from django.views.decorators.http import require_http_methods, require_POST
from django.db.models import OuterRef, Subquery, DecimalField, Case, When, F, Value, ExpressionWrapper

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

pool = PooledDB(
    creator=pyodbc,
    maxconnections=10, # Como vamos usar lotes, não precisamos de tantas conexões simultâneas
    mincached=2,
    blocking=True,
    driver='{ODBC Driver 17 for SQL Server}',
    server=f'{os.getenv("PROTHEUS_DB_HOST")}',
    database=f'{os.getenv("PROTHEUS_DB_DATABASE")}',
    uid=f'{os.getenv("PROTHEUS_DB_USER")}',
    pwd=f'{os.getenv("PROTHEUS_DB_PASSWORD")}'
)

pool_mysql = PooledDB(
    creator=pymysql,
    maxconnections=10, # Como vamos usar lotes, não precisamos de tantas conexões simultâneas
    mincached=2,
    blocking=True,
    host=f'{os.getenv("DB_HOST")}',
    port=int(os.getenv("DB_PORT")),
    user=f'{os.getenv("DB_USER")}',
    password=f'{os.getenv("DB_PASSWORD")}',
    database=f'{os.getenv("DB_NAME")}'
)

class NotasListView(LoginRequiredMixin, FilterView, ListView):
    mes_anterior = datetime.date.today() - relativedelta(month=1)
    mes_anterior = mes_anterior.strftime("%Y-%m")
    model = Nota
    # paginate_by = 25
    filterset_class = NotaFilter
    ordering = ['-data_emissao']
    login_url = 'login/' 
    context_object_name = 'nota_list'
    template_name = 'notas/nota_list.html'
    redirect_field_name = f'/?data_emissao_month={mes_anterior}' # Filtrar as notas do mes anterior depois do login
        
    def get_queryset(self):
        cfops_especiais = ['5101', '6101', '5116', '6116', '6107']
        # 1. Definimos a subquery para buscar o Custo relacionado
        # Filtramos onde a chave do Custo é igual à chave da Nota atual (OuterRef('pk'))
        custo_mais_recente = Custo.objects.filter(
            chave=OuterRef('pk')
        ).order_by('-data_cadastro').values('valor')[:1]
        
        ultimo_custo = Custo.objects.filter(
            chave=OuterRef('pk')
        ).order_by('data_cadastro').values('valor')[:1]
        
        justificativa_mais_recente = Nf_Has_Justificativa.objects.filter(
            nf=OuterRef('pk')
        ).order_by('-data_cadastro').values('justificativa')[:1]
        
        # 2. Anotamos o queryset principal da Nota com esse valor
        queryset = Nota.objects.annotate(
            custo_mais_recente=Coalesce(
                Subquery(custo_mais_recente), 
                0, 
                output_field=DecimalField()
            ),
            ultimo_custo_valor=Coalesce(
                Subquery(ultimo_custo), 
                0, 
                output_field=DecimalField()
            ),
            icms_calculado=Case(
                When(cfop__in=cfops_especiais, then=F('valor_icms') * 0.047), # 4,7% do ICMS
                default=F('valor_icms'),                                     # ICMS cheio
                output_field=DecimalField()
            ),
            margem_bruta=ExpressionWrapper(
                F('valor_contabil') 
                - F('custo_mais_recente') 
                - F('valor_ipi') 
                - F('valor_imp5') 
                - F('valor_imp6') 
                - F('valor_icms_difal') 
                - F('icms_calculado'),
                output_field=DecimalField(max_digits=18, decimal_places=2)
            ),
            margem_percentual=ExpressionWrapper(
                F('margem_bruta') 
                / F('valor_contabil') * 100,
                output_field=DecimalField(max_digits=4, decimal_places=2)
            ),
            justificativa=justificativa_mais_recente
        )
        
        filiais_selecionadas = self.request.GET.getlist('filial')
        if filiais_selecionadas:
            queryset = queryset.filter(filial__in=filiais_selecionadas)
            
        return queryset.order_by('-data_emissao')
    
    def get_context_data(self, **kwargs):
        # Chama a implementação base primeiro para pegar o contexto padrão (incluindo o queryset)
        context = super().get_context_data(**kwargs)
        
        # Adicionamos a lista de justificativas ativas para popular os selects no HTML
        context['lista_justificativas'] = Justificativa.objects.filter(ativo=True)
        
        # Monta lista de meses/anos disponíveis (format YYYY-MM)
        meses_qs = Nota.objects.annotate(
            year=ExtractYear('data_emissao'),
            month=ExtractMonth('data_emissao')
        ).values('year', 'month').distinct().order_by('-year', '-month')

        meses = []
        for m in meses_qs:
            y = int(m['year'])
            mo = int(m['month'])
            meses.append({
                'value': f"{y}-{mo:02d}",
                'label': f"{mo:02d}/{y}"
            })

        context['meses_disponiveis'] = meses
        # Valor selecionado atualmente (vindo dos GET params)
        selected = self.request.GET.get('data_emissao_month', '')

        filiais_selecionadas = self.request.GET.getlist('filial')
        context['filiais_selecionadas'] = filiais_selecionadas
        
        # Se não veio valor no GET, tenta preencher com mês atual (se houver dados),
        # caso contrário preenche com o mês mais recente disponível.
        if not selected:
            hoje = datetime.datetime.today()
            atual = f"{hoje.year}-{hoje.month:02d}"
            # se atual estiver na lista de meses disponiveis, usa; senão usa o primeiro disponível
            valores = [m['value'] for m in meses]
            if atual in valores:
                selected = atual
            elif len(valores) > 0:
                selected = valores[0]

        context['selected_month'] = selected
        
        # context['mes_anterior'] = (datetime.date.today() - relativedelta(month=1)).strftime("%Y-%m")
        filiais = Nota.objects.values('filial', 'nome_filial').distinct()
        
        context['filiais'] = filiais
        
        with pool_mysql.connection() as con:
            with con.cursor() as cursor:
                query = """SELECT cod_cliente, loja_cliente FROM analise_margem.cliente_parceiro;"""
                cursor.execute(query)
                clientes_parceiros_raw = cursor.fetchall()
        # 4459286002
        # 1. Cria um SET de tuplas para busca extremamente rápida.
        # Assumindo que o fetchall() retorna tuplas onde [0] é cod_cliente e [1] é loja.
        # (Se seu cursor retornar dicionários, use: c['cod_cliente'], c['loja'])
        set_parceiros = {(c[0], c[1]) for c in clientes_parceiros_raw}

        # 2. Verifica a lista de notas atual (object_list do ListView)
        # Se você usa um context_object_name diferente, troque 'object_list' pelo seu nome
        for nota in context['object_list']:
            # Cria um atributo dinâmico "is_parceiro" na nota
            # Se a tupla (cod_cliente, loja) da nota existir no SET, recebe True.
            if (nota.cod_cliente, nota.loja) in set_parceiros:
                nota.is_parceiro = True
                margem = Margem.objects.filter(chave=nota).last()
                if margem.margem_bruta_percentual > 0.15 and margem.margem_bruta_percentual < 0.27:
                    justificativa_nf = Nf_Has_Justificativa.objects.filter(nf=nota).last()
                    if not justificativa_nf:
                        Nf_Has_Justificativa.objects.create(
                            nf=nota,
                            justificativa=Justificativa.objects.get(texto='OK. Margem Parceiro.'),
                            usuario=CustomUser.objects.get(id=2)
                        )
            else:
                nota.is_parceiro = False

        return context
    
@login_required
@require_POST
def atualizar_custo_api(request):
    try:
        data = json.loads(request.body)
        nota_chave = data.get('chave')
        novo_valor_custo = data.get('valor')
        
        # 1. Validação básica
        if not nota_chave or novo_valor_custo is None:
            return JsonResponse({'error': 'Dados inválidos'}, status=400)

        valor_custo_decimal = Decimal(str(novo_valor_custo).replace(',', '.'))
        
        # 2. Busca a Nota
        nota = get_object_or_404(Nota, pk=nota_chave)

        # 3. Cria um NOVO registro de custo (preservando histórico)
        # Assumindo que você tem o request.user disponível. Se não, defina um usuário padrão ou trate isso.
        custo = Custo.objects.create(
            chave=nota,
            valor=valor_custo_decimal,
            usuario=request.user if request.user else 2  # usuário system
        )
               
        # Definição do ICMS baseada no CFOP
        cfops_especiais = ['5101', '6101', '5116', '6116', '6107']
        if nota.cfop in cfops_especiais:
            val_icms = nota.valor_icms * Decimal('0.047')
        else:
            val_icms = nota.valor_icms

        # Cálculo: Valor Contábil - Custo Novo - Impostos
        margem_bruta = (
            nota.valor_contabil 
            - valor_custo_decimal 
            - nota.valor_ipi 
            - nota.valor_imp5 
            - nota.valor_imp6 
            - nota.valor_icms_difal 
            - val_icms
        )

        # Cálculo da Margem Percentual (evitando divisão por zero)
        margem_percentual = 0
        if nota.valor_contabil > 0:
            margem_percentual = (margem_bruta / nota.valor_contabil)

        margem = Margem.objects.create(
            chave=nota,
            custo=custo,
            margem_bruta=margem_bruta,
            margem_bruta_percentual=margem_percentual
        )
        
        # 5. Retorna os dados formatados para o Frontend
        return JsonResponse({
            'success': True,
            'margem_bruta': f'{locale.currency(margem_bruta, grouping=True)}',
            'margem_percentual': f'{margem_percentual*100:.2f}%'
        })

    except Exception as e:
        raise e
        # return JsonResponse({'error': str(e)}, status=500)
    
@login_required
@require_POST
def atualizar_justificativa_api(request):
    try:
        data = json.loads(request.body)
        nota_chave = data.get('chave')
        nova_justificativa_id = data.get('justificativa')

        nota = get_object_or_404(Nota, pk=nota_chave)

        # 1. Buscamos a justificativa que o usuário clicou
        justificativa = get_object_or_404(Justificativa, id=nova_justificativa_id)

        # 2. VERIFICAÇÃO: Se o texto for 'Limpar / Sem Justificativa', 
        # nós apenas deletamos e retornamos.
        if 'Limpar' in justificativa.texto:
            Nf_Has_Justificativa.objects.filter(nf=nota).delete()
            return JsonResponse({'success': True, 'action': 'cleared'})

        # 3. Se NÃO for limpar, primeiro limpamos o que tinha antes (para não duplicar)
        # e depois criamos a nova relação
        Nf_Has_Justificativa.objects.filter(nf=nota).delete()
        
        Nf_Has_Justificativa.objects.create(
            usuario=request.user,
            nf=nota,
            justificativa=justificativa
        )

        return JsonResponse({'success': True, 'justificativa_id': justificativa.id})

    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@login_required
def dashboard_view(request):
    selected = request.GET.get('data_emissao_month', '')
    
    meses_qs = Nota.objects.annotate(
        year=ExtractYear('data_emissao'),
        month=ExtractMonth('data_emissao')
    ).values('year', 'month').distinct().order_by('-year', '-month')

    meses = []
    for m in meses_qs:
        y = int(m['year'])
        mo = int(m['month'])
        meses.append({
            'value': f"{y}-{mo:02d}",
            'label': f"{mo:02d}/{y}"
        })

    # Se não veio valor no GET, tenta preencher com mês atual (se houver dados),
    # caso contrário preenche com o mês mais recente disponível.
    if not selected:
        hoje = datetime.datetime.today()
        atual = f"{hoje.year}-{hoje.month:02d}"
        valores = [m['value'] for m in meses]
        if atual in valores:
            selected = atual
        elif len(valores) > 0:
            selected = valores[0]
        
    filiais = Nota.objects.values('filial', 'nome_filial').distinct()
    
    context = {
        'filiais': filiais,
        'meses_disponiveis': meses,
        'selected_month': selected
    }
    
    return render(request, 'notas/estatisticas.html', context)

@login_required
def dados_vendas_api(request):
    
    meses_str = request.GET.get('meses')
    filiais_str = request.GET.get('filiais')
        
    queryset = Nota.objects.all()
    
    if meses_str:
        meses_list = meses_str.split(',')
        queries = Q()
        for mes in meses_list:
            if '-' in mes:
                try:
                    y, m = mes.split('-')
                    queries |= Q(data_emissao__year=int(y), data_emissao__month=int(m))
                except ValueError:
                    pass
        if queries:
            queryset = queryset.filter(queries)
            
    if filiais_str:
        filiais_list = filiais_str.split(',')
        queryset = queryset.filter(filial__in=filiais_list)
    
    subquery_margem_percentual = Margem.objects.filter(
        chave=OuterRef('chave')
    ).values('margem_bruta_percentual').order_by('-custo__data_cadastro')[:1] # O Gemini disse: Para ordenar pelo atributo cadastro da tabela Custo (que possui a Foreign Key para Margem), você precisa utilizar a sintaxe de "follow relationship" do Django, que utiliza o duplo sublinhado (__).
    
    subquery_custo = Custo.objects.filter(
            chave=OuterRef('pk')
        ).order_by('-data_cadastro').values('valor')[:1]
    
    subquery_margem = Margem.objects.filter(
        chave=OuterRef('chave')
    ).values('margem_bruta').order_by('-custo__data_cadastro')[:1]

    subquery_justificaticas = Nf_Has_Justificativa.objects.filter(
        nf=OuterRef('pk')
    ).values('justificativa')[:1]
    
    subquery_justificativas_texto = Nf_Has_Justificativa.objects.filter(
        nf=OuterRef('pk')
    ).order_by('-data_cadastro').values('justificativa__texto')[:1]
    
    with pool_mysql.connection() as con:
        with con.cursor() as cursor:
            query = """SELECT cod_cliente, loja_cliente FROM analise_margem.cliente_parceiro;"""
            cursor.execute(query)
            clientes_parceiros_raw = cursor.fetchall()
            
    set_parceiros = {(c[0], c[1]) for c in clientes_parceiros_raw}

    notas_stats = queryset.annotate(
        margem_pct=Subquery(subquery_margem_percentual),
        just_texto=Subquery(subquery_justificativas_texto)
    ).values('cod_cliente', 'loja', 'valor_contabil', 'margem_pct', 'just_texto')
    
    total_vendas_periodo = Decimal('0.0')
    stats_just = {}
    total_notas_abaixo_margem_global = 0

    for nota in notas_stats:
        val_contabil = nota['valor_contabil'] or Decimal('0.0')
        total_vendas_periodo += val_contabil
        
        is_parceiro = (nota['cod_cliente'], nota['loja']) in set_parceiros
        limite_margem = 0.15 if is_parceiro else 0.27
        margem_atual = nota['margem_pct']
        
        abaixo_margem = False
        if margem_atual is not None and float(margem_atual) < limite_margem:
            abaixo_margem = True
            total_notas_abaixo_margem_global += 1
            
        j_texto = nota['just_texto']
        if j_texto:
            if j_texto not in stats_just:
                stats_just[j_texto] = {
                    'contagem': 0,
                    'vendas_total': Decimal('0.0'),
                    'abaixo_margem': 0
                }
            stats_just[j_texto]['contagem'] += 1
            stats_just[j_texto]['vendas_total'] += val_contabil
            if abaixo_margem:
                stats_just[j_texto]['abaixo_margem'] += 1

    estatisticas_justificativas = []
    total_vendas_float = float(total_vendas_periodo)
    
    for j_texto, data in stats_just.items():
        repres_vendas = (float(data['vendas_total']) / total_vendas_float * 100) if total_vendas_float > 0 else 0
        perc_abaixo = (data['abaixo_margem'] / total_notas_abaixo_margem_global * 100) if total_notas_abaixo_margem_global > 0 else 0
        
        estatisticas_justificativas.append({
            'justificativa': j_texto,
            'contagem': data['contagem'],
            'representatividade_vendas': round(repres_vendas, 2),
            'percentual_abaixo_margem': round(perc_abaixo, 2)
        })
        
    estatisticas_justificativas.sort(key=lambda x: x['contagem'], reverse=True)

    queryset = queryset.annotate(
        mes=TruncMonth('data_emissao')
    ).values('mes').annotate(
        total_vendas=Sum('valor_contabil'),
        margem_percentual=Avg(subquery_margem_percentual)*100,
        margem_total=Sum(subquery_margem),
        custo_total=Sum(subquery_custo)
    ).order_by('mes')
    
    labels = [item['mes'].strftime('%b/%Y') for item in queryset]
    
    total_vendas = [float(item['total_vendas']) for item in queryset]
    margem_percentual = [float(item['margem_percentual']) for item in queryset]
    margem_total = [float(item['margem_total']) for item in queryset]
    custo_total = [float(item['custo_total']) for item in queryset] 
        
    return JsonResponse({
        # 'margem': queryset['margem_total'],
        'labels': labels,
        'total_vendas': total_vendas,
        'margem_percentual': margem_percentual,
        'margem_total': margem_total,
        'custo_total': custo_total,
        'estatisticas_justificativas': estatisticas_justificativas
    })
    

# class OPListView(LoginRequiredMixin, FilterView, ListView):
#     model = OP
#        = ['-ord_producao', '-sequencial']
#     context_object_name = 'op_list'
#     template_name = 'notas/op_list.html'
#     # filterset_class = SeuFiltroDeOPs # Descomente se for usar o FilterView real

#     def get_queryset(self):
#         # 1. Recupera o queryset original com a ordenação definida na classe
#         qs = super().get_queryset()
        
#         # 2. Pega a 'chave' passada na URL
#         lote = self.kwargs.get('lote')
        
#         # 3. Busca a Nota e monta o prefixo
#         # nota = get_object_or_404(Nota, chave=chave)
#         # prefixo_busca = f"{nota.filial}{nota.nota}{nota.item}{nota.recno}"
        
#         # 4. Retorna o queryset filtrado
#         return qs.filter(id_op__startswith=lote)

#     def get_context_data(self, **kwargs):
#         # Adiciona o objeto 'nota' ao contexto para você usar no seu HTML (ex: {{ nota.lote }})
#         context = super().get_context_data(**kwargs)
#         chave = self.kwargs.get('chave')
#         context['nota'] = get_object_or_404(Nota, chave=chave)
#         return context

@login_required
# @cache_page(60 * 15)
def op_list_view(request, lote):
    with open(f'{settings.BASE_DIR}/notas/management/commands/querys/buscaOPs.sql', 'r') as f:
        query_ops = f.read()
        
    with pool.connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(query_ops, lote)
            colunas = [coluna[0] for coluna in cursor.description]
            linhas_op = [dict(zip(colunas, linha)) for linha in cursor.fetchall()]
    
    for op in linhas_op:
        if op.get('custo') is not None:
            op['custo'] = locale.currency(op['custo'], grouping=True)
        if op.get('custo_2') is not None:
            op['custo_2'] = locale.currency(op['custo_2'], grouping=True)

        if op.get('c_contabil') is None:
            op['c_contabil'] = '-'
            op['descricao_da_conta'] = '-'
        if op.get('centro_custo') is None:
            op['centro_custo'] = '-'
            op['desc_centro_de_custo'] = '-'
            
    return render(request, 'notas/op_list.html', {'linhas_op': linhas_op})

class JustificativaListView(LoginRequiredMixin, FilterView, ListView):
    login_url = 'login/'
    context_object_name = 'justificativas'
    template_name = 'notas/justificativas.html'

@login_required
def justificativa_admin_view(request):
    if not request.user.is_superuser:
        return render(request, '403.html', status=403) # Caso haja página de erro
    
    selected = request.GET.get('data_emissao_month', '')
    
    # Se não veio valor no GET, tenta preencher com mês atual (se houver dados),
    # caso contrário preenche com o mês mais recente disponível.
    if not selected:
        hoje = datetime.datetime.today()
        atual = f"{hoje.year}-{hoje.month:02d}"
        selected = atual
        
    justificativas = Justificativa.objects.all().order_by('-ativo', '-data_cadastro')
    return render(request, 'notas/admin.html', {'justificativas': justificativas, 'selected_month': selected})

@login_required
@require_POST
def justificativa_save(request):
    if not request.user.is_superuser:
        return JsonResponse({'success': False, 'error': 'Acesso negado'}, status=403)
    
    data = json.loads(request.body)
    j_id = data.get('id')
    texto = data.get('texto')
    
    if not texto:
        return JsonResponse({'success': False, 'error': 'Texto obrigatório'}, status=400)
        
    try:
        if j_id:
            justificativa = Justificativa.objects.get(id=j_id)
            justificativa.texto = texto
            justificativa.save()
        else:
            Justificativa.objects.create(texto=texto, usuario=request.user)
        return JsonResponse({'success': True})
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

@login_required
@require_POST
def justificativa_toggle_status(request):
    if not request.user.is_superuser:
        return JsonResponse({'success': False, 'error': 'Acesso negado'}, status=403)
        
    data = json.loads(request.body)
    j_id = data.get('id')
    
    try:
        justificativa = Justificativa.objects.get(id=j_id)
        justificativa.ativo = not justificativa.ativo
        justificativa.save()
        return JsonResponse({'success': True, 'ativo': justificativa.ativo})
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)
    

@login_required
def exportar_excel(request):
    # Obtém o queryset base com as pre-anotações e filial filter
    notas_view = NotasListView()
    notas_view.request = request
    base_qs = notas_view.get_queryset()

    # Aplica os filtros da URL
    f = NotaFilter(request.GET, queryset=base_qs)
    queryset = f.qs

    # 1. Criar o workbook e planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"

    # 2. Adicionar cabeçalhos (Substituído por campos existentes para não ocorrer Crash do framework)
    columns = [
        'chave',
        'filial',
        'nome_filial',
        'nota',
        'item',
        'no_pedido',
        'vendedor',
        'data_emissao',
        'lote',
        'cfop',
        'cfop_descri',
        'atualiza_estoque',
        'gera_duplicata',
        'cod_produto',
        'produto',
        'tipo_produto',
        'desc_tipo_produto',
        'armazem',
        'cod_cliente',
        'loja',
        'cliente',
        'grp_amar_ctb',
        'classificacao_produto',
        'estado_destino',
        'quantidade',
        'tabela_preco',
        'preco_tabela',
        'valor_contabil',
        'valor_unitario',
        'valor_ipi',
        'valor_imp5',
        'valor_imp6',
        'valor_icms_difal',
        'valor_icms',
        'aliq_icms',
    ]
    ws.append(columns)

    # 3. Adicionar dados (Dados reais baseados no Model para evitar crash na query)
    # Lembre-se de adaptar essa query de acordo com os dados exatos que desejar!
    
    # Obtém o queryset base com as pre-anotações e filial filter
    notas_view = NotasListView()
    notas_view.request = request
    base_qs = notas_view.get_queryset()

    # Aplica os filtros da URL
    f = NotaFilter(request.GET, queryset=base_qs)
    queryset = f.qs

    dados = list(queryset.values_list(*columns))
    for linha in dados:
        ws.append(linha)

    # 4. Formatos
    fmt_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fmt_header_font = Font(bold=True)
    fmt_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                        top=Side(style='thin'), bottom=Side(style='thin'))

    # Formatar o cabeçalho
    for cell in ws[1]:
        cell.font = fmt_header_font
        cell.fill = fmt_header_fill
        cell.border = fmt_border

    # 5. Aplica formatação de colunas (Largura e Números)
    for col_idx, col_name in enumerate(columns, 1):
        col_name_lower = str(col_name).lower()
        num_format = None
        
        if any(x in col_name_lower for x in ['valor_', 'custo', 'vlr_', 'margem_bruta']):
            num_format = 'R$ #,##0.00'
        elif any(x in col_name_lower for x in ['percent', 'aliquota', 'diff_percentual']):
            num_format = '0.00%'

        max_len = len(str(col_name))
        
        last_row = len(dados) + 1
        for row_idx in range(2, last_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            
            if num_format:
                cell.number_format = num_format
            
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2

    last_col_letter = get_column_letter(len(columns))
    last_row = len(dados) + 1

    # 6. Lógica para colorir a linha inteira se tp_movimento == '010'
    if 'tp_movimento' in columns:
        col_idx = columns.index('tp_movimento') + 1
        col_letter = get_column_letter(col_idx)
        
        fmt_destaque = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        
        # A fórmula é semelhante à do Excel. Passamos uma string numa lista e setamos stopIfTrue
        formula = [f'=${col_letter}2="010"']
        
        rule = FormulaRule(formula=formula, stopIfTrue=True, fill=fmt_destaque)
        ws.conditional_formatting.add(f"A2:{last_col_letter}{last_row}", rule)

    # Autofilter
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    # 7. Configurar a resposta HTTP para download
    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename="dados.xlsx"'

    # 8. Salvar o arquivo no response
    wb.save(response)
    return response


@login_required
def exportar_estatisticas_excel(request):
    import json
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    # 1. Obtém os dados chamando a API localmente
    resp = dados_vendas_api(request)
    if resp.status_code != 200:
        return HttpResponse("Erro ao gerar dados", status=500)
    
    dados_json = json.loads(resp.content)
    
    wb = Workbook()
    
    # Aba 1: Estatísticas Justificativas
    ws1 = wb.active
    ws1.title = "Estatísticas Justificativas"
    
    cols1 = ['Justificativa', 'Qtd Notas', 'Representatividade Vendas', 'Abaixo Margem Percentual']
    ws1.append(cols1)
    
    for item in dados_json.get('estatisticas_justificativas', []):
        ws1.append([
            item['justificativa'],
            item['contagem'],
            item['representatividade_vendas'] / 100.0,
            item['percentual_abaixo_margem'] / 100.0
        ])
        
    # Aba 2: Estatísticas Período
    ws2 = wb.create_sheet(title="Estatísticas Por Período")
    cols2 = ['Mês', 'Faturamento (valor_)', 'Custo Total', 'Margem Bruta', 'Margem Percentual']
    ws2.append(['Mês', 'Faturamento', 'Custo Total', 'Margem Total', 'Margem Percentual'])
    
    labels = dados_json.get('labels', [])
    vendas = dados_json.get('total_vendas', [])
    custos = dados_json.get('custo_total', [])
    margens = dados_json.get('margem_total', [])
    margens_pct = dados_json.get('margem_percentual', [])
    
    for i in range(len(labels)):
        ws2.append([
            labels[i],
            vendas[i],
            custos[i],
            margens[i],
            (margens_pct[i] / 100.0) if margens_pct[i] is not None else 0
        ])
    
    # Aplicar formato
    fmt_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fmt_header_font = Font(bold=True)
    fmt_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for ws, cols in [(ws1, cols1), (ws2, cols2)]:
        # Formatar cabeçalho
        for cell in ws[1]:
            cell.font = fmt_header_font
            cell.fill = fmt_header_fill
            cell.border = fmt_border

        for col_idx, col_name in enumerate(cols, 1):
            col_name_lower = str(col_name).lower()
            num_format = None
            if any(x in col_name_lower for x in ['valor_', 'custo', 'vlr_', 'margem_bruta', 'faturamento']):
                num_format = 'R$ #,##0.00'
            elif any(x in col_name_lower for x in ['percent', 'aliquota', 'diff_percentual', 'representatividade', 'abaixo margem']):
                num_format = '0.00%'

            max_len = len(str(ws.cell(row=1, column=col_idx).value))
            
            for row_idx in range(2, ws.max_row + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if num_format:
                    cell.number_format = num_format
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))

            ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 2
            
        last_col_letter = get_column_letter(len(cols))
        if ws.max_row > 1:
            ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"
        
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename="estatisticas.xlsx"'
    wb.save(response)
    return response

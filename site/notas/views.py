from setup import settings
from .filters import NotaFilter
from .models import Justificativa, Nota, Custo, Margem, Nf_Has_Justificativa, OP

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

        filial = self.request.GET.get('filial', '')
        context['filial_selecionada'] = filial
        
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
    
    # Se não veio valor no GET, tenta preencher com mês atual (se houver dados),
    # caso contrário preenche com o mês mais recente disponível.
    if not selected:
        hoje = datetime.datetime.today()
        atual = f"{hoje.year}-{hoje.month:02d}"
        selected = atual
        
    filiais = Nota.objects.values('filial', 'nome_filial').distinct()
    
    if filiais:
        return render(request, 'notas/estatisticas.html', {'filiais': filiais, 'selected_month': selected})
    
    return render(request, 'notas/estatisticas.html')

@login_required
def dados_vendas_api(request):
    
    data_inicio = request.GET.get('inicio')
    data_fim = request.GET.get('fim')
    filial = request.GET.get('filial')
        
    queryset = Nota.objects.all()
    
    if data_inicio:
        queryset = queryset.filter(data_emissao__gte=parse_date(data_inicio))
    if data_fim:
        queryset = queryset.filter(data_emissao__lte=parse_date(data_fim))
    if filial:
        queryset = queryset.filter(
            Q(filial__icontains=filial) | Q(nome_filial__icontains=filial)
        )
    
    subquery_margem = Margem.objects.filter(
        chave=OuterRef('chave')
    ).values('margem_bruta_percentual').order_by('-custo__data_cadastro')[:1] # O Gemini disse: Para ordenar pelo atributo cadastro da tabela Custo (que possui a Foreign Key para Margem), você precisa utilizar a sintaxe de "follow relationship" do Django, que utiliza o duplo sublinhado (__).
        
    queryset = queryset.annotate(
        mes=TruncMonth('data_emissao')
    ).values('mes').annotate(
        total_vendas=Sum('valor_contabil'),
        margem=Avg(subquery_margem)*100
    ).order_by('mes')
    
    labels = [item['mes'].strftime('%b/%Y') for item in queryset]
    
    total_vendas = [float(item['total_vendas']) for item in queryset]
    margem_por_mes = [float(item['margem']) for item in queryset] 
        
    return JsonResponse({
        # 'margem': queryset['margem_total'],
        'labels': labels,
        'total_vendas': total_vendas,
        'margens_por_mes': margem_por_mes
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
@cache_page(60 * 15)
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
            
    return render(request, 'notas/op_list.html', {'linhas_op': linhas_op})

class JustificativaListView(LoginRequiredMixin, FilterView, ListView):
    login_url = 'login/'
    context_object_name = 'justificativas'
    template_name = 'notas/justificativas.html'
    
from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.utils import timezone
from django.views.generic import ListView
from django.views.decorators.http import require_http_methods, require_POST
from django.db.models import OuterRef, Subquery, DecimalField, Case, When, F, Value, ExpressionWrapper
from django.db.models.functions import Coalesce, ExtractYear, ExtractMonth
from django_filters.views import FilterView
from decimal import Decimal

import json
import datetime

from .models import Justificativa, Nota, Custo, Margem, Nf_Has_Justificativa

from .filters import NotaFilter

class NotasListView(FilterView, ListView):
    model = Nota
    filterset_class = NotaFilter
    # paginate_by = 25
    ordering = ['-data_emissao'] 
    context_object_name = 'nota_list'
    template_name = 'notas/nota_list.html'
    
    def get_queryset(self):
        cfops_especiais = ['5101', '6101', '5116', '6116', '6107']
        # 1. Definimos a subquery para buscar o Custo relacionado
        # Filtramos onde a chave do Custo é igual à chave da Nota atual (OuterRef('pk'))
        custo_mais_recente = Custo.objects.filter(
            chave=OuterRef('pk')
        ).order_by('-data_cadastro').values('valor')[:1]
        
        justificativa_mais_recente = Nf_Has_Justificativa.objects.filter(
            nf=OuterRef('pk')
        ).order_by('-data_cadastro').values('justificativa')[:1]
        
        
        # 2. Anotamos o queryset principal da Nota com esse valor
        queryset = Nota.objects.annotate(
            ultimo_custo_valor=Coalesce(
                Subquery(custo_mais_recente), 
                0, 
                output_field=DecimalField()
            ),
            icms_calculado=Case(
                When(cfop__in=cfops_especiais, then=F('valor_icms') * 0.02), # 2% do ICMS
                default=F('valor_icms'),                                     # ICMS cheio
                output_field=DecimalField()
            ),
            margem_bruta=ExpressionWrapper(
                F('valor_contabil') 
                - F('ultimo_custo_valor') 
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

        # Se não veio valor no GET, tenta preencher com mês atual (se houver dados),
        # caso contrário preenche com o mês mais recente disponível.
        if not selected:
            from datetime import datetime
            hoje = datetime.today()
            atual = f"{hoje.year}-{hoje.month:02d}"
            # se atual estiver na lista de meses disponiveis, usa; senão usa o primeiro disponível
            valores = [m['value'] for m in meses]
            if atual in valores:
                selected = atual
            elif len(valores) > 0:
                selected = valores[0]

        context['selected_month'] = selected

        return context
    
    
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
            val_icms = nota.valor_icms * Decimal('0.02')
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
            'margem_bruta': f'{margem_bruta:.2f}',
            'margem_percentual': f'{margem_percentual*100:.2f}%'
        })

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    

@require_POST
def atualizar_justificativa_api(request):
    try:
        data = json.loads(request.body)
        nota_chave = data.get('chave')
        nova_justificativa_id = data.get('justificativa')

        if not nota_chave or not nova_justificativa_id:
            return JsonResponse({'success': False, 'error': 'Dados incompletos'}, status=400)

        nota = get_object_or_404(Nota, pk=nota_chave)
        justificativa = get_object_or_404(Justificativa, id=nova_justificativa_id)

        # Salva o vínculo (Nf_Has_Justificativa)
        # Usando update_or_create caso queira apenas uma justificativa por nota
        Nf_Has_Justificativa.objects.create(
            usuario=request.user if request.user.is_authenticated else None,
            nf=nota,
            justificativa=justificativa
        )

        return JsonResponse({
            'success': True,
            'justificativa_id': justificativa.id
        })

    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)}, status=500)

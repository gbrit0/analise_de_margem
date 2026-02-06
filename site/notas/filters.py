import django_filters
from django_filters import CharFilter, NumberFilter
from django.db.models import Q
from .models import Nota


class NotaFilter(django_filters.FilterSet):
    # Filtro por mês/ano no formato YYYY-MM (renderizamos um <input type="month"> no template)
    data_emissao_month = CharFilter(method='filter_by_month', label='Mês')

    lote = CharFilter(field_name='lote', lookup_expr='icontains', label='Lote')
    cfop = CharFilter(method='filter_cfop', label='CFOP (código ou descrição)')
    # Permite busca por nome do produto OU código do produto
    produto = CharFilter(method='filter_produto', label='Produto (nome ou código)')
    tipo_produto = CharFilter(field_name='tipo_produto', lookup_expr='icontains', label='Tipo do produto')
    classificacao_produto = CharFilter(field_name='classificacao_produto', lookup_expr='icontains', label='Classificação')

    margem_minima = NumberFilter(field_name='margem_percentual', lookup_expr='gte', label='Margem % mínima')
    margem_maxima = NumberFilter(field_name='margem_percentual', lookup_expr='lte', label='Margem % máxima')

    class Meta:
        model = Nota
        fields = [
            'filial', 'nota', 'grp_amar_ctb', 'produto', 'cod_produto', 'tipo_produto',
            'classificacao_produto', 'lote', 'cfop'
        ]

    def filter_produto(self, queryset, name, value):
        """Filtra notas cujo nome do produto OU código contenham o valor informado."""
        if not value:
            return queryset
        return queryset.filter(
            Q(produto__icontains=value) | Q(cod_produto__icontains=value)
        )

    def filter_by_month(self, queryset, name, value):
        """Filtra pelo mês informado no formato YYYY-MM."""
        if not value:
            return queryset
        try:
            parts = value.split('-')
            year = int(parts[0])
            month = int(parts[1])
        except Exception:
            return queryset

        return queryset.filter(data_emissao__year=year, data_emissao__month=month)

    def filter_cfop(self, queryset, name, value):
        """Filtra notas cuja descrição da CFOP OU código contenham o valor informado."""
        if not value:
            return queryset
        return queryset.filter(
            Q(cfop__icontains=value) | Q(cfop_descri__icontains=value)
        )
        
    def filter_filial(self, queryset, name, value):
        """Filtra notas por filial"""
        if not value:
            return queryset
        return queryset.filter(
            Q(filial__icontains=value) | Q(filial_descri__icontains=value)
        )


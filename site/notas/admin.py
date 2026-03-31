from django.contrib import admin
from .models import (
    Justificativa, 
    Nota, 
    Custo, 
    Nf_Has_Justificativa, 
    Margem, 
    OP
)

class CustomJustificativasAdmin(admin.ModelAdmin):
    list_display = ('texto', 'data_cadastro', 'ativo', 'data_desativa', 'usuario')
    list_display_links = ('texto', 'data_cadastro', 'ativo', 'data_desativa', 'usuario')
    search_fields = ('texto', 'data_cadastro', 'ativo', 'data_desativa', 'usuario')

class CustomNotasAdmin(admin.ModelAdmin):
    list_display = ['filial', 'nota', 'no_pedido', 'produto',]
    list_display_links = ['filial', 'nota', 'no_pedido', 'produto',]
    search_fields = ['filial', 'nota', 'no_pedido', 'produto',]
    readonly_fields = [
        'chave',
        'filial',
        'nome_filial',
        'nota',
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
        'armazem',
        'cod_cliente',
        'loja',
        'cliente',
        'grp_amar_ctb',
        'classificacao_produto',
        'estado_destino',
        'quantidade',
        'valor_contabil',
        'valor_unitario',
        'valor_ipi',
        'valor_imp5',
        'valor_imp6',
        'valor_icms_difal',
        'valor_icms',
        'aliq_icms',
        'delete',
    ]
    
class CustomCustoAdmin(admin.ModelAdmin):
    list_display = ['id', 'chave', 'valor', 'data_cadastro', 'usuario']
    list_display_links = ['id', 'chave', 'valor', 'data_cadastro', 'usuario']
    search_fields = ['id', 'chave', 'valor', 'data_cadstro', 'usuario']

class CustomNf_Has_JustificativaAdmin(admin.ModelAdmin):
    list_display = ['nf', 'justificativa', 'data_cadastro', 'usuario']
    list_display_links = ['nf', 'justificativa', 'data_cadastro', 'usuario']
    search_fields = ['nf', 'justificativa', 'data_cadastro', 'usuario']

class CustomMargemAdmin(admin.ModelAdmin):
    list_display = ["chave", "custo", "margem_bruta", "margem_bruta_percentual", ] # "cadastro"
    list_display_links = ["chave", "custo", "margem_bruta", "margem_bruta_percentual", ]
    search_fields = ["chave", "custo", "margem_bruta", "margem_bruta_percentual", ]
    
class CustomOPAdmin(admin.ModelAdmin):
    list_display = ['lote', 'documento', 'produto', 'ord_producao']
    list_display_links = ['lote', 'documento', 'produto', 'ord_producao']
    search_fields = ['lote', 'documento', 'produto', 'ord_producao']
    
admin.site.register(Justificativa, CustomJustificativasAdmin)
admin.site.register(Nota, CustomNotasAdmin)
admin.site.register(Custo, CustomCustoAdmin)
admin.site.register(Nf_Has_Justificativa, CustomNf_Has_JustificativaAdmin)
admin.site.register(Margem, CustomMargemAdmin)
admin.site.register(OP, CustomOPAdmin)
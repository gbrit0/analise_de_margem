from django.contrib import admin
from .models import Justificativa, Nota, Custo, Nf_Has_Justificativa, Margem

class CustomJustificativasAdmin(admin.ModelAdmin):
    list_display = ('texto', 'data_cadastro', 'ativo', 'data_desativa')
    list_display_links = ('texto', 'data_cadastro', 'ativo', 'data_desativa')
    search_fields = ('texto', 'data_cadastro', 'ativo', 'data_desativa')

class CustomNotasAdmin(admin.ModelAdmin):
    list_display = ['filial', 'nota', 'no_pedido', 'produto',]
    list_display_links = ['filial', 'nota', 'no_pedido', 'produto',]
    search_fields = ['filial', 'nota', 'no_pedido', 'produto',]
    
class CustomCustoAdmin(admin.ModelAdmin):
    list_display = ['id', 'chave', 'valor', 'data_cadastro', 'usuario']
    list_display_links = ['id', 'chave', 'valor', 'data_cadastro', 'usuario']
    search_fields = ['id', 'chave', 'valor', 'data_cadstro', 'usuario']

class CustomNf_Has_JustificativaAdmin(admin.ModelAdmin):
    list_display = ['nf', 'justificativa', 'data_cadastro', 'usuario']
    list_display_links = ['nf', 'justificativa', 'data_cadastro', 'usuario']
    search_fields = ['nf', 'justificativa', 'data_cadastro', 'usuario']

class CustomMargemAdmin(admin.ModelAdmin):
    list_display = ["chave", "custo", "margem_bruta", "margem_bruta_percentual", ]
    list_display_links = ["chave", "custo", "margem_bruta", "margem_bruta_percentual", ]
    search_fields = ["chave", "custo", "margem_bruta", "margem_bruta_percentual", ]
    
admin.site.register(Justificativa, CustomJustificativasAdmin)
admin.site.register(Nota, CustomNotasAdmin)
admin.site.register(Custo, CustomCustoAdmin)
admin.site.register(Nf_Has_Justificativa, CustomNf_Has_JustificativaAdmin)
admin.site.register(Margem, CustomMargemAdmin)
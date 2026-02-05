from django.db import models
from users.models import CustomUser

class Nota(models.Model):
    chave = models.CharField(primary_key=True, max_length=50)
    filial = models.CharField(default=None, blank=False, max_length=50)
    nota = models.CharField(default=None, blank=False, max_length=20)
    no_pedido = models.CharField(default=None, blank=False, max_length=15)
    vendedor = models.CharField(default=None, blank=False, null=True, max_length=60)
    data_emissao = models.DateField()
    lote = models.CharField(default=None, blank=False, max_length=20)
    cfop = models.CharField(default=None, blank=False, max_length=6)
    cfop_descri = models.CharField(default=None, blank=False, max_length=60)
    atualiza_estoque = models.CharField(default=None, blank=False, max_length=1)
    gera_duplicata = models.CharField(default=None, blank=False, max_length=1)
    cod_produto = models.CharField(default=None, blank=False, max_length=15)
    produto = models.CharField(default=None, blank=False, max_length=50)
    tipo_produto = models.CharField(default=None, blank=False, max_length=2)
    armazem = models.CharField(default=None, blank=False, max_length=60)
    cod_cliente = models.CharField(default=None, blank=False, max_length=10)
    loja = models.CharField(default=None, blank=False, max_length=4)
    cliente = models.CharField(default=None, blank=False, max_length=100)
    grp_amar_ctb = models.CharField(default=None, blank=False, max_length=10)
    classificacao_produto = models.CharField(default=None, null=True, max_length=100)
    estado_destino = models.CharField(default=None, blank=False, max_length=2)
    quantidade = models.DecimalField(max_digits=12, decimal_places=2)
    valor_contabil = models.DecimalField(max_digits=18, decimal_places=2)
    valor_unitario = models.DecimalField(max_digits=18, decimal_places=2)
    valor_ipi = models.DecimalField(max_digits=18, decimal_places=2)
    valor_imp5 = models.DecimalField(max_digits=18, decimal_places=2)
    valor_imp6 = models.DecimalField(max_digits=18, decimal_places=2)
    valor_icms_difal = models.DecimalField(max_digits=18, decimal_places=2)
    valor_icms = models.DecimalField(max_digits=18, decimal_places=2)
    aliq_icms = models.DecimalField(max_digits=5, decimal_places=2)
    delete = models.BooleanField(default=False)
    
    class Meta():
        verbose_name = "Nota"
        verbose_name_plural = "Notas"
        
    @property
    def justificativa_atual(self):
        relacao = self.nf_has_justificativa_set.last()
        return relacao.justificativa if relacao else None
    
    def __str__(self):
        return self.chave

class Custo(models.Model):
    chave = models.ForeignKey(to=Nota, on_delete=models.PROTECT)
    valor = models.DecimalField(max_digits=18, decimal_places=2)
    data_cadastro = models.DateTimeField(auto_now=True)
    # usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, null=True, blank=True)
    usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, default=2)
    
    class Meta():
        verbose_name = "Custo"
        verbose_name_plural = "Custos"

class Justificativa(models.Model):
    texto = models.CharField(max_length=150)
    data_cadastro = models.DateTimeField(auto_now=True)
    ativo = models.BooleanField(default=True)
    data_desativa = models.DateTimeField(blank=True, null=True)
    # usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, null=True, blank=True)
    usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, default=2)
    
    class Meta():
        verbose_name = "Justificativa"
        verbose_name_plural = "Justificativas"
        
    def __str__(self):
        return self.texto
    
class Nf_Has_Justificativa(models.Model):
    nf = models.ForeignKey(to=Nota, on_delete=models.PROTECT)
    justificativa = models.ForeignKey(to=Justificativa, on_delete=models.PROTECT)
    data_cadastro = models.DateTimeField(auto_now=True)
    # usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, null=True, blank=True)
    usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, default=2)

class Margem(models.Model):
    chave = models.ForeignKey(to=Nota, on_delete=models.PROTECT)
    custo = models.ForeignKey(to=Custo, on_delete=models.PROTECT)
    margem_bruta = models.DecimalField(max_digits=18, decimal_places=2)
    margem_bruta_percentual = models.DecimalField(max_digits=18, decimal_places=4)
    
    class Meta():
        verbose_name = "Margem"
        verbose_name_plural = "Margens"
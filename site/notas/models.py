from django.db import models
from users.models import CustomUser

import locale

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class Nota(models.Model):
    chave = models.CharField(primary_key=True, max_length=50)
    filial = models.CharField(default=None, blank=False, max_length=50)
    nome_filial = models.CharField(default=None, blank=False, max_length=80)
    nota = models.CharField(default=None, blank=False, max_length=20)
    item = models.CharField(default=None, blank=False, max_length=20)
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
    desc_tipo_produto = models.CharField(default=None, blank=False, max_length=50)
    armazem = models.CharField(default=None, blank=False, max_length=60)
    cod_cliente = models.CharField(default=None, blank=False, max_length=10)
    loja = models.CharField(default=None, blank=False, max_length=4)
    cliente = models.CharField(default=None, blank=False, max_length=100)
    grp_amar_ctb = models.CharField(default=None, blank=False, max_length=10)
    classificacao_produto = models.CharField(default=None, null=True, max_length=100)
    estado_destino = models.CharField(default=None, blank=False, max_length=2)
    quantidade = models.DecimalField(max_digits=12, decimal_places=2)
    tabela_preco = models.CharField(default=None, blank=False, max_length=10)
    preco_tabela = models.DecimalField(max_digits=18, decimal_places=2, default=None, blank=True)
    valor_contabil = models.DecimalField(max_digits=18, decimal_places=2)
    valor_unitario = models.DecimalField(max_digits=18, decimal_places=2)
    valor_ipi = models.DecimalField(max_digits=18, decimal_places=2)
    valor_imp5 = models.DecimalField(max_digits=18, decimal_places=2)
    valor_imp6 = models.DecimalField(max_digits=18, decimal_places=2)
    valor_icms_difal = models.DecimalField(max_digits=18, decimal_places=2)
    valor_icms = models.DecimalField(max_digits=18, decimal_places=2)
    aliq_icms = models.DecimalField(max_digits=5, decimal_places=2)
    recno = models.BigIntegerField()
    delete = models.BooleanField(default=False)
    
    class Meta():
        verbose_name = "Nota"
        verbose_name_plural = "Notas"
        
    @property
    def justificativa_atual(self):
        relacao = self.nf_has_justificativa_set.last()
        return relacao.justificativa if relacao else None
    
    @property
    def nome_curto_vendedor(self):
        if self.vendedor:
            return " ".join(self.vendedor.split()[:1])
        return "-"
    
    @property
    def valor_contabil_formatado(self):
        return locale.currency(self.valor_contabil, grouping=True)
    
    @property
    def margem_bruta_formatada(self):
        return locale.currency(self.margem_bruta, grouping=True)
    
    @property
    def valor_unitario_formatado(self):
        return locale.currency(self.valor_unitario, grouping=True)
    
    @property
    def valor_ipi_formatado(self):
        return locale.currency(self.valor_ipi, grouping=True)
    
    @property
    def valor_imp5_formatado(self):
        return locale.currency(self.valor_imp5, grouping=True)
    
    @property
    def valor_imp6_formatado(self):
        return locale.currency(self.valor_imp6, grouping=True)
    
    @property
    def valor_icms_difal_formatado(self):
        return locale.currency(self.valor_icms_difal, grouping=True)
    
    @property
    def valor_icms_formatado(self):
        return locale.currency(self.valor_icms, grouping=True)
    
    @property
    def relacao_lote_nota(self):
        return f"{self.filial}{self.nota}{self.item}{self.recno}"
    
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
        
    def __str__(self):
        return str(self.valor)


class Justificativa(models.Model):
    texto = models.CharField(max_length=150)
    data_cadastro = models.DateTimeField(auto_now=True)
    ativo = models.BooleanField(default=True)
    data_desativa = models.DateTimeField(blank=True, null=True)
    # usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, null=True, blank=True)
    usuario = models.ForeignKey(to=CustomUser, on_delete=models.PROTECT, default=2, editable=False)
    
    class Meta():
        verbose_name = "Justificativa"
        verbose_name_plural = "Justificativas"
        
    def __str__(self):
        return self.texto
    
    def save(self, *args, **kwargs):
        if not self.ativo:
            if not self.data_desativa:
                self.data_desativa = timezone.now()
        else:
            self.data_desativa = None
            
        super().save(*args, **kwargs)
    
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
    # cadastro = models.DateTimeField(auto_created=True, auto_now=True)
    
    class Meta():
        verbose_name = "Margem"
        verbose_name_plural = "Margens"
        
class OP(models.Model):
    id_op                = models.CharField(default=None, blank=True, max_length=80)
    filial               = models.CharField(default=None, blank=False, max_length=50)
    produto              = models.CharField(default=None, blank=False, max_length=15)
    armazem              = models.CharField(default=None, blank=True, max_length=60)
    tp_movimento         = models.CharField(default=None, blank=True, max_length=3)
    descricao_tm         = models.CharField(default=None, blank=True, null=True, max_length=80)
    descr_prod           = models.CharField(default=None, blank=True, null=True, max_length=80)
    unidade              = models.CharField(default=None, blank=True, max_length=2)
    quantidade           = models.DecimalField(max_digits=12, decimal_places=2)
    quant_2              = models.DecimalField(max_digits=12, decimal_places=2)
    custo                = models.DecimalField(max_digits=18, decimal_places=2)
    custo_2              = models.DecimalField(max_digits=18, decimal_places=2)
    ord_producao         = models.CharField(default=None, blank=True, max_length=12)
    lote                 = models.CharField(default=None, blank=True, max_length=20)
    os_ass_tecn          = models.CharField(default=None, blank=True, max_length=8)
    grupo                = models.CharField(default=None, blank=True, max_length=4)
    descricao_grupo      = models.CharField(default=None, blank=True, null=True, max_length=60)
    tipo_re_de	         = models.CharField(default=None, blank=True, null=True, max_length=3)
    ext_texto	         = models.CharField(default=None, blank=True, null=True, max_length=2)
    documento	         = models.CharField(default=None, blank=True, null=True, max_length=12)
    dt_emissao	         = models.DateField()
    c_contabil	         = models.CharField(default=None, blank=True, max_length=10)
    descricao_da_conta	 = models.CharField(default=None, blank=True, null=True, max_length=60)
    centro_custo	     = models.CharField(default=None, blank=True, max_length=8)
    desc_centro_de_custo = models.CharField(default=None, blank=True, null=True, max_length=60)
    parc_total	         = models.CharField(default=None, blank=True, null=True, max_length=1)
    estornado            = models.CharField(default=None, blank=True, null=True, max_length=1)
    sequencial	         = models.CharField(default=None, blank=True, null=True, max_length=6)
    tipo	             = models.CharField(default=None, blank=True, null=True, max_length=2)
    usuario	             = models.CharField(default=None, blank=True, null=True, max_length=40)
    nr_s_a	             = models.CharField(default=None, blank=True, null=True, max_length=6)
    item_s_a	         = models.CharField(default=None, blank=True, null=True, max_length=2)
    
    class Meta():
        verbose_name = "Op"
        verbose_name_plural = "Ops"
    
    @property
    def custo1_fmt(self):
        return locale.currency(self.custo, grouping=True)
    
    @property
    def custo2_fmt(self):
        return locale.currency(self.custo_2, grouping=True)
        
        
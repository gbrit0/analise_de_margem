import os
import time
import pyodbc
from datetime import date
from decimal import Decimal
from django.core.management.base import BaseCommand
from django.db import transaction
from django.db.models import Subquery, OuterRef
from notas.models import Nota, Custo, Margem, OP
from users.models import CustomUser
from dbutils.pooled_db import PooledDB
from setup import settings

def to_decimal(val):
    if val is None:
        return Decimal('0.00')
    if isinstance(val, Decimal):
        return val
    return Decimal(str(val))

def setup_protheus_connection(conn):
    conn.setdecoding(pyodbc.SQL_CHAR, encoding='cp1252')
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding='cp1252')
    conn.setdecoding(pyodbc.SQL_WMETADATA, encoding='cp1252')
    
pool = PooledDB(
    creator=pyodbc,
    # setsession=[setup_protheus_connection],
    maxconnections=10, # Como vamos usar lotes, não precisamos de tantas conexões simultâneas
    mincached=2,
    blocking=True,
    driver='{ODBC Driver 17 for SQL Server}',
    server=f'{os.getenv("PROTHEUS_DB_HOST")}',
    database=f'{os.getenv("PROTHEUS_DB_DATABASE")}',
    uid=f'{os.getenv("PROTHEUS_DB_USER")}',
    pwd=f'{os.getenv("PROTHEUS_DB_PASSWORD")}'
)

def dividir_em_lotes(lista, tamanho):
    """Gera pedaços menores de uma lista para não estourar os limites do SQL (ex: cláusula IN gigante)."""
    for i in range(0, len(lista), tamanho):
        yield lista[i:i + tamanho]

def protheus_delete_marcado(valor):
    return str(valor or '').strip() == '*'

class Command(BaseCommand):
    help = 'Sincroniza dados do Django com o Protheus de forma otimizada'

    def handle(self, *args, **options):
        inicio = time.time()
        self.stdout.write("Iniciando sincronização em lote...")
        
        with open(f'{settings.BASE_DIR}/notas/management/commands/querys/queryNotas.sql', 'r') as f:
            query_notas = f.read()
            
        with open(f'{settings.BASE_DIR}/notas/management/commands/querys/buscaOPs.sql', 'r') as f:
            query_ops_base = f.read()
        try:
            with pool.connection() as conn:
                with conn.cursor() as cursor:
                    cursor.execute(str(query_notas))
                    rows_notas = cursor.fetchall()
            
            self.stdout.write(f"{len(rows_notas)} notas encontradas no Protheus. Processando...")

            # 3. Mapeia o que já existe no Django para saber se é Insert ou Update
            chaves_protheus = [row[0] for row in rows_notas]
            ultimo_custo_subquery = Custo.objects.filter(chave=OuterRef('pk')).order_by('-data_cadastro').values('valor')[:1]
            existentes = Nota.objects.filter(chave__in=chaves_protheus).annotate(
                ultimo_custo=Subquery(ultimo_custo_subquery)
            ).values('chave', 'preco_tabela', 'ultimo_custo')
            
            existentes_dict = {
                n['chave']: {
                    'preco_tabela': n['preco_tabela'],
                    'ultimo_custo': n['ultimo_custo']
                }
                for n in existentes
            }
            chaves_existentes = set(existentes_dict.keys())
            
            usuario_sistema = CustomUser.objects.get(id=2)
            data_corte = date(2026, 2, 1)
            cfops_especiais = {'5101', '6101', '5116', '6116', '6107'}

            # Regra de bloqueio de atualização:
            # se for após dia 3, não atualizar notas do mês/ano anterior ao atual
            hoje = date.today()
            if hoje.month == 1:
                mes_anterior = 12
                ano_mes_anterior = hoje.year - 1
            else:
                mes_anterior = hoje.month - 1
                ano_mes_anterior = hoje.year
            bloquear_atualizacao_mes_anterior = hoje.day > 3

            notas_para_criar = {}
            notas_para_atualizar = {}
            notas_para_atualizar_delete = {}
            custos_para_criar = {}
            margens_para_criar = {}
            notas_ignoradas_regra_data = 0
            notas_ignoradas_sem_data = 0

            for row in rows_notas:
                chave, filial, nome_filial, nota, item, no_pedido, vendedor, data_emissao, lote, cfop, cfop_descri, atualiza_estoque, gera_duplicata, cod_produto, produto, tipo_produto, desc_tipo_produto, armazem, cod_cliente, loja, cliente, grp_amar_ctb, classificacao_produto, estado_destino, quantidade, tabela_preco, preco_tabela, valor_contabil, custo_valor, valor_unitario, valor_ipi, valor_imp5, valor_imp6, valor_icms_difal, valor_icms, base_icms,aliq_icms, recno, comentario, deletado = row
                delete_marcado = protheus_delete_marcado(deletado)
                
                # Garantir que campos de texto não nulos no banco de dados local não recebam None
                filial = filial or '-'
                nome_filial = nome_filial or '-'
                nota = nota or '-'
                item = item or '-'
                no_pedido = no_pedido or '-'
                lote = lote or '-'
                cfop = cfop or '-'
                cfop_descri = cfop_descri or '-'
                atualiza_estoque = atualiza_estoque or '-'
                gera_duplicata = gera_duplicata or '-'
                cod_produto = cod_produto or '-'
                produto = produto or '-'
                tipo_produto = tipo_produto or '-'
                desc_tipo_produto = desc_tipo_produto or '-'
                armazem = armazem or '-'
                cod_cliente = cod_cliente or '-'
                loja = loja or '-'
                cliente = cliente or '-'
                grp_amar_ctb = grp_amar_ctb or '-'
                estado_destino = estado_destino or '-'
                tabela_preco = tabela_preco or '-'

                if not data_emissao:
                    notas_ignoradas_sem_data += 1
                    self.stderr.write(
                        f"Nota ignorada sem data_emissao: chave={chave}, filial={filial}, nota={nota}, item={item}, recno={recno}"
                    )
                    continue
                
                # Guarda o lote se existir para a busca das OPs
                # if lote and lote.strip():
                #     lotes_unicos.add(lote.strip())
                #     relacao_lote_nota[lote.strip()] = lote.strip() # f"{filial}{nota}{item}{recno}"

                # Calcula a comissão se data_emissao >= 2026-06-01 e cod_produto começa com 'G0'
                is_intercompany = (cod_cliente, loja) in {
                    ('44592860', '0001'), # Agrogera BA
                    ('04675878', '0001'), # BRG Matriz
                    ('44592860', '0002'), # Agrogera BA Filial GO
                    ('27379581', '0004'), # GRID MS - INOCENCIA
                    ('27379581', '0001'), # GRID GO
                    ('27379581', '0002'), # GRID MG
                    ('27379581', '0003'), # GRID PA
                }
                vendedor_nome = (vendedor or '').strip()
                vendedor_sem_comissao = (
                    not vendedor_nome 
                    or vendedor_nome in ['DAVID MARTINS DOS SANTOS', 'CLEITON PAULO DE MOURA']
                )
                if is_intercompany or vendedor_sem_comissao:
                    comissao_calc = Decimal('0.00')
                else:
                    comissao_calc = Decimal('0.004') * to_decimal(valor_contabil) if (data_emissao >= date(2026, 6, 1) and cod_produto.startswith('G0')) else Decimal('0.00')

                nota_obj = Nota(
                    chave=chave, filial=filial, nome_filial=nome_filial, nota=nota, item=item,
                    no_pedido=no_pedido, vendedor=vendedor, data_emissao=data_emissao, lote=lote,
                    cfop=cfop, cfop_descri=cfop_descri, atualiza_estoque=atualiza_estoque,
                    gera_duplicata=gera_duplicata, cod_produto=cod_produto, produto=produto,
                    tipo_produto=tipo_produto, desc_tipo_produto=desc_tipo_produto, armazem=armazem,
                    cod_cliente=cod_cliente, loja=loja, cliente=cliente, grp_amar_ctb=grp_amar_ctb,
                    classificacao_produto=classificacao_produto, estado_destino=estado_destino,
                    quantidade=quantidade, tabela_preco=tabela_preco, preco_tabela=preco_tabela,
                    valor_contabil=valor_contabil, valor_unitario=valor_unitario, valor_ipi=valor_ipi,
                    valor_imp5=valor_imp5, valor_imp6=valor_imp6, valor_icms_difal=valor_icms_difal,
                    valor_icms=valor_icms, base_icms=base_icms, aliq_icms=aliq_icms, recno=recno, 
                    comentario=comentario, delete=delete_marcado, comissao=comissao_calc
                )

                if chave not in chaves_existentes:
                    notas_para_criar[chave] = nota_obj
                    
                    # Prepara Custo e Margem apenas para as notas novas
                    custo_obj = Custo(chave_id=chave, valor=custo_valor, usuario=usuario_sistema)
                    custos_para_criar[chave] = custo_obj
                    
                    pro_goias = Decimal('0.0477') if data_emissao >= data_corte else Decimal('0.02')
                    
                    icms_calc = to_decimal(base_icms) * pro_goias if cfop in cfops_especiais else to_decimal(valor_icms)
                    
                    margem_bruta = (
                        to_decimal(valor_contabil) 
                        - to_decimal(custo_valor) 
                        - to_decimal(valor_ipi) 
                        - to_decimal(valor_imp5) 
                        - to_decimal(valor_imp6) 
                        - to_decimal(valor_icms_difal) 
                        - to_decimal(icms_calc) 
                        - to_decimal(comissao_calc)
                    )
                    
                    margem_obj = Margem(
                        chave_id=chave,
                        custo=custo_obj, # O Django lida com a chave estrangeira em bulk_create se a PK estiver explícita (depende do setup, ver nota abaixo)
                        margem_bruta=margem_bruta,
                        margem_bruta_percentual=margem_bruta / to_decimal(valor_contabil) if to_decimal(valor_contabil) else Decimal('0.00')
                    )
                    margens_para_criar[chave] = margem_obj
                else:
                    if (
                        bloquear_atualizacao_mes_anterior
                        and data_emissao
                        and data_emissao.month == mes_anterior
                        and data_emissao.year == ano_mes_anterior
                    ):
                        # Mantém o preco_tabela atual do banco de dados para esta nota
                        nota_obj.preco_tabela = existentes_dict[chave]['preco_tabela']
                        notas_ignoradas_regra_data += 1

                    # Se o custo do Protheus mudou ou não existe no banco, criamos um novo Custo e Margem
                    ultimo_custo_db = existentes_dict[chave]['ultimo_custo']
                    c_protheus = float(custo_valor) if custo_valor is not None else 0.0
                    c_db = float(ultimo_custo_db) if ultimo_custo_db is not None else None

                    if c_db is None or abs(c_db - c_protheus) > 0.001:
                        custo_obj = Custo(chave_id=chave, valor=custo_valor, usuario=usuario_sistema)
                        custos_para_criar[chave] = custo_obj
                        
                        pro_goias = Decimal('0.0477') if data_emissao >= data_corte else Decimal('0.02')
                        icms_calc = to_decimal(base_icms) * pro_goias if cfop in cfops_especiais else to_decimal(valor_icms)
                        
                        margem_bruta = (
                            to_decimal(valor_contabil) 
                            - to_decimal(custo_valor) 
                            - to_decimal(valor_ipi) 
                            - to_decimal(valor_imp5) 
                            - to_decimal(valor_imp6) 
                            - to_decimal(valor_icms_difal) 
                            - to_decimal(icms_calc) 
                            - to_decimal(comissao_calc)
                        )
                        
                        margem_obj = Margem(
                            chave_id=chave,
                            custo=custo_obj,
                            margem_bruta=margem_bruta,
                            margem_bruta_percentual=margem_bruta / to_decimal(valor_contabil) if to_decimal(valor_contabil) else Decimal('0.00')
                        )
                        margens_para_criar[chave] = margem_obj

                    notas_para_atualizar[chave] = nota_obj

            # 4. Executa as operações em lote no Django
            with transaction.atomic():
                if notas_para_criar:
                    Nota.objects.bulk_create(notas_para_criar.values(), batch_size=1000)
                
                if custos_para_criar:
                    Custo.objects.bulk_create(custos_para_criar.values(), batch_size=1000)
                
                if margens_para_criar:
                    Margem.objects.bulk_create(margens_para_criar.values(), batch_size=1000)
                
                if notas_para_atualizar:
                    # Campos que devem ser atualizados
                    campos_update = [
                        'cfop', 'cfop_descri', 'estado_destino', 'quantidade', 'tabela_preco',
                        'valor_contabil', 'valor_unitario', 'valor_ipi', 'valor_imp5',
                        'valor_imp6', 'valor_icms_difal', 'valor_icms', 'base_icms','aliq_icms',
                        'delete', 'preco_tabela', 'comissao'
                    ] 
                    Nota.objects.bulk_update(notas_para_atualizar.values(), campos_update, batch_size=1000)

                if notas_para_atualizar_delete:
                    Nota.objects.bulk_update(notas_para_atualizar_delete.values(), ['delete'], batch_size=1000)

            self.stdout.write("Notas salvas. Buscando OPs no Protheus...")
            if notas_ignoradas_sem_data:
                self.stdout.write(
                    f"{notas_ignoradas_sem_data} notas ignoradas por data_emissao vazia no Protheus."
                )
            if notas_ignoradas_regra_data:
                self.stdout.write(
                    f"{notas_ignoradas_regra_data} notas com preco_tabela não atualizado pela regra de bloqueio (mês/ano anterior após dia 3)."
                )

            todas_ops_dict = {}
            
            # with pool.connection() as conn:
            #     with conn.cursor() as cursor:
            #         for pedaço_lotes in dividir_em_lotes(list(lotes_unicos), 500):
            #             if not pedaço_lotes:
            #                 continue
                        
            #             placeholders = ','.join(['?'] * len(pedaço_lotes))
            #             query_ops_ajustada = query_ops_base.replace('{in_clause}', placeholders)
                        
            #             cursor.execute(query_ops_ajustada, pedaço_lotes)
            #             rows_ops = cursor.fetchall()
                        
            #             for op_row in rows_ops:
            #                 filial, produto, armazem, tp_movimento, descricao_tm, descr_prod, unidade, quantidade, quant_2, custo, custo_2, ord_producao, lote_op, os_ass_tecn, grupo, descricao_grupo, tipo_re_de, ext_texto, documento, dt_emissao, c_contabil, descricao_da_conta, centro_custo, desc_centro_de_custo, parc_total, estornado, sequencial, tipo, usuario, nr_s_a, item_s_a = op_row
                            
            #                 lote_limpo = lote_op.strip()
            #                 # print(lote_limpo)
            #                 base_chave_nota = relacao_lote_nota.get(lote_limpo)
                            
            #                 # ID ÚNICO da OP
            #                 id_op_unico = f"{base_chave_nota}_{ord_producao.strip()}_{sequencial.strip()}"
                            
            #                 todas_ops_dict[id_op_unico] = OP(
            #                     id_op=id_op_unico, filial=filial, produto=produto, armazem=armazem,
            #                     tp_movimento=tp_movimento, descricao_tm=descricao_tm, descr_prod=descr_prod,
            #                     unidade=unidade, quantidade=quantidade, quant_2=quant_2, custo=custo,
            #                     custo_2=custo_2, ord_producao=ord_producao, lote=lote_op, os_ass_tecn=os_ass_tecn,
            #                     grupo=grupo, descricao_grupo=descricao_grupo, tipo_re_de=tipo_re_de,
            #                     ext_texto=ext_texto, documento=documento, dt_emissao=dt_emissao, c_contabil=c_contabil,
            #                     descricao_da_conta=descricao_da_conta, centro_custo=centro_custo, desc_centro_de_custo=desc_centro_de_custo,
            #                     parc_total=parc_total, estornado=estornado, sequencial=sequencial, tipo=tipo,
            #                     usuario=usuario, nr_s_a=nr_s_a, item_s_a=item_s_a
            #                 )

            # # 2. Agora fazemos a separação (Create vs Update)
            # if todas_ops_dict:
            #     self.stdout.write(f"Processando {len(todas_ops_dict)} OPs únicas...")
                
            #     # Busca no Django quais desses IDs já existem
            #     chaves_op_protheus = list(todas_ops_dict.keys())
            #     ids_op_existentes = set(OP.objects.filter(id_op__in=chaves_op_protheus).values_list('id_op', flat=True))
                
            #     ops_para_criar = []
            #     ops_para_atualizar = []
                
            #     for id_op, op_obj in todas_ops_dict.items():
            #         if id_op in ids_op_existentes:
            #             ops_para_atualizar.append(op_obj)
            #         else:
            #             ops_para_criar.append(op_obj)
                
            #     # 3. Salva no banco de forma otimizada
            #     with transaction.atomic():
            #         if ops_para_criar:
            #             OP.objects.bulk_create(ops_para_criar, batch_size=1000)
                        
            #         if ops_para_atualizar:
            #             # IMPORTANTE: Coloque aqui os campos que podem mudar caso a OP sofra alteração no Protheus
            #             campos_update_op = ['quantidade', 'custo', 'custo_2', 'estornado'] 
            #             OP.objects.bulk_update(ops_para_atualizar, campos_update_op, batch_size=1000)

            # 6. REGRA PARA VALIDAR EXCLUSÕES
            Nota.objects.exclude(chave__in=chaves_protheus).update(delete=True)

            fim = time.time()
            self.stdout.write(self.style.SUCCESS(f'Sincronização concluída! Tempo: {fim - inicio:.2f} segundos'))

        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Erro fatal: {str(e)}'))
            import traceback
            traceback.print_exc()

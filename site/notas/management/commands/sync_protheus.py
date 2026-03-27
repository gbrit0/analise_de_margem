import os
import pyodbc
from dotenv import load_dotenv
from django.db import transaction
from django.db.utils import IntegrityError, OperationalError
from notas.models import Nota, Custo, Margem, OP
from django.core.management.base import BaseCommand
from users.models import CustomUser
from dbutils.pooled_db import PooledDB

from setup import settings

from concurrent.futures import ThreadPoolExecutor, as_completed

import time 

# Cria o pool com controle exato no seu código Python
pool = PooledDB(
    creator=pyodbc,         # Módulo usado para conectar
    maxconnections=30,      # Limite MÁXIMO de conexões simultâneas (o que você queria)
    mincached=5,            # Mantém pelo menos 5 conexões abertas esperando trabalho
    blocking=True,          # Se bater 30 conexões, a thread 31 espera até uma liberar
    driver='{ODBC Driver 17 for SQL Server}',
    server=f'{os.getenv("PROTHEUS_DB_HOST")}',
    database=f'{os.getenv("PROTHEUS_DB_DATABASE")}',
    uid=f'{os.getenv("PROTHEUS_DB_USER")}',
    pwd=f'{os.getenv("PROTHEUS_DB_PASSWORD")}'
)

def carrega_ops(lote, nota, item, filial_nota):
    """Essa função recebe  o lote, nota fiscal, item da nota fiscal e a filial. A função deve carregar as OPs no banco de dados, relacionando-as com a nota fiscal correspondente.

    Args:
        lote (str): O lote da OP
        nota (str): A nota fiscal que vendeu o lote
        item (str): O item da nota fiscal que corresponde ao lote
        filial_nota (str): A filial da nota fiscal 
    """
    # print(f"Lote: '{lote}' - Nota: '{nota}' - Item: {item} - Filial: {filial_nota}")
    try:
        with pool.connection() as con:
            with con.cursor() as cursor:
                with open(f'{settings.BASE_DIR}/notas/management/commands/querys/buscaOPs.sql', 'r') as f:
                    query = f.read()
                
                cursor.execute(query, lote)
                
                ops = cursor.fetchall()
    except Exception as e:
        raise e
    
    tentativas = 3
    while tentativas > 0:
        try:
            for op in ops:
                filial, produto, armazem, tp_movimento, descricao_tm, descr_prod, unidade, quantidade, quant_2, custo, custo_2, ord_producao, lote, os_ass_tecn, grupo, descricao_grupo, tipo_re_de, ext_texto, documento, dt_emissao, c_contabil, descricao_da_conta, centro_custo, desc_centro_de_custo, parc_total, estornado, sequencial, tipo, usuario, nr_s_a, item_s_a = op
                
                op, created = OP.objects.update_or_create(
                    id_op=f'{filial_nota}{nota}{item}',
                    defaults={
                        'filial': filial,
                        'produto': produto,
                        'armazem': armazem,
                        'tp_movimento': tp_movimento,
                        'descricao_tm': descricao_tm,
                        'descr_prod': descr_prod,
                        'unidade': unidade,
                        'quantidade': quantidade,
                        'quant_2': quant_2,
                        'custo': custo,
                        'custo_2': custo_2,
                        'ord_producao': ord_producao,
                        'lote': lote,
                        'os_ass_tecn': os_ass_tecn,
                        'grupo': grupo,
                        'descricao_grupo': descricao_grupo,
                        'tipo_re_de': tipo_re_de,
                        'ext_texto': ext_texto,
                        'documento': documento,
                        'dt_emissao': dt_emissao,
                        'c_contabil': c_contabil,
                        'descricao_da_conta': descricao_da_conta,
                        'centro_custo': centro_custo,
                        'desc_centro_de_custo': desc_centro_de_custo,
                        'parc_total': parc_total,
                        'estornado': estornado,
                        'sequencial': sequencial,
                        'tipo': tipo,
                        'usuario': usuario,
                        'nr_s_a': nr_s_a,
                        'item_s_a': item_s_a
                    }
                )
                
                # print(created)
        
        except OperationalError as oe:
            if oe.args[0] == 1205:
                transaction.rollback()
                tentativas -= 1
                time.sleep(1)
                                        
        except Exception as e:
            # RETORNA O ERRO para o processo pai
            import traceback
            return {
                "sucesso": False,
                "lote": lote,
                "mensagem": f"Erro ao inserir/atualizar OPs relacionadas ao lote {lote}.",
                "traceback": traceback.format_exc()
            }
            
        else:
            # RETORNA O SUCESSO para o processo pai
            return {
                "sucesso": True,
                "lote": lote,
                "mensagem": f"OPs relacionadas ao lote {lote} inseridas/atualizadas com sucesso."
            }

class Command(BaseCommand):
    help = 'Sincroniza dados do Django com o Protheus'
    inicio = time.time()
    def handle(self, *args, **options):
        self.stdout.write("Iniciando sincronização...")
        
        try:
            with pool.connection() as conn:
                with conn.cursor() as cursor:
                    with open(f'{settings.BASE_DIR}/notas/management/commands/querys/queryNotas.sql', 'r') as f:
                        query = f.read()
                    
                    cursor.execute(query)
                    rows = cursor.fetchall()
                

            self.stdout.write(f"{len(rows)} registros encontrados. Processando...")            
            
            with ThreadPoolExecutor(max_workers=10) as executor:
                
                futuros_enviados = {}
                
                for row in rows:
                    chave, filial, nome_filial, nota, item, no_pedido, vendedor, data_emissao, lote, cfop, cfop_descri, atualiza_estoque, gera_duplicata, cod_produto, produto, tipo_produto, desc_tipo_produto, armazem, cod_cliente, loja, cliente, grp_amar_ctb, classificacao_produto, estado_destino, quantidade, tabela_preco, preco_tabela, valor_contabil, custo, valor_unitario, valor_ipi, valor_imp5, valor_imp6, valor_icms_difal, valor_icms, aliq_icms = row
                        
                    tentativas = 3
                    
                    # while tentativas > 0: 
                    try:
                        with transaction.atomic():
                            obj, created = Nota.objects.update_or_create(
                                chave=chave, # chave - campo único
                                defaults={
                                    'chave': chave,
                                    'filial': filial,
                                    'nome_filial': nome_filial,
                                    'nota': nota,
                                    'item': item,
                                    'no_pedido': no_pedido,
                                    'vendedor': vendedor,
                                    'data_emissao': data_emissao,
                                    'lote': lote,
                                    'cfop': cfop,
                                    'cfop_descri': cfop_descri,
                                    'atualiza_estoque': atualiza_estoque,
                                    'gera_duplicata': gera_duplicata,
                                    'cod_produto': cod_produto,
                                    'produto': produto,
                                    'tipo_produto': tipo_produto,
                                    'desc_tipo_produto': desc_tipo_produto,
                                    'armazem': armazem,
                                    'cod_cliente': cod_cliente,
                                    'loja': loja,
                                    'cliente': cliente,
                                    'grp_amar_ctb': grp_amar_ctb,
                                    'classificacao_produto': classificacao_produto,
                                    'estado_destino': estado_destino,
                                    'quantidade': quantidade,
                                    'tabela_preco': tabela_preco,
                                    'preco_tabela': preco_tabela,
                                    'valor_contabil': valor_contabil,
                                    'valor_unitario': valor_unitario,
                                    'valor_ipi': valor_ipi,
                                    'valor_imp5': valor_imp5,
                                    'valor_imp6': valor_imp6,
                                    'valor_icms_difal': valor_icms_difal,
                                    'valor_icms': valor_icms,
                                    'aliq_icms': aliq_icms,
                                },
                            )
                        
                            
                        if created:
                            # calcular custo e margem somente se for uma nova Nota
                                
                            with transaction.atomic():
                                custo, _ = Custo.objects.update_or_create(
                                    chave=obj,
                                    defaults = {
                                        'valor': custo,
                                        'usuario': CustomUser.objects.get(id=2), # ID do usuário system        
                                        }
                                )
                                
                            custo_valor = custo.valor                                  
                            
                            cfops_especiais = ['5101', '6101', '5116', '6116', '6107']

                            from datetime import date

                            data_corte = date(2026, 2, 1)
                            
                            if data_emissao >= data_corte:
                                pro_goias = 0.0477
                            else:
                                pro_goias = 0.02
                            
                            icms_calculado = valor_icms * pro_goias if cfop in cfops_especiais else valor_icms
                                                        
                            margem_bruta = valor_contabil - custo_valor - valor_ipi - valor_imp5 - valor_imp6 - valor_icms_difal - icms_calculado
                            
                            with transaction.atomic():
                                margem = Margem.objects.update_or_create(
                                    chave=obj,
                                    custo=custo,
                                    margem_bruta=margem_bruta,
                                    margem_bruta_percentual=margem_bruta/valor_contabil
                                )
                                    
                    # except IntegrityError:
                    except OperationalError as oe:
                        if oe.args[0] == 1205:
                            transaction.rollback()
                            tentativas -= 1
                            time.sleep(1)
                        else:
                            raise oe
                    
                    except Exception as e:
                        # self.stdout.write(self.style.ERROR(f"Erro de integridade para a chave {row[0]}. Verifique os dados.")) # Não é erro porquê a consulta não pode filtrar mais a última nota inserida por conta da necessidade de pagar notas de várias filiais.
                        # continue
                        raise e

                    
                    # if lote and lote.strip():
                    #     carrega_ops(lote, nota, item, filial)
                    
                    if lote and lote.strip():
                        futuro = executor.submit(carrega_ops, lote, nota, item, filial)
                        futuros_enviados[futuro] = lote
                        
                for futuro in as_completed(futuros_enviados):
                    lote_processado = futuros_enviados[futuro]
                    
                    try:
                        # Captura o retorno (o dicionário que configuramos na resposta anterior)
                        resultado = futuro.result()
                        
                        # Imprime no terminal do processo pai de forma segura
                        if resultado["sucesso"]:
                            self.stdout.write(self.style.SUCCESS(resultado["mensagem"]))
                            # self.stdout.write(self.style.SUCCESS(f"OPs relacionadas ao lote '{lote}' inseridas/atualizadas com sucesso."))
                        else:
                            self.stdout.write(self.style.ERROR(resultado["mensagem"]))
                            if "traceback" in resultado:
                                self.stdout.write(resultado["traceback"])
                                
                    except Exception as exc:
                        # Captura erros não tratados que fizeram a thread explodir
                        self.stdout.write(self.style.ERROR(f"Erro fatal na thread ao processar o lote {lote_processado}: {exc}"))
                
            # REGRA PARA VALIDAR EXCLUSÕES
            ids_externos = [row[0] for row in rows]
            
            Nota.objects.exclude(chave__in=ids_externos).update(delete=True)
                            
            self.stdout.write(self.style.SUCCESS('Sincronização concluída com sucesso!'))
            fim = time.time()
            self.stdout.write(self.style.SUCCESS(f'Tempo total de execução: {fim - inicio:.2f} segundos'))
            
        except Exception as e:
            raise e
            self.stdout.write(self.style.ERROR(f'Erro na sincronização: {e}'))
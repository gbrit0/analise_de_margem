import os
import pyodbc
from dotenv import load_dotenv
from django.db import transaction
from django.db.utils import IntegrityError
from notas.models import Nota, Custo, Margem
from django.core.management.base import BaseCommand
from users.models import CustomUser

class Command(BaseCommand):
    help = 'Sincroniza dados do Django com o Protheus'
    
    def handle(self, *args, **options):
        self.stdout.write("Iniciando sincronização...")
        
        conn_str = f"DRIVER={os.getenv('PROTHEUS_ODBC_DRIVER')};SERVER={os.getenv('PROTHEUS_DB_HOST')};DATABASE={os.getenv('PROTHEUS_DB_DATABASE')};UID={os.getenv('PROTHEUS_DB_USER')};PWD={os.getenv('PROTHEUS_DB_PASSWORD')};TrustServerCertificate=yes"
        
        try:
            with pyodbc.connect(conn_str) as conn:
                cursor = conn.cursor()
                
                query = f"""
                    SELECT DISTINCT
                        TRIM(D2_FILIAL) + TRIM(D2_DOC) + TRIM(D2_SERIE) + TRIM(D2_CLIENTE) + TRIM(D2_LOJA) + TRIM(D2_ITEM) AS [chave],
                        TRIM(D2_FILIAL) AS [filial],
                        TRIM(M0_FILIAL) AS [nome_filial],
                        LTRIM(D2_DOC, 0) AS [nota],
                        D2_PEDIDO AS [no_pedido],
                        TRIM(A3_NOME) AS [vendedor],
                        CAST(F2_EMISSAO AS DATE) AS [data_emissao],
                        TRIM(D2_LOTECTL) AS [lote],
                        TRIM(D2_CF) AS [cfop],
                        TRIM(X5_CF.X5_DESCRI) AS [cfop_descri],
                        TRIM(F4_ESTOQUE) AS [atualiza_estoque],
                        TRIM(F4_DUPLIC) AS [gera_duplicata],
                        TRIM(B1_COD) AS [cod_produto],
                        TRIM(SUBSTRING(B1_DESC, 1, 50)) AS [produto],
                        TRIM(B1_TIPO) AS [tipo_produto],
                        TRIM(X5.X5_DESCRI) AS [desc_tipo_produto],
                        TRIM(D2_LOCAL) + ' - ' + TRIM(NNR_DESCRI) AS [armazem],
                        TRIM(F2_CLIENTE) AS [cod_cliente],
                        TRIM(F2_LOJA) AS [loja],
                        TRIM(A1_NOME) AS [cliente],
                        LTRIM(B1_XGRPCTB, 0) AS [grp_amar_ctb],
                        TRIM(ZC2_DESCR) AS [classificacao_produto],
                        TRIM(D2_EST) AS [estado_destino],
                        D2_QUANT AS [quantidade],
                        CASE 
                            WHEN A1_TABELA IS NULL THEN '-'
                            ELSE A1_TABELA
                        END AS [tabela_preco],
                        CASE WHEN 
                            DA1.DA1_PRCVEN IS NULL THEN 0
                            ELSE DA1.DA1_PRCVEN * D2_QUANT 
                        END AS [preco_tabela],
                        D2_VALBRUT AS [valor_contabil],
                        D2_CUSTO1 AS [custo],
                        D2_PRCVEN AS [valor_unitario],
                        D2_VALIPI AS [valor_ipi],
                        D2_VALIMP5 AS [valor_imp5],
                        D2_VALIMP6 AS [valor_imp6],
                        D2_DIFAL AS [vlr_icms_difal],
                        D2_VALICM AS [valor_icms],
                        D2_PICM AS [aliq_icms]

                    FROM SD2010 AS D2 -- Itens de Venda da NF

                        INNER JOIN SF4010 F4 ON -- Tipos de Entrada e Saída
                            F4.D_E_L_E_T_ <> '*'
                            AND F4_CODIGO = D2_TES
                            AND F4_FILIAL = D2_FILIAL
                            AND F4_TEXTO LIKE 'VENDA%'

                        LEFT JOIN SB1010 B1 ON -- Cadastro de Produtos
                            B1.D_E_L_E_T_ <> '*'
                            AND B1_FILIAL = SUBSTRING(D2_FILIAL, 1, 2)
                            AND B1_COD = D2_COD

                        LEFT JOIN SF2010 AS F2 ON --  Cabeçalho das NF de Saída
                            F2.D_E_L_E_T_ <> '*'
                            AND TRIM(F2_DOC) = TRIM(D2_DOC)
                            AND TRIM(F2_SERIE) = TRIM(D2_SERIE)
                            AND TRIM(F2_CLIENTE) = TRIM(D2_CLIENTE)
                            AND TRIM(F2_LOJA) = TRIM(D2_LOJA)
                            AND TRIM(F2_FILIAL) = TRIM(D2_FILIAL)

                        LEFT JOIN SA1010 AS A1 ON -- Clientes
                            A1.D_E_L_E_T_ <> '*'
                            AND TRIM(A1_COD) = TRIM(F2_CLIENTE)
                            AND TRIM(A1_LOJA) = TRIM(F2_LOJA)
                            AND A1_FILIAL = SUBSTRING(F2_FILIAL, 1, 2)

                        LEFT JOIN DA1010 AS DA1 ON -- Itens da Tabela de preço
                            DA1.DA1_CODTAB  = A1_TABELA
                            AND DA1.DA1_CODPRO = B1.B1_COD
                            AND DA1.DA1_FILIAL = D2_FILIAL
                            AND DA1.D_E_L_E_T_ <> '*'
                        
                        LEFT JOIN SA3010 AS A3 ON -- Vendedores
                            A3.D_E_L_E_T_ <> '*'
                            AND A3_COD = F2_VEND1
                            AND A3_FILIAL = F2_FILIAL

                        -- Junção com a tabela ZC2 para obter a descrição do grupo de amarração contábil
                        LEFT JOIN ZC2010 AS ZC2 ON 
                            ZC2.D_E_L_E_T_ <> '*'
                            AND ZC2_GRP = B1_XGRPCTB
                            AND ZC2_FILIAL = B1_FILIAL

                        -- SX5 (Tabelas)	SD2 (Itens de Venda da NF)	X5_TABELA + X5_CHAVE	'13' + D2_CF
                        LEFT JOIN SX5010 AS X5_CF ON -- SX5 (Tabelas)	
                            X5_CF.D_E_L_E_T_ <> '*'
                            AND X5_CF.X5_TABELA = '13'
                            AND X5_CF.X5_FILIAL = D2_FILIAL
                            AND X5_CF.X5_CHAVE = D2_CF
                        
                        LEFT JOIN NNR010 ON -- NNR (Locais de Estoque)	
                            NNR010.D_E_L_E_T_ <> '*'
                            AND D2_FILIAL = NNR_FILIAL 
                            AND D2_LOCAL = NNR_CODIGO

                        LEFT JOIN SYS_COMPANY AS C ON -- TABELA DE FILIAIS
                            C.D_E_L_E_T_ <> '*'
                            AND C.M0_CODFIL = D2_FILIAL
                            AND TRIM(M0_NOME) = 'BRG Geradores'

                        LEFT JOIN SX5010 AS X5 ON  -- SX5 (Tabelas)	
                            X5.D_E_L_E_T_ <> '*'
                            AND X5.X5_TABELA = '02'
                            AND X5.X5_CHAVE = B1_TIPO
                            AND SUBSTRING(X5.X5_FILIAL, 1, 2) = B1_FILIAL

                    WHERE
                        D2.D_E_L_E_T_ <> '*'
                        AND D2_EMISSAO >= 20250901 -- AND 20250930
                        AND TRIM(D2_FILIAL) IN ('0101', '0501', '0502', '0503') -- BRG MATRIZ, GRID GO, GRID MG E GRID PA
                        AND TRIM(B1.B1_COD) NOT IN ('B0010046', 'E000H2P8') 
                        """
                
                cursor.execute(query)
                rows = cursor.fetchall()
                
                self.stdout.write(f"{len(rows)} registros encontrados. Processando...")
                
                with transaction.atomic():
                    for row in rows:
                        chave, filial, nome_filial, nota, no_pedido, vendedor, data_emissao, lote, cfop, cfop_descri, atualiza_estoque, gera_duplicata, cod_produto, produto, tipo_produto, desc_tipo_produto, armazem, cod_cliente, loja, cliente, grp_amar_ctb, classificacao_produto, estado_destino, quantidade, tabela_preco, preco_tabela, valor_contabil, custo, valor_unitario, valor_ipi, valor_imp5, valor_imp6, valor_icms_difal, valor_icms, aliq_icms = row
                        # print(
                        #     f'chave: {type(chave)}\n',
                        #     f'filial: {type(filial)}\n',
                        #     f'nome_filial: {type(nome_filial)}\n',
                        #     f'nota: {type(nota)}\n',
                        #     f'no_pedido: {type(no_pedido)}\n',
                        #     f'vendedor: {type(vendedor)}\n',
                        #     f'data_emissao: {type(data_emissao)}\n',
                        #     f'lote: {type(lote)}\n',
                        #     f'cfop: {type(cfop)}\n',
                        #     f'cfop_descri: {type(cfop_descri)}\n',
                        #     f'atualiza_estoque: {type(atualiza_estoque)}\n',
                        #     f'gera_duplicata: {type(gera_duplicata)}\n',
                        #     f'cod_produto: {type(cod_produto)}\n',
                        #     f'produto: {type(produto)}\n',
                        #     f'tipo_produto: {type(tipo_produto)}\n',
                        #     f'desc_tipo_produto: {type(desc_tipo_produto)}\n',
                        #     f'armazem: {type(armazem)}\n',
                        #     f'cod_cliente: {type(cod_cliente)}\n',
                        #     f'loja: {type(loja)}\n',
                        #     f'cliente: {type(cliente)}\n',
                        #     f'grp_amar_ctb: {type(grp_amar_ctb)}\n',
                        #     f'classificacao_produto: {type(classificacao_produto)}\n',
                        #     f'estado_destino: {type(estado_destino)}\n',
                        #     f'quantidade: {type(quantidade)}\n',
                        #     f'tabela_preco: {type(tabela_preco)}\n',
                        #     f'preco_tabela: {type(preco_tabela)}\n',
                        #     f'valor_contabil: {type(valor_contabil)}\n',
                        #     f'valor_unitario: {type(valor_unitario)}\n',
                        #     f'valor_ipi: {type(valor_ipi)}\n',
                        #     f'valor_imp5: {type(valor_imp5)}\n',
                        #     f'valor_imp6: {type(valor_imp6)}\n',
                        #     f'valor_icms_difal: {type(valor_icms_difal)}\n',
                        #     f'valor_icms: {type(valor_icms)}\n',
                        #     f'aliq_icms: {type(aliq_icms)}\n',
                        # )
                        try:
                            obj, created = Nota.objects.update_or_create(
                                chave=chave, # chave - campo único
                                defaults={
                                    'chave': chave,
                                    'filial': filial,
                                    'nome_filial': nome_filial,
                                    'nota': nota,
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
                        except IntegrityError:
                            # self.stdout.write(self.style.ERROR(f"Erro de integridade para a chave {row[0]}. Verifique os dados.")) # Não é erro porquê a consulta não pode filtrar mais a última nota inserida por conta da necessidade de pagar notas de várias filiais.
                            continue
                            
                        if created:
                            # calcular custo e margem somente se for uma nova Nota
                            
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
                            
                            margem = Margem.objects.update_or_create(
                                chave=obj,
                                custo=custo,
                                margem_bruta=margem_bruta,
                                margem_bruta_percentual=margem_bruta/valor_contabil
                            )

                    # REGRA PARA VALIDAR EXCLUSÕES
                    ids_externos = [row[0] for row in rows]
                    
                    Nota.objects.exclude(chave__in=ids_externos).update(delete=True)
                     
            self.stdout.write(self.style.SUCCESS('Sincronização concluída com sucesso!'))
        except Exception as e:
            raise e
            self.stdout.write(self.style.ERROR(f'Erro na sincronização: {e}'))
from django.core.management.base import BaseCommand
from django.db import transaction
from notas.models import Nota, Custo, Margem
import pyodbc
from dotenv import load_dotenv
import os

class Command(BaseCommand):
    help = 'Sincroniza dados do Django com o Protheus'
    
    def handle(self, *args, **options):
        self.stdout.write("Iniciando sincronização...")
        
        conn_str = f"DRIVER={os.getenv('PROTHEUS_ODBC_DRIVER')};SERVER={os.getenv('PROTHEUS_DB_HOST')};DATABASE={os.getenv('PROTHEUS_DB_DATABASE')};UID={os.getenv('PROTHEUS_DB_USER')};PWD={os.getenv('PROTHEUS_DB_PASSWORD')};TrustServerCertificate=yes"
        
        try:
            with pyodbc.connect(conn_str) as conn:
                cursor = conn.cursor()
                
                query = f"""
                SELECT
                    TRIM(D2_FILIAL) AS [filial],
                    TRIM(D2_FILIAL) + TRIM(D2_DOC) + TRIM(D2_SERIE) + TRIM(D2_CLIENTE) + TRIM(D2_LOJA) + TRIM(D2_ITEM) AS [chave],
                    LTRIM(D2_DOC, 0) AS [nota],
                    D2_PEDIDO AS [no_pedido],
                    TRIM(A3_NOME) AS [vendedor],
                    CAST(F2_EMISSAO AS DATE) AS [data_emissao],
                    TRIM(D2_LOTECTL) AS [lote],
                    TRIM(D2_CF) AS [cfop],
                    TRIM(X5_DESCRI) AS [cfop_descri],
                    TRIM(F4_ESTOQUE) AS [atualiza_estoque],
                    TRIM(F4_DUPLIC) AS [gera_duplicata],
                    TRIM(B1_COD) AS [cod_produto],
                    TRIM(SUBSTRING(B1_DESC, 1, 50)) AS [produto],
                    TRIM(B1_TIPO) AS [tipo_produto],
                    TRIM(D2_LOCAL) + ' - ' + TRIM(NNR_DESCRI) AS [armazem],
                    TRIM(F2_CLIENTE) AS [cod_cliente],
                    TRIM(F2_LOJA) AS [loja],
                    TRIM(A1_NOME) AS [cliente],
                    LTRIM(B1_XGRPCTB, 0) AS [grp_amar_ctb],
                    TRIM(ZC2_DESCR) AS [classificacao_produto],
                    TRIM(D2_EST) AS [estado_destino],
                    D2_QUANT AS [quantidade],
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
                    LEFT JOIN SX5010 ON
                        SX5010.D_E_L_E_T_ <> '*'
                        AND X5_TABELA = '13'
                        AND X5_FILIAL = D2_FILIAL
                        AND X5_CHAVE = D2_CF

                    -- NNR (Locais de Estoque)	SD2 (Itens de Venda da NF)	NNR_CODIGO	D2_LOCAL
                    LEFT JOIN NNR010 ON
                        NNR010.D_E_L_E_T_ <> '*'
                        AND D2_FILIAL = NNR_FILIAL 
                        AND D2_LOCAL = NNR_CODIGO

                WHERE
                    D2.D_E_L_E_T_ <> '*'
                    AND D2_EMISSAO >= 20250901 -- AND 20250930
                    AND D2_FILIAL IN ('0101', '0501', '0502', '0503')
                """
                
                cursor.execute(query)
                rows = cursor.fetchall()
                
                self.stdout.write(f"{len(rows)} registros encontrados. Processando...")
                
                with transaction.atomic():
                    for row in rows:
                        obj, created = Nota.objects.update_or_create(
                            chave=row[1], # chave - campo único
                            defaults={
                                'filial': row[0],
                                'chave': row[1],
                                'nota': row[2],
                                'no_pedido': row[3],
                                'vendedor': row[4],
                                'data_emissao': row[5],
                                'lote': row[6],
                                'cfop': row[7],
                                'cfop_descri': row[8],
                                'atualiza_estoque': row[9],
                                'gera_duplicata': row[10],
                                'cod_produto': row[11],
                                'produto': row[12],
                                'tipo_produto': row[13],
                                'armazem': row[14],
                                'cod_cliente': row[15],
                                'loja': row[16],
                                'cliente': row[17],
                                'grp_amar_ctb': row[18],
                                'classificacao_produto': row[19],
                                'estado_destino': row[20],
                                'quantidade': row[21],
                                'valor_contabil': row[22],
                                # 'custo': row[23],
                                'valor_unitario': row[24],
                                'valor_ipi': row[25],
                                'valor_imp5': row[26],
                                'valor_imp6': row[27],
                                'valor_icms_difal': row[28],
                                'valor_icms': row[29],
                                'aliq_icms': row[30],
                            }
                        )
                        if created:
                            # calcular custo e margem somente se for uma nova Nota
                            
                            custo = Custo.objects.update_or_create(
                                chave=obj,
                                valor=row[23]
                            )
                            
                            
                            cfops_especiais = ['5101', '6101', '5116', '6116', '6107']
                            
                            icms_calculado = row[29] * 0.047 if row[7] in cfops_especiais else row[29]
                            
                            margem_bruta = row[22] - row[23] - row[25] - row[26] - row[27] - row[28] - icms_calculado
                            
                            margem = Margem.objects.update_or_create(
                                chave=obj,
                                custo=custo[0],
                                margem_bruta=margem_bruta,
                                margem_bruta_percentual=margem_bruta/row[22]
                            )
                    
                    # REGRA PARA VALIDAR EXCLUSÕES
                    ids_externos = [row[0] for row in rows]
                    
                    Nota.objects.exclude(chave__in=ids_externos).update(delete=True)
                     
            self.stdout.write(self.style.SUCCESS('Sincronização concluída com sucesso!'))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f'Erro na sincronização: {e}'))
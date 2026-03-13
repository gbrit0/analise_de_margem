import os
import pyodbc
from dotenv import load_dotenv
from django.db import transaction
from django.db.utils import IntegrityError
from notas.models import Nota, Custo, Margem
from django.core.management.base import BaseCommand

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
                        TRIM(D2_FILIAL) + TRIM(D2_DOC) + TRIM(D2_SERIE) + TRIM(D2_CLIENTE) + TRIM(D2_LOJA) + TRIM(D2_ITEM) AS [chave],
                        TRIM(D2_FILIAL) AS [filial],
                        TRIM(M0_FILIAL) AS [nome_filial],
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

                        LEFT JOIN SYS_COMPANY AS C ON
                            C.D_E_L_E_T_ <> '*'
                            AND C.M0_CODFIL = D2_FILIAL
                            AND TRIM(M0_NOME) = 'BRG Geradores'
                    WHERE
                        D2.D_E_L_E_T_ <> '*'
                        AND D2_EMISSAO >= 20250901 -- AND 20250930
                        AND D2_FILIAL IN ('0101', '0501', '0502', '0503')
                        AND TRIM(B1.B1_COD) NOT IN ('B0010046', 'E000H2P8')
                """
                
                cursor.execute(query)
                rows = cursor.fetchall()
                
                self.stdout.write(f"{len(rows)} registros encontrados. Processando...")
                
                with transaction.atomic():
                    for row in rows:
                        try:
                            obj, created = Nota.objects.update_or_create(
                                chave=row[1], # chave - campo único
                                defaults={
                                    'chave': row[0],
                                    'filial': row[1],
                                    'nome_filial': row[2],
                                    'nota': row[3],
                                    'no_pedido': row[4],
                                    'vendedor': row[5],
                                    'data_emissao': row[6],
                                    'lote': row[7],
                                    'cfop': row[8],
                                    'cfop_descri': row[9],
                                    'atualiza_estoque': row[10],
                                    'gera_duplicata': row[11],
                                    'cod_produto': row[12],
                                    'produto': row[13],
                                    'tipo_produto': row[14],
                                    'armazem': row[15],
                                    'cod_cliente': row[16],
                                    'loja': row[17],
                                    'cliente': row[18],
                                    'grp_amar_ctb': row[19],
                                    'classificacao_produto': row[20],
                                    'estado_destino': row[21],
                                    'quantidade': row[22],
                                    'valor_contabil': row[23],
                                    # 'custo': row[23],
                                    'valor_unitario': row[25],
                                    'valor_ipi': row[26],
                                    'valor_imp5': row[27],
                                    'valor_imp6': row[28],
                                    'valor_icms_difal': row[29],
                                    'valor_icms': row[30],
                                    'aliq_icms': row[31],
                                },
                            )
                        except IntegrityError:
                            # self.stdout.write(self.style.ERROR(f"Erro de integridade para a chave {row[0]}. Verifique os dados.")) # Não é erro porquê a consulta não pode filtrar mais a última nota inserida por conta da necessidade de pagar notas de várias filiais.
                            continue
                            
                        if created:
                            # calcular custo e margem somente se for uma nova Nota
                            
                            custo = Custo.objects.update_or_create(
                                chave=obj,
                                valor=row[24]
                            )
                            
                            # print(f'custo: {custo}')                            
                            # print(f'custo[0]: {custo[0]}')                            
                            # print(f'custo[1]: {custo[1]}')                            
                            
                            cfops_especiais = ['5101', '6101', '5116', '6116', '6107']
                            
                            icms_calculado = row[30] * 0.02 if row[8] in cfops_especiais else row[30]
                            
                            margem_bruta = row[23] - row[24] - row[26] - row[27] - row[28] - row[29] - icms_calculado
                            
                            margem = Margem.objects.update_or_create(
                                chave=obj,
                                custo=custo[0],
                                margem_bruta=margem_bruta,
                                margem_bruta_percentual=margem_bruta/row[23]
                            )
                    
                    # REGRA PARA VALIDAR EXCLUSÕES
                    ids_externos = [row[0] for row in rows]
                    
                    Nota.objects.exclude(chave__in=ids_externos).update(delete=True)
                     
            self.stdout.write(self.style.SUCCESS('Sincronização concluída com sucesso!'))
        except Exception as e:
            raise e
            self.stdout.write(self.style.ERROR(f'Erro na sincronização: {e}'))
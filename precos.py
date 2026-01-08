# ==========================================================
#    Regra de precificação BRG implementada pelo Coutinho
# ==========================================================
def aplicar_formula(valor_extraido, cnpj, cliente_parceiro, filial_entrada):
    """
    Aplica a fórmula de cálculo do preço de venda conforme:
    - Filial de entrada (BRG MATRIZ ou GRID)
    - Cnpj (se for None, é venda dentro do estado)
    - Venda dentro ou fora do estado
    """

    valor = float(valor_extraido)
    # valor = float(valor_extraido.replace(',', '.'))
    
    cliente_parceiro = cliente_parceiro.strip().capitalize()

    filial_entrada = filial_entrada[:4]
    # Identifica o grupo (BRG MATRIZ ou GRID)
    if filial_entrada == "0101":
        grupo = "BRG MATRIZ"
    elif filial_entrada in ("0501", "0502", "0503"):
        grupo = "GRID"
    else:
        raise ValueError(f"Filial {filial_entrada} não reconhecida para cálculo.")

    # Define alíquotas conforme grupo e cenário
    if grupo == "GRID":
        # GRID GO, MG, PA
        if cliente_parceiro == "Sim" and cnpj is None: # venda_estado == "Dentro": # o cliente é contribuinte
            margem = 15
            icms = 18 if filial_entrada == "0502" else 19
            pis = 0.65
            cofins = 3
            outras = 7
        elif cliente_parceiro == "Sim" and cnpj is not None: # and venda_estado == "Fora": # o cliente é não contribuinte 
            margem = 15
            icms = 12
            pis = 0.65
            cofins = 3
            outras = 7
        elif cliente_parceiro == "Não" and cnpj is None:# and venda_estado == "Dentro": # o cliente é contribuinte
            margem = 27
            icms = 18 if filial_entrada == "0502" else 19
            pis = 0.65
            cofins = 3
            outras = 7
        elif cliente_parceiro == "Não" and cnpj is not None: # and venda_estado == "Fora": # o cliente não é contribuinte 
            margem = 27
            icms = 12
            pis = 0.65
            cofins = 3
            outras = 7
        else:
            raise ValueError(f"Cenário inválido: Parceiro={cliente_parceiro}")

    elif grupo == "BRG MATRIZ":
        # BRG MATRIZ
        if cliente_parceiro == "Sim" and cnpj is None: # venda_estado == "Dentro":
            margem = 15
            icms = 19
            pis = 1.65
            cofins = 7.60
            outras = 7
        elif cliente_parceiro == "Sim" and cnpj is not None: # and venda_estado == "Fora":
            margem = 15
            icms = 12
            pis = 1.65
            cofins = 7.60
            outras = 7
        elif cliente_parceiro == "Não" and cnpj is None: # and venda_estado == "Dentro":
            margem = 27
            icms = 19
            pis = 1.65
            cofins = 7.60
            outras = 7
        elif cliente_parceiro == "Não" and cnpj is not None: # and venda_estado == "Fora":
            margem = 27
            icms = 12
            pis = 1.65
            cofins = 7.60
            outras = 7
        else:
            raise ValueError(f"Cenário inválido: Parceiro={cliente_parceiro}")

    else:
        raise ValueError(f"Grupo {grupo} não mapeado.")

    # Cálculo principal
    divisor = 1 - ((margem + icms + pis + cofins + outras) / 100)
    preco_venda = valor / divisor


    print(f"""💰 Total final: R$ {preco_venda:.2f}""")

    return preco_venda

import pandas as pd

# 1. Configurações de busca (Critérios)
ARQUIVO = '2026.xlsx'
ABA = 'Janeiro'

criterio_DBA = "Vendedor"
criterio_FBA = "Amazon"
criterio_mes = input("digite o mes ")
criterio_reembolso = "Reembolso"
criterio_pacotes_extraviados = "Reembolso para pacotes extraviados"
criterio_tarifa_de_manuseio = "Tarifa de manuseio com base no peso Delivery by Amazon"
criterio_extorno_do_frete = "Estorno do frete pago pelo comprador"
criterio_tarifa_de_devolucao = "Tarifa de devolução de inventário pela Amazon"
criterio_tarifa_de_armazenagem = "Tarifa de armazenagem do programa FBA - Logística da Amazon"
criterio_frete_da_transportadora = "Frete da transportadora parceira da Amazon de FBA"
criterio_custo_de_publi = "Custo de Publicidade"
criterio_cadastro = "Cadastro"
criterio_transferir = "Transferir"
criterio_debito = "Débito"
criterio_ano = int(input("digite o ano ")) # Se no Excel for número, use sem aspas. Se for texto, use "2026"

# 2. Carregando os dados
try:
    df_dados = pd.read_excel(ARQUIVO, sheet_name=ABA)

    # Limpeza preventiva: remove espaços em branco no início/fim das células de texto
    # Isso evita que "Amazon " seja diferente de "Amazon"
    colunas_texto = ['atendimento', 'Nome do Mês', 'Ano']
    for col in colunas_texto:
        if col in df_dados.columns:
            df_dados[col] = df_dados[col].astype(str).str.strip()

    # 3. Tratamento da coluna de valores (N)
    # Se a coluna N vier como texto (ex: "R$ 10,50"), convertemos para número
    if df_dados['vendas do produto'].dtype == 'object':
        df_dados['vendas do produto'] = (
            df_dados['vendas do produto']
            .replace(r'[R\$\s.]', '', regex=True)  # Remove R$, espaços e pontos de milhar
            .replace(',', '.')  # Troca vírgula decimal por ponto
        )
    df_dados['vendas do produto'] = pd.to_numeric(df_dados['vendas do produto'], errors='coerce').fillna(0)

    df_dados["créditos de remessa"] = pd.to_numeric(
        df_dados["créditos de remessa"], errors="coerce"
    ).fillna(0)

    df_dados["créditos de embalagem de presente"] = pd.to_numeric(
        df_dados["créditos de embalagem de presente"], errors="coerce"
    ).fillna(0)

    df_dados["tarifas de venda"] = pd.to_numeric(
        df_dados["tarifas de venda"], errors="coerce"
    ).fillna(0)

    df_dados["outro"] = pd.to_numeric(
        df_dados["outro"], errors="coerce"
    ).fillna(0)

    df_dados["taxas de outras transações"] = pd.to_numeric(
        df_dados["taxas de outras transações"], errors="coerce"
    ).fillna(0)

    df_dados["descontos promocionais"] = pd.to_numeric(
        df_dados["descontos promocionais"], errors="coerce"
    ).fillna(0)

    df_dados["imposto de vendas coletados"] = pd.to_numeric(
        df_dados["imposto de vendas coletados"], errors="coerce"
    ).fillna(0)

    df_dados["créditos de remessa"] = pd.to_numeric(
        df_dados["créditos de remessa"], errors="coerce"
    ).fillna(0)

    # 4. Aplicação do Filtro (Equivalente ao SUMIFS)
    filtro = (
            (df_dados['atendimento'] == criterio_DBA) &
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano))  # Ajuste conforme o tipo de dado na sua planilha
    )

    filtro2 = (
            (df_dados['atendimento'] == criterio_FBA) &
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano))  # Ajuste conforme o tipo de dado na sua planilha
    )

    filtro3 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano))
    )

    filtro4 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['tipo'] == criterio_reembolso)
    )

    filtro5 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_pacotes_extraviados)
    )
    filtro6 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_tarifa_de_manuseio)
    )
    filtro7 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_extorno_do_frete)
    )

    filtro8 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_tarifa_de_devolucao)
    )

    filtro9 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_tarifa_de_armazenagem)
    )

    filtro10 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_frete_da_transportadora)
    )

    filtro11 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_custo_de_publi)
    )

    filtro12 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['descrição'] == criterio_cadastro)
    )

    filtro13 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['tipo'] == criterio_transferir)
    )

    filtro14 = (
            (df_dados['Nome do Mês'] == criterio_mes) &
            (df_dados['Ano'] == str(criterio_ano)) &
            (df_dados['tipo'] == criterio_debito)
    )




    # 5. Cálculo do Resultado
    resultado = df_dados.loc[filtro, 'vendas do produto'].sum()
    resultado2 = df_dados.loc[filtro2, 'vendas do produto'].sum()
    resultado3 = df_dados.loc[filtro3, "créditos de remessa"].sum()
    resultado4 = df_dados.loc[filtro3, "créditos de embalagem de presente"].sum()
    resultado5 = df_dados.loc[filtro4, "tarifas de venda"].sum()
    resultado6 = df_dados.loc[filtro5, "outro"].sum()
    resultado7 = df_dados.loc[filtro, "tarifas de venda"].sum()
    resultado8 = df_dados.loc[filtro2, "tarifas de venda"].sum()
    resultado9 = df_dados.loc[filtro6, "taxas de outras transações"].sum()
    resultado10 = df_dados.loc[filtro7, "taxas de outras transações"].sum()
    resultado11 = df_dados.loc[filtro3, "descontos promocionais"].sum()
    resultado12 = df_dados["imposto de vendas coletados"].sum()
    resultado13 = df_dados.loc[filtro4, "créditos de remessa"].sum()
    resultado14 = df_dados.loc[filtro4, "vendas do produto"].sum()
    resultado15 = df_dados.loc[filtro8, "outro"].sum()
    resultado16 = df_dados.loc[filtro9, "outro"].sum()
    resultado17 = df_dados.loc[filtro10, "outro"].sum()
    resultado18 = df_dados.loc[filtro11, "taxas de outras transações"].sum()
    resultado19 = df_dados.loc[filtro12, "outro"].sum()
    resultado20 = df_dados.loc[filtro13, "outro"].sum()
    resultado21 = df_dados.loc[filtro14, "outro"].sum()

    total = resultado + resultado2 + resultado3 + resultado4 + resultado5 + resultado6
    total_desconto = resultado7 + resultado8 + resultado9 + resultado10 + resultado11 + resultado12 + resultado13 + resultado14  + resultado15  + resultado16  + resultado17 + resultado18 + resultado19
    resultado_final = total + total_desconto

    print("-" * 30)
    print(f"Relatório: {criterio_DBA}")
    print(f"Período: {criterio_mes}/{criterio_ano}")
    print(f"Total Calculado DBA: R$ {resultado:,.2f}")
    print(f"Total Calculado FBA: R$ {resultado2:,.2f}")
    print("-" * 30)
    print(f"Total Créditos de Remessa: R$ {resultado3:,.2f}")
    print(f"Total créditos de embalagem de presente: R$ {resultado4:,.2f}")
    print(f"tarifas de venda: R$ {resultado5:,.2f}")
    print(f"Reembolso para pacotes extraviados: R$ {resultado6:,.2f}")
    print(f"total: R$ {total:,.2f}")
    print(f"total: R$ {resultado7:,.2f}")
    print(f"total: R$ {resultado8:,.2f}")
    print(f"total: R$ {resultado9:,.2f}")
    print(f"total: R$ {resultado10:,.2f}")
    print(f"total: R$ {resultado11:,.2f}")
    print(f"total: R$ {resultado12:,.2f}")
    print(f"total: R$ {resultado13:,.2f}")
    print(f"total: R$ {resultado14:,.2f}")
    print(f"total: R$ {resultado15:,.2f}")
    print(f"total: R$ {resultado16:,.2f}")
    print(f"total: R$ {resultado17:,.2f}")
    print(f"total: R$ {resultado18:,.2f}")
    print(f"total: R$ {resultado19:,.2f}")
    print(f"total: R$ {total_desconto:,.2f}")
    print(f"total: R$ {resultado_final:,.2f}")
    print(f"total: R$ {resultado20:,.2f}")
    print(f"total: R$ {resultado21:,.2f}")



except FileNotFoundError:
    print(f"Erro: O arquivo '{ARQUIVO}' não foi encontrado.")
except KeyError as e:
    print(f"Erro: A coluna {e} não existe na aba '{ABA}'. Verifique os cabeçalhos.")
except Exception as e:
    print(f"Ocorreu um erro inesperado: {e}")
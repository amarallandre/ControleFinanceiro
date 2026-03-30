import pandas as pd
import io


def limpeza_csv(arquivo_origem, arquivo_destino):
    # 1. Pular as 7 linhas iniciais
    with open(arquivo_origem, 'r', encoding='utf-8') as f_in:
        for _ in range(7):
            next(f_in)
        conteudo_util = f_in.read()

    # 2. Ler o conteúdo
    df = pd.read_csv(io.StringIO(conteudo_util), sep=None, engine='python')

    # 3. Tratamento da Data
    if "data/hora" in df.columns:
        def converter_para_data_br(texto):
            try:
                if pd.isna(texto): return None
                t = str(texto).lower()
                t = t.split(' gmt')[0].replace(' de ', ' ').replace('.', '').strip()

                meses_map = {
                    'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04',
                    'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08',
                    'set': '09', 'out': '10', 'nov': '11', 'dez': '12'
                }

                for nome, num in meses_map.items():
                    if nome in t:
                        t = t.replace(nome, num)
                        break

                return pd.to_datetime(t, format='%d %m %Y %H:%M:%S', errors='coerce')
            except:
                return None

        # Converte para objeto de data real
        df["data/hora"] = df["data/hora"].apply(converter_para_data_br)

        # Extração de Mês e Ano (serão adicionadas ao final do DataFrame por padrão)
        df['Nome do Mês'] = df['data/hora'].dt.month_name(locale='pt_BR')
        df['Ano'] = df['data/hora'].dt.year

    # 4. Filtros de Status
    if "Status" in df.columns:
        df = df[df["Status"] != "Cancelado"].copy()

    # 5. DEFINIÇÃO DA ORDEM DAS COLUNAS (Mês e Ano no FINAL)
    colunas_finais = [
        "data/hora", "id de liquidação", "tipo", "id do pedido", "sku", "descrição",
        "quantidade", "mercado", "tipo de conta", "atendimento", "cidade do pedido",
        "estado do pedido", "postal do pedido", "vendas do produto", "créditos de remessa",
        "créditos de embalagem de presente", "descontos promocionais",
        "imposto de vendas coletados", "tarifas de venda", "taxas fba",
        "taxas de outras transações", "outro", "total",
        "Nome do Mês", "Ano"  # <--- Aqui elas ficam no final da tabela
    ]

    # Garante que só pegamos colunas que realmente existem
    colunas_presentes = [c for c in colunas_finais if c in df.columns]
    df = df[colunas_presentes]

    # 6. Limpeza de duplicados e nulos
    df = df.dropna(subset=["data/hora"]).drop_duplicates()

    # 7. SALVAR NO FORMATO BRASILEIRO
    try:
        df.to_csv(arquivo_destino, index=False, encoding='utf-8-sig', date_format='%d/%m/%Y')
        print(f"✅ Sucesso! Arquivo '{arquivo_destino}' gerado com Mês e Ano no final da tabela.")
    except PermissionError:
        print(f"❌ Erro: Feche o arquivo '{arquivo_destino}' e tente novamente.")


# Execução
if __name__ == "__main__":
    limpeza_csv("Janeiro.csv", "tabela_final_limpa.csv")
# 📊 Automação de Relatórios Financeiros Amazon

Automação completa para processamento de relatórios financeiros da Amazon, incluindo limpeza de dados, cálculos e preenchimento automático de planilhas Excel.

---

## 🚀 Funcionalidades

* Processamento de arquivos CSV da Amazon
* Limpeza e padronização de dados
* Conversão de datas e valores financeiros
* Cálculo automático (equivalente ao `SUMIFS` do Excel)
* Geração de planilhas Excel formatadas
* Preenchimento automático de relatórios mensais

---

## 🧱 Estrutura

```
├── arrumar.py             # Limpeza e tratamento do CSV
├── arquivo_formatado.py   # Conversão e formatação para Excel
├── teste.py / teste2.py   # Cálculos financeiros
├── Tabela_final.py        # Preenchimento da planilha final
└── 2026.xlsx              # Planilha modelo
```

---

## ⚙️ Tecnologias

* Python 3
* pandas
* openpyxl

---

## ▶️ Como usar

### 1. Limpar o CSV

```bash
python arrumar.py
```

### 2. Gerar Excel formatado

```bash
python arquivo_formatado.py
```

### 3. Calcular os dados

```bash
python teste2.py
```

### 4. Preencher a planilha final

```bash
python Tabela_final.py
```

---

## 📌 Requisitos

```bash
pip install pandas openpyxl
```

---

## ⚠️ Observações

* O CSV deve seguir o padrão do relatório da Amazon
* A planilha `2026.xlsx` deve conter:

  * Meses na linha 4
  * Descrições na coluna B
* O sistema utiliza normalização de texto (ignora acentos e diferenças de escrita)

---

## 📈 Melhorias futuras

* Pipeline automatizado (execução em um único comando)
* Interface gráfica
* Integração com API da Amazon

---

## 👨‍💻 Autor

André Nascimento

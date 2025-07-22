# 📧 Relatório de Vendas Automatizado com Python

Este projeto tem como objetivo automatizar a geração e envio de um **relatório de vendas por loja** utilizando **Python**, **Pandas** e **Outlook**.

---

## ⚙️ O que o script faz?

1. Lê os dados da planilha `Vendas.xlsx`
2. Calcula:
   - Faturamento por loja
   - Quantidade de produtos vendidos por loja
   - Ticket médio por loja
3. Gera um relatório em HTML com as informações
4. Envia o relatório por e-mail automaticamente via Outlook

---

## 🛠️ Tecnologias Utilizadas

- Python
- Pandas
- win32com (para automação de e-mail com Outlook)
- Excel (`Vendas.xlsx`)

---

## 📄 Exemplo de saída

O e-mail gerado inclui:

- Tabelas formatadas com `pandas.to_html()`
- Faturamento total por loja
- Quantidade total vendida
- Ticket médio por produto

---

## 🚀 Como usar

1. Certifique-se de ter o Outlook instalado e aberto.
2. Coloque o arquivo `Vendas.xlsx` na mesma pasta do script.
3. Execute o script com Python 3:
   ```bash
   python relatorio_vendas.py

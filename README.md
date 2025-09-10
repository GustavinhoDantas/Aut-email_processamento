# 📂 Automação de E-mails de Apólices e Boletos

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Tkinter](https://img.shields.io/badge/Tkinter-GUI-green)
![Pandas](https://img.shields.io/badge/Pandas-Excel-yellow)

---

## 🔹 Sobre o Projeto

Ferramenta desenvolvida em **Python** com interface gráfica (**Tkinter**) para processar **e-mails (.eml)**, extrair informações de **apólices e boletos**, baixar PDFs e gerar planilhas Excel. Ideal para **automatizar tarefas repetitivas** e agilizar a gestão de documentos.

---

## 🔹 Funcionalidades

- Seleção de pasta contendo arquivos `.eml`.
- Extração de dados de tabelas HTML:
  - Número da apólice, controle interno, tomador, segurado, importância segurada, modalidade, vigência.
  - Número, vencimento, código e valor de boletos.
- Download automático de PDFs de apólices e boletos.
- Geração de planilha Excel (`dados_apolices_email.xlsx`) com todos os dados coletados.
- Barra de progresso e contador de arquivos processados.
- Botão para interromper execução de forma segura.
- Logs detalhados (`app.log`) de toda execução.

---

## 🔹 Tecnologias Utilizadas

- **Python 3.x**  
- **Tkinter** – interface gráfica  
- **BeautifulSoup** – parsing de HTML  
- **Pandas** – manipulação e exportação para Excel  
- **Requests** – download de PDFs  
- **Logging** – registro de ações e erros  

---

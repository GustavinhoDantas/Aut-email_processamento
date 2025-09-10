import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import logging
import os
import re
from email.parser import BytesParser
from email import policy
from bs4 import BeautifulSoup
import requests
import pandas as pd
import threading

# Configurações do log
logging.basicConfig(filename="app.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Lista para armazenar informações do excel
infos = []

# Variável global para parar execução
parar_execucao_flag = False

def parar_execucao():
    global parar_execucao_flag
    parar_execucao_flag = True
    messagebox.showinfo("Atenção", "Execução será interrompida após a operação atual.")

# Mensagem inicial
def mensagem_inicial(): 
    root = tk.Tk() 
    root.title("Aviso Importante") 
    root.geometry("600x300") 
    texto = ("Para essa automação preste atenção no que deve ser feito:\n"
             "1° - \n" 
             "2° - \n"
             "3° - \n"
             "4° - ") 
    tk.Label(root, text=texto, justify="center", padx=10, pady=20).pack() 
    tk.Button(root, text="OK", command=root.destroy).pack(pady=10) 
    root.mainloop()

# Tirar valor das tabelas pelo título <td> os que estão ao lado
def extrair_valor_por_titulo(soup, titulo):
    td = soup.find('td', string=re.compile(titulo))
    if td:
        proximo_td = td.find_next_sibling('td')
        if proximo_td:
            return proximo_td.get_text(strip=True).replace(u'\xa0','')
    return None

# Tirar valor das tabelas pelo título <td> os que estão embaixo
def extrair_valor_por_titulo_de_baixo(soup, titulo):
    td_titulo = soup.find('td', string=re.compile(titulo, re.IGNORECASE))
    if td_titulo:
        tr_atual = td_titulo.find_parent('tr')  # Sobe para a linha
        proxima_tr = tr_atual.find_next_sibling('tr')  # Pega a linha seguinte
        if proxima_tr:
            tds = proxima_tr.find_all('td')
            if tds:
                return [td.get_text(strip=True) for td in tds]  # Retorna lista com os valores
    return []

# Função principal de processamento de emails
def vendo_emails(progress, label_contador):
    global parar_execucao_flag
    if not hasattr(vendo_emails, 'pasta_selecionada'):
        messagebox.showerror("Erro", "Nenhuma pasta selecionada!")
        return
    
    # Criação de pastas para PDFs
    pasta_apolices = os.path.join(vendo_emails.pasta_selecionada, "apolices")
    pasta_boletos = os.path.join(vendo_emails.pasta_selecionada, "boletos")
    os.makedirs(pasta_apolices, exist_ok=True)
    os.makedirs(pasta_boletos, exist_ok=True)

    arquivos = [f for f in os.listdir(vendo_emails.pasta_selecionada) if f.endswith('.eml')]
    total = len(arquivos)
    processados = 0
    progress['maximum'] = total

    for arq in arquivos:
        if parar_execucao_flag:
            logging.info("Execução interrompida pelo usuário")
            break

        caminho_email = os.path.join(vendo_emails.pasta_selecionada, arq)
        logging.info(f"Email encontrado: {caminho_email}")

        # Abrir arquivo .eml
        with open(caminho_email, 'rb') as email:
            msg = BytesParser(policy=policy.default).parse(email)

        # Extrair HTML
        html_content = None
        if msg.is_multipart():
            print("O email tem varias partes")
            for parte in msg.walk():
                if parte.get_content_type() == 'text/html':
                    html_content = parte.get_content()
                    break
        else:
            html_content = msg.get_content()
            print("O email é uma parte unica")
        
        if html_content:
            soup = BeautifulSoup(html_content, 'html.parser')

            # Número da apólice
            num_apolice = extrair_valor_por_titulo(soup, r'Número da apólice')
            if num_apolice:
                nome_arquivo = f"{num_apolice}.pdf"
            else:
                nome_arquivo = f"{arq}_apolice.pdf"
                num_apolice = nome_arquivo

            # Procurar links
            link_apolice = None
            link_boleto = None
            for a in soup.find_all('a'):
                texto = a.get_text(strip=True).lower()
                href = a.get('href')
                if 'baixar apólice' in texto or 'baixar apolice' in texto:
                    link_apolice = href
                elif 'baixar boleto' in texto or 'baixar boleto(s)' in texto:
                    link_boleto = href

            # Baixar PDF apólice
            if link_apolice:
                caminho_pdf = os.path.join(pasta_apolices, nome_arquivo)
                try:
                    r = requests.get(link_apolice, timeout=20)
                    r.raise_for_status()
                    with open(caminho_pdf, 'wb') as f_pdf:
                        f_pdf.write(r.content)
                    logging.info(f"Apólice baixada: {caminho_pdf}")
                except Exception as e:
                    logging.error(f"Erro ao baixar apólice: {e}")

            # Baixar PDF boleto
            if link_boleto:
                caminho_pdf = os.path.join(pasta_boletos, nome_arquivo)
                try:
                    r = requests.get(link_boleto, timeout=20)
                    r.raise_for_status()
                    with open(caminho_pdf, 'wb') as f_pdf:
                        f_pdf.write(r.content)
                    logging.info(f"Boleto baixado: {caminho_pdf}")
                except Exception as e:
                    logging.error(f"Erro ao baixar boleto: {e}")

            # Pegando informações do email via BeautifulSoup
            numero_controle_interno = extrair_valor_por_titulo(soup, r'Controle interno')
            nome_tomador = extrair_valor_por_titulo(soup, r'Tomador')
            nome_segurador = extrair_valor_por_titulo(soup, r'Segurado')
            importancia_segurada = extrair_valor_por_titulo(soup, r'Importância segurada')
            modalidade = extrair_valor_por_titulo(soup, r'Modalidade')
            inicio_da_vigencia = extrair_valor_por_titulo(soup, r'Início de vigência')
            final_da_vigencia = extrair_valor_por_titulo(soup, r'Final de vigência')
            numero_boleto = extrair_valor_por_titulo_de_baixo(soup, r'Número')
            vencimento_boleto = extrair_valor_por_titulo_de_baixo(soup, r'Vencimento')
            cod_boleto = extrair_valor_por_titulo_de_baixo(soup, r'Boleto')
            valor_boleto = extrair_valor_por_titulo_de_baixo(soup, r'Valor')


            infos.append({
                'Numero Apolice': num_apolice,
                'Controle Interno': numero_controle_interno,
                'Tomador': nome_tomador,
                'Segurado': nome_segurador,
                'Importancia Segurada': importancia_segurada,
                'Modalidade': modalidade,
                'Inicio de Vigencia': inicio_da_vigencia,
                'Final de Vigencia': final_da_vigencia,
                'Link Apólice pdf': link_apolice,
                'Link Boleto pdf': link_boleto,
                'Numero do Boleto': numero_boleto,
                'Vencimento Boleto': vencimento_boleto,
                'Cod Boleto': cod_boleto,
                'Valor Boleto': valor_boleto
                
            })

        else:
            logging.warning(f"Este arquivo não é um e-mail: {arq}")

        # Atualiza barra e contador
        processados += 1
        progress['value'] = processados
        label_contador.config(text=f"{processados} / {total} processados")
        label_contador.update_idletasks()
        print(f"{processados} / {total} processados")

    # Após terminar, salvar Excel dentro da pasta
    excel_path = os.path.join(vendo_emails.pasta_selecionada, "dados_apolices_email.xlsx")
    df = pd.DataFrame(infos)
    df.to_excel(excel_path, index=False)
    logging.info(f"Excel gerado em: {excel_path}")
    messagebox.showinfo("Concluído", f"Processamento finalizado!\nExcel salvo em:\n{excel_path}")

# Interface GUI
def interface_usuario():
    root = tk.Tk()
    root.title("A Magia está acontecendo") 
    root.geometry("600x400")

    # Barra de progresso
    progress = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
    progress.pack(pady=10)
    label_contador = tk.Label(root, text="0 / 0 processados")
    label_contador.pack()

    # Botões iniciar/parar
    frame_btns = tk.Frame(root)
    frame_btns.pack(pady=10)
    tk.Button(frame_btns, text="Parar", command=parar_execucao, bg="red", fg="white").pack(side="left", padx=10)

    # Seleção de pasta
    label_arquivo = tk.Label(root, text="Nenhum arquivo selecionado")
    label_arquivo.pack(pady=10)

    def selecionar_arquivo():
        pasta = filedialog.askdirectory(title="Selecione a pasta com os e-mails")
        if pasta:
            vendo_emails.pasta_selecionada = pasta
            label_arquivo.config(text=f"Pasta selecionada:\n{pasta}")
            logging.info(f"Pasta selecionada: {pasta}")
            # Rodar processamento em thread para não travar GUI
            threading.Thread(target=vendo_emails, args=(progress, label_contador), daemon=True).start()
        else:
            label_arquivo.config(text="Nenhuma pasta selecionada")
            logging.warning("Nenhuma pasta foi selecionada")

    tk.Button(root, text="Selecione a Pasta", command=selecionar_arquivo).pack(pady=10)

    root.mainloop()

# Execução
mensagem_inicial()
interface_usuario()
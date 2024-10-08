from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from time import sleep
import win32com.client as win32
import re

# Lista de termos de pesquisa
termos = [
    "Sistema de Gestão de Documentos",
    "GSISTE Siga",
    "Subcomissão de Coordenação do Siga",
    "Subsiga",
    "Comissão de Coordenação do Siga",
    "comissão permanente de avaliação de documentos",
    "Edital Ciência de Eliminação de Documentos"
    "Gratificação Temporária das Unidades dos Sistemas Estruturantes da Administração Pública Federal - GSISTE"  # Novo termo adicionado
]

# Função para remover caracteres inválidos do nome do arquivo


def nome_arquivo_valido(termo):
    # Substitui caracteres inválidos por underline
    return re.sub(r'[<>:"/\\|?*]', '_', termo)

# Função para buscar e retornar resultados como HTML


def buscar_resultados(termo):
    # Configuração do WebDriver
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)

    # Acessa a página principal
    navegador.get(
        'https://www.in.gov.br/servicos/diario-oficial-da-uniao/destaques-do-diario-oficial-da-uniao')
    sleep(0.5)

    # Executa a busca avançada
    navegador.find_element(
        'xpath', '//*[@id="toggle-search-advanced"]').click()
    sleep(0.5)

    # Marca as opções das seções 1, 2 e 3
    navegador.find_element('xpath', '//*[@id="do1"]').click()  # Seção 1
    navegador.find_element('xpath', '//*[@id="do2"]').click()  # Seção 2
    navegador.find_element('xpath', '//*[@id="do3"]').click()  # Seção 3
    sleep(0.5)

    # Marca a opção "Dia"
    navegador.find_element('xpath', '//*[@id="dia"]').click()  # Opção "Dia"
    sleep(0.5)

    # Insere o termo de busca
    navegador.find_element(
        'xpath', '//*[@id="search-bar"]').send_keys(f'"{termo}"')
    navegador.find_element(
        'xpath', '//*[@id="div-search-bar"]/div/div/div/i').click()
    sleep(0.5)

    # Lista para armazenar os dados extraídos
    dados = []

    while True:
        # Obtém o conteúdo HTML da página atual
        conteudo = navegador.page_source
        site_do_soup = BeautifulSoup(conteudo, 'html.parser')

        # Extrai as portarias da página atual
        portarias = site_do_soup.findAll('div', attrs={'class': 'resultado'})

        for portaria in portarias:
            # Extrair o título
            titulo = portaria.find('h5', class_='title-marker').text.strip()

            # Extrair o link
            link = portaria.find('h5', class_='title-marker').find('a')['href']

            # Extrair o resumo
            resumo = portaria.find('p', class_='abstract-marker').text.strip()

            # Extrair a data
            data = portaria.find('p', class_='date-marker').text.strip()

            # Adiciona os dados extraídos à lista
            dados.append(
                {'Título': titulo, 'Link': f"https://www.dou.gov.br{link}", 'Resumo': resumo, 'Data': data})

        # Verifica se existe um botão para a próxima página
        try:
            # Tenta clicar no botão "Próxima Página"
            botao_proxima = navegador.find_element(
                'xpath', '//*[@id="rightArrow"]/span')
            botao_proxima.click()
            sleep(1)
        except Exception as e:
            print("Não há mais páginas.")
            break

    # Fecha o navegador
    navegador.quit()

    # Se não houver dados, retorna None
    if not dados:
        print(f"Nenhum dado encontrado para o termo: {termo}")
        return None

    # Gera o conteúdo HTML da tabela com os resultados
    conteudo_html = f"<h2>Resultados para o termo: {termo}</h2>"
    conteudo_html += "<table border='1'><tr><th>Título</th><th>Link</th><th>Resumo</th><th>Data</th></tr>"
    for dado in dados:
        conteudo_html += f"<tr><td>{dado['Título']}</td><td><a href='{
            dado['Link']}'>Link</a></td><td>{dado['Resumo']}</td><td>{dado['Data']}</td></tr>"
    conteudo_html += "</table>"

    return conteudo_html

# Função para enviar e-mail com resultados no corpo usando win32com.client


def enviar_email(conteudo_email):
    # Cria uma instância do Outlook
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    # Configurações do e-mail
    email.To = "renato.matos@gestao.an.gov.br"
    email.CC = 'fcosta@gestao.an.gov.br'
    email.BCC = 'leonardo.gati@gestao.an.gov.br'  # Adiciona CCO para você
    email.Subject = "Resultados das Portarias - DOU"
    email.HTMLBody = conteudo_email  # Usando HTML para o corpo do e-mail

    # Envia o e-mail
    email.Send()
    print(f'E-mail enviado com sucesso com os resultados.')


# Executa a busca e envia o e-mail para cada termo
for termo in termos:
    conteudo_html = buscar_resultados(termo)
    if conteudo_html:  # Verifica se houve resultados para o termo
        enviar_email(conteudo_html)


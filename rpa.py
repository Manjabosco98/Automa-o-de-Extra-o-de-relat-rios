# Criando ambiente virtual
# python -m venv 'igdRPA'
# ativando o ambiente virtual: igdRPA\Scripts\Activate.ps1

# Bibliotecas
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import service as ChromeService
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import pyautogui as pa
import os
from datetime import datetime, timedelta
import calendar
import shutil
import glob

# Variáveis
pasta_igd = r"caminho pasta"
url = r"site sistema cliente"
data_atual = datetime.now().day
mes_atual = f'{datetime.now().month:02}'
ano_atual = datetime.now().year
acesso = dict()

# Configuração navegador
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": pasta_igd,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

browser = webdriver.Chrome(options=chrome_options)
browser.implicitly_wait(5)
browser.maximize_window()
browser.get(url)

def mes_anterior():
    hoje = datetime.now()
    primeiro_do_mes_atual = hoje.replace(day=1)
    primeiro_do_mes_anterior = primeiro_do_mes_atual - timedelta(days=1)
    mes_anterior = primeiro_do_mes_anterior.strftime('%m')
    return mes_anterior


def ultimo_dia_do_mes(ano, mes):
    mes = int(mes)  # Converter a string do mês para um número inteiro
    ano = int(ano)
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return ultimo_dia


def ultimo_dia_do_mes_atual(ano, mes):
    mes = int(mes)  # Converter a string do mês para um número inteiro
    ano = int(ano)
    ultimo_dia = calendar.monthrange(ano, mes)[1]
    return ultimo_dia


def ano_anterior():
    ano_atual = datetime.now().year
    ano_anterior = ano_atual - 1
    return ano_anterior


mes_anterior_do_atual = mes_anterior()
ultimo_dia_mes_anterior = ultimo_dia_do_mes(ano_atual, mes_anterior_do_atual)
ultimo_dia_atual = ultimo_dia_do_mes_atual(ano_atual, mes_atual)
anoAnterior = ano_anterior()

print(f'O mês anterior ao mês atual é: {mes_anterior_do_atual}')
print(f'ùltimo dia mes atual - ({ultimo_dia_atual})')
print(f'ùltimo dia mes anterior - ({ultimo_dia_mes_anterior})')
print(f'Dia atual - ({data_atual})')
print(anoAnterior)

def dados_itag(browser, acesso):
    browser.find_element(By.XPATH, "//input[@name='email']").send_keys("login")   # Login
    browser.find_element(By.XPATH, "//input[@name='password']").send_keys("senha")  # Senha
    sleep(3)
    browser.find_element(By.XPATH, "//button[@type='submit']").click()  # Botao login
    sleep(5)
    browser.find_element(By.XPATH, "//h3[text()='Digital do cliente']").click()  # Digital do cliente
    sleep(10)
    browser.find_element(By.XPATH, "//input[@aria-label='ID Filter Input']").send_keys("239")
    sleep(10)
    browser.find_element(By.XPATH, "//div[@id='buttonedit']").click()
    sleep(3)
    browser.find_element(By.XPATH, "//span[@class='MuiTab-wrapper' and text()='Sistemas']").click()
    sleep(3)
    elemento = browser.find_elements(By.XPATH, "//*[@class='MuiSvgIcon-root']")
    if len(elemento) >= 4:
        elemento[2].click()
    sleep(4)

    # copiando login
    link = browser.find_element(By.XPATH, "//input[@name='link_sistema']")
    browser.implicitly_wait(3)
    acesso['link'] = link.get_attribute('value')
    sleep(2)

    # copiando login
    login = browser.find_element(By.XPATH, "//input[@name='usuario']")
    browser.implicitly_wait(3)
    acesso['login'] = login.get_attribute('value')
    sleep(2)

    # copiando senha
    senha = browser.find_element(By.XPATH, "//input[@name='senha']")
    browser.implicitly_wait(3)
    acesso['senha'] = senha.get_attribute('value')
    sleep(2)

dados_itag(browser, acesso)
print(acesso)

# Abrindo nova aba e acessando sistema REDE+
browser.switch_to.new_window("tab")

browser.get(acesso["link"])
sleep(4)

# colando login
browser.find_element(By.ID, "form:usuario").send_keys(acesso["login"])
# colando senha
browser.find_element(By.ID, "form:senha").send_keys(acesso["senha"])
# Botao entrar/logando
browser.find_element(By.ID, "form:loginBtn:loginBtn").click()
sleep(4)
browser.find_element(By.ID, "formPerfil:logarDiretamenteComoFuncionario:logarDiretamenteComoFuncionario").click()
sleep(8)

print('Primeiro relatório - relat-1.xlsx')
def categoriaDespesa(data_atual, ano_atual, mes_anterior_do_atual):
    # Relatorio Conta a pagar/pagamento - categoria Despesa
    browser.find_element(By.XPATH, "//a[@class='tituloCampos' and text()='Rel. Conta a Pagar/Pagamento - Categoria Despesa']").click()
    # pa.click(x=777, y=563)

    all_handles = browser.window_handles

    for handle in all_handles:
        browser.switch_to.window(handle)
        if "Rel. Conta a Pagar/Pagamento - Categoria Despesa - Google Chrome" in browser.title:
            break
    browser.maximize_window()

    sleep(10)

    if data_atual <= 15:
        data_inicio = f'01{mes_anterior_do_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_anterior_do_atual)
        data_final = f'{ultimo_dia}{mes_anterior_do_atual}{ano_atual}'
    else:
        data_inicio = f'01{mes_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_atual)
        data_final = f'{ultimo_dia}{mes_atual}{ano_atual}'

    # Período de
    inicio = browser.find_element(By.ID, 'form:data:data')
    inicio.send_keys(Keys.CONTROL + 'a')
    inicio.send_keys(Keys.DELETE)
    inicio.send_keys(data_inicio)
    sleep(2)

    # Período Fim
    final = browser.find_element(By.ID, 'form:dataFim:dataFim')
    final.send_keys(Keys.CONTROL + 'a')
    final.send_keys(Keys.DELETE)
    final.send_keys(data_final)
    sleep(2)

    # Filtrar Conta Paga Por Data = Data Pagamento
    filtro = Select(browser.find_element(By.ID, "form:filtroPago"))
    filtro.select_by_visible_text('Data Pagamento')
    sleep(2)

    # Gerar relatório (EXCEL)
    browser.find_element(By.ID, 'form:imprimirExcel:imprimirExcel').click()
    sleep(30)

    CAP = 'relat-1.xlsx'

    def rename_most_detailed_file_in_directory(directory, new_name):
        # Encontre todos os arquivos no diretório que correspondem ao padrão
        files = glob.glob(os.path.join(directory, '*'))

        # Se não houver arquivos no diretório, retorne
        if not files:
            return

        # Ordena os arquivos por data de modificação (o mais recente primeiro)
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Pega o arquivo mais recente
        most_detailed_file = None
        most_detailed_time = 0

        for file in files:
            mtime = os.path.getmtime(file)
            if mtime > most_detailed_time:
                most_detailed_time = mtime
                most_detailed_file = file

        if most_detailed_file is not None:
            # Caminho do novo arquivo
            new_file_path = os.path.join(directory, new_name)

            # Verifica se já existe um arquivo com o novo nome e o exclui
            if os.path.exists(new_file_path):
                os.remove(new_file_path)

            # Renomeia o arquivo mais detalhado para o novo nome
            os.rename(most_detailed_file, new_file_path)

            print(f"Arquivo mais detalhado renomeado como '{new_name}' em {directory}")

    rename_most_detailed_file_in_directory(pasta_igd, CAP)
    sleep(10)

    # Fechar a janela atual sem encerrar o navegador principal
    browser.close()

    # Voltar ao identificador da janela principal (caso você precise interagir com ela posteriormente)
    browser.switch_to.window(browser.window_handles[1])

categoriaDespesa(data_atual, ano_atual, mes_anterior_do_atual)
sleep(10)

browser.refresh()
sleep(4)


print('Segundo relatório - relat-2.xlsx')
def contas_a_pagar(data_atual, ano_atual, mes_anterior_do_atual):
    browser.find_element(By.XPATH, "//a[@class='tituloCampos' and text()='Rel. Conta a Pagar']").click()
    all_handles = browser.window_handles

    for handle in all_handles:
        browser.switch_to.window(handle)
        if "Relatório Conta a Pagar/Pagamento - Google Chrome" in browser.title:
            break
    browser.maximize_window()

    sleep(10)

    if data_atual <= 15:
        data_inicio = f'01{mes_anterior_do_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_anterior_do_atual)
        data_final = f'{ultimo_dia}{mes_anterior_do_atual}{ano_atual}'
    else:
        data_inicio = f'01{mes_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_atual)
        data_final = f'{ultimo_dia}{mes_atual}{ano_atual}'

    # Período de
    inicio = browser.find_element(By.ID, "form:data:data")
    inicio.send_keys(Keys.CONTROL + 'a')
    inicio.send_keys(Keys.DELETE)
    inicio.send_keys(data_inicio)
    sleep(2)

    # Período Fim
    final = browser.find_element(By.ID, "form:dataFim:dataFim")
    final.send_keys(Keys.CONTROL + 'a')
    final.send_keys(Keys.DELETE)
    final.send_keys(data_final) 
    sleep(2)

    # Filtrar Conta Paga Por Data = Data Pagamento
    filtro = Select(browser.find_element(By.ID, "form:filtroPago"))
    filtro.select_by_visible_text('Data Pagamento')

    # Conta Corrente
    browser.find_element(By.ID, 'form:j_idt438:0').click()
    sleep(8)

    # Gerar relatório (EXCEL)
    browser.find_element(By.ID, "form:imprimirExcel:imprimirExcel").click()
    sleep(30)

    CAP = 'relat-2.xlsx'

    def rename_most_detailed_file_in_directory(directory, new_name):
        # Encontre todos os arquivos no diretório que correspondem ao padrão
        files = glob.glob(os.path.join(directory, '*'))

        # Se não houver arquivos no diretório, retorne
        if not files:
            return

        # Ordena os arquivos por data de modificação (o mais recente primeiro)
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Pega o arquivo mais recente
        most_detailed_file = None
        most_detailed_time = 0

        for file in files:
            mtime = os.path.getmtime(file)
            if mtime > most_detailed_time:
                most_detailed_time = mtime
                most_detailed_file = file

        if most_detailed_file is not None:
            # Caminho do novo arquivo
            new_file_path = os.path.join(directory, new_name)

            # Verifica se já existe um arquivo com o novo nome e o exclui
            if os.path.exists(new_file_path):
                os.remove(new_file_path)

            # Renomeia o arquivo mais detalhado para o novo nome
            os.rename(most_detailed_file, new_file_path)

            print(f"Arquivo mais detalhado renomeado como '{new_name}' em {directory}")


    rename_most_detailed_file_in_directory(pasta_igd, CAP)
    sleep(10)

    # Fechar a janela atual sem encerrar o navegador principal
    browser.close()

    # Voltar ao identificador da janela principal (caso você precise interagir com ela posteriormente)
    browser.switch_to.window(browser.window_handles[1])


contas_a_pagar(data_atual, ano_atual, mes_anterior_do_atual)
sleep(10)

browser.refresh()
sleep(8)

print('Terceiro relatório - relat-3.xlsx')
def extrato_Conta_Corrente(data_atual, ano_atual, mes_anterior_do_atual):
    browser.find_element(By.XPATH, "//a[@class='tituloCampos' and text()='Extrato Conta Corrente']").click()
    sleep(10)

    #datas
    if data_atual <= 15:
        data_inicio = f'01{mes_anterior_do_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_anterior_do_atual)
        data_final = f'{ultimo_dia}{mes_anterior_do_atual}{ano_atual}'
    else:
        data_inicio = f'01{mes_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_atual)
        data_final = f'{ultimo_dia}{mes_atual}{ano_atual}'

    all_handles = browser.window_handles

    for handle in all_handles:
        browser.switch_to.window(handle)
        if "Extrato Conta Corrente - Google Chrome" in browser.title:
            break
    browser.maximize_window()
    sleep(3)

    # Período
    inicio = browser.find_element(By.ID, "form:valorConsultaData:valorConsultaData")
    inicio.send_keys(Keys.CONTROL + 'a')
    inicio.send_keys(Keys.DELETE)
    inicio.send_keys(data_inicio)
    sleep(2)

    # Período Fim
    final = browser.find_element(By.ID, "form:dataFim:dataFim")
    final.send_keys(Keys.CONTROL + 'a')
    final.send_keys(Keys.DELETE)
    final.send_keys(data_final) 
    sleep(2)

    # Conta Corrente
    filtro = Select(browser.find_element(By.ID, "form:contaCorrente"))
    filtro.select_by_visible_text('Todas')
    sleep(2)

    # Gerar Relatório (EXCEL)
    browser.find_element(By.ID, "form:imprimirExcel:imprimirExcel").click()
    sleep(20)

    CAP = 'relat-3.xlsx'

    def rename_most_detailed_file_in_directory(directory, new_name):
        # Encontre todos os arquivos no diretório que correspondem ao padrão
        files = glob.glob(os.path.join(directory, '*'))

        # Se não houver arquivos no diretório, retorne
        if not files:
            return

        # Ordena os arquivos por data de modificação (o mais recente primeiro)
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Pega o arquivo mais recente
        most_detailed_file = None
        most_detailed_time = 0

        for file in files:
            mtime = os.path.getmtime(file)
            if mtime > most_detailed_time:
                most_detailed_time = mtime
                most_detailed_file = file

        if most_detailed_file is not None:
            # Caminho do novo arquivo
            new_file_path = os.path.join(directory, new_name)

            # Verifica se já existe um arquivo com o novo nome e o exclui
            if os.path.exists(new_file_path):
                os.remove(new_file_path)

            # Renomeia o arquivo mais detalhado para o novo nome
            os.rename(most_detailed_file, new_file_path)

            print(f"Arquivo mais detalhado renomeado como '{new_name}' em {directory}")

    rename_most_detailed_file_in_directory(pasta_igd, CAP)
    sleep(10)

    # Fechar a janela atual sem encerrar o navegador principal
    browser.close()

    # Voltar ao identificador da janela principal (caso você precise interagir com ela posteriormente)
    browser.switch_to.window(browser.window_handles[1])

extrato_Conta_Corrente(data_atual, ano_atual, mes_anterior_do_atual)
sleep(10)

browser.refresh()
sleep(8)

print('Quarto relatório - relat-4.xlsx')
def Relatório_SEI_Decidir_Receita(data_atual, ano_atual, mes_anterior_do_atual):
    browser.find_element(By.XPATH, "//a[@class='tituloCampos' and text()='Relatório SEI Decidir (Receita)']").click()
    sleep(10)

    all_handles = browser.window_handles

    for handle in all_handles:
        browser.switch_to.window(handle)
        if "Relatório SEI Decidir (Receita) - Google Chrome" in browser.title:
            break
    browser.maximize_window()
        
    layout = Select(browser.find_element(By.XPATH, "//select[@id='form:tipoLayout']"))
    layout.select_by_visible_text('RELATÓRIO RECEITA')

    sleep(8)

    #datas
    if data_atual <= 15:
        data_inicio = f'01{mes_anterior_do_atual}{anoAnterior}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_anterior_do_atual)
        data_final = f'{ultimo_dia}{mes_anterior_do_atual}{ano_atual}'
    else:
        data_inicio = f'01{mes_atual}{anoAnterior}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_atual)
        data_final = f'{ultimo_dia}{mes_atual}{ano_atual}'

    # Filtra período por
    recebimento = Select(browser.find_element(By.ID, "form:tipoFiltroPeriodo")).select_by_value('DATA_RECEBIMENTO')
    sleep(3)

    # Data início
    inicio = browser.find_element(By.XPATH, "//input[@id='form:dataInicio:dataInicio']")
    inicio.send_keys(Keys.CONTROL + 'a')
    inicio.send_keys(Keys.DELETE)
    inicio.send_keys(data_inicio)

    # Data Final
    final = browser.find_element(By.XPATH, "//input[@id='form:dataFim:dataFim']")
    final.send_keys(Keys.CONTROL + 'a')
    final.send_keys(Keys.DELETE)
    final.send_keys(data_final)
    sleep(5)

    # Situação Conta Receber
    browser.find_element(By.XPATH, "//a[@id='form:j_idt847:j_idt847']").click()
    sleep(5)

    # Tipo Origem
    browser.find_element(By.XPATH, "//a[@id='form:j_idt914:j_idt914']").click()
    sleep(5)

    body = browser.find_element("tag name", "body")
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)

    sleep(3)

    # Filtros Situação Acadêmica
    browser.find_element(By.XPATH, "//a[@id='form:j_idt1005:j_idt1005']//*[name()='svg']//*[name()='path' and contains(@fill,'currentCol')]").click()
    sleep(4 )

    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)

    # Situação Financeira Matrícula -> confirmado
    sleep(2)
    confirmado =  browser.find_elements(By.XPATH, "//label[@class='flipswitch-label']")
    if len(confirmado) <= 32:
        confirmado[29].click()
    sleep(8)

    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)
    body.send_keys(Keys.ARROW_DOWN)

    # Gerar relatório (EXCEL)
    browser.find_element(By.XPATH, "//a[@id='form:imprimirExcel:imprimirExcel']").click()
    sleep(190)

    CAP = 'relat-4.xlsx'

    def rename_most_detailed_file_in_directory(directory, new_name):
        # Encontre todos os arquivos no diretório que correspondem ao padrão
        files = glob.glob(os.path.join(directory, '*'))

        # Se não houver arquivos no diretório, retorne
        if not files:
            return

        # Ordena os arquivos por data de modificação (o mais recente primeiro)
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Pega o arquivo mais recente
        most_detailed_file = None
        most_detailed_time = 0

        for file in files:
            mtime = os.path.getmtime(file)
            if mtime > most_detailed_time:
                most_detailed_time = mtime
                most_detailed_file = file

        if most_detailed_file is not None:
            # Caminho do novo arquivo
            new_file_path = os.path.join(directory, new_name)

            # Verifica se já existe um arquivo com o novo nome e o exclui
            if os.path.exists(new_file_path):
                os.remove(new_file_path)

            # Renomeia o arquivo mais detalhado para o novo nome
            os.rename(most_detailed_file, new_file_path)

            print(f"Arquivo mais detalhado renomeado como '{new_name}' em {directory}")

    rename_most_detailed_file_in_directory(pasta_igd, CAP)
    sleep(10)

    # Fechar a janela atual sem encerrar o navegador principal
    browser.close()

    # Voltar ao identificador da janela principal (caso você precise interagir com ela posteriormente)
    browser.switch_to.window(browser.window_handles[1])


Relatório_SEI_Decidir_Receita(data_atual, ano_atual, mes_anterior_do_atual)
sleep(30)

browser.refresh()
sleep(4)

print('Quinto relatório - relat-5.xlsx')
def Mapa_de_Pendências_Cartão_de_Crédito(data_atual, ano_atual, mes_anterior_do_atual):
    browser.find_element(By.XPATH, "//a[@class='tituloCampos' and text()='Mapa de Pendências Cartão de Crédito']").click()
    sleep(10)

    all_handles = browser.window_handles

    for handle in all_handles:
        browser.switch_to.window(handle)
        if "Mapa de Pendências Cartão - Google Chrome" in browser.title:
            break
    browser.maximize_window()

    # Filtrar por situação
    situacao = Select(browser.find_element(By.ID, "form:consulta"))
    situacao.select_by_visible_text('Recebida')
    sleep(10)

    # Filtrar por
    filtro = Select(browser.find_element(By.ID, "form:tipoData"))
    filtro.select_by_visible_text('Data Recebimento Operadora')
    sleep(10)

    #datas
    if data_atual <= 15:
        data_inicio = f'01{mes_anterior_do_atual}{ano_atual}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_anterior_do_atual)
        data_final = f'{ultimo_dia}{mes_anterior_do_atual}{ano_atual}'
    else:
        data_inicio = f'01{mes_atual}{anoAnterior}'
        ultimo_dia = ultimo_dia_do_mes(ano_atual, mes_atual)
        data_final = f'{ultimo_dia}{mes_atual}{ano_atual}'
        
    # data início
    inicio = browser.find_element(By.XPATH, "//input[@id='form:dataRecebimentoOperadoraInicial:dataRecebimentoOperadoraInicial']")
    inicio.send_keys(Keys.CONTROL + 'a')
    inicio.send_keys(Keys.DELETE)
    inicio.send_keys(data_inicio)
    sleep(5)

    # data final
    final = browser.find_element(By.XPATH, "//input[@id='form:dataRecebimentoOperadoraFinal:dataRecebimentoOperadoraFinal']")
    final.send_keys(Keys.CONTROL + 'a')
    final.send_keys(Keys.DELETE)
    final.send_keys(data_final)

    # consultar
    browser.find_element(By.XPATH, "//a[@id='form:consultar:consultar']").click()
    sleep(30)

    # imprimir (EXCEL)
    browser.find_element(By.XPATH, "//a[@id='form:imprimirExcel:imprimirExcel']").click()
    sleep(100)

    CAP = 'relat-5.xlsx'

    def rename_most_detailed_file_in_directory(directory, new_name):
        # Encontre todos os arquivos no diretório que correspondem ao padrão
        files = glob.glob(os.path.join(directory, '*'))

        # Se não houver arquivos no diretório, retorne
        if not files:
            return

        # Ordena os arquivos por data de modificação (o mais recente primeiro)
        files.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        # Pega o arquivo mais recente
        most_detailed_file = None
        most_detailed_time = 0

        for file in files:
            mtime = os.path.getmtime(file)
            if mtime > most_detailed_time:
                most_detailed_time = mtime
                most_detailed_file = file

        if most_detailed_file is not None:
            # Caminho do novo arquivo
            new_file_path = os.path.join(directory, new_name)

            # Verifica se já existe um arquivo com o novo nome e o exclui
            if os.path.exists(new_file_path):
                os.remove(new_file_path)

            # Renomeia o arquivo mais detalhado para o novo nome
            os.rename(most_detailed_file, new_file_path)

            print(f"Arquivo mais detalhado renomeado como '{new_name}' em {directory}")

    rename_most_detailed_file_in_directory(pasta_igd, CAP)

    # Fechar a janela atual sem encerrar o navegador principal
    browser.close()

    # Voltar ao identificador da janela principal (caso você precise interagir com ela posteriormente)
    browser.switch_to.window(browser.window_handles[1])


Mapa_de_Pendências_Cartão_de_Crédito(data_atual, ano_atual, mes_anterior_do_atual)
sleep(10)

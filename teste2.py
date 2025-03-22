from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook
import time


# Função para salvar os dados em uma planilha Excel
def salvar_dados_excel(dados, planilha):
    try:
        # Adicionar os dados à planilha
        planilha.append(dados)
        print("Dados salvos na planilha.")
    except Exception as e:
        print(f"Erro ao salvar os dados em Excel: {e}")


# Criar uma nova planilha Excel
workbook = Workbook()
sheet = workbook.active
sheet.append(["Objeto", "Razão Social", "Total Adjudicado"])  # Adicionar cabeçalhos


# Iterar sobre os processos na tabela
for i in range(3, 23):  # O primeiro processo começa em tr[3], o último em tr[22]
    try:

        # Configurações do ChromeDriver
        options = Options()
        options.add_experimental_option("detach", True)  # Mantém o navegador aberto após o script terminar
        driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))

        # Acessar o site
        url = "https://transparencia.pb.gov.br/relatorios/?rpt=licitacoes"
        driver.get(url)

        # Aguardar o carregamento da página
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            print("Página carregada.")
        except Exception as e:
            print(f"Erro ao carregar a página: {e}")

        # Fechar o pop-up de cookies (se existir)
        try:
            botao_ok = WebDriverWait(driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'OK')]"))
            )
            botao_ok.click()
            print("Pop-up de cookies fechado.")
        except:
            print("Nenhum pop-up de cookies encontrado.")

        # Preencher os filtros
        try:
            # Selecionar o ano de abertura
            select_ano = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='RPTRender_ctl08_ctl03_ddValue']"))
            )
            Select(select_ano).select_by_value("2019")  # Seleciona o ano 2019
            print("Ano de abertura selecionado: 2019")
            time.sleep(1)  # Aguarda 2 segundos

            # Selecionar o mês inicial
            select_mes_inicial = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='RPTRender_ctl08_ctl07_ddValue']"))
            )
            Select(select_mes_inicial).select_by_value("1")  # Seleciona JANEIRO
            print("Mês inicial selecionado: JANEIRO")
            time.sleep(1)  # Aguarda 2 segundos

            # Selecionar o mês final
            select_mes_final = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='RPTRender_ctl08_ctl11_ddValue']"))
            )
            Select(select_mes_final).select_by_value("12")  # Seleciona DEZEMBRO
            print("Mês final selecionado: DEZEMBRO")
            time.sleep(1)  # Aguarda 2 segundos

            # Selecionar a modalidade
            select_modalidade = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='RPTRender_ctl08_ctl15_ddValue']"))
            )
            Select(select_modalidade).select_by_value("1")  # Seleciona DISPENSA DE LICITAÇÃO
            print("Modalidade selecionada: DISPENSA DE LICITAÇÃO")
            time.sleep(1)  # Aguarda 2 segundos

            # Selecionar a situação
            select_situacao = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//*[@id='RPTRender_ctl08_ctl19_ddValue']"))
            )
            Select(select_situacao).select_by_value("7")  # Seleciona PROCESSO FINALIZADO DISPENSA
            print("Situação selecionada: PROCESSO FINALIZADO DISPENSA")
            time.sleep(1)  # Aguarda 2 segundos
        except Exception as e:
            print(f"Erro ao preencher os filtros: {e}")

        # Clicar no botão "Exibir Relatório"
        try:
            botao_exibir = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='RPTRender_ctl08_ctl00']"))
            )
            botao_exibir.click()
            print("Botão 'Exibir Relatório' clicado.")
        except Exception as e:
            print(f"Erro ao clicar no botão 'Exibir Relatório': {e}")

        # Aguardar o carregamento da tabela
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[1]/table/tbody/tr[3]/td[1]/div/a/div"))
            )
            print("Tabela carregada.")
        except Exception as e:
            print(f"Erro ao carregar a tabela: {e}")

        # Clicar no número do processo na mesma janela
        try:
            # Localizar o número do processo usando o full XPath
            numero_processo = driver.find_element(By.XPATH, f"/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[1]/table/tbody/tr[{i}]/td[1]/div/a/div")
            numero_processo.click()
        except Exception as e:
            print(f"Erro ao clicar no número do processo: {e}")

        # Aguardar o carregamento da página do processo
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            print("Página do processo carregada.")
        except Exception as e:
            print(f"Erro ao carregar a página do processo: {e}")

        # Extrair os dados da página do processo
        try:
            # Aguardar o carregamento dos elementos
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[1]/div/div"))
            )

            # Extrair o objeto
            objeto = driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[1]/div/div").text
            print(f"Objeto: {objeto}")

            # Extrair a razão social
            razao_social = driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[9]/td/table/tbody/tr/td/table/tbody/tr[3]/td[5]/div/div").text
            print(f"Razão Social: {razao_social}")

            # Extrair o total adjudicado
            total_adjudicado = driver.find_element(By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[4]/div/div").text
            print(f"Total Adjudicado: {total_adjudicado}")

            # Salvar os dados na planilha (linha por linha)
            if objeto and razao_social and total_adjudicado:
                salvar_dados_excel([objeto, razao_social, total_adjudicado], sheet)

            try:
                workbook.save("dados_processos.xlsx")
                print("Planilha final salva em 'dados_processos.xlsx'.")
            except Exception as e:
                print(f"Erro ao salvar a planilha: {e}")

        except Exception as e:
            print(f"Erro ao extrair os dados: {e}")
        finally:
            # Fechar a janela atual e reabrir o navegador para o próximo processo
            driver.quit()
            print("Navegador reiniciado para o próximo processo.")
    except Exception as e:
        print(f"Erro ao processar o processo {i - 2}: {e}")

# Salvar os dados em uma planilha Excel após o término de todo o loop


# Manter o navegador aberto para inspeção
print("Script concluído. O navegador permanecerá aberto.")
time.sleep(10)   # Mantém o navegador aberto por 10 segundos após o término do script
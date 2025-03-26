from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException
import time
from selenium.common.exceptions import NoSuchElementException, TimeoutException


def extrair_com_fallback(xpath, nome_campo):
    try:
        elemento = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        valor = elemento.text.strip() if elemento.text else "VAZIO"
        print(f"{nome_campo}: {valor}")
        return valor
    except Exception:
        print(f"{nome_campo} não encontrado")
        return "NÃO ENCONTRADO"


def salvar_dados_excel(dados, planilha):
    planilha.append(dados)

def get_process_xpath(pagina, i):
    """Retorna o XPath correto baseado no número da página"""
    if pagina == 1:
        return f"/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[1]/table/tbody/tr[{i}]/td[1]/div/a/div"
    else:
        return f"/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr[{i-1}]/td[1]/div/a/div"


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
sheet.append(["Número do Processo", "Objeto", "Razão Social", "Total Adjudicado"])  # Adicionar cabeçalhos



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



# Loop para 4 páginas (de 1 a 4)
for pagina in range(1, 5):
    print(f"==================== PÁGINA {pagina} ====================")
    
    # Determina o range de linhas baseado na página
    start_row = 3 if pagina == 1 else 2
    end_row = 23 if pagina == 1 else 22
    
    # Loop para os processos de cada página
    for i in range(start_row, end_row):
        print(f"--------------------- PROCESSO nº {i-start_row+1} ---------------------")
        
        try:
            # Localizar e clicar no número do processo com XPath dinâmico
            process_xpath = get_process_xpath(pagina, i)
            
            try:
                numero_processo = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, process_xpath))
                )
                texto_numero_processo = numero_processo.text
                print(f"Número do processo: {texto_numero_processo}")
                
                # Scroll para o elemento antes de clicar
                driver.execute_script("arguments[0].scrollIntoView();", numero_processo)
                numero_processo.click()
            except Exception as e:
                print(f"Erro ao clicar no processo {i-start_row+1}: {e}")
                continue

            # Aguardar carregamento da página do processo
            try:
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.TAG_NAME, "body")))
            except TimeoutException:
                print("Timeout ao carregar página do processo")
                continue

            # Extração dos dados (mesmo método para todas páginas)
            try:
                objeto = extrair_com_fallback(
                    "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[1]/div/div",
                    "Objeto"
                )
                
                razao_social = extrair_com_fallback(
                    "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[9]/td/table/tbody/tr/td/table/tbody/tr[3]/td[5]/div/div",
                    "Razão Social"
                )
                
                total_adjudicado = extrair_com_fallback(
                    "/html/body/form/div[3]/div/div/table/tbody/tr[4]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[2]/td[3]/table/tbody/tr[5]/td[4]/div/div",
                    "Total Adjudicado"
                )

                # Salvar dados
                if texto_numero_processo:
                    salvar_dados_excel([texto_numero_processo, objeto, razao_social, total_adjudicado], sheet)
                    workbook.save("dados_processos.xlsx")

            except Exception as e:
                print(f"Erro na extração de dados: {e}")
                continue

            # Voltar para lista de processos
            try:
                btn_voltar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div[1]/div/div[3]/table/tbody/tr/td/div/div[1]/table/tbody/tr/td/span"))
                )
                btn_voltar.click()
                
                # Aguardar recarregamento com XPath dinâmico
                wait_xpath = get_process_xpath(pagina, start_row)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, wait_xpath)))
                
            except Exception as e:
                print(f"Erro ao voltar para lista: {e}")
                driver.refresh()
                time.sleep(5)
                continue

        except Exception as e:
            print(f"Erro geral no processo {i-start_row+1}: {e}")
            continue

    # Navegar para próxima página (exceto na última iteração)
    if pagina < 4:
        try:
            print("Indo para próxima página...")
            btn_proximo = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[3]/div/div/table/tbody/tr[3]/td/div[1]/div/div[1]/table/tbody/tr/td[11]/div/div[1]/table/tbody/tr/td/span"))
            )
            driver.execute_script("arguments[0].scrollIntoView();", btn_proximo)
            btn_proximo.click()
            
            # Aguardar carregamento com timeout maior
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, get_process_xpath(pagina+1, start_row)))
            )
            time.sleep(5)  # Espera adicional
            
        except Exception as e:
            print(f"Falha ao mudar de página: {e}")
            break

print("Extração concluída para todas as páginas!")
driver.quit()

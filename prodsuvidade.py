import time
import getpass # Importa a biblioteca para senhas seguras
import os
import glob
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta # Para lidar com datas (hoje e ontem)
from selenium.webdriver.support.ui import Select # Para caixas de <select>


# Precisa fazer pip install openpyxl

# --- Bibliotecas do EXCEL ---
try:
  import openpyxl
  from openpyxl.styles import PatternFill, Font, Border, Side
  from openpyxl.utils import get_column_letter
except ImportError:
  print("Biblioteca 'openpyxl' não encontrada.")
  print("Por favor, instale-a com: pip install openpyxl")
  exit()


# --- Configuração do Driver ---
try:
  driver = webdriver.Chrome() 
  driver.get("https://podio.com/login")
  driver.maximize_window()
except Exception as e:
  print(f"Erro ao iniciar o Chrome. Verifique seu chromedriver. {e}")
  exit()
# --- Etapa 1: Aceitar Cookies E Clicar no Login da Microsoft ---
try:
  print("Aguardando página de login...")
  
  # --- ETAPA 1.A (NOVA): Aceitar o banner de Cookies ---
  try:
    print("Procurando o banner de cookies (OneTrust)...")
    # Este é o ID padrão do botão "Aceitar" do OneTrust
    accept_cookies_id = "onetrust-accept-btn-handler"
    
    # Espera o botão de cookies aparecer e ser clicável
    cookie_button = WebDriverWait(driver, 10).until(
      EC.element_to_be_clickable((By.ID, accept_cookies_id))
    )
    
    print("Banner de cookies encontrado. Clicando em 'Aceitar'...")
    cookie_button.click()
    time.sleep(1) # Espera 1s para o banner fechar
  except TimeoutException:
    # Se o banner não aparecer (ex: já foi aceito), apenas avisa e continua
    print("Banner de cookies não encontrado ou já aceito. Continuando...")
  # --- FIM DA ETAPA 1.A ---

  # --- ETAPA 1.B (Original): Clicar no Login da Microsoft ---
  microsoft_login_xpath = "//a[@data-provider='live']"
  
  print("Procurando o botão de login da Microsoft...")
  microsoft_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, microsoft_login_xpath))
  )
  
  print("Clicando no botão da Microsoft...")
  microsoft_button.click()

except Exception as e:
  print(f"Erro na Etapa 1: {e}")
  driver.save_screenshot("erro_etapa_1_cookies.png")
  driver.quit()
  exit()
# --- Etapa 1.5: Lidar com o login da Microsoft (Versão Final) ---
try:
  print("Aguardando a nova janela/aba de login da Microsoft...")
  
  # Espera o pop-up abrir (total de 2 janelas)
  WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))

  # Pega o ID da janela pop-up
  popup_window = None
  original_window = driver.current_window_handle # Guarda a original (que vai fechar)
  for window_handle in driver.window_handles:
    if window_handle != original_window:
      popup_window = window_handle
      break
      
  driver.switch_to.window(popup_window)
  print("Foco mudado para a janela de login da Microsoft.")
      
  print("Aguardando a tela de login da Microsoft...")
  
  # Preenche o e-mail (usando a variável segura)
  email_field_microsoft = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "i0116")) 
  )
  print("Preenchendo e-mail da Microsoft...")
  email_field_microsoft.send_keys("pedro.henrsilva@mrv.com.br")
  WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
  
  # Preenche a senha (usando a variável segura)
  password_field_microsoft = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "i0118")) 
  )
  print("Preenchendo senha da Microsoft...")
  password_field_microsoft.send_keys(" ")#Sua senha aqui
  
  # Tenta clicar no botão "Entrar" (com loop anti-stale)
  print("Procurando o botão 'Entrar'...")
  tentativas = 0
  clicado_entrar = False
  while not clicado_entrar and tentativas < 5:
    try:
      entrar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
      entrar_button.click()
      clicado_entrar = True
      print("Botão 'Entrar' clicado.")
    except StaleElementReferenceException:
      tentativas += 1; time.sleep(0.5)
  if not clicado_entrar: raise Exception("Falha ao clicar em Entrar")

  # --- ESPERA PELO MFA MANUAL ---
  print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
  print("Aguardando aprovação do MFA no seu celular (até 180s)...")
  
  tentativas = 0
  clicado_manter = False
  while not clicado_manter and tentativas < 5:
    try:
      keep_logged_in_button = WebDriverWait(driver, 180).until( 
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
      )
      keep_logged_in_button.click() # Clica "Sim"
      clicado_manter = True
      print("MFA Aprovado! Botão 'Manter conectado' clicado.")
    except StaleElementReferenceException:
      tentativas += 1; time.sleep(0.5)
    except TimeoutException:
      print("Erro: Timeout após 180s. Você não aprovou o MFA a tempo?")
      clicado_manter = False; break
  if not clicado_manter: raise Exception("Falha ao clicar em Manter Conectado")

  print("Login da Microsoft concluído na janela pop-up.")
  
  # --- LÓGICA CORRIGIDA (Baseada na sua análise) ---
  
  # 1. Espera a janela pop-up fechar sozinha
  print("Aguardando janela pop-up fechar...")
  WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))
  
  # 2. Pega o handle da ÚNICA janela que sobrou (a nova janela principal)
  nova_janela_principal = driver.window_handles[0]
  driver.switch_to.window(nova_janela_principal)
  print("Foco retornado para a janela principal do Podio.")

  ''''''
  # 3. Força a navegação para a home page
  print("Forçando navegação para https://podio.com/home")
  driver.get("https://podio.com/home")
  

  print("Página principal carregada com sucesso.")
  
except Exception as e:
  print(f"Erro durante o login na Microsoft (Etapa 1.5): {e}")
  driver.save_screenshot("erro_etapa_1-5.png") 
  driver.quit()
  exit()

# --- Início da Navegação no Podio (Etapas 2-12) ---

try:
  # --- ETAPA 2 CORRIGIDA ---
  print("Etapa 2: Procurando 'Vá para uma área de trabalho'...")
  
  # 1. Encontra o elemento 'pai' (a caixa que você passa o mouse por cima)
  parent_element_xpath = "//div[contains(@class, 'space-switcher-wrapper')]"
  parent_element = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, parent_element_xpath))
  )
  
  # 2. Simula o mouse passando por cima dele
  print("Simulando passagem do mouse (hover)...")
  actions = ActionChains(driver)
  actions.move_to_element(parent_element).perform()
  
  # 3. AGORA, espera o botão/texto de dentro aparecer
  
  # Etapa 3: Clicar em "ADM - Núcleo Contratos"
  print("Etapa 3: Aguardando a lista de áreas de trabalho...")
  adm_link_xpath = "//a[contains(text(), 'ADM - Núcleo Contratos')]"
  adm_link = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, adm_link_xpath))
  )
  print("Clicando em 'ADM - Núcleo Contratos'...")
  adm_link.click()


  # --- ETAPA 4 (Usando data-app-id) ---
  print("Etapa 4: Procurando o app 'Mensageria'...")
  
  mensageria_app_xpath = "//li[@data-app-id='22830484']"
  
  mensageria_app = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, mensageria_app_xpath))
  )
  print("Clicando no app 'Mensageria'...")
  mensageria_app.click()
  
  # --- FIM DA ETAPA 4 ---


  # --- ETAPA 5 ---
  print("Etapa 5: Aguardando a página 'Mensageria' carregar...") 

  x = 0
  while x <=2: # Tenta por 2 vezes
    try:
      # 1. Espera o <ul> (pai) carregar
      parent_filter_xpath = "//ul[@class='app-filter-tools']"
      parent_filter_element = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.XPATH, parent_filter_xpath))
      )
      print(f"Tentativa {tentativas+1}: Container <ul> (pai) encontrado.")

      # 2. Encontra TODOS os <li> dentro dele
      child_items_xpath = ".//li"
      child_items = parent_filter_element.find_elements(By.XPATH, child_items_xpath)
      
      if not child_items:
        raise Exception("Container <ul> encontrado, mas nenhum <li> filho foi encontrado.")
      
      print(f"Encontrados {len(child_items)} itens. Passando o mouse sobre eles...")

      # 3. Simula a ação "humana":
      #  Move o mouse sobre cada item da lista para o menu
      actions_filter = ActionChains(driver)
      for item in child_items:
        actions_filter.move_to_element(item)
      actions_filter.perform() # Executa a sequência de "hovers"
      
      # 4. Agora que o menu está ativo, espera o <li> "Filtros"
      target_filter_xpath = ".//li[@data-original-title='Filtros']" 
      target_filter = WebDriverWait(parent_filter_element, 2).until(
        EC.element_to_be_clickable((By.XPATH, target_filter_xpath))
      )
      
      # 5. Clica
      print("Ícone 'Filtros' acordado e clicável. Clicando...")
      target_filter.click()
      
      clicado_filtro = True # Sucesso! Sai do loop.
      print("Ícone 'Filtros' clicado com sucesso.")
      x += 1
    except (StaleElementReferenceException, TimeoutException) as e:
      # 6. Se der "Stale" ou "Timeout", espera e tenta a etapa 5 inteira de novo
      print(f"Tentativa {x} falhou (Stale ou Timeout). Página recarregando... Tentando de novo.")
      time.sleep(1) # Espera 1 segundo para a página estabilizar
      
  # 7. Se o loop terminar sem clicar, força um erro
  if not clicado_filtro:
    raise Exception("Falha ao encontrar o ícone 'Filtros' após 2 tentativas.")


      
  print("Ação concluída com sucesso!")
  time.sleep(1)

# --- FIM DA ETAPA 5 ---

# --- ETAPA 6 ---

  print("Etapa 6: Procurando o 'Criado em'...")
  
  Criado_em_xpath = "//li[@data-id='created_on']"
  
  Criado_em = WebDriverWait(driver, 3).until(
    EC.element_to_be_clickable((By.XPATH, Criado_em_xpath))
  )
  print("Clicando no 'Criado em'...")
  Criado_em.click()




# --- FIM DA ETAPA 6 ---

# --- ETAPA 7 ---

  print("Etapa 7: Procurando o 'Hoje'...")
  
  Criado_no_mes_passado = "//li[@data-id='-1mr:-1mr']"
  
  Criado_mes = WebDriverWait(driver, 5).until(
    EC.element_to_be_clickable((By.XPATH, Criado_no_mes_passado))
  )
  print("Clicando no 'Hoje'...")
  Criado_mes.click()


# --- FIM DA ETAPA 7 ---


# --- ETAPA 8 ---

  try:
  # 1. Define o seletor. 
  #  Usar CSS_SELECTOR é mais fácil para classes.
  #  O ponto (.) significa "classe".
    seletor_css = ".app-header__app-menu"
  
    print(f"Procurando todos os elementos com a classe: {seletor_css}")
# 2. Espera até que PELO MENOS 2 elementos estejam presentes
    #  (Você pode mudar o '2' para quantos você espera)
    WebDriverWait(driver, 10).until(
      lambda d: len(d.find_elements(By.CSS_SELECTOR, seletor_css)) >= 2
    )

    # 3. Pega a LISTA de todos os elementos
    elementos = driver.find_elements(By.CSS_SELECTOR, seletor_css)
  
    print(f"Encontrados {len(elementos)} elementos.")

    # 4. Clica no primeiro elemento (índice 0)
    if len(elementos) > 0:
      print("Clicando no primeiro elemento (índice 0)...")
      elementos[0].click()
  
    # 5. Espera a página reagir
    #  (MUITO IMPORTANTE: Clicar em algo pode mudar a página)
    print("Aguardando 2 segundos para a página/menu reagir...")
    time.sleep(2) 

    # 6. RE-ENCONTRA a lista de elementos
    #  (É a forma mais segura, caso o primeiro clique tenha
    #  recarregado os elementos - evita o erro 'stale element')
  
    print("Re-encontrando os elementos (para segurança)...")
    elementos = driver.find_elements(By.CSS_SELECTOR, seletor_css)

    # 7. Clica no segundo elemento (índice 1)
    if len(elementos) > 1:
      print("Clicando no segundo elemento (índice 1)...")
      elementos[1].click()
    else:
      print("Erro: Não foi possível encontrar o segundo elemento após o primeiro clique.")
    
    print("Ações nos dois elementos concluídas!")
    time.sleep(3)

  except Exception as e:
    print(f"Ocorreu um erro: {e}")
    # driver.save_screenshot("erro_multiplos.png")


# --- FIM DA ETAPA 8 ---


# --- ETAPA 9 ---
  print("Etapa 9: Aguardando o menu dropdown abrir...")
  
  # Usando o seletor CSS (mais limpo) que discutimos
  exportar_excel_selector = "a.app-box-supermenu-v2__link.app-export-excel"
  
  exportar_link = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, exportar_excel_selector))
  )
  
  print("Link 'Exportar Excel' encontrado. Clicando...")
  exportar_link.click()
  # --- FIM DA ETAPA 9 ---

  time.sleep(3)

# --- ETAPA 10 ---
  try:
    print("Procurando o ícone de 'Notificação' (Inbox)...")
  
    # Usando o Seletor CSS (recomendado por ser mais limpo)
    notificacao_selector = "li.navigation-link.inbox"
  
    # Espera o ícone estar presente e ser clicável
    notificacao_icon = WebDriverWait(driver, 5).until(
      EC.element_to_be_clickable((By.CSS_SELECTOR, notificacao_selector))
    )
  
    print("Ícone de 'Notificação' encontrado. Clicando...")
    notificacao_icon.click()
  
    time.sleep(1) # Espera o menu de notificação abrir

  except Exception as e:
    print(f"Erro ao tentar clicar no ícone de Notificação: {e}")
    driver.save_screenshot("erro_notificacao.png")
  # --- FIM DA ETAPA 10 ---


# --- ETAPA 11 ---
  css_corrigido = "a.PodioUI__Notifications__NotificationGroup"
  item_notificacao = WebDriverWait(driver, 3).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, css_corrigido))
  )
  item_notificacao.click()
# --- FIM DA ETAPA 11 ---

# --- ETAPA 12: Esperar o processamento do Excel ---
  print("Etapa 12: Aguardando a página de exportação carregar e o status ser 'Completado'...")
  
  # Espera até 3 minutos (180s) para o Excel ser processado
  # A espera procura por um <span> verde que contém "Completado"
  status_completado_xpath = "//div[contains(@class, 'field-type-text')]"
  
  WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, status_completado_xpath))
  )
  
  print("Exportação 'Completado'!")
  # --- FIM DA ETAPA 12 ---
  
  # --- ETAPA 13 (NOVA): Clicar no link de download ---
  print("Etapa 13: Procurando o link de download do arquivo...")
  
  # O seletor By.LINK_TEXT é perfeito para isso, pois busca o texto exato do link
  nome_do_arquivo = "Mensageria - Última vista usada.xlsx"
  
  link_download = WebDriverWait(driver, 30).until(
    EC.element_to_be_clickable((By.LINK_TEXT, nome_do_arquivo))
  )
  
  print("Link encontrado! Clicando para baixar...")
  link_download.click()
  
  print("Ação final concluída com sucesso! O download deve começar.")
  time.sleep(10) # Espera 10s para o download iniciar

except Exception as e:
  print(f"Erro durante a navegação no Podio (Etapas 2-5): {e}")
  print("Verifique os seletores XPath. Um deles pode ter mudado.")
  driver.save_screenshot("erro_de_navegacao.png")


  '''  print("\n--- INICIANDO PROCESSAMENTO DO EXCEL ---")
  
  # 1. Encontra o arquivo baixado
  download_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
  
  # Procura por "Mensageria - Última vista usada*.xlsx"
  list_of_files = glob.glob(os.path.join(download_dir, 'Mensageria - Última vista usada*.xlsx'))
  if not list_of_files:
    raise Exception("Arquivo Excel não encontrado na pasta Downloads.")
    
  # Pega o mais recente (caso tenha ' (1)', ' (2)', etc.)
  latest_file = max(list_of_files, key=os.path.getctime)
  print(f"Arquivo encontrado: {latest_file}")
  '''


driver = webdriver.Chrome()
wait = WebDriverWait(driver, 70) # Define um tempo de espera máximo (70 seg)


# ----------- Parte 1 Agilis ----------- # 
# DEFINA SEUS DADOS DE LOGIN E A URL INICIAL
URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
try:
  # 0. ABRIR A PÁGINA E FAZER LOGIN
  driver.get(URL_INICIAL)
  print(f"Página aberta: {URL_INICIAL}")
  print("Aguardando tela de login...")

  # O seletor mais provável para esse botão é pelo texto.
  # Tentativa 1: Usando By.LINK_TEXT (se for uma tag <a>)
  try:
    selector_login_integrado = (By.LINK_TEXT, "Login Integrado Microsoft")
    wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
    
  # Tentativa 2: Usando By.XPATH (funciona para <button>, <div>, <span>, etc.)
  except:
    print("Não encontrou por LINK_TEXT. Tentando por XPATH...")
    # Este XPATH procura QUALQUER elemento que tenha o texto exato.
    selector_login_integrado = (By.XPATH, "//*[text()='Login Integrado Microsoft']")
    wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()

  print("0. Cliquei em 'Login Integrado Microsoft'.")
  print("Aguardando autenticação SSO e carregamento da página principal...")

  # Preenche o e-mail (usando a variável segura)
  email_field_microsoft = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "i0116")) 
  )
  print("Preenchendo e-mail da Microsoft...")
  email_field_microsoft.send_keys("pedro.henrsilva@mrv.com.br")
  WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
  
  # Preenche a senha (usando a variável segura)
  password_field_microsoft = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "i0118")) 
  )
  print("Preenchendo senha da Microsoft...")
  password_field_microsoft.send_keys(" ")#Coloque a sua senha aqui
  
  # Tenta clicar no botão "Entrar" (com loop anti-stale)
  print("Procurando o botão 'Entrar'...")
  tentativas = 0
  clicado_entrar = False
  while not clicado_entrar and tentativas < 5:
    try:
      entrar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
      entrar_button.click()
      clicado_entrar = True
      print("Botão 'Entrar' clicado.")
    except StaleElementReferenceException:
      tentativas += 1; time.sleep(0.5)
  if not clicado_entrar: raise Exception("Falha ao clicar em Entrar")

  # --- ESPERA PELO MFA MANUAL ---
  print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
  print("Aguardando aprovação do MFA no seu celular (até 180s)...")
  
  tentativas = 0
  clicado_manter = False
  while not clicado_manter and tentativas < 5:
    try:
      keep_logged_in_button = WebDriverWait(driver, 180).until( 
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
      )
      keep_logged_in_button.click() # Clica "Sim"
      clicado_manter = True
      print("MFA Aprovado! Botão 'Manter conectado' clicado.")
    except StaleElementReferenceException:
      tentativas += 1; time.sleep(0.5)
    except TimeoutException:
      print("Erro: Timeout após 180s. Você não aprovou o MFA a tempo?")
      clicado_manter = False; break
  if not clicado_manter: raise Exception("Falha ao clicar em Manter Conectado")

  print("Login da Microsoft concluído na janela pop-up.")

  # --- O RESTO DO SEU SCRIPT CONTINUA DAQUI ---

  # 1. CLICAR EM 'RELATÓRIOS'
  # Aguarda o menu 'Relatórios' ficar visível após o login
  
  selector_relatorios = (By.LINK_TEXT, "Relatórios")
  wait.until(EC.element_to_be_clickable(selector_relatorios)).click()
  print("1. Cliquei em 'Relatórios'.")
  print("2. Navegando no menu...")

  # 2. IR EM "CONTRATOS - ADM" -> "PRODUTIVIDADE CONTRATOS - ADM"
  # (Provavelmente precisa clicar no primeiro para o segundo aparecer)
  selector_contratos_adm = (By.LINK_TEXT, "Contratos - ADM")
  wait.until(EC.element_to_be_clickable(selector_contratos_adm)).click()
  print("  - Cliquei em 'Contratos - ADM'.")

  selector_produtividade = (By.LINK_TEXT, "Produtividade Contratos - ADM")
  wait.until(EC.element_to_be_clickable(selector_produtividade)).click()
  print("  - Cliquei em 'Produtividade Contratos - ADM'.")

  # 3. APERTAR "EDITAR"
  # (Pode ser por ID, NOME, ou texto. 'name' é um bom chute em apps Zoho)
  selector_editar = (By.CLASS_NAME, "linkborder") # Chute, pode ser "Editar"
  wait.until(EC.element_to_be_clickable(selector_editar)).click()
  print("3. Cliquei em 'Editar'.")

  # 4. PASSO 1 - SELECIONAR "COLETOR DE CUSTO ADM" E MOVER
  # Selecionar o item na lista da esquerda (pelo texto)
  selector_coletor = (By.XPATH, "//option[text()='Coletor de custo ADM']")
  wait.until(EC.element_to_be_clickable(selector_coletor)).click()
  print("4. Selecionei 'Coletor de custo ADM'.")
  

  # O seletor dele provavelmente é um 'class' ou 'onclick'
  selector_seta_direita = (By.CLASS_NAME, "moverightButton") # CHUTE!
  driver.find_element(*selector_seta_direita).click()
  print("  - Cliquei na seta para mover.")

  print("4.5. Expandindo 'Passo 2: Opções de filtragem'...")
  try:
    # O ID 'reportstep2' foi confirmado pela sua imagem do Inspecionar
    selector_opcoes_filtragem = (By.ID, "rcstep2src")
    # Encontra o elemento que você quer clicar (use "reportstep2" que é melhor)
    elemento_clique = wait.until(EC.presence_of_element_located((By.ID, "rcstep2src")))

    # Usa JavaScript para forçar o clique
    driver.execute_script("arguments[0].click();", elemento_clique)

    print("  - SUCESSO: Cliquei em 'Opções de filtragem' (via JavaScript).")
    
    # Pequena pausa para a animação de expandir terminar
    time.sleep(1) 

  except TimeoutException:
    print("  - FALHA: Não foi possível encontrar 'Passo 2: Opções de filtragem' (ID: reportstep2).")
    # Se este passo falhar, o próximo (clicar no rádio) também vai falhar.
    raise # 'raise' vai parar o script e pular para o bloco 'except Exception'

  # --- PASSO 5: Selecionar o rádio 'Durante' ---
  print("5. Selecionando o filtro 'Durante'...")
  try: 
    # Usando CSS_SELECTOR para encontrar pelo atributo [value='predefined']
    selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
    
    # Espera o rádio ficar clicável
    wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
    print("  - SUCESSO: Filtro 'Durante' selecionado.")

  except TimeoutException:
    print("  - FALHA: Não foi possível encontrar o rádio 'Durante' (CSS_SELECTOR: input[value='predefined']).")
    raise # Para o script se não encontrar

  # 6. APERTAR "EXECUTAR RELATÓRIO"
  selector_executar = (By.ID, "addnew223222")
  wait.until(EC.element_to_be_clickable(selector_executar)).click()
  print("6. Cliquei em 'Executar relatório'.")
  print("--- Relatório executado, aguardando 10s para carregar...")
  time.sleep(3) # Pausa importante para o relatório carregar

  # --- 7. Enviar Relatório por E-mail ---
  print("7. Iniciando envio de e-mail...")
  try:  
    # 7. APERTAR "ENVIAR POR EMAIL ESTE ARQUIVO"
    SELECIONAR_ARQ_EMAIL = (By.ID, "sendmaillink") 
    wait.until(EC.element_to_be_clickable(SELECIONAR_ARQ_EMAIL)).click()
    print("7. Cliquei em 'Enviar este arquivo por email'.")

    # 7.1 - Aguardar o pop-up (modal) de e-mail aparecer
    # (Estamos esperando o botão "Enviar" de DENTRO do pop-up ficar visível)
    selector_enviar_modal = (By.CSS_SELECTOR, "input[value='Enviar']")
    wait.until(EC.element_to_be_clickable(selector_enviar_modal))
    print("  - Modal de e-mail aberto.")
    
    # 7.2 - Selecionar "XLS" no dropdown "Formato"
    select_element = wait.until(EC.element_to_be_clickable((By.ID, "file_type")))
    select_obj = Select(select_element)
    
    select_obj.select_by_visible_text("XLS")
    print("  - Formato 'XLS' selecionado.")

    # 7.3 - Preencher o campo "Para"
    email_para = "pedro.henrsilva@mrv.com.br"
    wait.until(EC.element_to_be_clickable((By.ID, "toEmailSearch"))).send_keys(email_para)
    print(f"  - E-mail preenchido: {email_para}")

    # 7.4 - Clicar no botão "Enviar" final
    driver.find_element(*selector_enviar_modal).click()
    print("  - E-mail enviado com sucesso!")

    # Pausa para ver o resultado antes de fechar
    time.sleep(20) 

  except TimeoutException:
    print("  - FALHA: Não foi possível encontrar um dos elementos do modal de e-mail.")
    print("  - Verifique os seletores ")
    pass # Continua mesmo se falhar
  
  print("--- Automação concluída com sucesso! ---")
  time.sleep(20) # Pausa para você ver o resultado

except Exception as e:
  print(f"ERRO: A automação falhou.")
  print(e)

finally:
  driver.quit() # Comente "driver.quit()" para o navegador não fechar no final 
  print("Script finalizado.")
import time
import os
import glob
import win32com.client
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta # Para lidar com datas (hoje e ontem)
from selenium.webdriver.support.ui import Select # Para caixas de <select>
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
from selenium.webdriver.support.ui import WebDriverWait, Select

# --- CONFIGURAÇÕES ---
EMAIL_USER = "pedro.henrsilva@mrv.com.br"
SENHA_USER = " " 
WAIT_TIME = 20

# --- FUNÇÃO DE APOIO: LOGIN MICROSOFT ---
def fazer_login_microsoft(driver, wait, email, senha):
  """Lida com o login da MS. Retorna True se logou, False se der erro."""
  print("--- Iniciando rotina de Login Microsoft ---")
  try:
    try:
      email_field = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
      print("Preenchendo e-mail...")
      email_field.send_keys(email)
      wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
      
      password_field = wait.until(EC.presence_of_element_located((By.ID, "i0118")))
      print("Preenchendo senha...")
      password_field.send_keys(senha)
      
      clicked = False
      for _ in range(3):
        try:
          wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
          clicked = True
          break
        except StaleElementReferenceException:
          time.sleep(1)
      if not clicked: raise Exception("Não clicou em Entrar")

      print("!!! AGUARDANDO APROVAÇÃO MFA (Se necessário) !!!")
      wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click() 
      print("Login Microsoft efetuado.")
    except TimeoutException:
      print("Campo de login não apareceu. Assumindo que já estamos logados (SSO).")
    return True
  except Exception as e:
    print(f"Erro no Login Microsoft: {e}")
    return False

# --- INÍCIO DO SCRIPT ---
try:
  driver = webdriver.Chrome()
  driver.maximize_window()
  wait = WebDriverWait(driver, WAIT_TIME)

  # ==============================================================================
  # PARTE 1: PODIO
  # ==============================================================================
  print("\n=== INICIANDO PARTE 1: PODIO ===")
  driver.get("https://podio.com/login")

  try:
    wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))).click()
  except: pass 

  wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-provider='live']"))).click()

  janela_principal = driver.current_window_handle
  wait.until(EC.number_of_windows_to_be(2))
  
  for handle in driver.window_handles:
    if handle != janela_principal:
      driver.switch_to.window(handle)
      break

  fazer_login_microsoft(driver, wait, EMAIL_USER, SENHA_USER)
  driver.switch_to.window(janela_principal)
  
  print("Navegando no Podio...")
  menu_area = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'space-switcher-wrapper')]")))
  ActionChains(driver).move_to_element(menu_area).perform()
  wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'ADM - Núcleo Contratos')]"))).click()
  wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-app-id='22830484']"))).click()

  # --- CORREÇÃO DO FILTRO PODIO ---
  print("Aplicando filtros (Método Robusto)...")
  time.sleep(3) # Espera script da página carregar
  
  # 1. Pega o container UL
  ul_filtros = wait.until(EC.presence_of_element_located((By.XPATH, "//ul[@class='app-filter-tools']")))
  
  # 2. Pega todos os itens LI dentro dele
  itens_lista = ul_filtros.find_elements(By.TAG_NAME, "li")
  
  # 3. Passa o mouse em CADA item para garantir que o menu acorde
  actions = ActionChains(driver)
  for item in itens_lista:
    actions.move_to_element(item)
  actions.perform()
  
  # 4. Agora clica no filtro
  wait.until(EC.element_to_be_clickable((By.XPATH, ".//li[@data-original-title='Filtros']"))).click()
  wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-id='created_on']"))).click() 
  wait.until(EC.element_to_be_clickable((By.XPATH, "//li[@data-id='-1mr:-1mr']"))).click() 

  # Seleção de Views
  print("Ajustando visualização...")
  elementos_menu = wait.until(lambda d: d.find_elements(By.CSS_SELECTOR, ".app-header__app-menu"))
  if len(elementos_menu) >= 1: elementos_menu[0].click()
  time.sleep(2)
  elementos_menu = driver.find_elements(By.CSS_SELECTOR, ".app-header__app-menu") 
  if len(elementos_menu) > 1: elementos_menu[1].click()
  time.sleep(2)

  print("Exportando Excel...")
  wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.app-box-supermenu-v2__link.app-export-excel"))).click()
  wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "li.navigation-link.inbox"))).click()

  time.sleep(10)

  wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.PodioUI__Notifications__NotificationGroup"))).click()
  wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'field-type-text')]"))) 
  wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Mensageria - Última vista usada.xlsx"))).click()
  print("Download Podio iniciado!")
  time.sleep(5) 

  # ==============================================================================
  # PARTE 2: AGILIS
  # ==============================================================================
  print("\n=== INICIANDO PARTE 2: AGILIS ===")
  driver.get("https://agilis.mrv.com.br/HomePage.do?view_type=my_view")

  try:
    btn_login = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='Login Integrado Microsoft']")))
    btn_login.click()
    fazer_login_microsoft(driver, wait, EMAIL_USER, SENHA_USER)
  except TimeoutException:
    print("Botão de login não apareceu, seguindo...")

  print("Navegando menus Agilis...")
  wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Relatórios"))).click()
  wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Contratos - ADM"))).click()
  wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Produtividade Contratos - ADM"))).click()
  
  wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "linkborder"))).click() 
  wait.until(EC.element_to_be_clickable((By.XPATH, "//option[text()='Coletor de custo ADM']"))).click()
  driver.find_element(By.CLASS_NAME, "moverightButton").click()

  try:
    expand_btn = wait.until(EC.presence_of_element_located((By.ID, "rcstep2src")))
    driver.execute_script("arguments[0].click();", expand_btn)
    time.sleep(1)
  except: pass

  # --- SELEÇÃO DE DATA (DROPDOWN + RÁDIO) ---
  print("Selecionando Data 'Mês Passado' no dropdown...")
  select_elem = wait.until(EC.presence_of_element_located((By.ID, "dateFilterType")))
  Select(select_elem).select_by_visible_text("Mês passado")
  
  # [INSERIDO] Código do Rádio 'Durante' que você pediu
  print("Selecionando o rádio 'Durante' (Ajuste obrigatório)...")
  try:
    selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
    wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
    print(" - SUCESSO: Filtro 'Durante' selecionado.")
  except TimeoutException:
    print(" - FALHA: Rádio 'Durante' não encontrado.")
    raise

  # Executar Relatório
  wait.until(EC.element_to_be_clickable((By.ID, "addnew223222"))).click() 
  print("Relatório gerando. Aguardando 10 segundos...")
  time.sleep(10) 

  # --- CORREÇÃO DO CLIQUE "ENVIAR POR E-MAIL" ---
  print("Tentando clicar em 'Enviar por e-mail'...")
  try:
    # 1. Localiza o botão
    btn_email = wait.until(EC.presence_of_element_located((By.ID, "sendmaillink")))
    
    # 2. Rola a tela até ele (garantia)
    driver.execute_script("arguments[0].scrollIntoView(true);", btn_email)
    time.sleep(1)

    # 3. CLIQUE FORÇADO VIA JAVASCRIPT (Resolve o "Element click intercepted")
    driver.execute_script("arguments[0].click();", btn_email)
    print(" - Botão clicado com JavaScript!")
  
  except Exception as e:
    print(f"Erro ao clicar em enviar email: {e}")
    raise

  # Modal Email
  wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[value='Enviar']"))) 
  Select(driver.find_element(By.ID, "file_type")).select_by_visible_text("XLS")
  driver.find_element(By.ID, "toEmailSearch").send_keys(EMAIL_USER)
  driver.find_element(By.CSS_SELECTOR, "input[value='Enviar']").click()
  
  print("E-mail Agilis enviado com sucesso!")
  time.sleep(5)

except Exception as e:
  print(f"\nCRITICAL ERROR: {e}")
  driver.save_screenshot("erro_recente.png")

finally:
  print("Fim.")
  # driver.quit()
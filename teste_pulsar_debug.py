import os
import time
import re
from dotenv import load_dotenv

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

load_dotenv()

# --- CONFIGURAÇÕES ---
PULSAR_EMAIL = os.getenv("PULSAR_EMAIL", "fiscal.pulsar@4cta.eb.mil.br")
PULSAR_PASSWORD = os.getenv("PULSAR_PASSWORD", "K809(F4a[?")
LOGIN_URL = "https://sport.pulsarconnect.io/login"
STARLINK_URL = "https://sport.pulsarconnect.io/starlink/starlinkMap"

def create_driver():
    options = Options()
    # --- MODO HEADLESS (VM) ---
    options.add_argument("--headless=new") 
    # Tamanho de janela é CRUCIAL para o Hover funcionar em headless
    options.add_argument("--window-size=1920,1080")
    
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--log-level=3")
    
    # User agent para evitar bloqueios
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    
    try:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    except:
        return webdriver.Chrome(options=options)

def perform_login(driver):
    print(">>> Módulo Pulsar: Fazendo login...")
    try:
        driver.get(LOGIN_URL)
        time.sleep(5)

        email_input = driver.find_element(By.NAME, "userName")
        pass_input = driver.find_element(By.NAME, "password")
        
        email_input.clear()
        email_input.send_keys(PULSAR_EMAIL)
        pass_input.clear()
        pass_input.send_keys(PULSAR_PASSWORD)

        login_button = driver.find_element(By.XPATH, "//button[contains(@class, 'loginButton')]")
        login_button.click()
        time.sleep(15)
        return True
    except Exception as e:
        print(f"❌ Erro Login Pulsar: {e}")
        return False

def apply_date_filter(driver):
    print(">>> Módulo Pulsar: Aplicando filtro 'Last 1 Day'...")
    try:
        time.sleep(8)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//button")))
        
        date_buttons = driver.find_elements(By.XPATH, "//button")
        clicked_filter = False
        
        for btn in date_buttons:
            btn_text = btn.text.strip()
            if 'Day' in btn_text or 'MTD' in btn_text:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.5)
                    btn.click()
                    time.sleep(2)
                    clicked_filter = True
                    break
                except: continue
        
        if clicked_filter:
            try:
                time.sleep(1)
                last_1_day = driver.find_element(By.XPATH, "//*[text()='Last 1 Day']")
                last_1_day.click()
                time.sleep(1.5)
                
                apply_btns = driver.find_elements(By.XPATH, "//button[text()='Apply' or contains(text(), 'Apply')]")
                for apply_btn in apply_btns:
                    if apply_btn.is_displayed():
                        apply_btn.click()
                        break
                
                time.sleep(8) 
                return True
            except:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    except Exception as e:
        print(f"⚠ Erro filtro: {e}")
    return False

def scrape_data(driver):
    """
    Lógica do seu script 'EXTRATOR STARLINK' com a detecção de paginação via Regex.
    """
    all_data = []
    kit_ids_processed = set()
    identifiers_processed = set()
    
    print(">>> Módulo Pulsar: Iniciando extração (Modo Original)...")

    # 1. Zoom Out
    try: driver.execute_script("document.body.style.zoom='75%'")
    except: pass
    time.sleep(2)

    # 2. Contar total de itens (Lógica do seu script)
    print("    📊 Contando total de itens na tabela...")
    total_items_expected = 0
    try:
        # Procura textos como "1-25 of 44"
        pagination_texts = driver.find_elements(By.XPATH, "//*[contains(text(), 'of') or contains(text(), 'de') or contains(text(), '–')]")
        for text_elem in pagination_texts:
            text = text_elem.text
            # Seus padrões de regex originais
            patterns = [
                r'of\s+(\d+)', 
                r'de\s+(\d+)', 
                r'(\d+)\s*[-–]\s*(\d+)\s+of\s+(\d+)', 
                r'(\d+)\s*[-–]\s*(\d+)\s+de\s+(\d+)'
            ]
            for pattern in patterns:
                match = re.search(pattern, text)
                if match:
                    # Pegar o último número capturado (que é o total)
                    numbers = [int(n) for n in match.groups() if n]
                    if numbers:
                        total_items_expected = max(numbers)
                        print(f"    ✔ Total de itens detectado: {total_items_expected}")
                        break
            if total_items_expected > 0: break
    except Exception as e:
        print(f"    ⚠ Erro ao detectar total: {e}")

    # 3. Scroll Topo
    try:
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)
    except: pass

    current_page = 1
    has_next_page = True

    # --- LOOP DE PAGINAÇÃO (Seu Script) ---
    while has_next_page:
        print(f"    Processando página {current_page}...", end="\r")
        
        # Buscar linhas
        rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
        if not rows:
            rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'MuiTableRow') and .//td]")
        
        if not rows:
            print("\n    ❌ Nenhuma linha encontrada nesta página.")
            break

        for idx, row in enumerate(rows):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) < 2: continue

                om_name = cells[0].text.strip()
                if not om_name or "SERVICE LINE" in om_name.upper() or "NO SERVICE" in om_name.upper() or len(om_name) < 3:
                    continue

                status_cell = cells[1]
                # Seletor universal de SVG
                svgs = status_cell.find_elements(By.XPATH, ".//svg | .//*[name()='svg']")
                
                kit_id = ""
                status_cor = "GRAY"
                status_text = "UNKNOWN"

                if len(svgs) >= 2:
                    segundo_svg = svgs[1]
                    
                    # --- Scroll Suave ---
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", segundo_svg)
                        time.sleep(0.8)
                        
                        tooltip_found = False
                        max_tentativas = 5 # Seu script usa 5
                        
                        for _ in range(max_tentativas):
                            if tooltip_found: break
                            
                            actions = ActionChains(driver)
                            actions.move_to_element(segundo_svg).perform()
                            time.sleep(3.5)
                            
                            # --- Estratégias de Tooltip (Seu script) ---
                            tooltip_text = ""
                            
                            # 1. Role Tooltip
                            try:
                                tooltips = driver.find_elements(By.XPATH, "//div[@role='tooltip']")
                                visibles = [t for t in tooltips if t.is_displayed() and t.size['height'] > 0]
                                if visibles: 
                                    tooltip_text = visibles[-1].text.strip()
                                    if "KIT" in tooltip_text:
                                        for word in tooltip_text.replace("\n", " ").split():
                                            if word.startswith("KIT"):
                                                kit_id = word
                                                tooltip_found = True
                                                break
                            except: pass

                            # 2. Classes
                            if not tooltip_found:
                                try:
                                    tooltips = driver.find_elements(By.XPATH, "//*[contains(@class, 'tooltip') or contains(@class, 'Popper')]")
                                    visibles = [t for t in tooltips if t.is_displayed()]
                                    if visibles: 
                                        tooltip_text = visibles[-1].text.strip()
                                        if "KIT" in tooltip_text:
                                            for word in tooltip_text.replace("\n", " ").split():
                                                if word.startswith("KIT"):
                                                    kit_id = word
                                                    tooltip_found = True
                                                    break
                                except: pass
                            
                            if not tooltip_found:
                                # Move mouse e tenta de novo
                                ActionChains(driver).move_by_offset(200, 0).perform()
                                time.sleep(0.5)

                        # Análise de COR
                        svg_html = segundo_svg.get_attribute('outerHTML').lower()
                        if any(c in svg_html for c in ['green', '#00ff00', '#0f0', '#00e676', '#4caf50']):
                            status_cor = "GREEN"
                            status_text = "ONLINE"
                        elif any(c in svg_html for c in ['red', '#ff0000', '#f00', '#f44336', '#e53935']):
                            status_cor = "RED"
                            status_text = "OFFLINE"
                        else:
                            status_cor = "YELLOW"
                            status_text = "WARNING"
                        
                        # Limpa mouse
                        ActionChains(driver).move_by_offset(100, 100).perform()
                        
                    except: pass

                # --- DEDUPLICAÇÃO (Seu script) ---
                # Se KIT ID já foi processado, ignora
                if kit_id and kit_id in kit_ids_processed:
                    continue
                
                # Identificador
                identificador = f"{om_name}|{kit_id}" if kit_id else f"{om_name}|{idx}"
                if identificador in identifiers_processed:
                    continue
                
                identifiers_processed.add(identificador)
                if kit_id: kit_ids_processed.add(kit_id)

                # Adiciona dados
                all_data.append({
                    "om": om_name,
                    "pop": kit_id if kit_id else "N/A",
                    "status": status_text,
                    "cor": status_cor
                })

            except Exception: continue

        # --- PAGINAÇÃO (Seu Script) ---
        try:
            # Procura botão Next
            next_btn = driver.find_element(By.XPATH, "//button[@aria-label='Next page' or contains(@aria-label, 'next') or contains(@class, 'next')]")
            
            if next_btn.get_attribute('disabled'):
                has_next_page = False
                print(f"\n    ✔ Última página alcançada (página {current_page})")
            else:
                # Clica e espera
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
                next_btn.click()
                print(f"\n    ➡️ Indo para página {current_page + 1}...")
                time.sleep(8) # Tempo generoso para carregar
                current_page += 1
                
                # Wait explícito
                try:
                    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//table//tbody//tr[td]")))
                except: time.sleep(5)
                
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(2)
        except:
            has_next_page = False
            print(f"\n    ✔ Não há mais páginas (total: {current_page})")

    return all_data, total_items_expected

# --- FUNÇÃO PRINCIPAL ---

def collect_pulsar_data():
    """
    Função chamada pelo main.py.
    """
    driver = None
    data = []
    
    print(">>> Módulo Pulsar: Iniciando...")
    try:
        driver = create_driver()
        if perform_login(driver):
            print(">>> Módulo Pulsar: Acessando Mapa...")
            driver.get(STARLINK_URL)
            apply_date_filter(driver)
            
            # Recebe dados e total
            data, total_esperado = scrape_data(driver)
            
            print(f">>> Pulsar: Extraídos {len(data)} de {total_esperado} esperados.")
            
            # Validação simples de total
            if total_esperado > 0 and len(data) < total_esperado:
                print(f"    ⚠ AVISO: Faltaram {total_esperado - len(data)} itens!")
        else:
            data = [{"om": "FALHA LOGIN PULSAR", "pop": "-", "status": "ERRO", "cor": "RED"}]
            
    except Exception as e:
        print(f"Erro crítico Pulsar: {e}")
        data = [{"om": "ERRO DE EXECUÇÃO", "pop": str(e)[:30], "status": "ERRO", "cor": "RED"}]
    finally:
        if driver: driver.quit()
    
    return data

if __name__ == "__main__":
    collect_pulsar_data()
"""
EXTRATOR STARLINK - VERSÃO AUTOMÁTICA COM PAGINAÇÃO
Otimizado para Ubuntu Server (sem display)
"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from dotenv import load_dotenv

USER_EMAIL = os.getenv("PULSAR_EMAIL")
USER_PASSWORD = os.getenv("PULSAR_PASSWORD")
LOGIN_URL = "https://sport.pulsarconnect.io/login"
STARLINK_URL = "https://sport.pulsarconnect.io/starlink/starlinkMap"

load_dotenv()

def get_chrome_options(headless=True):
    """Configura opções do Chrome para ambiente server"""
    options = Options()
    
    if headless:
        print("   🖥️  Modo headless ativado")
        
        # Localizar Chromium no Ubuntu (ANTES de configurar opções)
        possible_paths = [
            '/usr/bin/chromium',
            '/usr/bin/chromium-browser',
            '/snap/bin/chromium'
        ]
        
        chromium_found = False
        for path in possible_paths:
            if os.path.exists(path):
                options.binary_location = path
                print(f"   ✓ Chromium: {path}")
                chromium_found = True
                break
        
        if not chromium_found:
            print("   ⚠️  Chromium não encontrado nos caminhos padrão")
        
        # Opções ESSENCIAIS para headless em server (testadas e funcionando)
        options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--single-process')
        options.add_argument('--disable-setuid-sandbox')
        
        # Configurações de janela
        options.add_argument('--window-size=1920,1080')
        
        # Otimizações adicionais
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')
        
    else:
        options.add_argument('--start-maximized')
    
    return options

def get_chrome_driver():
    """Localiza o ChromeDriver no sistema"""
    possible_paths = [
        '/usr/bin/chromedriver',
        '/usr/local/bin/chromedriver',
        '/snap/bin/chromedriver'
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            print(f"   ✓ ChromeDriver: {path}")
            return path
    
    return None

def check_chrome_installation():
    """Verifica se Chrome/Chromium e ChromeDriver estão instalados"""
    errors = []
    
    # Verificar Chromium
    chromium_paths = ['/usr/bin/chromium', '/usr/bin/chromium-browser', '/snap/bin/chromium']
    chromium_found = any(os.path.exists(p) for p in chromium_paths)
    
    if not chromium_found:
        errors.append("Chromium não encontrado. Instale: sudo apt install chromium-browser")
    
    # Verificar ChromeDriver
    driver_paths = ['/usr/bin/chromedriver', '/usr/local/bin/chromedriver', '/snap/bin/chromedriver']
    driver_found = any(os.path.exists(p) for p in driver_paths)
    
    if not driver_found:
        errors.append("ChromeDriver não encontrado. Instale: sudo apt install chromium-chromedriver")
    
    return errors

def extrair_dados_starlink(headless=False):
    print("="*75)
    print(" "*20 + "EXTRATOR STARLINK")
    print("="*75)

    # Verificação prévia
    if headless:
        check_errors = check_chrome_installation()
        if check_errors:
            print("\n   ❌ PROBLEMAS DETECTADOS:")
            for error in check_errors:
                print(f"   • {error}")
            print("\n   💡 Execute o teste primeiro: python3 test_chrome_ubuntu.py")
            raise RuntimeError("Chrome/ChromeDriver não está configurado corretamente")

    # Configurar opções
    options = get_chrome_options(headless)
    
    # Iniciar driver
    try:
        driver_path = get_chrome_driver()
        
        if driver_path:
            service = Service(executable_path=driver_path)
            driver = webdriver.Chrome(service=service, options=options)
        else:
            print("   ⚠️  ChromeDriver não encontrado no sistema")
            print("   📥 Tentando baixar automaticamente...")
            from webdriver_manager.chrome import ChromeDriverManager
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=options
            )
        
        print("   ✅ Navegador iniciado com sucesso!")
        
    except Exception as e:
        error_msg = str(e)
        print(f"\n   ❌ ERRO ao iniciar navegador:")
        print(f"   {error_msg[:200]}...\n")
        
        print("   💡 SOLUÇÕES:")
        print("   1. Teste primeiro: python3 test_chrome_ubuntu.py")
        print("   2. Instale dependências:")
        print("      sudo apt install chromium-browser chromium-chromedriver")
        print("      sudo apt install libgbm1 libnss3 libnspr4 libatk1.0-0")
        print("   3. Verifique instalação:")
        print("      which chromium || which chromium-browser")
        print("      which chromedriver")
        
        driver = None
        raise

    actions = ActionChains(driver)

    print("\n[1/5] Fazendo login...")
    driver.get(LOGIN_URL)
    time.sleep(5)

    try:
        email_input = driver.find_element(By.NAME, "userName")
        pass_input = driver.find_element(By.NAME, "password")
        email_input.send_keys(USER_EMAIL)
        pass_input.send_keys(USER_PASSWORD)

        login_button = driver.find_element(By.XPATH, "//button[contains(@class, 'loginButton')]")
        login_button.click()
        print("   ✓ Login realizado")
        time.sleep(15)
    except Exception as e:
        print(f"   ❌ Erro no login: {e}")
        driver.quit()
        exit(1)

    print("\n[2/5] Navegando para Starlink...")
    driver.get(STARLINK_URL)
    time.sleep(15)

    print("\n[3/5] Aplicando filtro 'Last 1 Day'...")
    try:
        time.sleep(8)
        
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//button"))
        )
        
        date_buttons = driver.find_elements(By.XPATH, "//button")
        
        clicked_filter = False
        for btn in date_buttons:
            btn_text = btn.text.strip()
            if 'Day' in btn_text or 'MTD' in btn_text:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                    time.sleep(0.5)
                    btn.click()
                    print(f"   ✓ Filtro clicado: '{btn_text}'")
                    time.sleep(2)
                    clicked_filter = True
                    break
                except:
                    continue
        
        if clicked_filter:
            try:
                time.sleep(1)
                last_1_day = driver.find_element(By.XPATH, "//*[text()='Last 1 Day']")
                last_1_day.click()
                print("   ✓ 'Last 1 Day' selecionado")
                time.sleep(1.5)
                
                apply_btns = driver.find_elements(By.XPATH, "//button[text()='Apply' or contains(text(), 'Apply')]")
                for apply_btn in apply_btns:
                    if apply_btn.is_displayed():
                        apply_btn.click()
                        print("   ✓ Apply clicado")
                        break
                
                time.sleep(8)
                print("   ✅ Filtro aplicado!")
                
            except Exception as e2:
                print(f"   ⚠️  Erro ao aplicar filtro: {e2}")
                actions.send_keys(Keys.ESCAPE).perform()
                time.sleep(2)
        else:
            print("   ⚠️  Continuando com filtro padrão...")
            time.sleep(5)
        
    except Exception as e:
        print(f"   ❌ ERRO CRÍTICO NO FILTRO: {e}")
        driver.quit()
        exit(1)

    print("\n[4/5] Extraindo dados...")

    print("   ⏳ Aguardando tabela...")
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//table//tbody//tr"))
        )
        time.sleep(5)
    except Exception as e:
        print(f"   ❌ Tabela não carregou: {str(e)[:100]}")
        driver.quit()
        exit(1)

    print("   🔍 Aplicando zoom 50%...")
    try:
        driver.execute_script("document.body.style.zoom='50%'")
        time.sleep(2)
    except:
        pass

    try:
        driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)
    except:
        pass

    all_data = []
    kit_ids_processados = set()
    identificadores_processados = set()
    current_page = 1
    has_next_page = True

    while has_next_page:
        print(f"\n   📄 Página {current_page}...")
        
        rows = driver.find_elements(By.XPATH, "//table//tbody//tr[td]")
        
        if not rows or len(rows) == 0:
            rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'MuiTableRow') and .//td]")
        
        print(f"   ✓ {len(rows)} linhas encontradas\n")
        
        if len(rows) == 0:
            print("   ❌ Nenhuma linha!")
            break
        
        for idx, row in enumerate(rows):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                
                if len(cells) < 2:
                    continue
                
                om = cells[0].text.strip()
                
                if not om or "SERVICE LINE" in om.upper() or "NO SERVICE" in om.upper() or len(om) < 3:
                    continue
                
                status_cell = cells[1]
                svgs = status_cell.find_elements(By.XPATH, ".//svg | .//*[name()='svg']")
                
                kit_id = ""
                status_cor = "DESCONHECIDO"
                
                if len(svgs) >= 2:
                    segundo_svg = svgs[1]
                    
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", segundo_svg)
                        time.sleep(0.8)
                        
                        tooltip_found = False
                        max_tentativas = 5
                        
                        for tentativa in range(max_tentativas):
                            if tooltip_found:
                                break
                            
                            actions = ActionChains(driver)
                            actions.move_to_element(segundo_svg).perform()
                            time.sleep(3.5)
                            
                            try:
                                tooltips = driver.find_elements(By.XPATH, "//div[@role='tooltip']")
                                visible_tooltips = [t for t in tooltips if t.is_displayed() and t.size['height'] > 0 and t.size['width'] > 0]
                                
                                if visible_tooltips:
                                    tooltip = visible_tooltips[-1]
                                    tooltip_text = tooltip.text.strip()
                                    if tooltip_text and "KIT" in tooltip_text:
                                        words = tooltip_text.replace("\n", " ").split()
                                        for word in words:
                                            if word.startswith("KIT"):
                                                kit_id = word
                                                tooltip_found = True
                                                break
                                        if tooltip_found:
                                            break
                            except:
                                pass
                            
                            if not tooltip_found:
                                try:
                                    tooltips = driver.find_elements(By.XPATH, "//*[contains(@class, 'tooltip') or contains(@class, 'Tooltip') or contains(@class, 'Popper')]")
                                    for tooltip in tooltips:
                                        if tooltip.is_displayed() and tooltip.size['height'] > 0:
                                            tooltip_text = tooltip.text.strip()
                                            if tooltip_text and "KIT" in tooltip_text:
                                                words = tooltip_text.replace("\n", " ").split()
                                                for word in words:
                                                    if word.startswith("KIT"):
                                                        kit_id = word
                                                        tooltip_found = True
                                                        break
                                                if tooltip_found:
                                                    break
                                except:
                                    pass
                            
                            if not tooltip_found:
                                try:
                                    parent = segundo_svg.find_element(By.XPATH, "..")
                                    title_attr = parent.get_attribute('title')
                                    aria_label_attr = parent.get_attribute('aria-label')
                                    data_tip = parent.get_attribute('data-tip')
                                    
                                    for attr in [title_attr, aria_label_attr, data_tip]:
                                        if attr and "KIT" in attr:
                                            words = attr.split()
                                            for word in words:
                                                if word.startswith("KIT"):
                                                    kit_id = word
                                                    tooltip_found = True
                                                    break
                                            if tooltip_found:
                                                break
                                except:
                                    pass
                            
                            if not tooltip_found:
                                title_attr = segundo_svg.get_attribute('title')
                                aria_label_attr = segundo_svg.get_attribute('aria-label')
                                data_tip = segundo_svg.get_attribute('data-tip')
                                
                                for attr in [title_attr, aria_label_attr, data_tip]:
                                    if attr and "KIT" in attr:
                                        words = attr.split()
                                        for word in words:
                                            if word.startswith("KIT"):
                                                kit_id = word
                                                tooltip_found = True
                                                break
                                        if tooltip_found:
                                            break
                            
                            if not tooltip_found and tentativa < max_tentativas - 1:
                                actions = ActionChains(driver)
                                actions.move_by_offset(200, 0).perform()
                                time.sleep(0.5)
                        
                        svg_html = segundo_svg.get_attribute('outerHTML').lower()
                        
                        if 'green' in svg_html or '#00ff00' in svg_html or '#0f0' in svg_html or '#00e676' in svg_html or '#4caf50' in svg_html:
                            status_cor = "VERDE"
                        elif 'red' in svg_html or '#ff0000' in svg_html or '#f00' in svg_html or '#f44336' in svg_html or '#e53935' in svg_html:
                            status_cor = "VERMELHO"
                        
                        actions.move_by_offset(100, 100).perform()
                        time.sleep(0.1)
                        
                    except Exception as e:
                        pass
                
                usage = cells[2].text.strip() if len(cells) >= 3 else ""
                
                if kit_id and kit_id in kit_ids_processados:
                    continue
                
                identificador = f"{om}|{kit_id}" if kit_id else f"{om}|{idx}"
                
                if identificador in identificadores_processados:
                    continue
                
                identificadores_processados.add(identificador)
                if kit_id:
                    kit_ids_processados.add(kit_id)
                
                all_data.append({
                    'om': om,
                    'pop': kit_id if kit_id else "N/A",
                    'status': status_cor,
                    'OCORRÊNCIA': ''
                })
                
                emoji = "🟢" if status_cor == "VERDE" else ("🔴" if status_cor == "VERMELHO" else "⚪")
                print(f"   {idx+1:2d}. {om[:32]:<32} | {kit_id:<17} | {emoji}")
                
            except Exception as e:
                continue
        
        try:
            next_button = driver.find_element(By.XPATH, "//button[@aria-label='Next page' or contains(@aria-label, 'next') or contains(@class, 'next')]")
            
            is_disabled = next_button.get_attribute('disabled')
            
            if is_disabled:
                has_next_page = False
                print(f"\n   ✓ Última página ({current_page})")
            else:
                next_button.click()
                print(f"\n   ➡️  Página {current_page + 1}...")
                time.sleep(8)
                current_page += 1
                
                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.XPATH, "//table//tbody//tr[td]"))
                    )
                    time.sleep(3)
                except:
                    time.sleep(5)
                
                driver.execute_script("window.scrollTo(0, 0);")
                time.sleep(2)
                
        except:
            # Se deu erro ao tentar mudar de página, não assuma que acabou. Aborte.
            print(f"\n   ❌ ERRO AO MUDAR DE PÁGINA (Tentando reiniciar...)")
            driver.quit()
            exit(1)

    driver.quit()

    if not all_data:
        print("\n❌ NENHUM DADO EXTRAÍDO!")
        exit(1)

    print(f"\n✅ {len(all_data)} registros extraídos!")
    print(all_data)
    return all_data



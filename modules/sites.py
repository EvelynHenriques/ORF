import hashlib
import re
import json
import os
from urllib.parse import urlparse
from pathlib import Path

# Selenium Imports
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

try:
    from bs4 import BeautifulSoup
    _BS4_AVAILABLE = True
except ImportError:
    _BS4_AVAILABLE = False

HASH_DIR = Path("output/hashes")
HASH_DIR.mkdir(parents=True, exist_ok=True)
JSON_FILE = "sites.json"

def load_sites_list():
    """Lê o arquivo sites.json e retorna a lista de URLs"""
    if not os.path.exists(JSON_FILE):
        print(f"ERRO CRÍTICO: Arquivo {JSON_FILE} não encontrado!")
        return []
    
    try:
        with open(JSON_FILE, 'r', encoding='utf-8') as f:
            lista = json.load(f)
            return lista
    except Exception as e:
        print(f"ERRO ao ler JSON: {e}")
        return []

def extract_om_name(url):
    try:
        parsed = urlparse(url)
        if "licitacoeseb" in parsed.netloc:
            return "LICITAÇÕES 12RM"
            
        domain = parsed.netloc.replace('.eb.mil.br', '').replace('www.', '')
        if len(parsed.path) > 1: domain += " " + parsed.path.replace('/', '').upper()
        return domain.upper()
    except: return "SITE"

def sanitize_filename(url):
    p = urlparse(url)
    name = (p.netloc + p.path).strip()
    name = re.sub(r"[^0-9a-zA-Z.-]", "_", name)
    return name[:120]

def get_hash_file_path(url):
    return HASH_DIR / f"{sanitize_filename(url)}.txt"

def normalize_html(html):
    if _BS4_AVAILABLE:
        soup = BeautifulSoup(html, "html.parser")
        for tag in soup(["script", "style"]): tag.decompose()
        text = soup.get_text(separator=" ", strip=True)
        lines = [line.strip() for line in text.split('\n') if not re.search(r'(\d{1,2}[:/]\d{1,2}[:/]\d{2,4}|\d+:\d+|visualizaç|views?:|acess|visit|online)', line, re.I)]
        text = '\n'.join(lines)
    else:
        text = re.sub(r"<script[\s\S]*?</script>", "", html, flags=re.I)
        text = re.sub(r"<style[\s\S]*?</style>", "", text, flags=re.I)
        text = re.sub(r"<[^>]+>", " ", text)
    return " ".join(text.split())

def extract_image_sets(html):
    image_sets = {}
    if _BS4_AVAILABLE:
        soup = BeautifulSoup(html, "html.parser")
        for selector in ['[class*="carousel"]', '[class*="slider"]', '[class*="banner"]', '[id*="banner"]']:
            for i, el in enumerate(soup.select(selector)):
                imgs = [img.get('src') or img.get('data-src') for img in el.find_all('img')]
                imgs = [u for u in imgs if u and not u.startswith('data:')]
                if imgs:
                    sig = '|'.join(sorted(set(imgs)))
                    image_sets[f"{selector}_{i}"] = hashlib.md5(sig.encode()).hexdigest()
    return image_sets

def create_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--log-level=3")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
    try:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    except:
        return webdriver.Chrome(options=options)

def collect_sites_data():
    """
    Carrega sites do JSON, roda Selenium e retorna dados para o PDF.
    Retorno: (dados_tabela, dados_ocorrencias)
    """
    urls_do_arquivo = load_sites_list()
    
    if not urls_do_arquivo:
        return [], [["-", "ERRO: Lista de sites vazia ou arquivo JSON não encontrado", "-"]]

    SITES_PARA_MONITORAR = [(extract_om_name(url), url) for url in urls_do_arquivo]
    
    driver = None
    resultados = []
    lista_ocorrencias = []
    ocorrencias_unicas = {}
    proximo_id = 1

    print(f">>> Módulo Sites: Iniciando verificação de {len(SITES_PARA_MONITORAR)} endereços...")
    
    try:
        driver = create_driver()
        
        for idx, (om, url) in enumerate(SITES_PARA_MONITORAR, start=1):
            print(f"Checking [{idx}] {om}...", end="\r")
            status_text = "S/A"
            cor_status = "GREEN"
            id_oc = "-"
            msg_erro = None

            try:
                original_url = url
                driver.set_page_load_timeout(30)
                driver.get(url)
                
                try: WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                except: pass 

                final_url = driver.current_url
                orig_p = urlparse(original_url)
                fin_p = urlparse(final_url)
                
                # Lógica Redirecionamento
                if orig_p.netloc.lower() != fin_p.netloc.lower():
                    msg_erro = f"Redirecionado para externo: {fin_p.netloc}"
                else:
                    path_lower = fin_p.path.lower()
                    if "/error" in path_lower or "/404" in path_lower or "pagina-nao-encontrada" in path_lower:
                         msg_erro = f"Página de erro: {fin_p.path}"

                # Lógica Hash/Conteúdo
                if not msg_erro:
                    html = driver.page_source
                    if "404 - File or directory not found" in html:
                        msg_erro = "Erro 404 detectado"
                    else:
                        norm = normalize_html(html)
                        imgs = extract_image_sets(html)
                        comb = f"TEXT:{norm}|IMAGES:{sorted(imgs.items())}"
                        curr_hash = hashlib.sha256(comb.encode("utf-8")).hexdigest()
                        
                        hf = get_hash_file_path(url)
                        if hf.exists():
                            if curr_hash != hf.read_text(encoding="utf-8").strip():
                                msg_erro = "Alteração visual detectada"
                                hf.write_text(curr_hash, encoding="utf-8")
                        else:
                            hf.write_text(curr_hash, encoding="utf-8")

            except Exception as e:
                msg_erro = "Site inacessível / Timeout"
            
            if msg_erro:
                status_text = "C.O."
                cor_status = "RED"
                if msg_erro in ocorrencias_unicas:
                    id_oc = ocorrencias_unicas[msg_erro]
                else:
                    id_oc = str(proximo_id)
                    ocorrencias_unicas[msg_erro] = id_oc
                    lista_ocorrencias.append([id_oc, msg_erro, "-"])
                    proximo_id += 1

            resultados.append([str(idx), om, url, status_text, id_oc, cor_status])

    except Exception as e:
        print(f"\nErro Crítico Selenium: {e}")
    finally:
        if driver: driver.quit()

    print("\n>>> Módulo Sites: Finalizado.")
    return resultados, lista_ocorrencias
import requests
import urllib3
import re
import io
import os
from dotenv import load_dotenv

# Carrega variáveis de ambiente
load_dotenv()

# --- CONFIGURAÇÕES GERAIS DE IMAGEM ---
WIDTH = 1200 
HEIGHT = 200
PERIOD_STRING = "now-24h"

# --- SERVIDOR 1 (Links Principais) ---
# URL: http://10.78.151.97/zabbix (Padrão do .env)
ZABBIX1_URL = os.getenv("ZABBIX_URL")
ZABBIX1_USER = os.getenv("ZABBIX_USER")
ZABBIX1_PASS = os.getenv("ZABBIX_PASS")

GRAPHS_CONFIG_1 = [
    {"id": 3260, "title": "Empresa CLARO (INTERNET)"},
    {"id": 3227, "title": "Empresa NWHERE (INTERNET)"},
    {"id": 3247, "title": "BBI (Metro Ethernet via 6 CTA)"},
    {"id": 3228, "title": "BBI (Metro Ethernet via 7 CTA)"},
    {"id": 3253, "title": "BBI (Metro Ethernet via 41 CT)"}
]

# --- SERVIDOR 2 (EBNET - .210) ---
ZABBIX2_URL = os.getenv("ZABBIX2_URL")
ZABBIX2_USER = os.getenv("ZABBIX_USER") 
ZABBIX2_PASS = os.getenv("ZABBIX_PASS") 

GRAPHS_CONFIG_2 = [
    {"id": 2751, "title": "Tráfego EBNET área CMA", "hostid": 10598}
]

# --- SERVIDOR 3 (Novo - .208) ---
ZABBIX3_URL = os.getenv("ZABBIX3_URL")
ZABBIX3_USER = os.getenv("ZABBIX2_USER") 
ZABBIX3_PASS = os.getenv("ZABBIX2_PASS") 

GRAPHS_CONFIG_3 = [
    {"id": 4833, "title": "Tráfego Vila militar 3º BIS", "hostid": 10318}
]

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- FUNÇÕES DE CONEXÃO GENÉRICAS ---

def get_csrf_token(session, base_url):
    """Varre rotas para encontrar token CSRF no servidor especificado."""
    login_paths = ["/index.php", "/", "/zabbix.php?action=signin", "/index_sso.php"]
    
    base_url = base_url.rstrip('/')
    
    for path in login_paths:
        try:
            resp = session.get(f"{base_url}{path}", verify=False, timeout=10)
            if resp.status_code != 200: continue

            match = re.search(r'name="csrf_token" value="([^"]+)"', resp.text)
            if match:
                return match.group(1)
            
            if 'name="enter"' in resp.text:
                return None
                
        except Exception as e:
            print(f"Erro conexão Zabbix ({base_url}): {e}")
    return None

def create_authenticated_session(base_url, user, password):
    """Cria sessão autenticada para um servidor Zabbix específico."""
    if not base_url or not user or not password:
        return None

    session = requests.Session()
    base_url = base_url.rstrip('/')
    
    print(f">>> Autenticando em {base_url}...")
    csrf = get_csrf_token(session, base_url)
    
    payload = {
        "name": user,
        "password": password,
        "enter": "Sign in"
    }
    if csrf: payload["csrf_token"] = csrf

    try:
        resp = session.post(
            f"{base_url}/index.php",
            data=payload,
            verify=False,
            allow_redirects=True
        )
        
        cookies = session.cookies.get_dict()
        if "zbx_session" in cookies or "zbx_sessionid" in cookies:
            return session
        else:
            print(f"❌ Falha no login Zabbix ({base_url}).")
            return None
    except Exception as e:
        print(f"❌ Erro crítico no login Zabbix ({base_url}): {e}")
        return None

def download_graphs_from_server(session, base_url, graph_list):
    """Baixa uma lista de gráficos usando uma sessão autenticada."""
    images = []
    base_url = base_url.rstrip('/')
    
    for item in graph_list:
        gid = item["id"]
        title = item["title"]
        
        chart_url = (
            f"{base_url}/chart2.php?"
            f"graphid={gid}"
            f"&width={WIDTH}&height={HEIGHT}"
            f"&from={PERIOD_STRING}&to=now"
            f"&profileIdx=web.charts.filter&resolve_macros=1"
        )
        
        # Verifica se há hostid específico e adiciona à URL
        if "hostid" in item:
            chart_url += f"&hostids[0]={item['hostid']}"
        
        try:
            resp = session.get(chart_url, stream=True, verify=False)
            if resp.status_code == 200:
                images.append((title, io.BytesIO(resp.content)))
            else:
                print(f"⚠ Erro HTTP {resp.status_code} no gráfico: {title}")
        except Exception as e:
            print(f"⚠ Erro de conexão no gráfico {title}: {e}")
            
    return images

# --- FUNÇÃO PRINCIPAL (INTERFACE) ---

def collect_graph_images():
    """
    Orquestra o download de múltiplos servidores Zabbix.
    """
    all_images = []

    # 1. Servidor Principal (Links Internet)
    if ZABBIX1_URL and ZABBIX1_USER and ZABBIX1_PASS:
        session1 = create_authenticated_session(ZABBIX1_URL, ZABBIX1_USER, ZABBIX1_PASS)
        if session1:
            print(f">>> Baixando {len(GRAPHS_CONFIG_1)} gráficos do Servidor 1...")
            imgs = download_graphs_from_server(session1, ZABBIX1_URL, GRAPHS_CONFIG_1)
            all_images.extend(imgs)
    else:
        print("⚠ Configuração do Servidor 1 incompleta no .env")

    # 2. Servidor EBNET (.210)
    if ZABBIX2_USER and ZABBIX2_PASS:
        session2 = create_authenticated_session(ZABBIX2_URL, ZABBIX2_USER, ZABBIX2_PASS)
        if session2:
            print(f">>> Baixando {len(GRAPHS_CONFIG_2)} gráficos do Servidor 2...")
            imgs = download_graphs_from_server(session2, ZABBIX2_URL, GRAPHS_CONFIG_2)
            all_images.extend(imgs)
    else:
        print("⚠ Pular Servidor 2: ZABBIX2_USER/PASS ausentes.")

    # 3. Servidor Novo (.208)
    if ZABBIX3_USER and ZABBIX3_PASS:
        session3 = create_authenticated_session(ZABBIX3_URL, ZABBIX3_USER, ZABBIX3_PASS)
        if session3:
            print(f">>> Baixando {len(GRAPHS_CONFIG_3)} gráficos do Servidor 3...")
            imgs = download_graphs_from_server(session3, ZABBIX3_URL, GRAPHS_CONFIG_3)
            all_images.extend(imgs)
    else:
        print("⚠ Pular Servidor 3: ZABBIX3_USER/PASS ausentes no .env")

    return all_images

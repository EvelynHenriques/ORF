import requests
import os
import urllib3
import unicodedata
from dotenv import load_dotenv

load_dotenv() 

# --- CONFIGURAÇÕES ---
# Tenta pegar do .env, se não tiver, usa o padrão do seu script
ZABBIX_API_URL = os.getenv("ZABBIX_API_URL")
GRAFANA_URL = os.getenv("GRAFANA_URL")
GRAFANA_TOKEN = os.getenv("GRAFANA_TOKEN")
DASHBOARD_UID = "ngTcCD04z"

USERNAME = os.getenv("ZABBIX_USER")
PASSWORD = os.getenv("ZABBIX_PASS")

urllib3.disable_warnings()

# --- FUNÇÕES AUXILIARES (Lógica Robusta do seu script original) ---

def get_best_ip(interfaces):
    """
    Seleciona o melhor IP da lista de interfaces do Zabbix.
    Prioriza IPs de gerência (10.78... 10.83) e evita IPs de Docker.
    """
    if not interfaces: return "-"
    # Filtra localhost e 0.0.0.0
    lista_ips = [iface['ip'] for iface in interfaces if iface['ip'] not in ['127.0.0.1', '0.0.0.0']]
    
    # 1. Prioridade para faixas de Gerência
    for ip in lista_ips:
        if ip.startswith(("10.78.", "10.79.", "10.80.", "10.81.", "10.82.", "10.83.")):
            return ip
    
    # 2. Evita IPs de containers/internos se possível (Ex: 10.147)
    for ip in lista_ips:
        if not ip.startswith("10.147."): return ip
            
    # 3. Retorna o primeiro que sobrou ou traço
    if lista_ips: return lista_ips[0]
    return "-"

def clean_host_key(text):
    """
    Normaliza o nome do host (remove acentos, prefixos SW_, 4CTA_) 
    para facilitar o cruzamento de dados entre Grafana e Zabbix.
    """
    if not text: return ""
    try:
        nfkd = unicodedata.normalize('NFKD', str(text))
        # Remove acentos
        s = "".join([c for c in nfkd if not unicodedata.combining(c)]).upper()
        # Remove prefixos comuns de infraestrutura
        s = s.replace("SW_", "").replace("4CTA_", "").replace("RTR_", "").replace("FG_", "")
        # Mantém apenas alfanuméricos
        return "".join(filter(str.isalnum, s)) 
    except:
        return str(text).upper()

def download_zabbix_ips():
    """
    Conecta na API do Zabbix e cria um mapa {nome_host: ip}.
    Usa várias chaves (nome original, nome limpo) para garantir que o IP seja encontrado.
    """
    sess = requests.Session()
    try:
        # 1. Autenticação
        resp = sess.post(ZABBIX_API_URL, json={
            "jsonrpc": "2.0", 
            "method": "user.login", 
            "params": {"user": USERNAME, "password": PASSWORD}, 
            "id": 1
        }, verify=False)
        
        auth = resp.json().get('result')
        if not auth: 
            print("Erro auth Zabbix")
            return {}

        # 2. Busca Hosts + Interfaces
        payload = {
            "jsonrpc": "2.0",
            "method": "host.get",
            "params": {"output": ["host", "name"], "selectInterfaces": ["ip"]},
            "auth": auth,
            "id": 2
        }
        hosts = sess.post(ZABBIX_API_URL, json=payload, verify=False).json().get('result', [])
        
        # 3. Cria Mapa Inteligente
        mapa = {}
        for h in hosts:
            ip_final = get_best_ip(h.get('interfaces', []))
            if ip_final != "-":
                # Mapeia todas as variações possíveis do nome
                mapa[h['host']] = ip_final
                mapa[h['name']] = ip_final
                mapa[clean_host_key(h['host'])] = ip_final 
                mapa[clean_host_key(h['name'])] = ip_final
        return mapa
    except Exception as e:
        print(f"Erro ao baixar IPs: {e}")
        return {}

def query_grafana_status(target_config):
    """
    Consulta a API do Grafana para obter o valor numérico e o nome do item.
    """
    # Tenta descobrir qual host estamos consultando
    host_configurado = target_config.get("host", {}).get("filter", "")
    if not host_configurado: 
        host_configurado = target_config.get("zabbix", {}).get("host", {}).get("filter", "")
    
    if not host_configurado: return None, None, None

    # Monta query
    query = target_config.copy()
    query['refId'] = 'A'
    payload = {"from": "now-24h", "to": "now", "queries": [query]}
    
    headers = {"Authorization": f"Bearer {GRAFANA_TOKEN}", "Content-Type": "application/json"}
    
    try:
        resp = requests.post(f"{GRAFANA_URL}/api/ds/query", headers=headers, json=payload, verify=False)
        # Navega no JSON de resposta do Grafana (Frames -> Data -> Values)
        frames = resp.json().get("results", {}).get("A", {}).get("frames", [])
        if frames:
            vals = frames[0].get("data", {}).get("values", [])
            # vals[0] costuma ser tempo, vals[1] o valor
            if len(vals) > 1 and vals[1]:
                # Pega valores não nulos
                validos = [v for v in vals[1] if v is not None]
                if validos:
                    valor_final = float(validos[-1])
                    # Tenta descobrir o nome do item (para saber se é latência ou status)
                    item = target_config.get("item", {}).get("filter", "")
                    if not item: item = target_config.get("zabbix", {}).get("item", {}).get("filter", "")
                    return host_configurado, valor_final, str(item)
        
        return host_configurado, None, None
    except: 
        return None, None, None

def get_status_tuple(valor, item_nome):
    """
    Define o Texto e a Cor do status baseando-se no valor e no tipo de item.
    Retorna: (Texto, Cor_Para_PDF)
    """
    if valor is None: 
        return "SEM DADOS", "GRAY"
    
    item_lower = str(item_nome).lower()
    
    # Lógica Específica para Latência (BBS)
    if "latencia" in item_lower or "latência" in item_lower:
        if valor == 0: 
            return "DOWN", "RED"
        elif valor <= 0.2: 
            return "Via BBS", "GREEN"
        else: 
            return "Via BBS", "YELLOW"
            
    # Lógica Padrão (Status UP/DOWN)
    else:
        if valor == 1: 
            return "UP", "GREEN"
        else: 
            return "DOWN", "RED"

# --- FUNÇÃO PRINCIPAL (Chamada pelo main.py) ---

def collect_reme_data():
    """
    Orquestra a coleta de dados:
    1. Baixa IPs do Zabbix.
    2. Lê Dashboard do Grafana.
    3. Processa painéis e seções.
    4. Retorna estrutura de dados para o PDF.
    """
    print(">>> Módulo REME: Iniciando coleta de dados...")
    
    mapa_ips = download_zabbix_ips()
    
    headers = {"Authorization": f"Bearer {GRAFANA_TOKEN}"}
    try:
        resp = requests.get(f"{GRAFANA_URL}/api/dashboards/uid/{DASHBOARD_UID}", headers=headers, verify=False)
        dashboard_data = resp.json()
        panels = dashboard_data.get('dashboard', {}).get('panels', [])
    except Exception as e:
        print(f"Erro ao baixar dashboard Grafana: {e}")
        return []

    relatorio_final = []
    secao_atual = None

    for panel in panels:
        # A lógica do seu script usa o campo 'noValue' para identificar Cabeçalhos de Cidade
        no_value = panel.get('fieldConfig', {}).get('defaults', {}).get('noValue', '')
        
        # Se for um cabeçalho de seção (Ex: "MANAUS-AM")
        if no_value and len(no_value) > 2:
            titulo_raw = no_value.upper()
            
            # Formata o título para ficar bonito no relatório
            if "INTERNET" in titulo_raw or "BBI" in titulo_raw:
                display_title = f"STATUS DE PROVEDORES DE {titulo_raw}"
            else:
                display_title = f"STATUS DOS PoP REME {titulo_raw}"
            
            secao_atual = {
                "titulo": display_title, 
                "dados": []
            }
            relatorio_final.append(secao_atual)
            continue

        # Se for um painel de dados (Status)
        targets = panel.get('targets', [])
        if targets and secao_atual is not None:
            target = targets[0]
            nome_om_grafana = panel.get('title', 'Sem Nome')
            
            # 1. Consulta Status
            host_tec, valor, item_nome = query_grafana_status(target)
            
            # 2. Define Texto e Cor
            status_txt, status_cor = get_status_tuple(valor, item_nome)

            # 3. Descobre o IP correto
            # Tenta pelo nome técnico, depois pelo nome visível, depois pela chave limpa
            ip = mapa_ips.get(host_tec, 
                    mapa_ips.get(nome_om_grafana, 
                        mapa_ips.get(clean_host_key(host_tec), "-")
                    )
                 )
            
            # Adiciona linha de dados na seção atual
            secao_atual["dados"].append({
                "om": nome_om_grafana,
                "ip": ip,
                "status": status_txt,
                "cor": status_cor
            })
            
    return relatorio_final
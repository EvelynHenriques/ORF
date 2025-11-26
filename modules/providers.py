#!/usr/bin/env python3
import subprocess

# Configurações
PROVEDORES = [
    {
        "link": "EMPRESA CLARO*",
        "teste": "ping 200.213.232.73",
        "link_wan": "INTERNET",
        "timeout": 5, "count": 4
    },
    {
        "link": "EMPRESA NWHERE*",
        "teste": "ping 186.233.94.33",
        "link_wan": "INTERNET",
        "timeout": 5, "count": 4
    },
    {
        "link": "ANÚNCIO BGP*",
        "teste": "mtr 177.8.82.1",
        "link_wan": "ENW",
        "timeout": 5, "count": 4
    },
    {
        "link": "4ªCTA ⇔ 7ªCTA**",
        "teste": "mtr 10.67.4.35",
        "link_wan": "METRO",
        "timeout": 5, "count": 4,
        "observacao": "(deve passar por 172.30.192.129)"
    },
    {
        "link": "4ªCTA ⇔ 6ªCTA**",
        "teste": "mtr 10.56.67.163",
        "link_wan": "METRO",
        "timeout": 5, "count": 4,
        "observacao": "(deve passar por 172.30.192.101)"
    },
    {
        "link": "4ªCTA ⇔ 41ªCT**",
        "teste": "mtr 10.89.36.95",
        "link_wan": "METRO",
        "timeout": 5, "count": 4,
        "observacao": "(deve passar por 172.30.192.194)"
    }
]

TUNEIS = [
    {
        "tunel": "4ª CTA ⇔ HMAM (Cisco)",
        "teste": "mtr 10.79.1.46",
        "observacao": "(observar se chega ao destino com apenas 3 saltos)",
        "timeout": 5, "count": 4
    }
]

def executar_ping(host, count=4, timeout=5):
    try:
        # -W em ping espera em segundos no Linux
        cmd = ['ping', '-c', str(count), '-W', str(timeout), host]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout+5)
        return result.returncode == 0
    except Exception:
        return False

def executar_mtr(host, count=4, timeout=5):
    try:
        # Tentar mtr
        cmd = ['mtr', '-r', '-c', str(count), '-w', host]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout+10)
        return result.returncode == 0
    except FileNotFoundError:
        # Fallback para traceroute
        try:
            cmd = ['traceroute', '-m', '15', '-w', str(timeout), host]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout+10)
            return result.returncode == 0
        except Exception:
            return False
    except Exception:
        return False

def executar_teste_cmd(teste_cmd, timeout=5, count=4):
    partes = teste_cmd.split()
    comando = partes[0]
    host = partes[1] if len(partes) > 1 else ""
    
    if comando == "ping":
        return executar_ping(host, count, timeout)
    elif comando == "mtr":
        return executar_mtr(host, count, timeout)
    return False

def collect_providers_data():
    """
    Executa os testes e retorna dois dicionários com os resultados.
    """
    print("   [Providers] Iniciando testes de conectividade...")
    
    # Processar Provedores
    resultados_provedores = []
    for prov in PROVEDORES:
        sucesso = executar_teste_cmd(prov['teste'], prov['timeout'], prov['count'])
        
        item = {
            'link': prov['link'],
            'teste_str': prov['teste'],
            'link_wan': prov['link_wan'],
            'observacao': prov.get('observacao', ''),
            'status': 'GREEN' if sucesso else 'RED'
        }
        resultados_provedores.append(item)

    # Processar Túneis
    resultados_tuneis = []
    for tun in TUNEIS:
        sucesso = executar_teste_cmd(tun['teste'], tun['timeout'], tun['count'])
        
        item = {
            'tunel': tun['tunel'],
            'teste_str': tun['teste'],
            'observacao': tun.get('observacao', ''),
            'status': 'GREEN' if sucesso else 'RED'
        }
        resultados_tuneis.append(item)
        
    print("   [Providers] Testes concluídos.")
    return resultados_provedores, resultados_tuneis

#if __name__ == "__main__":
    # Teste rápido se rodar o arquivo direto
    #p, t = collect_providers_data()
    #print(p)
    #print(t)
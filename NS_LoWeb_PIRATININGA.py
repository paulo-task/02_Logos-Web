import re
import os
import time
import pandas as pd
from datetime import datetime
import warnings 
from playwright.sync_api import Playwright, sync_playwright, expect

# Silencia avisos de forma universal
warnings.filterwarnings("ignore", message=".*SettingWithCopyWarning.*")

def get_download_path():
    """Define o caminho de download baseado no ambiente"""
    if os.getenv('GITHUB_ACTIONS') == 'true':
        # Ambiente GitHub Actions
        download_path = os.path.join(os.getcwd(), 'downloads')
        os.makedirs(download_path, exist_ok=True)
        return download_path
    else:
        # Ambiente Local
        return r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\16_Notas_Servico"

def upload_to_sharepoint(conteudo_bytes, nome_arquivo, pasta_sharepoint):
    """Envia arquivo para o SharePoint via Microsoft Graph API"""
    try:
        import requests
        from urllib.parse import urlparse
        
        SP_CLIENT_ID = os.getenv("SP_CLIENT_ID", "").strip()
        SP_CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "").strip()
        SP_TENANT_ID = os.getenv("SP_TENANT_ID", "").strip()
        SITE_URL = "https://engelmigproject.sharepoint.com/sites/LEC_ENGELMIG"
        
        if not SP_CLIENT_ID:
            print("   ⚠️ SharePoint: credenciais não configuradas (verifique as variáveis de ambiente)")
            return False
        
        print("   🔑 Autenticando no SharePoint...")
        # Obter token
        url_token = f'https://login.microsoftonline.com/{SP_TENANT_ID}/oauth2/v2.0/token'
        data = {
            'grant_type': 'client_credentials',
            'client_id': SP_CLIENT_ID,
            'client_secret': SP_CLIENT_SECRET,
            'scope': 'https://graph.microsoft.com/.default'
        }
        r = requests.post(url_token, data=data)
        r.raise_for_status()
        token = r.json()['access_token']
        
        # Obter Site ID
        parsed = urlparse(SITE_URL)
        host = parsed.netloc
        site_path = parsed.path.strip("/")
        url_site = f"https://graph.microsoft.com/v1.0/sites/{host}:/{site_path}"
        headers = {"Authorization": f"Bearer {token}"}
        r = requests.get(url_site, headers=headers)
        r.raise_for_status()
        site_id = r.json()["id"]
        
        # Obter Drive Workspace
        url_drives = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        r = requests.get(url_drives, headers={"Authorization": f"Bearer {token}"})
        r.raise_for_status()
        
        drive_id = None
        for drive in r.json().get('value', []):
            if drive.get('name') == 'Workspace':
                drive_id = drive.get('id')
                break
        
        if not drive_id:
            raise Exception("Pasta raiz (Drive) 'Workspace' não encontrada no SharePoint")
        
        # Upload
        print(f"   ☁️ Enviando '{nome_arquivo}' para '{pasta_sharepoint}'...")
        url_upload = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{pasta_sharepoint}/{nome_arquivo}:/content"
        headers_upload = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream"
        }
        r = requests.put(url_upload, headers=headers_upload, data=conteudo_bytes)
        r.raise_for_status()
        
        print(f"   ✅ Upload SharePoint concluído com sucesso!")
        return True
        
    except Exception as e:
        print(f"   ❌ Erro no upload para o SharePoint: {e}")
        return False

def run(playwright: Playwright) -> None:
    # --- DETECÇÃO DE AMBIENTE ---
    is_github_actions = os.getenv('GITHUB_ACTIONS') == 'true'
    
    # --- CONFIGURAÇÕES DE RETENTATIVA ---
    max_tentativas = 30 if not is_github_actions else 10
    intervalo_segundos = 3 if not is_github_actions else 5
    tentativa_atual = 0
    logado = False
    # ------------------------------------

    # --- CONFIGURAÇÕES DO BROWSER ---
    browser_options = {
        "headless": is_github_actions,  # Headless no CI, visível no local
        "slow_mo": 100 if is_github_actions else 300  # Mais lento no CI
    }
    
    # Se estiver no GitHub Actions, adiciona argumentos para estabilidade
    if is_github_actions:
        browser_options["args"] = [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-dev-shm-usage',
            '--disable-accelerated-2d-canvas',
            '--disable-gpu'
        ]

    browser = playwright.chromium.launch(**browser_options)
    context = browser.new_context(
        viewport={'width': 1280, 'height': 720},
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    )
    page = context.new_page()
    
    print("--- Iniciando Script Piratininga ---")

    while tentativa_atual < max_tentativas and not logado:
        tentativa_atual += 1
        print(f"\nTentativa de login {tentativa_atual} de {max_tentativas}...")
        
        try:
            page.goto("https://contratadas.cpfl.com.br/account/login.aspx")
            
            # Preenchimento de Credenciais
            page.locator("#MainContent_txtLogin").fill("CP14111")
            page.locator("#MainContent_txtSenha").fill("Vini@2015")
            page.get_by_role("button", name="Logar").click()
            
            # Aguarda um momento para o servidor processar o redirecionamento
            time.sleep(5)
            
            # VALIDAÇÃO 1: Se a URL ainda contém 'login.aspx', o login falhou ou a sessão está presa
            if "login.aspx" in page.url.lower():
                raise Exception("Acesso negado ou Sessão já ocupada por outro usuário")

            # VALIDAÇÃO 2: Tenta localizar o link usando Regex (ignora ícones e espaços extras)
            link_semaforo = page.get_by_role("link", name=re.compile(r"Consulta Semáforo de Notas"))
            
            # Se não estiver visível pelo nome, tenta pelo seletor de texto parcial
            if not link_semaforo.is_visible():
                link_semaforo = page.locator("a:has-text('Semáforo')")

            link_semaforo.wait_for(state="visible", timeout=15000)
            link_semaforo.click()
            
            # VALIDAÇÃO 3: Teste de Estabilidade (espera ver se o sistema desloga após o clique)
            time.sleep(3)
            if "login.aspx" in page.url.lower():
                raise Exception("O sistema deslogou automaticamente logo após o acesso")

            # Confirmação de entrada na tela de consulta
            page.locator("#MainContent_btnConsultarJS").wait_for(state="visible", timeout=15000)
            
            logado = True
            print("--- SUCESSO: Login estabilizado! ---")
            
        except Exception as e:
            print(f"⚠️ Erro na tentativa {tentativa_atual}: {e}")
            if tentativa_atual < max_tentativas:
                print(f"Aguardando {intervalo_segundos} segundos para nova tentativa...")
                time.sleep(intervalo_segundos)
            else:
                print("❌ Limite de tentativas atingido. Encerrando.")
                browser.close()
                return

    # --- INÍCIO DO PROCESSO DE EXPORTAÇÃO (SÓ EXECUTA SE LOGADO) ---
    print("Iniciando filtragem e exportação...")
    page.locator("#MainContent_btnConsultarJS").click()
    
    contratos = ["CTLEC074", "CTLEC073"]
    for c in contratos:
        try: 
            page.get_by_role("checkbox", name=re.compile(c)).check(timeout=3000)
        except: 
            pass
    
    print("Selecionando Cidades...")
    cidades = ["ITU", "BOITUVA", "PORTO FELIZ", "ALUMINIO", "ARACARIGUAMA", "IBIUNA", "MAIRINQUE", "SAO ROQUE",
               "ARACOIABA DA SERRA", "CAPELA DO ALTO", "IPERO", "SALTO DE PIRAPORA", "SOROCABA", "VOTORANTIM",
               "INDAIATUBA", "SALTO", "CAMPO LIMPO PAULISTA", "ITUPEVA", "JUNDIAI", "LOUVEIRA", "VARZEA PAULISTA", "VINHEDO"]

    for cidade in cidades:
        try: 
            page.get_by_role("checkbox", name=cidade, exact=True).check(timeout=1000)
        except: 
            continue

    page.get_by_role("row", name="TODOS", exact=True).get_by_label("TODOS").check()
    
    try:
        with page.expect_download(timeout=0) as download_info:
            page.locator("#MainContent_btnExportExcel").click()
        
        download = download_info.value
        
        # Define caminhos baseado no ambiente
        pasta_destino = get_download_path()
        caminho_final = os.path.join(pasta_destino, "Nota_Servico_Piratininga.xlsx")
        caminho_temp = os.path.join(pasta_destino, "temp_pira.xls")

        if os.path.exists(caminho_final):
            try:
                os.rename(caminho_final, caminho_final)
            except OSError:
                print(f"❌ ERRO: O arquivo '{caminho_final}' está aberto. Feche-o!")
                return 

        download.save_as(caminho_temp)
        
        # Tratamento de Dados
        tabelas = pd.read_html(caminho_temp, flavor='lxml')
        df = tabelas[0].copy()

        if "0" in str(df.columns[0]) or df.columns[0] == 0:
            df.columns = df.iloc[0]
            df = df[1:].copy()

        if 'QTDHORAS' in df.columns:
            df['QTDHORAS'] = df['QTDHORAS'].astype(str).str.replace(',', '.')
            df['QTDHORAS'] = pd.to_numeric(df['QTDHORAS'], errors='coerce')
            if df['QTDHORAS'].abs().max() > 1000:
                df['QTDHORAS'] = df['QTDHORAS'] / 100

        df['DT_RELATORIO'] = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        df.to_excel(caminho_final, index=False)
        print(f"✓ Arquivo excel local salvo temporariamente")
        
        # Faz o envio para o SharePoint
        with open(caminho_final, "rb") as f:
            conteudo_bytes = f.read()
        
        upload_to_sharepoint(conteudo_bytes, "Nota_Servico_Piratininga.xlsx", "BI_LEC/16_Notas_Servico")
        
        if os.path.exists(caminho_temp): os.remove(caminho_temp)
        print(f"--- SUCESSO FINAL: Arquivo Piratininga Gerado! ---")
        
    except Exception as e:
        print(f"ERRO NA EXPORTAÇÃO: {e}")

    time.sleep(2)
    context.close()
    browser.close()

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
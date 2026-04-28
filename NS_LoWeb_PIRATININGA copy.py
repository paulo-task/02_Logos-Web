import re
import os
import time
import pandas as pd
from datetime import datetime
import warnings 
from playwright.sync_api import Playwright, sync_playwright, expect

# Silencia avisos de forma universal
warnings.filterwarnings("ignore", message=".*SettingWithCopyWarning.*")

def run(playwright: Playwright) -> None:
    # --- CONFIGURAÇÕES DE RETENTATIVA ---
    max_tentativas = 30      # Quantidade de vezes que vai tentar logar
    intervalo_segundos = 3 # Tempo de espera recomendado (60s ajuda o servidor a liberar a sessão)
    tentativa_atual = 0
    logado = False
    # ------------------------------------

    browser = playwright.chromium.launch(headless=False, slow_mo=300)
    context = browser.new_context()
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
        pasta_destino = r"C:\Users\paulo.janio\ENGELMIG ENERGIA LTDA\LEC ENGELMIG - Workspace\BI_LEC\16_Notas_Servico"
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
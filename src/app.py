import os
import time
import yagmail
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC
import schedule


# configurações do selenium
download_dir = os.path.expanduser("~/Downloads") # diretorio de downloads do pc
chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\PC\Downloads",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "profile.default_content_settings.popups": 0
})

# Inicializando o WebDriver
driver = webdriver.Chrome(service=Service(), options=chrome_options)
                          
# func find file 
def find_latest_file():
    files = [f for f in os.listdir(download_dir) if f.startswith("Portal") and f.endswith(".xlsx")]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)
    return files[0] if files else None

# Função para verificar se o arquivo foi completamente baixado
def wait_for_download(filename, timeout=60):
    start_time = time.time()
    file_path = os.path.join(download_dir, filename)
    while time.time() - start_time < timeout:
        if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
            # Verifica se o arquivo não está sendo baixado
            if not filename.endswith('.crdownload'):
                return file_path
        time.sleep(1)
    return None

# func de enviar emails
def send_email():
    recipients = ["cassandra@controllersbr.com", "ti2.controllersbr@gmail.com", "carlos.junior@controllersbr.com", "joas@controllersbr.com", "lucas@controllersbr.com", "juliocesar@controllersbr.com"]
    sender_email = "redstarenzo@gmail.com"
    yag = yagmail.SMTP(sender_email, "evms sijx toii aqxx")
    
    # Encontrar o arquivo mais recente
    filename = find_latest_file()
    if filename:
        # Esperar o arquivo ser completamente baixado
        file_path = wait_for_download(os.path.basename(filename))
        if file_path:
            yag.send(
                to=recipients,
                subject="Relatório Automático(Certificado Digital)",
                contents="Por favor, veja a planilha em anexo.",
                attachments=file_path,
            )
            print(f"E-mail enviado com sucesso com o anexo {file_path}")
        else:
            print("Arquivo não encontrado ou não está completo.")
    else:
        print("Nenhum arquivo encontrado para enviar.")
        
def run_automation():
    # Setup
    driver.get("https://08978175000180.portal-veri.com.br")
    driver.set_window_size(1366, 768)

    # Actions in Login
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[2]/input'))).send_keys("ti2.controllersbr@gmail.com")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[3]/input'))).send_keys("senha123")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="kt_sign_in_form"]/div[5]/button'))).click()

    # Actions in Dashboard
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(1) > .col-md-3:nth-child(3) .ms-3"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".row:nth-child(6) > .col-md-3:nth-child(4) .text-gray-700"))).click()
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".buttons-excel > span"))).click()
    
    # Aguardar download ser concluído
    time.sleep(10)
    
    # Envia o e-mail já com docmuento baixado
    send_email()
    time.sleep(10)
    # Fechar
    driver.quit()

# Agendamento
def job():
    run_automation()

schedule.every().monday.at("09:00").do(job)

# Mantém o script rodando para cumprir agendamento
while True:
    schedule.run_pending()
    time.sleep(1)

# Teste imediato
run_automation()


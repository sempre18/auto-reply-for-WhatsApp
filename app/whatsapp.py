import os
import time
import urllib.parse
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WhatsAppSender:
    def __init__(self, chromedriver_path: str = ""):
        self.chromedriver_path = chromedriver_path.strip() if chromedriver_path else ""
        self.driver = None
        self.wait = None
        self.profile_dir = self._get_profile_dir()

    def _get_profile_dir(self) -> str:
        base_dir = os.getenv("LOCALAPPDATA") or str(Path.home())
        profile_dir = os.path.join(base_dir, "WhatsAppCobranca", "chrome_profile")
        os.makedirs(profile_dir, exist_ok=True)
        return profile_dir

    def start(self):
        options = Options()

        self.profile_dir = self._get_profile_dir()

        options.add_argument(f"--user-data-dir={self.profile_dir}")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--remote-debugging-port=9222")

        try:
            if self.chromedriver_path:
                print("[INFO] Usando ChromeDriver manual...")
                service = Service(self.chromedriver_path)
                self.driver = webdriver.Chrome(service=service, options=options)
            else:
                print("[INFO] Usando Selenium Manager (automático)...")
                self.driver = webdriver.Chrome(options=options)

            self.wait = WebDriverWait(self.driver, 30)

        except Exception as e:
            print("[ERRO] Falha ao iniciar Chrome:", e)
            raise

    def open_whatsapp(self):
        if not self.driver:
            raise Exception("Chrome não foi iniciado. Execute start() antes.")
        self.driver.get("https://web.whatsapp.com/")

    def wait_for_login(self):
        if not self.wait:
            raise Exception("WebDriverWait não foi iniciado.")
        self.wait.until(EC.presence_of_element_located((By.ID, "side")))

    import os
import time
import urllib.parse
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


class WhatsAppSender:
    def __init__(self, chromedriver_path: str = ""):
        self.chromedriver_path = chromedriver_path.strip() if chromedriver_path else ""
        self.driver = None
        self.wait = None
        self.profile_dir = None

    def _get_profile_dir(self) -> str:
        base_dir = os.getenv("LOCALAPPDATA") or str(Path.home())
        profile_dir = os.path.join(base_dir, "WhatsAppCobranca", "chrome_profile")
        os.makedirs(profile_dir, exist_ok=True)
        return profile_dir

    def start(self):
        options = Options()

        self.profile_dir = self._get_profile_dir()

        options.add_argument(f"--user-data-dir={self.profile_dir}")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--remote-debugging-port=9222")

        try:
            if self.chromedriver_path:
                print("[INFO] Usando ChromeDriver manual...")
                service = Service(self.chromedriver_path)
                self.driver = webdriver.Chrome(service=service, options=options)
            else:
                print("[INFO] Usando Selenium Manager (automático)...")
                self.driver = webdriver.Chrome(options=options)

            self.wait = WebDriverWait(self.driver, 30)

        except Exception as e:
            print("[ERRO] Falha ao iniciar Chrome:", e)
            raise

    def open_whatsapp(self):
        if not self.driver:
            raise Exception("Chrome não foi iniciado. Execute start() antes.")
        self.driver.get("https://web.whatsapp.com/")

    def wait_for_login(self):
        if not self.wait:
            raise Exception("WebDriverWait não foi iniciado.")
        self.wait.until(EC.presence_of_element_located((By.ID, "side")))

    def _find_message_box(self):
        selectors = [
            '//footer//div[@contenteditable="true"]',
            '//div[@contenteditable="true"][@role="textbox"]',
            '//div[@contenteditable="true"][@data-tab="10"]',
            '//div[@contenteditable="true"][@data-tab="6"]',
        ]

        for xpath in selectors:
            elements = self.driver.find_elements(By.XPATH, xpath)
            for el in elements:
                try:
                    if el.is_displayed() and el.is_enabled():
                        return el
                except Exception:
                    pass
        return None

    def _has_invalid_number_error(self):
        error_xpaths = [
            '//div[contains(text(),"não está no WhatsApp")]',
            '//div[contains(text(),"not on WhatsApp")]',
            '//div[contains(text(),"Phone number shared via url is invalid")]',
        ]

        for xpath in error_xpaths:
            if self.driver.find_elements(By.XPATH, xpath):
                return True
        return False

    def send_message(self, phone: str, message: str):
        if not self.driver or not self.wait:
            raise Exception("WhatsApp não foi iniciado corretamente.")

        if not phone:
            raise Exception("Telefone vazio.")

        if not message:
            raise Exception("Mensagem vazia.")

        encoded = urllib.parse.quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded}"
        self.driver.get(url)

        # Espera a conversa carregar de verdade ou aparecer erro
        timeout_seconds = 60
        start_time = time.time()
        box = None

        while time.time() - start_time < timeout_seconds:
            if self._has_invalid_number_error():
                raise Exception("Número não está no WhatsApp")

            box = self._find_message_box()
            if box:
                break

            time.sleep(0.5)

        if not box:
            raise Exception("Tempo esgotado aguardando a conversa carregar.")

        # Espera o texto pré-preenchido aparecer na caixa
        prefill_ok = False
        start_prefill = time.time()

        while time.time() - start_prefill < 20:
            try:
                text_now = (box.text or "").strip()
                if text_now:
                    prefill_ok = True
                    break
            except Exception:
                pass
            time.sleep(0.3)

        if not prefill_ok:
            # mesmo sem texto visível, ainda tenta enviar
            pass

        # Conta mensagens antes do envio
        before_count = len(self.driver.find_elements(By.XPATH, '//div[contains(@class,"message-out")]'))

        box.send_keys(Keys.ENTER)

        # Aguarda uma nova mensagem enviada aparecer
        sent_ok = False
        start_sent = time.time()

        while time.time() - start_sent < 25:
            if self._has_invalid_number_error():
                raise Exception("Número não está no WhatsApp")

            after_count = len(self.driver.find_elements(By.XPATH, '//div[contains(@class,"message-out")]'))
            if after_count > before_count:
                sent_ok = True
                break

            time.sleep(0.4)

        if not sent_ok:
            raise Exception("Mensagem não foi confirmada como enviada.")

        return True

    def close(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None

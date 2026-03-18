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
    def __init__(self, chromedriver_path: str):
        self.chromedriver_path = chromedriver_path
        self.driver = None
        self.wait = None

    def _get_profile_dir(self) -> str:
        """
        Cria um diretório fixo e seguro para o perfil do Chrome no Windows.
        """
        base_dir = os.getenv("LOCALAPPDATA") or str(Path.home())
        profile_dir = os.path.join(base_dir, "WhatsAppCobranca", "chrome_profile")

        os.makedirs(profile_dir, exist_ok=True)
        return profile_dir

    def start(self):
        profile_dir = self._get_profile_dir()

        options = Options()
        options.add_argument("--start-maximized")
        options.add_argument(f"--user-data-dir={profile_dir}")
        options.add_argument("--lang=pt-BR")
        options.add_argument("--disable-notifications")

        service = Service(self.chromedriver_path)
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 30)

    def open_whatsapp(self):
        self.driver.get("https://web.whatsapp.com/")

    def wait_for_login(self):
        self.wait.until(
            EC.presence_of_element_located((By.ID, "side"))
        )

    def send_message(self, phone: str, message: str):
        encoded = urllib.parse.quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded}"
        self.driver.get(url)

        box = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="10" or @data-tab="6"]')
            )
        )

        time.sleep(2)
        box.send_keys(Keys.ENTER)
        time.sleep(2)

        return True

    def close(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None
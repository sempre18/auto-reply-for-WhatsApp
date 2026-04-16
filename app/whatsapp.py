import os
import time
import urllib.parse
from pathlib import Path
from typing import Callable, Optional

from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    from webdriver_manager.chrome import ChromeDriverManager

    _WDM_AVAILABLE = True
except ImportError:
    _WDM_AVAILABLE = False


_LOGIN_SELECTORS = [
    (By.CSS_SELECTOR, 'div[data-testid="chat-list"]'),
    (By.CSS_SELECTOR, 'div[data-testid="conversation-panel-wrapper"]'),
    (By.XPATH, '//div[@aria-label="Lista de conversas"]'),
    (By.XPATH, '//div[@aria-label="Chats"]'),
    (By.XPATH, '//div[@aria-label="Chat list"]'),
    (By.CSS_SELECTOR, 'div[role="grid"]'),
    (By.CSS_SELECTOR, '#app .two'),
]

_MESSAGE_BOX_SELECTORS = [
    (By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'),
    (By.XPATH, '//div[@contenteditable="true"][@data-tab="6"]'),
    (By.CSS_SELECTOR, 'div[data-testid="conversation-compose-box-input"]'),
    (By.XPATH, '//footer//div[@contenteditable="true"]'),
]

_INVALID_NUMBER_SELECTORS = [
    (By.XPATH, '//*[contains(text(), "número de telefone compartilhado através de url é inválido")]'),
    (By.XPATH, '//*[contains(text(), "Phone number shared via url is invalid")]'),
    (By.XPATH, '//*[contains(text(), "não está no WhatsApp")]'),
    (By.XPATH, '//*[contains(text(), "isn\'t on WhatsApp")]'),
]

_SENT_INDICATOR_SELECTORS = [
    (By.CSS_SELECTOR, 'span[data-testid="msg-time"]'),
    (By.CSS_SELECTOR, 'span[data-icon="msg-dtime"]'),
    (By.CSS_SELECTOR, 'span[data-icon="msg-check"]'),
]


class WhatsAppError(Exception):
    pass


class WhatsAppSender:
    PROFILE_DIR_NAME = "WhatsAppPro"
    LOGIN_TIMEOUT = 180
    SEND_TIMEOUT = 30
    PAGE_LOAD_SLEEP = 2.5

    def __init__(self, chromedriver_path: Optional[str] = None):
        self.chromedriver_path = chromedriver_path
        self.driver: Optional[webdriver.Chrome] = None
        self.wait: Optional[WebDriverWait] = None
        self._screenshot_dir = "logs"

    def _get_profile_dir(self) -> str:
        base = os.getenv("LOCALAPPDATA") or str(Path.home())
        path = os.path.join(base, self.PROFILE_DIR_NAME, "chrome_profile")
        os.makedirs(path, exist_ok=True)
        return path

    def _build_service(self) -> Service:
        if self.chromedriver_path:
            return Service(self.chromedriver_path)
        if _WDM_AVAILABLE:
            return Service(ChromeDriverManager().install())
        raise WhatsAppError(
            "webdriver-manager não instalado e nenhum chromedriver informado.\n"
            "Execute: pip install webdriver-manager"
        )

    def _build_options(self) -> Options:
        opts = Options()
        opts.add_argument("--start-maximized")
        opts.add_argument(f"--user-data-dir={self._get_profile_dir()}")
        opts.add_argument("--lang=pt-BR")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--disable-popup-blocking")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("useAutomationExtension", False)
        opts.add_argument("--disable-blink-features=AutomationControlled")
        return opts

    def start(self) -> None:
        service = self._build_service()
        options = self._build_options()
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        self.wait = WebDriverWait(self.driver, self.SEND_TIMEOUT)

    def open_whatsapp(self) -> None:
        if not self.driver:
            raise WhatsAppError("Navegador não iniciado. Chame start() primeiro.")
        self.driver.get("https://web.whatsapp.com/")

    def wait_for_login(self, log_fn: Optional[Callable] = None) -> None:
        if not self.driver:
            raise WhatsAppError("Navegador não iniciado.")

        long_wait = WebDriverWait(self.driver, self.LOGIN_TIMEOUT)

        for by, selector in _LOGIN_SELECTORS:
            try:
                long_wait.until(EC.presence_of_element_located((by, selector)))
                if log_fn:
                    log_fn("[OK] Login detectado.")
                return
            except TimeoutException:
                continue
            except Exception:
                continue

        self._save_screenshot("login_timeout")
        raise WhatsAppError(
            "Login não detectado após aguardar. Verifique o QR Code e o screenshot nos logs."
        )

    def send_message(
        self,
        phone: str,
        message: str,
        typing_delay: float = 0.0,
        retries: int = 2,
    ) -> bool:
        if not self.driver:
            raise WhatsAppError("Navegador não iniciado.")

        last_err = None
        for attempt in range(1, retries + 2):
            try:
                self._send_once(phone=phone, message=message, typing_delay=typing_delay)
                return True
            except WhatsAppError as e:
                last_err = e
                self._save_screenshot(f"send_fail_{phone}_try_{attempt}")
                if attempt <= retries:
                    try:
                        self.driver.refresh()
                    except Exception:
                        pass
                    time.sleep(2)
                    continue
                break

        raise WhatsAppError(str(last_err) if last_err else f"Falha ao enviar para {phone}")

    def _send_once(self, phone: str, message: str, typing_delay: float = 0.0) -> None:
        encoded = urllib.parse.quote(message)
        url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded}"
        self.driver.get(url)
        time.sleep(self.PAGE_LOAD_SLEEP)

        self._ensure_chat_ready(phone)

        invalid = self._detect_invalid_number()
        if invalid:
            raise WhatsAppError(invalid)

        box = self._find_message_box()
        if box is None:
            raise WhatsAppError(f"Caixa de mensagem não encontrada para {phone}.")

        try:
            ActionChains(self.driver).move_to_element(box).click(box).perform()
        except Exception:
            box.click()

        if typing_delay > 0:
            time.sleep(typing_delay)

        box.send_keys(Keys.ENTER)

        if not self._confirm_send():
            raise WhatsAppError(f"Mensagem não confirmada visualmente para {phone}.")

    def _ensure_chat_ready(self, phone: str) -> None:
        started = time.time()
        while time.time() - started < self.SEND_TIMEOUT:
            if self._detect_invalid_number():
                return
            box = self._find_message_box(raise_on_fail=False)
            if box is not None:
                return
            time.sleep(0.8)
        raise WhatsAppError(f"Conversa não ficou pronta a tempo para {phone}.")

    def _detect_invalid_number(self) -> str:
        for by, selector in _INVALID_NUMBER_SELECTORS:
            try:
                elements = self.driver.find_elements(by, selector)
                if elements:
                    text = elements[0].text.strip() or "Número inválido ou não encontrado no WhatsApp."
                    return text
            except Exception:
                continue
        return ""

    def _find_message_box(self, raise_on_fail: bool = True):
        for by, selector in _MESSAGE_BOX_SELECTORS:
            try:
                return WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((by, selector))
                )
            except (TimeoutException, NoSuchElementException):
                continue
            except Exception:
                continue

        if raise_on_fail:
            raise WhatsAppError("Caixa de mensagem não encontrada.")
        return None

    def _confirm_send(self, timeout: int = 10) -> bool:
        started = time.time()
        while time.time() - started < timeout:
            for by, selector in _SENT_INDICATOR_SELECTORS:
                try:
                    els = self.driver.find_elements(by, selector)
                    if els:
                        return True
                except Exception:
                    continue
            time.sleep(0.5)
        return False

    def _save_screenshot(self, name: str) -> None:
        try:
            os.makedirs(self._screenshot_dir, exist_ok=True)
            path = os.path.join(self._screenshot_dir, f"{name}.png")
            self.driver.save_screenshot(path)
        except Exception:
            pass

    def is_connected(self) -> bool:
        if not self.driver:
            return False
        try:
            for by, sel in _LOGIN_SELECTORS[:3]:
                el = self.driver.find_elements(by, sel)
                if el:
                    return True
            return False
        except WebDriverException:
            return False

    def close(self) -> None:
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None
            self.wait = None
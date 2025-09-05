import logging
import asyncio
import re
import time
import signal
import sys
import threading
from typing import Optional, Dict, Any
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException
import json
import os
import socket
import psutil
from pathlib import Path
from logging.handlers import RotatingFileHandler

# Konfigurationsdateien
CONFIG_FILE = "bot_config.json"
LOG_FILE = "bot.log"
STATS_FILE = "bot_stats.json"
PID_FILE = "bot.pid"


def setup_logging():
    """Logging mit automatischer Rotation einrichten"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            RotatingFileHandler(LOG_FILE, maxBytes=10*1024*1024, backupCount=5, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

setup_logging()
logger = logging.getLogger(__name__)


class BotStats:
    def __init__(self):
        self.stats = {
            'start_time': datetime.now().isoformat(),
            'messages_processed': 0,
            'responses_sent': 0,
            'errors': 0,
            'last_activity': None,
            'uptime': 0
        }
        self.load_stats()

    def load_stats(self):
        try:
            if os.path.exists(STATS_FILE):
                with open(STATS_FILE, 'r', encoding='utf-8') as f:
                    self.stats.update(json.load(f))
        except Exception as e:
            logger.warning(f"Could not load stats: {e}")

    def save_stats(self):
        try:
            self.stats['uptime'] = (datetime.now() - datetime.fromisoformat(self.stats['start_time'])).total_seconds()
            with open(STATS_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.stats, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.warning(f"Could not save stats: {e}")

    def increment(self, key: str):
        self.stats[key] = self.stats.get(key, 0) + 1
        self.stats['last_activity'] = datetime.now().isoformat()
        self.save_stats()


class NetworkChecker:
    @staticmethod
    def is_connected() -> bool:
        try:
            socket.create_connection(("8.8.8.8", 53), timeout=5)
            return True
        except OSError:
            return False

    @staticmethod
    def wait_for_network(timeout: int = 300) -> bool:
        logger.info("Waiting for network connection...")
        start_time = time.time()

        while time.time() - start_time < timeout:
            if NetworkChecker.is_connected():
                logger.info("Network connection established")
                return True
            time.sleep(5)

        logger.error("Network connection timeout")
        return False


class ProcessManager:
    def __init__(self):
        self.running = True
        self.setup_signal_handlers()
        self.write_pid()

    def setup_signal_handlers(self):
        def signal_handler(signum, frame):
            logger.info(f"Received signal {signum}, shutting down...")
            self.running = False

        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)

    def write_pid(self):
        try:
            with open(PID_FILE, 'w') as f:
                f.write(str(os.getpid()))
        except Exception as e:
            logger.warning(f"Could not write PID file: {e}")

    def cleanup_pid(self):
        try:
            if os.path.exists(PID_FILE):
                os.remove(PID_FILE)
        except Exception as e:
            logger.warning(f"Could not remove PID file: {e}")

    @staticmethod
    def is_already_running() -> bool:
        try:
            if os.path.exists(PID_FILE):
                with open(PID_FILE, 'r') as f:
                    pid = int(f.read().strip())

                try:
                    process = psutil.Process(pid)
                    if 'teams_chatgpt_bot' in ' '.join(process.cmdline()):
                        return True
                except psutil.NoSuchProcess:
                    pass
        except Exception as e:
            logger.warning(f"Could not check if already running: {e}")

        return False


class TeamsBot:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.teams_driver = None
        self.chatgpt_driver = None
        self.is_running = False
        self.processed_messages = set()
        self.stats = BotStats()
        self.process_manager = ProcessManager()
        self.last_health_check = datetime.now()
        self.health_check_interval = timedelta(minutes=5)
        self.max_retries = config.get('max_retries', 3)
        self.retry_delay = config.get('retry_delay', 30)

    def detect_captcha(self, driver) -> bool:
        try:
            if driver.find_elements(By.CSS_SELECTOR, "iframe[src*='recaptcha']") or \               driver.find_elements(By.CSS_SELECTOR, "iframe[src*='hcaptcha']") or \               driver.find_elements(By.CSS_SELECTOR, "div.g-recaptcha"):
                logger.warning("⚠️ Captcha detected – please solve it manually in the browser!")
                return True
        except Exception:
            pass
        return False

    async def login_to_teams(self) -> bool:
        try:
            logger.info("Starte Teams-Login...")
            self.teams_driver = self.setup_driver()
            self.teams_driver.get("https://teams.microsoft.com")

            if self.detect_captcha(self.teams_driver):
                while self.detect_captcha(self.teams_driver):
                    time.sleep(5)

            email_input = WebDriverWait(self.teams_driver, 15).until(
                EC.presence_of_element_located((By.ID, "i0116"))
            )
            email_input.clear()
            email_input.send_keys(self.config['teams_email'])

            self.teams_driver.find_element(By.ID, "idSIButton9").click()

            password_input = WebDriverWait(self.teams_driver, 15).until(
                EC.presence_of_element_located((By.ID, "i0118"))
            )
            password_input.clear()
            password_input.send_keys(self.config['teams_password'])
            self.teams_driver.find_element(By.ID, "idSIButton9").click()

            try:
                stay_signed_in = WebDriverWait(self.teams_driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "idSIButton9"))
                )
                stay_signed_in.click()
            except TimeoutException:
                pass

            WebDriverWait(self.teams_driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-tid='app-bar-chat']"))
            )
            logger.info("Successfully logged into Teams")
            return True

        except Exception as e:
            logger.error(f"Teams login error: {e}")
            return False

    async def login_to_chatgpt(self) -> bool:
        try:
            logger.info("Starte ChatGPT-Login...")
            self.chatgpt_driver = self.setup_driver()
            self.chatgpt_driver.get("https://chat.openai.com")

            if self.detect_captcha(self.chatgpt_driver):
                while self.detect_captcha(self.chatgpt_driver):
                    time.sleep(5)

            login_button = WebDriverWait(self.chatgpt_driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Log in')]"))
            )
            login_button.click()

            email_input = WebDriverWait(self.chatgpt_driver, 15).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            email_input.clear()
            email_input.send_keys(self.config['chatgpt_email'])

            continue_button = self.chatgpt_driver.find_element(By.XPATH, "//button[@type='submit']")
            continue_button.click()

            password_input = WebDriverWait(self.chatgpt_driver, 15).until(
                EC.presence_of_element_located((By.ID, "password"))
            )
            password_input.clear()
            password_input.send_keys(self.config['chatgpt_password'])

            continue_button = self.chatgpt_driver.find_element(By.XPATH, "//button[@type='submit']")
            continue_button.click()

            WebDriverWait(self.chatgpt_driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "textarea[placeholder*='Message']"))
            )
            logger.info("Successfully logged into ChatGPT")
            return True

        except Exception as e:
            logger.error(f"ChatGPT login error: {e}")
            return False

    def setup_driver(self) -> webdriver.Chrome:
        chrome_options = Options()
        if self.config.get('headless', False):
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")

        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(30)
        driver.implicitly_wait(10)
        return driver

    async def start(self):
        logger.info("Starting Teams-ChatGPT Bot...")
        teams_login = await self.login_to_teams()
        if not teams_login:
            logger.error("Failed to login to Teams")
            return

        chatgpt_login = await self.login_to_chatgpt()
        if not chatgpt_login:
            logger.error("Failed to login to ChatGPT")
            return

        logger.info("Bot successfully started and ready!")
        self.is_running = True
        while self.is_running:
            await asyncio.sleep(self.config.get('check_interval', 10))

    async def stop(self):
        logger.info("Stopping bot...")
        self.is_running = False
        if self.teams_driver:
            self.teams_driver.quit()
        if self.chatgpt_driver:
            self.chatgpt_driver.quit()
        logger.info("Bot stopped")


def load_config() -> Dict[str, Any]:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        config = {
            "teams_email": "",
            "teams_password": "",
            "chatgpt_email": "",
            "chatgpt_password": "",
            "bot_trigger": "@bot",
            "check_interval": 10,
            "headless": False,
            "max_retries": 3,
            "retry_delay": 30
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"Konfigurationsdatei '{CONFIG_FILE}' erstellt. Bitte ausfüllen!")
        return config


async def main():
    config = load_config()
    missing = [f for f in ['teams_email', 'teams_password', 'chatgpt_email', 'chatgpt_password'] if not config.get(f)]
    if missing:
        print(f"Bitte folgende Felder in {CONFIG_FILE} ausfüllen: {', '.join(missing)}")
        return

    bot = TeamsBot(config)
    try:
        await bot.start()
    except KeyboardInterrupt:
        logger.info("Bot interrupted by user")
    except Exception as e:
        logger.error(f"Bot crashed: {e}")
    finally:
        await bot.stop()


if __name__ == "__main__":
    asyncio.run(main())

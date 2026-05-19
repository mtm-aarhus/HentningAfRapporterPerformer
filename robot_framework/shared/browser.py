import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


def setup_driver(downloads_folder: str) -> webdriver.Chrome:
    """Opsætter og returnerer en Chrome WebDriver."""
    chrome_options = Options()
    chrome_options.add_argument('--remote-debugging-pipe')
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--safebrowsing-disable-download-protection")
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": downloads_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })

    chrome_service = Service()
    driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
    # Sæt timeout højere end standard 120 sekunder
    driver.command_executor.set_timeout(60 * 20)  # 20 minutter
    return driver


def login(driver: webdriver.Chrome, url: str, username: str, password: str) -> None:
    """Logger ind på Opus-portalen."""
    print("Navigerer til Opus login-side")
    driver.get(url)
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
    driver.find_element(By.ID, "logonuidfield").send_keys(username)
    driver.find_element(By.ID, "logonpassfield").send_keys(password)
    driver.find_element(By.ID, "buttonLogon").click()
    print("Logget ind på Opus")
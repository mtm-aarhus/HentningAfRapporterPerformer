import os
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from robot_framework.shared.file_utils import (
    get_downloads_folder,
    wait_for_download,
    convert_xls_to_xlsx,
    close_extra_windows,
)
from robot_framework.shared.sharepoint import upload_file
from datetime import datetime
from dateutil.relativedelta import relativedelta

def get_end_date(today: datetime = None) -> str:
    if today is None:
        today = datetime.now()
    
    if today.day < 6:
        # Ultimo to måneder tilbage
        end = today.replace(day=1) - relativedelta(months=1) - relativedelta(days=1)
    else:
        # Ultimo sidste måned
        end = today.replace(day=1) - relativedelta(days=1)
    
    return end.strftime("%d.%m.%Y")

def _build_runs() -> list[dict]:
    today = datetime.now()
    start = (today - relativedelta(months=3)).replace(day=1)
    end = get_end_date(today)
    return [
        {
            "name": "run1",
            "start": start.strftime("%d.%m.%Y"),
            "end": end,
            "Filename": "run1file",
        }
    ]

def _navigate_to_report(driver, url: str) -> None:
    """Navigerer til rapport-URL'en med op til 3 forsøg."""
    for attempt in range(3):
        try:
            print(f"Forsøg {attempt + 1}")
            driver.get(url)
            wait = WebDriverWait(driver, 30)
            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            wait.until(lambda d: d.current_url.startswith(url))
            return
        except Exception as e:
            print(f"Fejl på forsøg {attempt + 1}: {e}")
    raise Exception(f"Kunne ikke nå {url} efter 3 forsøg")


def _download_run(driver, run: dict, url: str, downloads_folder: str) -> str:
    """
    Kører Flex Total rapport i Opus og downloader filen.
    Returnerer stien til den downloadede .xls-fil.
    """
    _navigate_to_report(driver, url)

    wait = WebDriverWait(driver, 300)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.NAME, "Flex total")
    ))

    # Vent på at siden er færdig med at loade
    wait.until(lambda d: len(d.find_element(By.TAG_NAME, "body").text.strip()) > 0)

    # Step 1: Klik på Varianter
    element = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//div[@ct='B' and .//span[normalize-space(text())='Varianter']]"
    )))
    element.click()
    print("Varianter-knap klikket")

    driver.switch_to.default_content()

    # Vent på at popup-iframen dukker op
    wait.until(EC.presence_of_element_located((By.NAME, "URLSPW-0")))

    # Switch til den
    wait.until(EC.frame_to_be_available_and_switch_to_it(
            (By.NAME, "URLSPW-0")
        ))

    celle = driver.find_element(By.XPATH, 
        "//td[@acf='CSEL' and .//span[normalize-space(text())='MTM ALLE']]"
    )

    # Prøv mousedown + mouseup events
    driver.execute_script("""
        var el = arguments[0];
        var rect = el.getBoundingClientRect();
        var x = rect.left + rect.width / 2;
        var y = rect.top + rect.height / 2;
        
        el.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, cancelable: true, clientX: x, clientY: y}));
        el.dispatchEvent(new MouseEvent('mouseup', {bubbles: true, cancelable: true, clientX: x, clientY: y}));
        el.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, clientX: x, clientY: y}));
    """, celle)
    print("MTM ALLE mousedown/mouseup/click afsendt")
    # Klik "Overfør til skærm"
    overfør = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//div[@ct='B' and .//span[normalize-space(text())='Overfør til skærm']]"
    )))
    overfør.click()
    print("Overfør til skærm klikket")

    #Skifter tilbage til original iframe
    driver.switch_to.default_content()
    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.NAME, "Flex total")
    ))
    time.sleep(5)

    # Step 3: Tjek at "Inkluder underliggende" er markeret
    inkluder = driver.find_element(By.XPATH,
        "//span[contains(text(), 'Inkludér underliggende')]/ancestor::td/following-sibling::td//input[@type='checkbox']"
    )
    if not inkluder.is_selected():
        inkluder.click()
        print("Inkludér underliggende markeret")
    else:
        print("Inkludér underliggende allerede markeret")

    # Step 4: Sæt skæringsdato
    skæringsdato = wait.until(EC.presence_of_element_located((
        By.XPATH,
        "//span[contains(text(), 'Skæringsdato')]/ancestor::td/following-sibling::td//input[@type='text']"
    )))
    skæringsdato.clear()
    skæringsdato.send_keys(run["end"])
    print(f"Skæringsdato sat til {run['end']}")

    # Step 5: Klik Søg
    søg = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//div[@ct='B' and .//span[normalize-space(text())='Søg']]"
    )))
    søg.click()
    print("Søg klikket")

    # Vent på at data er loadet - venter på at tabellen ikke længere viser "Tabel indeholder ingen data"
    wait.until_not(EC.presence_of_element_located((
        By.XPATH, "//span[contains(text(), 'Tabel indeholder ingen data')]"
    )))
    print("Data er loadet")

    # Step 6: Klik Eksport
    initial_file_count = len(os.listdir(downloads_folder))

    eksport = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//div[@ct='B' and .//span[normalize-space(text())='Eksport']]"
    )))
    eksport.click()
    print("Eksport klikket")
    excel_span = wait.until(EC.presence_of_element_located((
        By.XPATH, "//span[normalize-space(text())='Eksport til Excel']"
    )))

    driver.execute_script("""
        var el = arguments[0].parentElement;
        var rect = el.getBoundingClientRect();
        var x = rect.left + rect.width / 2;
        var y = rect.top + rect.height / 2;
        
        el.dispatchEvent(new MouseEvent('mousedown', {bubbles: true, cancelable: true, clientX: x, clientY: y}));
        el.dispatchEvent(new MouseEvent('mouseup', {bubbles: true, cancelable: true, clientX: x, clientY: y}));
        el.dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true, clientX: x, clientY: y}));
    """, excel_span)
    print("Eksport til Excel klikket")
    xls_path = wait_for_download(downloads_folder, initial_file_count, run["Filename"])
    close_extra_windows(driver)
    return xls_path


def run(driver, orchestrator_connection, data: dict) -> None:
    navn = data["Navn"]
    sti = data["Sti"]
    sharepoint_url = data["SharePointMappeLink"]

    downloads_folder = get_downloads_folder()
    run_config = _build_runs()[0]

    print(f"\n--- Starter {run_config['name']} ---")
    xlsx_path = _download_run(driver, run_config, sti, downloads_folder)

    upload_file(
        local_file_path=xlsx_path,
        sharepoint_url=sharepoint_url,
        orchestrator_connection=orchestrator_connection,
        remote_filename=f"{navn}.xlsx",
    )
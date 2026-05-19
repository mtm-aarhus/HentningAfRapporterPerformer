import os
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
            print("Siden er klar")
            return
        except Exception as e:
            print(f"Fejl på forsøg {attempt + 1}: {e}")
    raise Exception(f"Kunne ikke nå {url} efter 3 forsøg")


def _download_run(driver, run: dict, url: str, downloads_folder: str) -> str:
    """
    Kører én rapport-periode i Opus og downloader filen.
    Returnerer stien til den downloadede .xls-fil.
    """
    _navigate_to_report(driver, url)

    wait = WebDriverWait(driver, 6000)

    wait.until(EC.frame_to_be_available_and_switch_to_it(
        (By.CSS_SELECTOR, "iframe[title='Flex total']")
    ))

    element = wait.until(EC.element_to_be_clickable((
    By.XPATH, "//div[@ct='B' and .//span[normalize-space(text())='Varianter']]"
    )))
    element.click()
    print("Varianter-knap klikket")

    arbejdsdato_input = wait.until(EC.presence_of_element_located((
        By.XPATH,
        "//span[contains(text(), 'Arbejdsdato')]/ancestor::td/following-sibling::td//input[@type='text']"
    )))

    initial_file_count = len(os.listdir(downloads_folder))
    print(f"Antal filer før download: {initial_file_count}")

    arbejdsdato_input.clear()
    arbejdsdato_input.send_keys(f"{run['start']} - {run['end']}")

    ok_knap = wait.until(EC.element_to_be_clickable((By.ID, "DLG_VARIABLE_dlgBase_BTNOK")))
    ok_knap.click()

    print("Venter på eksport-knap...")
    WebDriverWait(driver, timeout=60 * 15).until(
        EC.presence_of_element_located((By.ID, "BUTTON_EXPORT_btn1_acButton"))
    )
    driver.find_element(By.ID, "BUTTON_EXPORT_btn1_acButton").click()
    print("Eksport-knap klikket")

    xls_path = wait_for_download(downloads_folder, initial_file_count, run["Filename"])
    close_extra_windows(driver)
    return xls_path


def _find_table_start(file: str) -> int:
    """Finder rækkenummer for tabelstarten (0-indekseret til pandas)."""
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Operationsnr. og tekst":
                return cell.row - 1
    raise ValueError(f"Tabelstart ikke fundet i {file}")


def _merge_xlsx_files(file_paths: list[str]) -> pd.DataFrame:
    """Læser og sammensætter flere xlsx-filer til én dataframe."""
    dfs = []
    for i, file in enumerate(file_paths):
        start_row = _find_table_start(file)
        df = pd.read_excel(file, skiprows=start_row)
        if i > 0:
            df = df.iloc[1:]
        df.columns = [col if col else "" for col in df.columns]
        df = df.dropna(how='all')
        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)


def _write_excel(result: pd.DataFrame, output_path: str) -> None:
    """Skriver den samlede dataframe til en formateret Excel-fil."""
    with pd.ExcelWriter(output_path, engine='xlsxwriter', datetime_format='dd-mm-yyyy') as writer:
        result.to_excel(writer, sheet_name='YKMD_STD', index=False)
        workbook = writer.book
        worksheet = writer.sheets['YKMD_STD']

        max_row, max_col = result.shape
        col_letter = chr(65 + max_col - 1)
        table_range = f"A1:{col_letter}{max_row + 1}"

        worksheet.add_table(table_range, {
            'name': 'SamletTabel',
            'columns': [{'header': col if col else ""} for col in result.columns],
        })

        for i, column in enumerate(result.columns):
            max_len = max(
                result[column].astype(str).map(len).max(),
                len(str(column)),
            )
            worksheet.set_column(i, i, max_len + 2)

    print(f"Excel skrevet til {output_path}")


def run(driver, orchestrator_connection, data: dict) -> None:
    navn = data["Navn"]
    sti = data["Sti"]
    sharepoint_url = data["SharePointMappeLink"]

    downloads_folder = get_downloads_folder()
    run_config = _build_runs()[0]

    print(f"\n--- Starter {run_config['name']} ---")
    xls_path = _download_run(driver, run_config, sti, downloads_folder)
    xlsx_path = convert_xls_to_xlsx(xls_path, sheet_name="Sheet1")

    upload_file(
        local_file_path=xlsx_path,
        sharepoint_url=sharepoint_url,
        orchestrator_connection=orchestrator_connection,
        remote_filename=f"{navn}.xlsx",
    )
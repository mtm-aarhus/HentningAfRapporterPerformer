import os
import gc
import shutil
import time
import subprocess
import win32com.client as win32


ILLEGAL_FILENAMES = [
    "YKMD_STD.xls", "run1file.xls", "run2file.xls", "run3file.xls", "run4file.xls",
    "YKMD_STD.xlsx", "run1file.xlsx", "run2file.xlsx", "run3file.xlsx", "run4file.xlsx",
    "samlet.xlsx", "samlet_pænt.xlsx",
    "PSA011, Mangler Timeregistrering.xlsx ",
    "PSA011, Manglende Timeregistrering.xlsx ",
]

_conversion_in_progress = set()


def get_downloads_folder() -> str:
    return os.path.join(os.path.expanduser("~"), "Downloads")


def delete_temp_files(downloads_folder: str = None) -> None:
    """Sletter midlertidige filer fra tidligere kørsler."""
    if downloads_folder is None:
        downloads_folder = get_downloads_folder()
    for filename in ILLEGAL_FILENAMES:
        full_path = os.path.join(downloads_folder, filename)
        if os.path.exists(full_path):
            os.remove(full_path)
            print(f"{filename} slettet")


def convert_xls_to_xlsx(path: str, sheet_name: str = "Sheet1") -> str:
    """
    Konverterer en .xls-fil til .xlsx via Excel COM-automatisering.
    Returnerer stien til den nye .xlsx-fil.
    """
    absolute_path = os.path.abspath(path)

    if absolute_path in _conversion_in_progress:
        print(f"Konvertering allerede i gang for {absolute_path}. Springer over.")
        return os.path.splitext(absolute_path)[0] + ".xlsx"

    _conversion_in_progress.add(absolute_path)
    try:
        print(f"Konverterer {absolute_path}")
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(absolute_path)
        wb.Sheets(1).Name = sheet_name

        new_path = os.path.splitext(absolute_path)[0] + ".xlsx"
        wb.SaveAs(new_path, FileFormat=51)
        wb.Close()
        excel.Application.Quit()
        del wb
        del excel
        return new_path

    except AttributeError as e:
        if "CLSIDToClassMap" in str(e):
            print("Korrupt gen_py cache fundet, rydder op...")
            shutil.rmtree(win32.gencache.GetGeneratePath(), ignore_errors=True)
            return convert_xls_to_xlsx(path, sheet_name)
        raise

    except Exception as e:
        print(f"Uventet fejl under konvertering: {e}")
        gc.collect()
        subprocess.call("taskkill /im excel.exe /f >nul 2>&1", shell=True)
        time.sleep(2)
        raise

    finally:
        _conversion_in_progress.discard(absolute_path)


def wait_for_download(downloads_folder: str, initial_file_count: int, filename: str, timeout: int = 3600) -> str:
    """
    Venter på at en ny .xls-fil dukker op i downloads-mappen,
    omdøber den til `filename`.xls og returnerer den nye sti.
    """
    start_time = time.time()
    while True:
        files = os.listdir(downloads_folder)
        if len(files) > initial_file_count:
            latest_file = max(
                [os.path.join(downloads_folder, f) for f in files],
                key=os.path.getctime,
            )
            if latest_file.endswith(".xls"):
                new_path = os.path.join(downloads_folder, f"{filename}.xls")
                os.rename(latest_file, new_path)
                print(f"Fil downloadet og omdøbt til {new_path}")
                return new_path

        if time.time() - start_time > timeout:
            raise TimeoutError("Fil-download fuldførtes ikke inden for tidsgrænsen.")
        time.sleep(1)


def close_extra_windows(driver) -> None:
    """Lukker eventuelle ekstra browser-vinduer og skifter tilbage til det primære."""
    for handle in driver.window_handles:
        if handle != driver.current_window_handle:
            driver.switch_to.window(handle)
            driver.close()
            print("Ekstra vindue lukket")
    driver.switch_to.window(driver.window_handles[0])
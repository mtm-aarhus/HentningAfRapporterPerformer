import json
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement

from robot_framework.shared import browser
from robot_framework.shared.file_utils import get_downloads_folder, delete_temp_files
from robot_framework.processes import manglertidsregistrering
from robot_framework.processes import manglertidsregistreringmtm
from robot_framework.processes import mtmflexrapport
# Mapning fra QueueName til den proces der skal køre
PROCESS_MAP = {
    "PSA011, Mangler Timeregistrering": manglertidsregistrering.run,
    "PSA011, Mangler Timeregistrering, MTM": manglertidsregistreringmtm.run,
    "MTMFlexRapport": mtmflexrapport.run
}


def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    data = json.loads(queue_element.data)
    queue_name = data.get("QueueName")

    if queue_name not in PROCESS_MAP:
        orchestrator_connection.log_error('Kønavnet har ikke en tilsvarende process')

    downloads_folder = get_downloads_folder()
    delete_temp_files(downloads_folder)

    opus_login = orchestrator_connection.get_credential("OpusBruger")
    opus_url = orchestrator_connection.get_constant("OpusAdgangUrl").value

    driver = browser.setup_driver(downloads_folder)
    try:
        browser.login(driver, opus_url, opus_login.username, opus_login.password)
        PROCESS_MAP[queue_name](driver, orchestrator_connection, data)
    finally:
        driver.quit()
        delete_temp_files(downloads_folder)
        print("Browser lukket og midlertidige filer slettet")

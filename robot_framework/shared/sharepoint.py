import os
from urllib.parse import urlparse, parse_qs, unquote
from office365.sharepoint.client_context import ClientContext
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection


def upload_file(
    local_file_path: str,
    sharepoint_url: str,
    orchestrator_connection: OrchestratorConnection,
    remote_filename: str = None,
) -> None:
    """
    Uploader en lokal fil til en SharePoint-mappe.

    Args:
        local_file_path:         Sti til den lokale fil der skal uploades.
        sharepoint_url:          URL til den SharePoint-mappe der skal uploades til.
        orchestrator_connection: OrchestratorConnection til at hente credentials.
        remote_filename:         Filnavn på SharePoint. Bruger local_file_path's navn hvis None.
    """
    parsed_url = urlparse(sharepoint_url)
    base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"

    if "/Teams/" in sharepoint_url:
        teamsite = sharepoint_url.split('Teams/')[1].split('/')[0]
        base_url = f"{base_url}/Teams/{teamsite}"
    elif "/Sites/" in sharepoint_url:
        sitename = sharepoint_url.split('Sites/')[1].split('/')[0]
        base_url = f"{base_url}/Sites/{sitename}"
    else:
        print("ADVARSEL: Kunne ikke afgøre om URL er Teams eller Sites. Bruger standard base_url.")

    certification = orchestrator_connection.get_credential("SharePointCert")
    api = orchestrator_connection.get_credential("SharePointAPI")
    cert_credentials = {
        "tenant": api.username,
        "client_id": api.password,
        "thumbprint": certification.username,
        "cert_path": certification.password,
    }
    ctx = ClientContext(base_url).with_client_certificate(**cert_credentials)

    query_params = parse_qs(parsed_url.query)
    id_param = query_params.get("id", [None])[0]

    if id_param:
        decoded_path = unquote(id_param).rstrip('/')
    else:
        if "/r/" in sharepoint_url:
            decoded_path = sharepoint_url.split('/r/', 1)[1].split('?', 1)[0]
        else:
            decoded_path = parsed_url.path.lstrip('/')

    decoded_path = decoded_path.replace("%20", " ")
    if not decoded_path.startswith("/"):
        decoded_path = "/" + decoded_path

    target_folder = ctx.web.get_folder_by_server_relative_path(decoded_path)
    ctx.load(target_folder)
    ctx.execute_query()

    file_name = remote_filename or os.path.basename(local_file_path)
    with open(local_file_path, "rb") as f:
        target_folder.upload_file(file_name, f.read()).execute_query()

    print(f"Fil uploadet til SharePoint: {file_name}")
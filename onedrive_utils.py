import os
from msal import ConfidentialClientApplication
import requests
from io import BytesIO

def enviar_para_onedrive(df, nome_arquivo="dados_formulario.xlsx"):
    client_id = os.environ['CLIENT_ID']
    client_secret = os.environ['CLIENT_SECRET']
    tenant_id = os.environ['TENANT_ID']
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scopes = ["https://graph.microsoft.com/.default"]

    app = ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    token_response = app.acquire_token_for_client(scopes=scopes)
    access_token = token_response.get("access_token")

    if not access_token:
        print("Erro ao obter token:", token_response.get("error_description"))
        return False

    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    upload_url = f"https://graph.microsoft.com/v1.0/users/estatistica.analisededados@defesacivil.am.gov.br/drive/root:/Comite-Enfrentamento/{nome_arquivo}:/content"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    response = requests.put(upload_url, headers=headers, data=buffer.read())
    return response.status_code == 200

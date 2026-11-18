"""Gera um refresh token OAuth para o Google Drive do Outlieria.

Execute localmente, nunca no Streamlit Cloud. O script não salva as credenciais
em arquivo e lê o Client ID e o Client Secret de variáveis de ambiente ou do
terminal.
"""

import getpass
import os

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def main():
    client_id = os.getenv("GOOGLE_CLIENT_ID") or input("Cole o GOOGLE_CLIENT_ID: ").strip()
    client_secret = os.getenv("GOOGLE_CLIENT_SECRET") or getpass.getpass("Cole o GOOGLE_CLIENT_SECRET (não será exibido): ").strip()

    if not client_id or not client_secret:
        raise SystemExit("Client ID e Client Secret são obrigatórios.")

    client_config = {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": ["http://localhost"],
        }
    }

    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    credentials = flow.run_local_server(
        port=0,
        access_type="offline",
        prompt="consent",
    )

    if not credentials.refresh_token:
        raise SystemExit(
            "O Google não retornou um refresh token. Execute novamente e mantenha prompt=consent."
        )

    print("\nGOOGLE_REFRESH_TOKEN (copie somente o valor abaixo):")
    print(credentials.refresh_token)
    print("\nNão salve esse token no GitHub nem o envie por mensagem.")


if __name__ == "__main__":
    main()

"""
GigaChat client (Sber).

Env vars:
    GIGACHAT_CREDENTIALS — base64(client_id:client_secret)
    GIGACHAT_SCOPE       — GIGACHAT_API_PERSONAL | GIGACHAT_API_CORP
"""
import os
from gigachat import GigaChat
from dotenv import load_dotenv

load_dotenv()


def get_client() -> GigaChat:
    credentials = os.getenv("GIGACHAT_CREDENTIALS")
    scope = os.getenv("GIGACHAT_SCOPE", "GIGACHAT_API_PERSONAL")
    if not credentials:
        raise EnvironmentError("GIGACHAT_CREDENTIALS is not set in .env")
    return GigaChat(credentials=credentials, scope=scope, verify_ssl_certs=False)


def ask(prompt: str, model: str = "GigaChat") -> str:
    """Send a single prompt and return the reply text."""
    with get_client() as client:
        response = client.chat(prompt)
        return response.choices[0].message.content


if __name__ == "__main__":
    reply = ask("Привет! Кто ты?")
    print(reply)

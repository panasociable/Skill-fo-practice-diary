"""
Yandex GPT client (Yandex Cloud ML SDK).

Env vars:
    YANDEX_FOLDER_ID — ID папки в Yandex Cloud
    YANDEX_API_KEY   — API-ключ сервисного аккаунта
"""
import os
from yandex_cloud_ml_sdk import YCloudML
from dotenv import load_dotenv

load_dotenv()


def get_client() -> YCloudML:
    folder_id = os.getenv("YANDEX_FOLDER_ID")
    api_key = os.getenv("YANDEX_API_KEY")
    if not folder_id or not api_key:
        raise EnvironmentError("YANDEX_FOLDER_ID and YANDEX_API_KEY must be set in .env")
    return YCloudML(folder_id=folder_id, auth=api_key)


def ask(prompt: str, model: str = "yandexgpt", temperature: float = 0.5) -> str:
    """Send a single prompt and return the reply text."""
    sdk = get_client()
    result = (
        sdk.models.completions(model)
        .configure(temperature=temperature)
        .run(prompt)
    )
    return result.alternatives[0].text


if __name__ == "__main__":
    reply = ask("Привет! Кто ты?")
    print(reply)

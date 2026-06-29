#!/usr/bin/env python3
"""
MCP-сервер для дневника практики ПГТУ.
Транспорт: stdio (для Claude Desktop).

Инструменты:
  fill_diary       — заполнить шаблон .docx по данным студента
  ask_gigachat     — задать вопрос GigaChat
  ask_yandexgpt    — задать вопрос Яндекс GPT

Запуск:
  python mcp_server.py

Добавление в Claude Desktop (~/.config/claude/claude_desktop_config.json):
  {
    "mcpServers": {
      "dnevnik-praktiki": {
        "command": "python",
        "args": ["/полный/путь/до/mcp_server.py"]
      }
    }
  }
"""

import sys
import json
import os
import tempfile

# ── пути ───────────────────────────────────────────────────────────────────
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from skills.dnevnik_praktiki.scripts.fill_diary import generate, InputError  # noqa: E402


# ══════════════════════════════════════════════════════════════════════════
#  Вспомогательные функции MCP (stdio JSON-RPC 2.0)
# ══════════════════════════════════════════════════════════════════════════

def send(obj: dict):
    """Отправить JSON-объект в stdout и сразу сбросить буфер."""
    print(json.dumps(obj, ensure_ascii=False), flush=True)


def ok(req_id, result):
    send({"jsonrpc": "2.0", "id": req_id, "result": result})


def err(req_id, code: int, message: str):
    send({"jsonrpc": "2.0", "id": req_id,
          "error": {"code": code, "message": message}})


# ══════════════════════════════════════════════════════════════════════════
#  Описание инструментов
# ══════════════════════════════════════════════════════════════════════════

TOOLS = [
    {
        "name": "fill_diary",
        "description": (
            "Заполняет официальный «Дневник практики» ПГТУ по данным студента "
            "и возвращает путь к готовому .docx файлу."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "fio":        {"type": "string", "description": "ФИО студента полностью"},
                "kafedra":    {"type": "string", "description": "Кафедра / факультет"},
                "spec":       {"type": "string", "description": "Специальность с кодом"},
                "kurs":       {"type": "integer", "description": "Курс (число)"},
                "gruppa":     {"type": "string", "description": "Группа"},
                "mesto":      {"type": "string", "description": "Место прохождения практики"},
                "start_date": {"type": "string", "description": "Начало практики ГГГГ-ММ-ДД"},
                "end_date":   {"type": "string", "description": "Конец практики ГГГГ-ММ-ДД"},
                "zadanie":    {"type": "string", "description": "Индивидуальное задание / тема"},
                "vid":        {"type": "string", "description": "Вид практики (необяз.)"},
                "tip":        {"type": "string", "description": "Тип практики (необяз.)"},
                "forma":      {"type": "string", "description": "Форма обучения (необяз.)"},
                "prikaz_n":   {"type": "string", "description": "Номер приказа (необяз.)"},
                "prikaz_date":{"type": "string", "description": "Дата приказа ГГГГ-ММ-ДД (необяз.)"},
                "dog_n":      {"type": "string", "description": "Номер договора (необяз.)"},
                "dog_date":   {"type": "string", "description": "Дата договора ГГГГ-ММ-ДД (необяз.)"},
                "instr_vuz":  {"type": "string", "description": "Руководитель от вуза (необяз.)"},
                "instr_org":  {"type": "string", "description": "Руководитель от организации (необяз.)"},
                "output_path":{"type": "string", "description": "Путь для сохранения .docx (необяз.)"},
            },
            "required": ["fio", "kafedra", "spec", "kurs", "gruppa",
                         "mesto", "start_date", "end_date", "zadanie"],
        },
    },
    {
        "name": "ask_gigachat",
        "description": (
            "Отправляет запрос в GigaChat (Сбер) и возвращает ответ. "
            "Требует GIGACHAT_CREDENTIALS и GIGACHAT_SCOPE в .env."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "prompt": {"type": "string", "description": "Текст запроса"},
                "model":  {"type": "string", "description": "Модель (по умолч. GigaChat)",
                           "default": "GigaChat"},
            },
            "required": ["prompt"],
        },
    },
    {
        "name": "ask_yandexgpt",
        "description": (
            "Отправляет запрос в Яндекс GPT и возвращает ответ. "
            "Требует YANDEX_FOLDER_ID и YANDEX_API_KEY в .env."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "prompt":      {"type": "string", "description": "Текст запроса"},
                "model":       {"type": "string", "description": "Модель (по умолч. yandexgpt)",
                                "default": "yandexgpt"},
                "temperature": {"type": "number", "description": "Температура 0-1 (по умолч. 0.5)",
                                "default": 0.5},
            },
            "required": ["prompt"],
        },
    },
]


# ══════════════════════════════════════════════════════════════════════════
#  Обработчики инструментов
# ══════════════════════════════════════════════════════════════════════════

def handle_fill_diary(args: dict) -> str:
    out_path = args.pop("output_path", None)
    if not out_path:
        tmp = tempfile.NamedTemporaryFile(
            suffix=".docx",
            prefix="Дневник_",
            delete=False,
        )
        tmp.close()
        out_path = tmp.name
    try:
        generate(args, out_path)
    except InputError as ex:
        raise ValueError(str(ex))
    return f"Документ сохранён: {out_path}"


def handle_ask_gigachat(args: dict) -> str:
    from providers.gigachat_client import ask
    return ask(args["prompt"], model=args.get("model", "GigaChat"))


def handle_ask_yandexgpt(args: dict) -> str:
    from providers.yandex_gpt_client import ask
    return ask(
        args["prompt"],
        model=args.get("model", "yandexgpt"),
        temperature=float(args.get("temperature", 0.5)),
    )


HANDLERS = {
    "fill_diary":    handle_fill_diary,
    "ask_gigachat":  handle_ask_gigachat,
    "ask_yandexgpt": handle_ask_yandexgpt,
}


# ══════════════════════════════════════════════════════════════════════════
#  Главный цикл
# ══════════════════════════════════════════════════════════════════════════

def main():
    for raw_line in sys.stdin:
        raw_line = raw_line.strip()
        if not raw_line:
            continue

        try:
            req = json.loads(raw_line)
        except json.JSONDecodeError as ex:
            send({"jsonrpc": "2.0", "id": None,
                  "error": {"code": -32700, "message": f"Parse error: {ex}"}})
            continue

        req_id = req.get("id")
        method = req.get("method", "")
        params = req.get("params", {})

        # ── стандартные методы MCP ──────────────────────────────────────
        if method == "initialize":
            ok(req_id, {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {}},
                "serverInfo": {"name": "dnevnik-praktiki", "version": "1.0.0"},
            })

        elif method == "tools/list":
            ok(req_id, {"tools": TOOLS})

        elif method == "tools/call":
            tool_name = params.get("name", "")
            tool_args  = params.get("arguments", {})
            handler = HANDLERS.get(tool_name)
            if handler is None:
                err(req_id, -32601, f"Инструмент не найден: {tool_name}")
                continue
            try:
                result_text = handler(tool_args)
                ok(req_id, {
                    "content": [{"type": "text", "text": result_text}],
                    "isError": False,
                })
            except Exception as ex:
                ok(req_id, {
                    "content": [{"type": "text", "text": f"Ошибка: {ex}"}],
                    "isError": True,
                })

        elif method == "notifications/initialized":
            pass  # уведомление, ответ не нужен

        else:
            err(req_id, -32601, f"Метод не поддерживается: {method}")


if __name__ == "__main__":
    main()

"""Client for Gemini proxy (OpenAI-compatible API)."""

from __future__ import annotations

import json
import requests

API_URL = "http://127.0.0.1:2048/v1/chat/completions"
MODEL = "gemini-flash-lite-latest"
TIMEOUT = 120

SYSTEM_PROMPT = (
    "You are a professional translator for a Nintendo Switch homebrew app called DBI. "
    "You will receive a JSON object with a Russian source string and a list of target language codes.\n\n"
    "Rules:\n"
    "1. Return ONLY a valid JSON object: {\"<lang_code>\": \"<translation>\", ...}\n"
    "2. Tokens like [[LF]], [[CR]], [[TAB]], [[ESC]] are special placeholders. "
    "Keep them EXACTLY as-is in every translation — same count, same order, same case.\n"
    "3. If the source text is 'ru' (lowercase, exactly two letters), return the language code itself "
    "(e.g. 'en' for English, 'ua' for Ukrainian).\n"
    "4. If the source text is pure ASCII (Latin characters, numbers, punctuation only), "
    "return it unchanged for every language.\n"
    "5. Preserve all format specifiers like {}, {:02}, {:<10}, \\x1b sequences (shown as [[ESC]]), etc.\n"
    "6. Do NOT add any explanation, markdown, or extra text — only the JSON object.\n"
)


def translate_batch(text: str, target_langs: list[str]) -> dict[str, str]:
    """Send a single source string for translation into multiple languages.

    Returns dict like {"en": "...", "de": "...", ...}.
    Raises on network/parse errors.
    """
    user_msg = json.dumps({"text": text, "languages": target_langs}, ensure_ascii=False)

    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_msg},
        ],
        "stream": False,
    }

    resp = requests.post(API_URL, json=payload, timeout=TIMEOUT)
    resp.raise_for_status()

    data = resp.json()
    content = data["choices"][0]["message"]["content"]

    # Strip markdown fences if the model wraps output
    content = content.strip()
    if content.startswith("```"):
        content = content.split("\n", 1)[1]
    if content.endswith("```"):
        content = content.rsplit("```", 1)[0]
    content = content.strip()

    return json.loads(content)

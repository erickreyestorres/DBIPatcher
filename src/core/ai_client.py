"""Gemini Proxy AI Client with Continuous Chat support."""

import json
import re
import requests
from pathlib import Path
from datetime import datetime
from typing import Optional

# ── Logging ──────────────────────────────────────────────────────────

LOG_DIR = Path(__file__).resolve().parent.parent.parent / "logs"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "ai_proxy.log"

def _log_interaction(payload: dict, response_text: str, row_id: Optional[int] = None) -> None:
    """Append request/response pair to the log file."""
    timestamp = datetime.now().isoformat()
    row_label = f"ROW: {row_id}" if row_id is not None else "GENERAL"
    try:
        log_entry = (
            f"\n{'='*20} {row_label} {'='*20}\n"
            f"TIMESTAMP: {timestamp}\n"
            f"REQUEST PAYLOAD:\n{json.dumps(payload, ensure_ascii=False, indent=2)}\n"
            f"{'-'*40}\n"
            f"RESPONSE TEXT:\n{response_text}\n"
            f"{'='*80}\n"
        )
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception as e:
        print(f"  [!] Failed to write to log: {e}")

# ── Configuration ────────────────────────────────────────────────────

API_URL                = "http://127.0.0.1:2048/v1/chat/completions"
NEW_CHAT_URL           = "http://127.0.0.1:2048/api/new-chat"
SYSTEM_INSTRUCTIONS_URL = "http://127.0.0.1:2048/api/system-instructions"
SWITCH_MODEL_URL       = "http://127.0.0.1:2048/api/switch-model"
MODEL                  = "gemini-flash-lite-latest"
TIMEOUT                = 180

SYSTEM_PROMPT = """\
You are a translation engine for a Nintendo Switch app called DBI.

INPUT FORMAT (JSON):
{"text": "<source_text_in_russian>", "languages": ["<lang_code>", ...]}

OUTPUT FORMAT (JSON only, nothing else):
{"<lang_code>": "<translated_text>", ...}

STRICT RULES:
1. Output ONLY the JSON object. No explanations, no thinking, no markdown, no comments.
2. Tokens [[LF]], [[CR]], [[TAB]], [[ESC]] are formatting placeholders — keep them exactly as-is.
3. Preserve all format specifiers: {}, {:02}, {:>3}, \\x1b, etc. Do NOT translate or modify them.
4. If the source text is the literal string "ru", return the 2-letter language code for each language (e.g. "en", "ua", "fr").
5. If the source text is pure ASCII/English, return it unchanged for every language.
6. Do NOT wrap output in ```json``` blocks. Just raw JSON.

EXAMPLE:
Input:  {"text": "Привет, мир!", "languages": ["en", "de", "ua"]}
Output: {"en": "Hello, world!", "de": "Hallo, Welt!", "ua": "Привіт, світе!"}
"""

# ── Public API ───────────────────────────────────────────────────────

def init_session() -> None:
    """Full initialization:
      1. POST /api/new-chat
      2. POST /api/switch-model
      3. POST /api/system-instructions
    """
    # Clear log
    print("  [INIT] Clearing log file...")
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"--- NEW SESSION STARTED AT {datetime.now().isoformat()} ---\n")

    # New chat
    print("  [INIT] Creating new chat...")
    try:
        resp = requests.post(NEW_CHAT_URL, timeout=15)
        print(f"  [INIT] New chat: OK ({resp.status_code})")
    except Exception as e:
        print(f"  [INIT] WARNING: new-chat failed: {e}")

    # Switch model first
    print(f"  [INIT] Switching model -> {MODEL}...")
    try:
        resp = requests.post(
            SWITCH_MODEL_URL,
            json={"model": MODEL},
            timeout=30,
        )
        print(f"  [INIT] Model switch: OK ({resp.status_code})")
    except Exception as e:
        print(f"  [INIT] WARNING: switch-model failed: {e}")

    # Then set system instructions
    print(f"  [INIT] Setting system instructions...")
    try:
        resp = requests.post(
            SYSTEM_INSTRUCTIONS_URL,
            json={"content": SYSTEM_PROMPT},
            timeout=15,
        )
        print(f"  [INIT] System instructions: OK ({resp.status_code})")
    except Exception as e:
        print(f"  [INIT] WARNING: system-instructions failed: {e}")

    print(f"  [INIT] Ready!")


def translate_batch(text: str, target_langs: list[str], row_id: Optional[int] = None) -> dict[str, str]:
    """Translate text with automatic retry and session recovery.
    
    On failure:
      1. Retry up to 2 times in the current session.
      2. If both retries fail, reinitialize the session (new chat + system instructions)
         and retry once more.
    """
    import time

    user_content = json.dumps(
        {"text": text, "languages": target_langs}, ensure_ascii=False
    )

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": user_content}],
        "stream": False,
    }

    MAX_RETRIES = 2
    last_error = None

    for attempt in range(MAX_RETRIES + 1):
        resp_text = "N/A"
        try:
            resp = requests.post(API_URL, json=payload, timeout=TIMEOUT)
            resp_text = resp.text

            if resp.status_code != 200:
                print(f"  [Row {row_id}] Proxy error {resp.status_code} (attempt {attempt + 1})")
                _log_interaction(payload, resp_text, row_id=row_id)
                raise requests.HTTPError(f"Status {resp.status_code}")

            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            _log_interaction(payload, resp_text, row_id=row_id)
            return _extract_json(content)

        except Exception as e:
            last_error = e
            if resp_text == "N/A" and "resp" in locals():
                resp_text = getattr(resp, 'text', 'N/A')
            _log_interaction(payload, resp_text, row_id=row_id)

            if attempt < MAX_RETRIES:
                wait = 2 * (attempt + 1)
                print(f"  [Row {row_id}] Retry {attempt + 1}/{MAX_RETRIES} in {wait}s...")
                time.sleep(wait)
            else:
                # Both retries failed — session is likely broken
                print(f"  [Row {row_id}] All {MAX_RETRIES} retries failed. Reinitializing session...")
                try:
                    init_session()
                    time.sleep(1)
                    # One final attempt with fresh session
                    resp = requests.post(API_URL, json=payload, timeout=TIMEOUT)
                    resp_text = resp.text
                    if resp.status_code == 200:
                        data = resp.json()
                        content = data["choices"][0]["message"]["content"]
                        _log_interaction(payload, resp_text, row_id=row_id)
                        return _extract_json(content)
                except Exception as recovery_error:
                    print(f"  [Row {row_id}] Session recovery also failed: {recovery_error}")
                    _log_interaction(payload, str(recovery_error), row_id=row_id)

    raise RuntimeError(f"Translation failed after {MAX_RETRIES} retries + session recovery: {last_error}")


def refine(correction: str, target_langs: list[str], row_id: Optional[int] = None) -> dict[str, str]:
    """Send a correction into the existing chat, with retry logic."""
    import time

    user_content = (
        f"{correction}\n\n"
        f"Return the corrected full JSON object with ALL languages: "
        f"{', '.join(target_langs)}."
    )

    payload = {
        "model": MODEL,
        "messages": [{"role": "user", "content": user_content}],
        "stream": False,
    }

    MAX_RETRIES = 2
    last_error = None

    for attempt in range(MAX_RETRIES + 1):
        resp_text = "N/A"
        try:
            resp = requests.post(API_URL, json=payload, timeout=TIMEOUT)
            resp_text = resp.text

            if resp.status_code != 200:
                _log_interaction(payload, resp_text, row_id=row_id)
                raise requests.HTTPError(f"Status {resp.status_code}")

            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            _log_interaction(payload, resp_text, row_id=row_id)
            return _extract_json(content)

        except Exception as e:
            last_error = e
            if resp_text == "N/A" and "resp" in locals():
                resp_text = getattr(resp, 'text', 'N/A')
            _log_interaction(payload, resp_text, row_id=row_id)

            if attempt < MAX_RETRIES:
                wait = 2 * (attempt + 1)
                print(f"  [Row {row_id}] Refine retry {attempt + 1}/{MAX_RETRIES} in {wait}s...")
                time.sleep(wait)
            else:
                print(f"  [Row {row_id}] Refine retries exhausted. Reinitializing session...")
                try:
                    init_session()
                    time.sleep(1)
                    resp = requests.post(API_URL, json=payload, timeout=TIMEOUT)
                    resp_text = resp.text
                    if resp.status_code == 200:
                        data = resp.json()
                        content = data["choices"][0]["message"]["content"]
                        _log_interaction(payload, resp_text, row_id=row_id)
                        return _extract_json(content)
                except Exception as recovery_error:
                    print(f"  [Row {row_id}] Refine recovery failed: {recovery_error}")

    raise RuntimeError(f"Refine failed after {MAX_RETRIES} retries + recovery: {last_error}")


# ── Helpers ──────────────────────────────────────────────────────────

def _extract_json(content: str) -> dict[str, str]:
    """Robustly extract a JSON object from AI response text.
    
    Handles cases where:
    - AI wraps JSON in ```json ... ``` code blocks
    - AI adds "thinking" text before the JSON
    - Translations contain {} format specifiers that look like empty JSON
    """
    content = content.strip()

    # Strategy 1: Look for ```json ... ``` code blocks
    code_blocks = re.findall(r"```(?:json)?\s*(\{.*?\})\s*```", content, re.DOTALL)
    if code_blocks:
        json_str = code_blocks[-1]
        return _parse_json_safe(json_str)

    # Strategy 2: Find the LAST large JSON-like block (the one with quoted keys)
    # This avoids matching small {} format specifiers in translations
    # Look for blocks that contain at least one "key": "value" pair
    json_blocks = re.findall(r'(\{[^{}]*(?:"[^"]*"\s*:\s*"[^"]*"[^{}]*)+\})', content, re.DOTALL)
    if json_blocks:
        # Take the last one (most likely the actual translation JSON)
        return _parse_json_safe(json_blocks[-1])

    # Strategy 3: Find the biggest { ... } block (greedy, last resort)
    # Use rfind to start from the end
    last_close = content.rfind("}")
    if last_close >= 0:
        # Walk backwards to find matching opening brace
        depth = 0
        for i in range(last_close, -1, -1):
            if content[i] == "}":
                depth += 1
            elif content[i] == "{":
                depth -= 1
                if depth == 0:
                    json_str = content[i:last_close + 1]
                    return _parse_json_safe(json_str)

    raise json.JSONDecodeError("No JSON object found in AI response", content, 0)


def _parse_json_safe(json_str: str) -> dict[str, str]:
    """Parse JSON string with common AI mistake fixes."""
    # Fix trailing commas
    json_str = re.sub(r",\s*}", "}", json_str)
    json_str = re.sub(r",\s*]", "]", json_str)
    return json.loads(json_str)

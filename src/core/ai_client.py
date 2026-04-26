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

# Providers: "GEMINI_PROXY" or "OMNIROAD"
PROVIDER               = "OMNIROAD"

# Gemini Proxy Config
API_URL                = "http://127.0.0.1:2048/v1/chat/completions"
NEW_CHAT_URL           = "http://127.0.0.1:2048/api/new-chat"
SYSTEM_INSTRUCTIONS_URL = "http://127.0.0.1:2048/api/system-instructions"
SWITCH_MODEL_URL       = "http://127.0.0.1:2048/api/switch-model"
MODEL_GEMINI           = "gemini-flash-lite-latest"

# OmniRoad Config
OMNIROAD_URL           = "http://localhost:20128/v1/chat/completions"
MODEL_OMNI             = "kr/claude-sonnet-4.5"

# Active Model (will be chosen based on PROVIDER)
MODEL = MODEL_OMNI if PROVIDER == "OMNIROAD" else MODEL_GEMINI

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
7. Use ONLY ASCII punctuation: () NOT （）, : NOT ：, [] NOT ［］, ! NOT ！. Preserve the exact same count of brackets and colons as in the source text.

EXAMPLE:
Input:  {"text": "Привет, мир!", "languages": ["en", "de", "ua"]}
Output: {"en": "Hello, world!", "de": "Hallo, Welt!", "ua": "Привіт, світе!"}
"""

# ── Public API ───────────────────────────────────────────────────────

def init_session() -> None:
    """Initialize session based on provider."""
    # Clear log
    print("  [INIT] Clearing log file...")
    with open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(f"--- NEW SESSION STARTED AT {datetime.now().isoformat()} | PROVIDER: {PROVIDER} ---\n")

    if PROVIDER == "OMNIROAD":
        print(f"  [INIT] OmniRoad ({MODEL}) selected. Ready!")
        return

    # Gemini Proxy initialization
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


SHADOK_SYSTEM_PROMPT = """\
You are a literary translator. You will receive a Russian text that is a single cohesive literary passage (an easter egg from a Nintendo Switch app).

INPUT FORMAT (JSON):
{"text": "<full_russian_text>", "languages": ["<lang_code>", ...], "max_line_length": <number>}

OUTPUT FORMAT (JSON only, nothing else):
{"<lang_code>": "<translated_full_text>", ...}

STRICT RULES:
1. Translate the ENTIRE text as ONE literary passage. Preserve the narrative flow, humor, and style.
2. Output ONLY the JSON object. No explanations, no thinking, no markdown.
3. Each translated text must be a single string with newlines (\\n) separating lines.
4. Each line in the translation MUST NOT exceed max_line_length characters.
5. Do NOT add extra lines. Keep line count equal to or less than the original.
6. Preserve proper names: Shadoks=Шадоки, Gibis=Гібі (adapt to target language).
7. Preserve numbers (999999) as-is.
8. Do NOT wrap output in ```json``` blocks.
"""


def init_session_shadok() -> None:
    """Initialize a NEW chat session with shadok-specific system instructions."""
    if PROVIDER == "OMNIROAD":
        print(f"  [SHADOK-INIT] OmniRoad ({MODEL}) selected. Ready!")
        return

    print("  [SHADOK-INIT] Creating new chat for Shadok block...")
    try:
        requests.post(NEW_CHAT_URL, timeout=15)
    except Exception as e:
        print(f"  [SHADOK-INIT] WARNING: new-chat failed: {e}")

    print(f"  [SHADOK-INIT] Switching model -> {MODEL}...")
    try:
        requests.post(SWITCH_MODEL_URL, json={"model": MODEL}, timeout=30)
    except Exception as e:
        print(f"  [SHADOK-INIT] WARNING: switch-model failed: {e}")

    print(f"  [SHADOK-INIT] Setting Shadok system instructions...")
    try:
        requests.post(
            SYSTEM_INSTRUCTIONS_URL,
            json={"content": SHADOK_SYSTEM_PROMPT},
            timeout=15,
        )
    except Exception as e:
        print(f"  [SHADOK-INIT] WARNING: system-instructions failed: {e}")

    print(f"  [SHADOK-INIT] Ready!")


def translate_shadok_block(full_text: str, target_langs: list[str], max_line_length: int) -> dict[str, str]:
    """Translate the full Shadok text as one literary block."""
    import time

    user_content = json.dumps(
        {"text": full_text, "languages": target_langs, "max_line_length": max_line_length},
        ensure_ascii=False
    )

    if PROVIDER == "OMNIROAD":
        messages = [
            {"role": "system", "content": SHADOK_SYSTEM_PROMPT},
            {"role": "user", "content": user_content}
        ]
        url = OMNIROAD_URL
    else:
        messages = [{"role": "user", "content": user_content}]
        url = API_URL

    payload = {
        "model": MODEL,
        "messages": messages,
        "stream": False,
    }

    MAX_RETRIES = 2
    last_error = None

    for attempt in range(MAX_RETRIES + 1):
        resp_text = "N/A"
        try:
            resp = requests.post(url, json=payload, timeout=300)
            resp_text = resp.text

            if resp.status_code != 200:
                _log_interaction(payload, resp_text, row_id="SHADOK")
                raise requests.HTTPError(f"Status {resp.status_code}")

            data = resp.json()
            content = data["choices"][0]["message"]["content"]
            _log_interaction(payload, resp_text, row_id="SHADOK")
            return _extract_json(content)

        except Exception as e:
            last_error = e
            _log_interaction(payload, resp_text, row_id="SHADOK")

            if attempt < MAX_RETRIES:
                wait = 3 * (attempt + 1)
                print(f"  [SHADOK] Retry {attempt + 1}/{MAX_RETRIES} in {wait}s...")
                time.sleep(wait)
            else:
                print(f"  [SHADOK] All retries failed. Reinitializing...")
                try:
                    init_session_shadok()
                    time.sleep(1)
                    resp = requests.post(url, json=payload, timeout=300)
                    resp_text = resp.text
                    if resp.status_code == 200:
                        data = resp.json()
                        content = data["choices"][0]["message"]["content"]
                        _log_interaction(payload, resp_text, row_id="SHADOK")
                        return _extract_json(content)
                except Exception as recovery_error:
                    print(f"  [SHADOK] Recovery failed: {recovery_error}")

    raise RuntimeError(f"Shadok translation failed: {last_error}")


def translate_batch(text: str, target_langs: list[str], row_id: Optional[int] = None) -> dict[str, str]:
    """Translate text with automatic retry and session recovery."""
    import time

    user_content = json.dumps(
        {"text": text, "languages": target_langs}, ensure_ascii=False
    )

    if PROVIDER == "OMNIROAD":
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content}
        ]
        url = OMNIROAD_URL
    else:
        messages = [{"role": "user", "content": user_content}]
        url = API_URL

    payload = {
        "model": MODEL,
        "messages": messages,
        "stream": False,
    }

    MAX_RETRIES = 2
    last_error = None

    for attempt in range(MAX_RETRIES + 1):
        resp_text = "N/A"
        try:
            resp = requests.post(url, json=payload, timeout=TIMEOUT)
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
                    resp = requests.post(url, json=payload, timeout=TIMEOUT)
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

    if PROVIDER == "OMNIROAD":
        # For OmniRoad, we'd ideally want to keep history, but since we're simulating 
        # stateless calls here, we'll just send it as a follow-up.
        # However, without session management in OMNIROAD_URL, we'll just send it as user message.
        # Actually, OmniRoad should support chat history if we manage it. 
        # For now, let's keep it simple and send common SYSTEM_PROMPT.
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_content}
        ]
        url = OMNIROAD_URL
    else:
        messages = [{"role": "user", "content": user_content}]
        url = API_URL

    payload = {
        "model": MODEL,
        "messages": messages,
        "stream": False,
    }

    MAX_RETRIES = 2
    last_error = None

    for attempt in range(MAX_RETRIES + 1):
        resp_text = "N/A"
        try:
            resp = requests.post(url, json=payload, timeout=TIMEOUT)
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
                    resp = requests.post(url, json=payload, timeout=TIMEOUT)
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
    """Parse JSON string with common AI mistake fixes.
    
    Handles:
    - Trailing commas
    - Unescaped double quotes inside values (e.g. Chinese using " instead of ')
    """
    # Fix trailing commas
    json_str = re.sub(r",\s*}", "}", json_str)
    json_str = re.sub(r",\s*]", "]", json_str)
    
    try:
        return json.loads(json_str)
    except json.JSONDecodeError:
        pass
    
    # Fallback: manually extract "key": "value" pairs
    # This handles cases where AI puts unescaped " inside values
    result = {}
    # Match: "lang_code" : "...text..." followed by , or }
    # Use a state-machine approach to find key-value boundaries
    pairs = re.finditer(r'"([a-z]{2,6})"\s*:\s*"', json_str)
    
    for match in pairs:
        key = match.group(1)
        val_start = match.end()  # position right after the opening quote of value
        
        # Find the closing quote of this value:
        # Look for " followed by , or } or end of string
        # This handles unescaped " inside the value
        val_end = None
        for end_match in re.finditer(r'"(?:\s*[,}]|\s*$)', json_str[val_start:]):
            candidate = val_start + end_match.start()
            # Make sure this isn't the start of the next key
            remaining = json_str[candidate + 1:].lstrip()
            if remaining.startswith(',') or remaining.startswith('}') or not remaining:
                val_end = candidate
                break
        
        if val_end is not None:
            value = json_str[val_start:val_end]
            # Replace any unescaped double quotes with single quotes in the value
            value = value.replace('"', "'")
            result[key] = value
    
    if result:
        return result
    
    raise json.JSONDecodeError("Failed to parse JSON even with fallback", json_str, 0)

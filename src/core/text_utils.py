import re

# Мапа токенів для заміни
TOKEN_MAP = {
    r'\n': '[[LF]]',
    r'\r': '[[CR]]',
    r'\t': '[[TAB]]',
    r'\x1b': '[[ESC]]',
}

# Регулярний вираз для пошуку як літеральних представлень (\\n), так і реальних байтів (\n)
# Також враховуємо ESC (\\x1b або \x1b)
RE_TOKENS = re.compile(r'(\\n|\n|\\r|\r|\\t|\t|\\x1b|\x1b)')

def tokenize(text: str) -> str:
    """Перетворює керуючі символи та їх текстові представлення в токени."""
    if not text:
        return ""
    
    def replace_match(match):
        m = match.group(0)
        if m in ('\\n', '\n'): return '[[LF]]'
        if m in ('\\r', '\r'): return '[[CR]]'
        if m in ('\\t', '\t'): return '[[TAB]]'
        if m in ('\\x1b', '\x1b'): return '[[ESC]]'
        return m

    return RE_TOKENS.sub(replace_match, text)

def detokenize(text: str) -> str:
    """Перетворює токени назад у літеральні представлення (\\n і т.д.) для CSV білдера."""
    if not text:
        return ""
    text = text.replace('[[LF]]', '\\n')
    text = text.replace('[[CR]]', '\\r')
    text = text.replace('[[TAB]]', '\\t')
    text = text.replace('[[ESC]]', '\\x1b')
    return text

def normalize_tokens_out(text: str) -> str:
    """Виправляє випадки, коли AI повернув реальний символ замість токена."""
    if not text:
        return ""
    # Якщо AI видав справжній перенос рядка замість токена - фіксимо
    return tokenize(text)

# Regex для ANSI escape-послідовностей (наприклад \x1b[32;1m)
RE_ANSI = re.compile(r'(\[\[ESC\]\]\[[^\]]*\]|\\x1b\[[^\]]*\]|\x1b\[[^\x1b]*?[a-zA-Z])')

def visual_length(text: str) -> int:
    """Обчислює візуальну довжину рядка на екрані Nintendo Switch.
    
    Видаляє:
    - Токени [[LF]], [[CR]], [[TAB]], [[ESC]] (нульова ширина)
    - ANSI escape-послідовності типу [[ESC]][32;1m (кольори/форматування)
    - Літеральні \\n, \\r, \\t, \\x1b
    """
    if not text:
        return 0
    
    s = text
    # 1. Видаляємо ANSI escape-послідовності (з токенами або без)
    s = RE_ANSI.sub('', s)
    
    # 2. Видаляємо залишені токени (без ANSI-хвоста)
    s = s.replace('[[LF]]', '')
    s = s.replace('[[CR]]', '')
    s = s.replace('[[TAB]]', '')
    s = s.replace('[[ESC]]', '')
    
    # 3. Видаляємо літеральні представлення
    s = s.replace('\\n', '')
    s = s.replace('\\r', '')
    s = s.replace('\\t', '')
    s = s.replace('\\x1b', '')
    
    return len(s)


def normalize_fullwidth(text: str) -> str:
    """Replace CJK full-width characters with ASCII equivalents.
    
    DBI uses ASCII chars in its binary format, so full-width
    brackets, colons etc. from CJK translations must be normalized.
    """
    if not text:
        return text
    replacements = {
        "（": "(", "）": ")",
        "［": "[", "］": "]",
        "｛": "{", "｝": "}",
        "：": ":",
        "；": ";",
        "，": ",",
        "！": "!",
        "？": "?",
        "．": ".",
        "＝": "=",
        "＋": "+",
        "－": "-",
        "／": "/",
        "＼": "\\",
        "＊": "*",
        "＆": "&",
        "＠": "@",
        "＃": "#",
        "％": "%",
    }
    for fw, ascii_char in replacements.items():
        text = text.replace(fw, ascii_char)
    return text

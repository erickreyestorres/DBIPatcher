import re

def is_ascii_only(text: str) -> bool:
    """Перевіряє, чи складається рядок тільки з ASCII символів."""
    return all(ord(c) < 128 for c in text)

def get_tokens_list(text: str) -> list:
    """Витягує список токенів [[...]] у порядку їх появи."""
    return re.findall(r'(\[\[[A-Z]+\]\])', text)

def validate(original: str, translation: str, target_lang: str) -> tuple[bool, str]:
    """
    Повертає (success, error_message).
    """
    # 1. NotEmpty
    if not translation or not translation.strip():
        return False, "Translation is empty"

    # 2. RuPlaceholder
    if original.strip().lower() == "ru":
        if translation.strip().lower() != target_lang.lower():
            return False, f"Rule 'ru' failed: expected {target_lang}, got {translation}"
        return True, ""

    # 3. EnglishPreservation
    if is_ascii_only(original):
        # Allow slight whitespace differences if they were in the original, but usually should be identical
        if translation.strip() != original.strip():
            # Ми дозволяємо AI перекладати, але правило каже "AI не має права його змінювати"
            # Тому ми примусово повертаємо оригінал (це буде зроблено в пайплайні), 
            # але валідатор має сигналізувати, якщо переклад відхилився.
            return False, "English preservation failed: original is ASCII but translation differs"

    # 4. TokenPreservation
    orig_tokens = get_tokens_list(original)
    trans_tokens = get_tokens_list(translation)
    
    if orig_tokens != trans_tokens:
        return False, f"Token mismatch: original {orig_tokens}, translation {trans_tokens}"

    return True, ""

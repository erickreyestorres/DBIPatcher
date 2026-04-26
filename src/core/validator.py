import re
import json
from pathlib import Path

class Validator:
    def __init__(self, blocks_config_path=None):
        self.blocks = {}
        self.compiled_blocks = {}  # block_id -> [(pattern_str, compiled_regex)]
        
        if blocks_config_path and Path(blocks_config_path).exists():
            with open(blocks_config_path, "r", encoding="utf-8") as f:
                self.blocks = json.load(f)
        
        # Compile all regex patterns from blocks, grouped by block
        for block_id, patterns in self.blocks.items():
            self.compiled_blocks[block_id] = []
            for p in patterns:
                try:
                    self.compiled_blocks[block_id].append(
                        (p, re.compile(p, re.IGNORECASE))
                    )
                except re.error:
                    pass

    def check_placeholders(self, original, translation):
        """Checks if all {…} tags and %x specifiers are preserved exactly."""
        pattern = r"(\{[^}]*\}|%[a-zA-Z])"
        orig_placeholders = re.findall(pattern, original)
        trans_placeholders = re.findall(pattern, translation)
        
        if sorted(orig_placeholders) != sorted(trans_placeholders):
            return f"Placeholder mismatch: expected {orig_placeholders}, found {trans_placeholders}"
        return None

    def check_square_bracket_tags(self, original, translation):
        """Checks if all [content] tags are preserved exactly.
        
        Excludes [[TOKEN]] patterns (handled by check_tokens).
        Matches single-bracket tags like [+], [-1], [NSP], etc.
        """
        # Match [content] but NOT [[content]]
        pattern = r"(?<!\[)\[([^\[\]]+)\](?!\])"
        orig_tags = re.findall(pattern, original)
        trans_tags = re.findall(pattern, translation)
        
        if sorted(orig_tags) != sorted(trans_tags):
            return f"Square bracket tag mismatch: expected {['['+t+']' for t in orig_tags]}, found {['['+t+']' for t in trans_tags]}"
        return None

    def check_tokens(self, original, translation):
        """Checks if internal tokens like [[LF]], [[TAB]] are preserved."""
        pattern = r"\[\[[A-Z]+\]\]"
        orig_tokens = re.findall(pattern, original)
        trans_tokens = re.findall(pattern, translation)
        
        if sorted(orig_tokens) != sorted(trans_tokens):
            return f"Token mismatch: expected {orig_tokens}, found {trans_tokens}"
        return None

    def check_colon(self, original, translation):
        """Checks if colon presence is preserved (ignoring colons inside format specifiers)."""
        # Strip format specifiers like {:02X}, {:>3}, {:s} before checking
        strip_fmt = lambda s: re.sub(r'\{[^}]*\}', '', s)
        orig_clean = strip_fmt(original)
        trans_clean = strip_fmt(translation)
        
        if ":" in orig_clean and ":" not in trans_clean:
            return "Missing colon in translation"
        if ":" not in orig_clean and ":" in trans_clean:
            return "Unexpected colon in translation"
        return None

    def check_parentheses_count(self, original, translation):
        """Checks round bracket count: translation must have >= original count.
        
        Some languages (Korean, Japanese) use grammatical constructions
        like 이(가), を(は) that add legitimate parentheses.
        Translation can have MORE brackets, but not FEWER.
        """
        for char in "()":
            orig_count = original.count(char)
            trans_count = translation.count(char)
            if trans_count < orig_count:
                return f"Missing parenthesis '{char}': expected at least {orig_count}, found {trans_count}"
        return None

    def validate_row(self, original, translation):
        """Runs all checks for a single translation row."""
        if not translation:
            return ["Translation is empty"]
        
        errors = []
        
        err = self.check_placeholders(original, translation)
        if err: errors.append(err)
        
        # We intentionally SKIP check_square_bracket_tags for DBI. 
        # In DBI, square brackets are used for translatable statuses like [ОШИБКА], [ОТСУТСТВУЕТ].
        
        err = self.check_tokens(original, translation)
        if err: errors.append(err)
        
        err = self.check_colon(original, translation)
        if err: errors.append(err)
        
        err = self.check_parentheses_count(original, translation)
        if err: errors.append(err)
        
        return errors

    def validate_blocks(self, originals: dict[int, str]) -> list[str]:
        """
        Validates that every regex pattern in blocks.json matches exactly one
        row in the given originals dict (row_idx -> original_value).
        
        Call this AFTER alignment to ensure nothing was broken.
        Returns a list of error strings (empty = all OK).
        """
        errors = []
        
        for block_id, compiled_list in self.compiled_blocks.items():
            for pattern_str, regex in compiled_list:
                matches = []
                for row, val in originals.items():
                    if regex.search(val):
                        matches.append((row, val))
                
                if len(matches) == 0:
                    errors.append(
                        f"[{block_id}] Pattern NOT FOUND: {pattern_str[:60]}..."
                    )
                elif len(matches) > 1:
                    rows_info = ", ".join(f"row {r}" for r, _ in matches)
                    errors.append(
                        f"[{block_id}] AMBIGUOUS ({len(matches)} matches): "
                        f"{pattern_str[:40]}... -> {rows_info}"
                    )
        
        return errors


_validator_instance = None

def validate(original: str, translation: str, lang_code: str = "ua") -> tuple[bool, str]:
    """
    Compatibility wrapper for Validator class.
    Returns (success, error_message or "OK").
    Uses a module-level singleton to avoid repeated instantiation.
    """
    global _validator_instance
    if _validator_instance is None:
        _validator_instance = Validator()
    errs = _validator_instance.validate_row(original, translation)
    if errs:
        return False, "; ".join(errs)
    return True, "OK"


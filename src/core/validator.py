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
        """Checks if all technical placeholders like {} or %d are preserved."""
        pattern = r"(\{[^}]*\}|%[a-zA-Z])"
        orig_placeholders = re.findall(pattern, original)
        trans_placeholders = re.findall(pattern, translation)
        
        if sorted(orig_placeholders) != sorted(trans_placeholders):
            return f"Placeholder mismatch: expected {orig_placeholders}, found {trans_placeholders}"
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
        """Checks if colon presence is preserved. Accepts full-width colon ： as equivalent."""
        orig_has = ":" in original
        trans_has = ":" in translation or "：" in translation
        if orig_has and not trans_has:
            return "Missing colon in translation"
        if not orig_has and trans_has:
            return "Unexpected colon in translation"
        return None

    def check_bracket_counts(self, original, translation):
        """Checks bracket counts. Accepts CJK full-width brackets as equivalents."""
        # Normalize full-width brackets to ASCII for counting
        def normalize(text):
            return text.replace("（", "(").replace("）", ")").replace("［", "[").replace("］", "]")
        
        norm_orig = normalize(original)
        norm_trans = normalize(translation)
        
        for char in "()[]":
            if norm_orig.count(char) != norm_trans.count(char):
                return f"Bracket count mismatch for '{char}': expected {norm_orig.count(char)}, found {norm_trans.count(char)}"
        return None

    def validate_row(self, original, translation):
        """Runs all checks for a single translation row."""
        if not translation:
            return ["Translation is empty"]
        
        errors = []
        
        err = self.check_placeholders(original, translation)
        if err: errors.append(err)
        
        err = self.check_tokens(original, translation)
        if err: errors.append(err)
        
        err = self.check_colon(original, translation)
        if err: errors.append(err)
        
        err = self.check_bracket_counts(original, translation)
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


def validate(original: str, translation: str, lang_code: str = "ua") -> tuple[bool, str]:
    """
    Compatibility wrapper for Validator class.
    Returns (success, error_message or "OK").
    """
    v = Validator()
    errs = v.validate_row(original, translation)
    if errs:
        return False, "; ".join(errs)
    return True, "OK"


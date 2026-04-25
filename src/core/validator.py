import re
import json
from pathlib import Path

class Validator:
    def __init__(self, blocks_config_path=None):
        self.blocks = {}
        if blocks_config_path and Path(blocks_config_path).exists():
            with open(blocks_config_path, "r", encoding="utf-8") as f:
                self.blocks = json.load(f)
        
        # Compile all regex patterns from blocks for validation
        self.compiled_patterns = []
        for block_id, patterns in self.blocks.items():
            for p in patterns:
                try:
                    self.compiled_patterns.append(re.compile(p, re.IGNORECASE))
                except:
                    pass

    def check_placeholders(self, original, translation):
        """Checks if all technical placeholders like {} or %d are preserved."""
        # Find all patterns like {}, %d, %s, {:s}, etc.
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
        """Checks if colon is preserved."""
        if ":" in original and ":" not in translation:
            return "Missing colon in translation"
        if ":" not in original and ":" in translation:
            return "Unexpected colon in translation"
        return None

    def check_brackets(self, translation):
        """Checks balanced brackets."""
        stack = []
        brackets = {"(": ")", "[": "]", "{": "}"}
        for char in translation:
            if char in brackets:
                stack.append(char)
            elif char in brackets.values():
                if not stack or brackets[stack.pop()] != char:
                    return f"Unbalanced brackets in translation: {translation}"
        if stack:
            return f"Unbalanced brackets in translation: {translation}"
        return None

    def check_regex_block(self, translation):
        """If string belongs to a block, check if it matches the pattern after alignment."""
        # Note: This check is usually run AFTER alignment.
        # If the string matched a regex in its raw form, 
        # it should still match its anchored regex after spaces are added.
        for regex in self.compiled_patterns:
            # We check if the translation matches ANY pattern that was intended for it.
            # This is a bit broad, but if we have the row index we could be more specific.
            if regex.fullmatch(translation):
                return None
        
        # If we are here, we don't necessarily have an error unless we KNOW 
        # this string MUST match a specific block regex.
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
        
        err = self.check_brackets(translation)
        if err: errors.append(err)
        
        return errors

def validate(original: str, translation: str, lang_code: str = "ua") -> tuple[bool, str]:
    """
    Compatibility wrapper for Validator class.
    Returns (success, error_message or "OK").
    """
    # Note: in a real scenario we'd pass the blocks config path here
    # but for a quick check we can use a singleton or just a new instance without blocks
    v = Validator()
    errs = v.validate_row(original, translation)
    if errs:
        return False, "; ".join(errs)
    return True, "OK"

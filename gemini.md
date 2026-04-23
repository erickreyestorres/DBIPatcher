# Gemini Activity Log

## [2026-04-23] Початкова ініціалізація

### Виконані дії:
1.  **Реорганізація**: Створено папки `scripts/` та `src/core/`.
2.  **Переміщення**: `build_translation_bin.py` перенесено в `scripts/`.
3.  **Документація**: Створено `implementation_plan.md` та `task.md`.
4.  **Інсталяція**: Встановлено `openpyxl` та `requests`.

### Комміти:
- `Init`: Початковий стан проекту та нова структура.

## [2026-04-23] Розробка Core та Pipeline

### Виконані дії:
1.  **text_utils.py**: Реалізовано токенізацію (`tokenize`, `detokenize`, `normalize_tokens_out`).
2.  **validator.py**: Реалізовано 4 правила валідації (NotEmpty, RuPlaceholder, EnglishPreservation, TokenPreservation).
3.  **ai_client.py**: Клієнт для Gemini Proxy (`http://127.0.0.1:2048`), модель `gemini-flash-lite-latest`. Системний промпт з усіма правилами перекладу.
4.  **main.py**: Головний пайплайн з командами `sync`, `translate`, `validate`, `export`, `build`, `all`.
5.  **.gitignore**: Оновлено — виключено `*.nro`, `*.bin`, `output/`, `dictionary.xlsx`.
6.  **Видалено** `DBI.892.ru_patched.nro` з репозиторію (бінарник, не повинен бути в Git).

### Комміти:
- `feat: core modules (text_utils, validator, ai_client) and main pipeline`

### Фінальна структура:
```
dbi_patcher/
├── .gitignore
├── gemini.md
├── implementation_plan.md
├── task.md
├── languages.json
├── ua.csv
├── scripts/
│   └── build_translation_bin.py
└── src/
    ├── __init__.py
    ├── main.py
    └── core/
        ├── __init__.py
        ├── text_utils.py
        ├── validator.py
        └── ai_client.py
```

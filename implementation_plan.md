# DBI Patcher Pipeline Implementation Plan

Це комплексний план створення системи автоматизованого перекладу для DBI. Основна мета — забезпечити цілісність даних через токенізацію та валідацію за допомогою AI (Gemini Proxy).

## User Review Required

> [!IMPORTANT]
> **Gemini Proxy URL**: Використовується `http://127.0.0.1:2048/v1/chat/completions`.
> **Модель**: `gemini-flash-lite-latest`.
> **Бібліотека openpyxl**: Необхідна для роботи з Excel.

## Стовпці словника (dictionary.xlsx)
1. **Original** (Ключ - RU токенізований)
2. **ua** (Українська)
3. **en** (English)
... інші мови з languages.json.

## Токенізація (text_utils.py)
Заміна наступних послідовностей на токени:
- `\n` та `\\n` ➡️ `[[LF]]`
- `\r` та `\\r` ➡️ `[[CR]]`
- `\t` та `\\t` ➡️ `[[TAB]]`
- `\x1b` (HEX) ➡️ `[[ESC]]`

## Правила валідації (validator.py)
1. **NotEmpty**: Переклад не може бути порожнім.
2. **RuPlaceholder**: Якщо RU == `ru`, переклад == код мови (`en`, `ua` і т.д.).
3. **EnglishPreservation**: Якщо оригінал — ASCII, переклад ідентичний.
4. **TokenPreservation**: Кількість та порядок токенів `[[...]]` мають збігатися.

## Вихідні дані
- Для кожної мови генерується CSV.
- Викликається `python scripts/build_translation_bin.py <lang>.csv -o translation_<lang>.bin`.

## Версійність
- Версія зберігається в `dictionary.xlsx` (MetaData).
- Кожен комміт/великий крок ітерує версію.
